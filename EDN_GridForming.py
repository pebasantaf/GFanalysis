# Import PowerFactory module
import sys
sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8')
import powerfactory as pf

# Import modules
import functions.dynamisation.PowerfactoryFunctions as pff
import functions.dynamisation.ImportDynamicModels as idm
import functions.dynamisation.SimBenchDynamisation as sbd

# General modules
import os
from datetime import datetime

# Start PowerFactory connection
try:
    app = pf.GetApplication()
    app.SetWriteCacheEnabled(1)
except:
    print('PowerFactory connection failed')
# Set modes
GF_modell = 'Virtual Synchronous Machine' # Synchronveter, Droop Controlled Converter, Virtual Synchronous Machine

# set criteria to apply converters

criteria = 'fullname' # 'Vlevel': voltage level; 'Gentype': type of generation; 'fullname': Full name of the converter

# application list, that will change depending on the criteria
app_list = ['MV1.101 SGen 1'] #MV1.101 SGen 1

# Dynamic converter model project
path_model = r"I:\05_Basanta Franco\Masterarbeit_local\Models\DEAModel_20210209_PQ_Limits.pfd"

# SimBench PowerFactory Project
path_griddata = r"I:\05_Basanta Franco\Masterarbeit_local\01_MVLV-Netze_vollständig"
grid_dirlist = ['1-MVLV-rural-all-0-sw']

# Location for exports
path_export = os.path.join(r"I:\05_Basanta Franco\Masterarbeit_local\Results", f'{datetime.now().strftime("%Y%m%d-%H-%M-%S")}')
os.mkdir(path_export)

#available scenarios
operation_scenarios = ['hPV.IntScenario', 'hW.IntScenario', 'lPV.IntScenario',
                           'lW.IntScenario']

cosgini = 1

# Iterate over all pfd files
for grid_dir in grid_dirlist:

    # Import SimBench PowerFactory project
    print(f'{datetime.now().strftime("%H:%M:%S")}: Processing {grid_dir}')
    pfd_file = [f for f in os.listdir(os.path.join(path_griddata, grid_dir)) if '.pfd' in f][0]
    prj_sb = pff.ImportPfdFile(app, os.path.join(path_griddata, grid_dir, pfd_file))
    prj_sb.Activate()

    # Create subdirectory for project
    path_export_prj = os.path.join(path_export, pfd_file.replace('.pfd',''))
    os.mkdir(path_export_prj)

    # Save base grid and base study case to variables
    ElmNet_base = app.GetProjectFolder('netdat').GetContents('*.ElmNet')[0]
    IntCase_base = app.GetProjectFolder('study').GetContents('*.IntCase')[0]

    # Get task automation command
    ComTasks = pff.GetPowerFactoryObject('ComTasks', app.GetProjectFolder('study'), **{})

    # Create folder in library for equivalent line parameters
    IntFolder_eqLines = app.GetProjectFolder('equip').CreateObject('IntFolder', 'Equivalent Line Types')

    # Get dynamic converter model
    print(f'{datetime.now().strftime("%H:%M:%S")}: Starting grid dynamisation...')
    ModelDict = {}
    ModelDict, DERModel_params = idm.ImportDynamicsConverters(app, prj_sb, path_model, app.GetProjectFolder('blk'),
                                                              ModelDict, copy_VS=True)

    # Create voltage source
    ElmNet_base.GetContents('*.ElmXnet')[0].Delete()
    sbd.AddVoltageSource(ElmNet_base, ModelDict, app.GetCalcRelevantObjects('HV*.ElmTerm')[0], 110)

    # Adjust path of characteristics
    for ChaTime in app.GetProjectFolder('chars').GetContents('*.ChaTime'):
        ChaTime.SetAttribute('f_name', ChaTime.GetAttribute('f_name').replace(r'U:\SimBench_Dataset_final', path_griddata))

    # Activate characteristics
    for ChaRef in app.GetProjectFolder('netdat').GetContents('*.ChaRef', 1):
        ChaRef.SetAttribute('outserv', 1)

    # Add a dynamic model to all ElmGenstat Objects ToDo: differentiate between different types

    # initialize dictionary for storing scenario data
    Dict_IntScenario_ElmGenstat_GFor = {}
    Dict_IntScenario_ElmGenstat_GFol = {}

    allgens = ElmNet_base.GetContents('*.ElmGenstat', 1)

    exceptions = [s for s in ElmNet_base.GetContents('*.ElmGenstat', 1) if
                 'LV2' in s.GetFullName() and 'SGen 4' in s.GetFullName()]

    for ElmGenstat in ElmNet_base.GetContents('*.ElmGenstat', 1):

        convname = ElmGenstat.GetAttribute('loc_name')

        # Get Voltage level
        u_ElmGenstat = ElmGenstat.GetAttribute('bus1').GetAttribute('cterm').GetAttribute('uknom')

        # LV settings
        if u_ElmGenstat < 10:
            av_mode = 'constc'
            PQLimit = 'VDE_AR_N_4105'

            if ElmGenstat.GetAttribute('sgn') / 0.95 > 0.0046:
                cosn = 0.9
            else:
                cosn = 0.95

        # MV settings
        elif u_ElmGenstat < 110:
            av_mode = 'constc'
            PQLimit = 'VDE_AR_N_4110'
            cosn = 0.95

        # HV settings
        elif u_ElmGenstat >= 110:
            av_mode = 'constc'
            PQLimit = 'VDE_AR_N_4120_Var1'
            cosn = 0.9

        # selecte goal converter depending on criteria selected

        if criteria == 'Vlevel':

            goal = ElmGenstat.GetAttribute('loc_name')[0:2]

        elif criteria == 'Gentype':

            goal =ElmGenstat.GetAttribute('cCategory')

        elif criteria == 'fullname':

            goal = ElmGenstat.GetAttribute('loc_name')

        else:

            sys.exit('Wrong criteria')

        # Add converter type according to category

        exception = []

        if 'LV2' in convname and 'SGen 4' in convname:

            exception = [convname]

        if (goal in app_list) and (convname not in exception):

            sbd.AddGridformingConverter(ElmNet_base, ElmGenstat, app.GetGlobalLibrary(), GF_modell, 'constc', cosn, Dict_IntScenario_ElmGenstat_GFol, dynamisation=True,
                                        PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'), dispatchcosn=cosgini,
                                        **{'Virtual impedance': {'r': 0,
                                                                 'x': 0.1}}, #**{'Grid-forming control': {'Ta': 3}},
                                       )
        else:
            sbd.AddConverterModell(prj_sb, ElmGenstat, av_mode, cosn, ModelDict, DERModel_params, Dict_IntScenario_ElmGenstat_GFor, dynamisation=True, PCR=False, qv_ref=1,
                                   PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'), dispatchcosn=cosgini)
    app.WriteChangesToDb()


    # Apply operational data to all relevant scenarios
    for id_scen in operation_scenarios:

        if not isinstance(id_scen, datetime):

            app.GetProjectFolder('scen').GetContents(id_scen)[0].Activate()

            #remove characteristics of generators

            for ChaRef in app.GetProjectFolder('netdat').GetContents('*.ChaRef', 1):
                ChaRef.SetAttribute('outserv', 1)

            for ElmGenstat, params in Dict_IntScenario_ElmGenstat_GFol.items():

                for param, val in params.items():

                    ElmGenstat.SetAttribute(param, val)

            app.GetProjectFolder('scen').GetContents(id_scen)[0].Save()

            app.WriteChangesToDb()

            for ElmGenstat, params in Dict_IntScenario_ElmGenstat_GFor.items():

                for param, val in params.items():

                    ElmGenstat.SetAttribute(param, val)

            app.GetProjectFolder('scen').GetContents(id_scen)[0].Save()
            app.WriteChangesToDb()

app.WriteChangesToDb()
app.SetWriteCacheEnabled(0)