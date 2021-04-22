import sys
import functions.dynamisation.ImportDynamicModels as idm
import functions.dynamisation.SimBenchDynamisation as sbd

def execComRes(IntCase, ElmRes, f_name, **kwargs):

    try:
        ComRes = IntCase.GetContents('*.ComRes', 1)[0]
    except:
        ComRes = IntCase.CreateObject('ComRes', 'ASCII - Ergebnis Export')
    ComRes.SetAttribute('pResult', ElmRes)

    ComRes.SetAttribute('f_name', f_name)

    for key, var in kwargs.items():
        ComRes.SetAttribute(key, var)

    ComRes.Execute()


def ActivateVariation(SelVariation, app):

    Scheme = app.GetProjectFolder('scheme')
    Variations = Scheme.GetContents('*.IntScheme')

    ActVar = app.GetActiveNetworkVariations()

    # make sure there is not active variation

    if not ActVar:

        print("No variation activated. Proceeding to activate selected variation")

    else:

        ActVar[0].Deactivate()

    for varis in Variations:

        VarName = varis.GetFullName().split('\\')[-1][:-10]

        if VarName == SelVariation:

            varis.Activate()


        return varis

def CreateSimpleStabilityStudy(app, CreateScens):

    MyCases = app.GetProjectFolder('study')
    MyScenarios = app.GetProjectFolder('scen')
    AktCase = app.GetActiveStudyCase()
    CaseNames = ['Frequency Ramp', 'Voltage Step', 'Voltage Ramp']

    if len(MyCases.GetContents('*.IntCase')) == 1:

        if CaseNames[0] not in AktCase.GetFullName():
            AktCase.SetAttribute('loc_name', CaseNames[0])
            print('Changing name of active study case')
        print('Procceeding to create the rest of cases')

        for n in range(len(CaseNames)-1):
            MyCases.AddCopy(AktCase, CaseNames[n + 1])

        for case in MyCases.GetContents('*.IntCase'): #for every object intcase

            print('Setting events for ' + CaseNames[MyCases.GetContents('*.IntCase').index(case)])
            Events = case.GetContents('*.IntEvt')[0]

            if len(Events[0].GetContents()) != 0:  # Delete whatever existing event parameters. This could be improved to check if desired events exist

                for evt in Events[0].GetContents():
                    evt.Delete()

            if CaseNames[0] in case.GetFullName():


                Events[0].CreateObject('*.EvtParam', 'ftime')
                Events[0].GetContents('ftime')
                Events[0].CreateObject('*.EvtParam', 'fslope')

            elif CaseNames[1] in case.GetFullName():

                Events[0].CreateObject('*.EvtParam', 'vdip')
                Events[0].CreateObject('*.EvtParam', 'vtime')

            elif CaseNames[2] in case.GetFullName():

                Events[0].CreateObject('*.EvtParam', 'vdip')
                Events[0].CreateObject('*.EvtParam', 'tinic')
                Events[0].CreateObject('*.EvtParam', 'tdip')
                Events[0].CreateObject('*.EvtParam', 'tramp')












    if CreateScens == 1:

        if not MyScenarios.GetContents():

            print('Scenarios object empty. Creating SM and VS scenarios')
            MyScenarios.CreateObject('IntScenario', 'SM')
            MyScenarios.CreateObject('IntScenario', 'VS')

        else:

            s = 0
            t = 0
            for scen in MyScenarios.GetContents():

                if 'SM' in scen.GetFullName():

                    s += 1

                    if s == 0:
                        print("Creating SM scenario")
                        MyScenarios.CreateObject('IntScenario', 'SM')

                elif 'VS' in scen.GetFullName():

                    t += 1

                    if t == 0:

                        print("Creating VS scenario")
                        MyScenarios.CreateObject('IntScenario', 'VS')
    elif CreateScens == 0:

        print("No scenarios created")

    else:

        raise ValueError("Flag value not valid. Choose 1 or 0")


def SetAttributesforFaultEvent(app, faultname, **kwargs):

    faultcase = app.GetFromStudyCase('IntEvt')
    event = faultcase.GetContents(faultname)[0]

    for key, var in kwargs.items():
        event.SetAttribute(key, var)


def SetFaulEvent(app,event, tinit, tend, value):

    if event == 'Frequency Ramp':

        SetAttributesforFaultEvent(app,'ftime',value=str(tend), time=tinit)
        SetAttributesforFaultEvent(app, 'fslope', value=str(value), time=tinit)

    elif event == 'Voltage Step':

        SetAttributesforFaultEvent(app, 'vtime', value=str(tend), time=tinit)
        SetAttributesforFaultEvent(app, 'vdip', value=str(value), time=tinit)

    elif event == 'Voltage Ramp':

        SetAttributesforFaultEvent(app, 'tinic', value=str(tinit), time=tinit)
        SetAttributesforFaultEvent(app, 'Vldip', value=str(value), time=tinit)
        SetAttributesforFaultEvent(app, 'tramp', value=str(tend), time=tinit)
        SetAttributesforFaultEvent(app, 'tdip', value=str(tend-tinit), time=tinit)

    else:

        sys.exit('Wrong event name')


def RunNSave(app, saveResults, **kwargs):
    # EXECUTING OF SIMULATION

    ComInc = app.GetActiveStudyCase().GetContents('*.ComInc')[0]
    ComSim = app.GetActiveStudyCase().GetContents('*.ComSim')[0]


    # these functions optimize performnance
    app.WriteChangesToDb()
    app.SetWriteCacheEnabled(0)

    # change simulation time

    ComSim.SetAttribute(list(kwargs)[0], kwargs.get(list(kwargs)[0]))

    # excecute simulation

    ComInc.Execute() # execute initial conditions
    ComSim.Execute() # execute simulation


    # execute the object to export results

    ElmRes = app.GetFromStudyCase("ElmRes") #results object

    stucase = app.GetFromStudyCase('IntCase')  # study case

    # execute results and save as csv

    if saveResults:
        execComRes(stucase,ElmRes,kwargs.get('path'),
                       iopt_exp=6,
                       iopt_sep=0,
                       dec_Sep='.')
        app.SetWriteCacheEnabled(0)



def ApplyConverters(app, activeproject, GFol_model, gentype, convtype, **kwargs):

    ElmNet_base = app.GetProjectFolder('netdat').GetContents('*.ElmNet')[0]
    ModelDict = {}
    ModelDict, DERModel_params = idm.ImportDynamicsConverters(app, activeproject, GFol_model, app.GetProjectFolder('blk'),
                                                              ModelDict, copy_VS=True)


    for ElmGenstat in ElmNet_base.GetContents('*.ElmGenstat', 1): # for every grid element

        if ElmGenstat.GetAttribute('cCategory') in gentype:

            if ElmGenstat.GetAttribute('cCategory') == 'Photovoltaic':

                if kwargs.get('convname') in ElmGenstat.GetAttribute('loc_name'):

                    pass
                else:

                    continue

            pass
        else:

            continue

        # delete current converter model

        if ElmGenstat.GetAttribute('c_pmod') != None:

            ElmGenstat.GetAttribute('c_pmod').Delete()

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
            av_mode = 'qvchar'
            PQLimit = 'VDE_AR_N_4120_Var1'
            cosn = 0.9

        cosgini = kwargs.get('cosgini')
        # Add converter type according to category

        # initialize dictionary for storing scenario data
        Dict_IntScenario_ElmGenstat = {}

        if convtype == 'GridFollowing':

            sbd.AddConverterModell(app, activeproject, ElmGenstat, av_mode, cosn, ModelDict, DERModel_params,Dict_IntScenario_ElmGenstat, dynamisation=False, PCR=False, qv_ref=1,
                                   PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'))

        elif convtype == 'Synchronverter':

            sbd.AddGridformingConverter(app, ElmNet_base, ElmGenstat, app.GetGlobalLibrary(), 'Synchronverter', 'constc', cosn, Dict_IntScenario_ElmGenstat, dynamisation=False,
                                    PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'), dispatchcosn=cosgini,
                                    **{'Synchronverter control': {'Ta': 3}}, **{'Virtual impedance': {'r': 0}}, **{'Virtual impedance': {'x': 0.1}})

        elif convtype == 'Droop Controlled Converter':

            sbd.AddGridformingConverter(app, ElmNet_base, ElmGenstat, app.GetGlobalLibrary(), 'Droop Controlled Converter', 'constc', cosn, Dict_IntScenario_ElmGenstat, dynamisation=False,
                                        PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'), dispatchcosn=cosgini, **{'Virtual impedance': {'r': 0}}, **{'Virtual impedance': {'x': 0.1}})


        elif convtype == 'Virtual Synchronous Machine':

            sbd.AddGridformingConverter(app, ElmNet_base, ElmGenstat, app.GetGlobalLibrary(), 'Virtual Synchronous Machine', 'constc', cosn, Dict_IntScenario_ElmGenstat, dynamisation=False,
                                    PQLimit=PQLimit, IntFolder_PQLimitsLF=app.GetProjectFolder('mvar'), dispatchcosn=cosgini,
                                    **{'Grid-forming control': {'Ta': 3}}, **{'Virtual impedance': {'r': 0}}, **{'Virtual impedance': {'x': 0.1}})

        else:

            sys.exit('Invalid converter name: '+ convtype)

    app.WriteChangesToDb()