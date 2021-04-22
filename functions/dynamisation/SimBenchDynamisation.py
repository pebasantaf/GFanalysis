import functions.dynamisation.CreatePowerfactoryObjects as cpo
import functions.dynamisation.CreatePowerFactoryObjects_v2 as cpo2
import math
import datetime

# Adds a dynamic model to a ElmGenstat object
def AddConverterModell(app,prj, ElmGenstat, av_mode, cosn, ModelDict, DERModel_params, Dict_IntScenario_ElmGenstat, dynamisation=False, PCR=False, qv_ref=1, PQLimit=None, IntFolder_PQLimitsLF = None, dispatchcosn=0.95):

    # Dictionary for values to be changed in IntScenario objects

    Dict_IntScenario_ElmGenstat[ElmGenstat] = {}

    Dict_IntScenario_ElmGenstat[ElmGenstat]['mode_inp'] = ElmGenstat.GetAttribute('mode_inp')
    Dict_IntScenario_ElmGenstat[ElmGenstat]['cosgini'] = ElmGenstat.GetAttribute('cosgini')
    Dict_IntScenario_ElmGenstat[ElmGenstat]['pf_recap'] = ElmGenstat.GetAttribute('pf_recap')

    # operation scenarios

    operation_scenarios = ['lPV.IntScenario', 'hL.IntScenario', 'hPV.IntScenario', 'hW.IntScenario', 'lPV.IntScenario', 'lW.IntScenario']

    # Apply operational data to all relevant scenarios
    for id_scen in operation_scenarios:
        if not isinstance(id_scen, datetime.datetime):
            app.GetProjectFolder('scen').GetContents(id_scen)[0].Activate()

            for ElmGenstat, params in Dict_IntScenario_ElmGenstat.items():
                for param, val in params.items():
                    ElmGenstat.SetAttribute(param, val)

            app.GetProjectFolder('scen').GetContents(id_scen)[0].Save()
            app.WriteChangesToDb()

    # Get Point of Coupling (ElmTerm)
    ElmTerm = ElmGenstat.GetAttribute('bus1').GetAttribute('cterm')

    # Get Voltage level
    uknom = ElmTerm.GetAttribute('uknom')

    if dynamisation:

        # Set RMS settings
        ElmGenstat.SetAttribute('umin', 0)
        ElmGenstat.SetAttribute('uonthr',10)
        ElmGenstat.SetAttribute('iAstabint', 1)

        # ToDo: Add Switches?



        # Change active power limits
        ElmGenstat.SetAttribute('cosn', cosn)
        ElmGenstat.SetAttribute('Pmax_uc', ElmGenstat.GetAttribute('sgn'))
        ElmGenstat.SetAttribute('sgn', ElmGenstat.GetAttribute('sgn')/cosn)

    # Set stationary reactive power provision
    ElmGenstat.SetAttribute('av_mode', av_mode)

    # Create Controller
    ElmComp_DER = cpo.create_ElmComp(ElmGenstat.GetAttribute('loc_name') + '_Controller', prj.SearchObject('\\'.join(ElmGenstat.GetFullName().split('\\')[:-1])),
                                      ModelDict['DER_Frame'])
    ElmComp_DER.SetAttribute('Converter', ElmGenstat)

    # Create active power control
    if PCR == True:
        ElmDsl_PCR = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Control_PCR', ElmComp_DER,
                                        ModelDict[DERModel_params['PCR'][0]],
                                        params_dict=DERModel_params['PCR'][1])
        ElmComp_DER.SetAttribute('PCR', ElmDsl_PCR)

    # Create limited frequency sensitivy object
    ElmDsl_LFSM = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Control_LFSM', ElmComp_DER,
                                     ModelDict[DERModel_params['LFSM'][0]],
                                     params_dict=DERModel_params['LFSM'][1])
    ElmComp_DER.SetAttribute('LFSM', ElmDsl_LFSM)

    # Create reactive power control
    if av_mode == 'qvchar':
        ElmDsl_Q = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Control_Q', ElmComp_DER,
                                      ModelDict[DERModel_params['Q_Control'][0]],
                                      params_dict=DERModel_params['Q_Control'][1])
        ElmDsl_Q.Vref = qv_ref
        ElmComp_DER.SetAttribute('Reactive Power Control', ElmDsl_Q)

        # Set load flow behavior
        ElmGenstat.SetAttribute('ddroop', ElmDsl_Q.GetAttribute('ddroop'))
        ElmGenstat.SetAttribute('udeadblow', ElmDsl_Q.GetAttribute('Vref') + ElmDsl_Q.GetAttribute('Vdead'))
        ElmGenstat.SetAttribute('udeadbup', ElmDsl_Q.GetAttribute('Vref') - ElmDsl_Q.GetAttribute('Vdead'))
    elif av_mode == 'constc':
        ElmGenstat.SetAttribute('mode_inp', 'PC')
        ElmGenstat.SetAttribute('cosgini', dispatchcosn)
        ElmGenstat.SetAttribute('pf_recap', 1)


    # Set values for default controller
    ElmDsl_Layer = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Input_Layer', ElmComp_DER,
                                      ModelDict[DERModel_params['Input_Layer'][0]],
                                      params_dict=DERModel_params['Input_Layer'][1])
    ElmComp_DER.SetAttribute('Input', ElmDsl_Layer)

    # Create current injection control
    ElmDsl_Dyn_I = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Control_Dyn_I', ElmComp_DER,
                                      ModelDict[DERModel_params['Iq_Dyn'][0]],
                                      params_dict=DERModel_params['Iq_Dyn'][1])
    ElmComp_DER.SetAttribute('Dyn. Current Injection', ElmDsl_Dyn_I)

    # Create current limiter control
    ElmDsl_CurCont = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_I_Control', ElmComp_DER,
                                       ModelDict[DERModel_params['Current_Controller'][0]],
                                       params_dict=DERModel_params['Current_Controller'][1])
    ElmComp_DER.SetAttribute('Current Controller', ElmDsl_CurCont)

    # Create frequency protection
    ElmDsl_FreqProt = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Freq_Prot', ElmComp_DER,
                                         ModelDict[DERModel_params['fprot'][0]],
                                         params_dict=DERModel_params['fprot'][1])
    ElmComp_DER.SetAttribute('Frequency-Protection', ElmDsl_FreqProt)

    # Create voltage protection
    # ElmDsl_VoltProt = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_Volt_Prot', ElmComp_DER,
    #                                      ModelDict[DERModel_params['vprot'][0]],
    #                                      params_dict=DERModel_params['vprot'][1])
    # ElmComp_DER.SetAttribute('Voltage-Protection', ElmDsl_VoltProt)

    # Create PQ-Limits
    if PQLimit != None:


        if 'VDE_AR_N_4105' in PQLimit:
            ElmDsl_PQLimits = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_PQ_Limits', ElmComp_DER,
                                                ModelDict['PQ_Limits_' + PQLimit], {'cosn': cosn})
        else:
            ElmDsl_PQLimits = ElmComp_DER.PasteCopy(ModelDict['PQ_Limits_' + PQLimit].GetContents('PQ_Limits_' + PQLimit + '.ElmDsl')[0])[1]
        ElmComp_DER.SetAttribute('PQ-Limits', ElmDsl_PQLimits)

        # Set load flow limits
        if 'VDE_AR_N_4105' in PQLimit:
            ElmGenstat.SetAttribute('pQlimType', IntFolder_PQLimitsLF.GetContents(PQLimit + '_PF_' + str(cosn))[0])
        else:
            ElmGenstat.SetAttribute('pQlimType', IntFolder_PQLimitsLF.GetContents(PQLimit)[0])

    # Create FRT protection
    if uknom < 1:
        FRT_setting = 'FRT_LV'
    elif uknom >= 1 and uknom < 50:
        FRT_setting = 'FRT_MV'
    elif uknom >= 50 and uknom < 100:
        FRT_setting = 'FRT_HV'
    else:
        FRT_setting = 'FRT_UHV'

    ElmDsl_FRT_Prot = cpo.create_ElmDsl(ElmGenstat.GetAttribute('loc_name') + '_FRT_Prot', ElmComp_DER,
                                         ModelDict[DERModel_params[FRT_setting][0]],
                                         params_dict=DERModel_params[FRT_setting][1])
    ElmComp_DER.SetAttribute('FRT-Protection', ElmDsl_FRT_Prot)

    # Create measurement objects
    ElmPLL = cpo.create_ElmPLL(ElmGenstat.GetAttribute('loc_name') + '_PLL', ElmComp_DER,
                                {'pbusbar': ElmTerm, 'Kp': 50, 'Ki': 10})
    ElmComp_DER.SetAttribute('PLL', ElmPLL)

    ElmVmeas = cpo.create_VMeas(ElmGenstat.GetAttribute('loc_name') + '_Vmeas', ElmComp_DER,
                                 {'pbusbar': ElmTerm, 'i_mode': 1, 'nphase': 3})
    ElmComp_DER.SetAttribute('Vmeas', ElmVmeas)

    ElmImeas = cpo.create_IMeas(ElmGenstat.GetAttribute('loc_name') + '_Imeas', ElmComp_DER,
                                 {'pcubic': ElmGenstat.GetCubicle(0), 'i_mode': 1, 'nphase': 3})
    ElmComp_DER.SetAttribute('CurMeas', ElmImeas)

    ElmPQmeas = cpo.create_PQMeas(ElmGenstat.GetAttribute('loc_name') + '_PQmeas', ElmComp_DER,
                                   {'pcubic': ElmGenstat.GetCubicle(0), 'i_mode': 1, 'nphase': 3, 'i_orient': 0})
    ElmComp_DER.SetAttribute('PQmeas', ElmPQmeas)

    # Activate astabile integration option
    for ElmDsl in ElmComp_DER.GetContents('*.ElmDsl'):
        ElmDsl.SetAttribute('iAstabint', 1)


def AddVoltageSource(ElmNet, ModelDict, ElmTerm, Unom):
    ElmVac = cpo.create_ElmVac('ExtGrid', ElmNet, ElmTerm)
    ElmVac.SetAttribute('Unom', Unom)
    ElmComp_VS = ElmNet.CreateObject('ElmComp', 'VS_model')
    ElmComp_VS.SetAttribute('typ_id', ModelDict['VoltageSource'])
    ElmDsl_VS = ElmComp_VS.CreateObject('ElmDsl', 'VS_signal')
    ElmDsl_VS.SetAttribute('typ_id', ModelDict['VoltageSource'].GetContents('models.IntFolder')[0].GetContents(
        'sig_conditioning.BlkDef')[0])
    ElmDsl_VS.SetAttribute('f', 1)
    ElmDsl_VS.SetAttribute('tdip', 0.2)
    ElmDsl_VS.SetAttribute('tinic', 0)
    ElmDsl_VS.SetAttribute('Vldip', -0.5)
    ElmDsl_VS.SetAttribute('tramp', 0)
    StaVmea_VS = ElmComp_VS.CreateObject('StaVmea', 'VS_mea')
    StaVmea_VS.SetAttribute('pbusbar', ElmTerm)
    ElmComp_VS.SetAttribute('VSource', ElmVac)
    ElmComp_VS.SetAttribute('sig_conditioning', ElmDsl_VS)
    ElmComp_VS.SetAttribute('Vmeas', StaVmea_VS)

    return ElmVac

# Add grid forming converter modell from PowerFactory library
# Possible Models: Droop Controlled Converter, Synchronverter, Virtual Synchronous Machine
def AddGridformingConverter(app,target, ElmGenstat, IntLibrary, GF_modell, av_mode, cosn,Dict_IntScenario_ElmGenstat, dynamisation=False, PQLimit=None,
                            IntFolder_PQLimitsLF=None, dispatchcosn = 0.95,
                            **kwargs):
    # Dictionary for values to be changed in IntScenario objects

    Dict_IntScenario_ElmGenstat[ElmGenstat] = {}

    Dict_IntScenario_ElmGenstat[ElmGenstat]['mode_inp'] = ElmGenstat.GetAttribute('mode_inp')
    Dict_IntScenario_ElmGenstat[ElmGenstat]['cosgini'] = ElmGenstat.GetAttribute('cosgini')
    Dict_IntScenario_ElmGenstat[ElmGenstat]['pf_recap'] = ElmGenstat.GetAttribute('pf_recap')

    # operation scenarios

    operation_scenarios = ['lPV.IntScenario', 'hL.IntScenario', 'hPV.IntScenario', 'hW.IntScenario', 'lPV.IntScenario',
                           'lW.IntScenario']

    # Apply operational data to all relevant scenarios
    for id_scen in operation_scenarios:
        if not isinstance(id_scen, datetime.datetime):
            app.GetProjectFolder('scen').GetContents(id_scen)[0].Activate()

            for ElmGenstat, params in Dict_IntScenario_ElmGenstat.items():
                for param, val in params.items():
                    ElmGenstat.SetAttribute(param, val)

            app.GetProjectFolder('scen').GetContents(id_scen)[0].Save()
            app.WriteChangesToDb()

    if dynamisation:
        # Get Point of Cupling (ElmTerm)
        ElmTerm = ElmGenstat.GetAttribute('bus1').GetAttribute('cterm')

        # Get Voltage level
        uknom = ElmTerm.GetAttribute('uknom')

        # Set RMS settings
        ElmGenstat.SetAttribute('umin', 0)
        ElmGenstat.SetAttribute('uonthr',10)
        ElmGenstat.SetAttribute('iAstabint', 1)

        # ToDo: Add Switches?



        # Change active power limits
        ElmGenstat.SetAttribute('cosn', cosn)
        ElmGenstat.SetAttribute('Pmax_uc', ElmGenstat.GetAttribute('sgn'))
        ElmGenstat.SetAttribute('sgn', ElmGenstat.GetAttribute('sgn')/cosn)

    # Set stationary reactive power provision
    ElmGenstat.SetAttribute('av_mode', av_mode)

    #dispatch power factor

    ElmGenstat.SetAttribut('cosgini', dispatchcosn)

    # Create PQ-Limits
    if PQLimit != None:

        # Set load flow limits
        if 'VDE_AR_N_4105' in PQLimit:
            ElmGenstat.SetAttribute('pQlimType', IntFolder_PQLimitsLF.GetContents(PQLimit + '_PF_' + str(cosn))[0])
        else:
            ElmGenstat.SetAttribute('pQlimType', IntFolder_PQLimitsLF.GetContents(PQLimit)[0])

    # Get template
    IntTemplate = IntLibrary.GetContents('Templ')[0].GetContents('TemplGfc')[0].GetContents(GF_modell)[0]

    # Get virtual impedance

    vim = IntTemplate.GetContents('Library')[0].GetContents('Dynamic Models')[0].GetContents(
        'vim_constant virtual impedance')[0]

    # Copy composite model
    ElmComp = target.PasteCopy(IntTemplate.GetContents('*.ElmComp')[0])[1]

    #set virtual impedance
    ElmComp.GetAttribute('Virtual impedance').SetAttribute('typ_id', vim)

    # Rename model
    ElmComp.SetAttribute('loc_name', f'{ElmGenstat.GetAttribute("loc_name")}_GF_Controller')

    # Reset converter block
    ElmComp.SetAttribute('Converter', ElmGenstat)

    # Reset measurement objects
    for elm in ElmComp.GetContents():
        if elm.GetClassName() == 'StaImea':
            elm.SetAttribute('pcubic', ElmGenstat.GetAttribute('bus1'))
        elif elm.GetClassName() == 'StaVmea':
            elm.SetAttribute('pbusbar', ElmGenstat.GetAttribute('bus1').GetAttribute('cterm'))

    # Set PQ-Limits


    # Optional: Change settings of ElmDsl objects
    for ElmDsl_blockname, attrdict in kwargs.items():
        ElmDsl = ElmComp.GetAttribute(ElmDsl_blockname)

        for attr, val in attrdict.items():
            ElmDsl.SetAttribute(attr, val)


def GetIntScenarioData(ElmNet, IntScenario):

    IntScenarioData_dict = {}

    # Activate operation scenario
    IntScenario.Activate()

    # Get relevant base data for generators
    for ElmGenstat in ElmNet.GetContents('*.ElmGenstat', 1):
        IntScenarioData_dict[f'{ElmGenstat.GetAttribute("loc_name")}.ElmGenstat'] = {'pgini': ElmGenstat.GetAttribute('pgini')}

    # Get relevant base data for loads
    for ElmLod in ElmNet.GetContents('*.ElmLod', 1):
        IntScenarioData_dict[f'{ElmLod.GetAttribute("loc_name")}.ElmLod'] = {'plini': ElmLod.GetAttribute('plini'),
                                                                             'qlini': ElmLod.GetAttribute('qlini')}

    # Deactivate operation scenario
    IntScenario.Deactivate()

    return IntScenarioData_dict