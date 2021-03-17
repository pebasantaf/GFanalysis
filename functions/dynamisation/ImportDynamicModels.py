import functions.dynamisation.PowerfactoryFunctions as pff

def ImportDynamicsConverters(app, prj, path_model, UserModFolder, ModelDict, copy_VS=False):

    # Import DER model
    prj_DER = pff.ImportPfdFile(app, path_model)

    # Get model from imported project
    prj_DER.Activate()
    UserModFolder_DER = app.GetProjectFolder('blk')
    DERModel = UserModFolder_DER.GetContents('Generic Model_VAR-Idq')[0]
    DERModel_name = DERModel.GetAttribute('loc_name')
    PQLimits_LF = app.GetProjectFolder('mvar').GetContents()

    # Get modell for voltage source if necessary
    if copy_VS == True:
        VS_modell = UserModFolder_DER.GetContents('Sources')[0]
        VS_modell_name = VS_modell.GetAttribute('loc_name')

    # Copy model to main project
    prj.Activate()
    UserModFolder.PasteCopy(DERModel)
    DERModel = UserModFolder.GetContents(DERModel_name)[0]
    if copy_VS == True:
        UserModFolder.PasteCopy(VS_modell)
        VS_modell = UserModFolder.GetContents(VS_modell_name)[0]

    # Copy load flow PQ-Limits
    app.GetProjectFolder('mvar').PasteCopy(PQLimits_LF)

    # Delete project with model
    prj_DER.Delete()

    # Components for DER model
    ModelDict.update({'DER_Frame': DERModel.GetContents('idq_Converter_Frame')[0]})

    ModelDict.update({'Input': DERModel.GetContents('Input', 1)[0]})

    ModelDict.update({'Current_Controller': DERModel.GetContents('Current_Controller', 1)[0]})

    ModelDict.update({'iq_Injection': DERModel.GetContents('iq_Injection', 1)[0]})

    ModelDict.update({'QV_Control': DERModel.GetContents('QV_Control', 1)[0]})

    ModelDict.update({'PCR': DERModel.GetContents('PCR', 1)[0]})

    ModelDict.update({'LFSM': DERModel.GetContents('LFSM', 1)[0]})

    ModelDict.update({'f-Prot_MinMax': DERModel.GetContents('f-Prot_MinMax', 1)[0]})

    ModelDict.update({'V-Prot_4Step': DERModel.GetContents('V-Prot_4Step', 1)[0]})

    ModelDict.update({'FRT_VDE_AR_N_4105': DERModel.GetContents('FRT_VDE_AR_N_4105', 1)[0]})
    ModelDict.update({'FRT_VDE_AR_N_4110': DERModel.GetContents('FRT_VDE_AR_N_4110', 1)[0]})
    ModelDict.update({'FRT_VDE_AR_N_4120': DERModel.GetContents('FRT_VDE_AR_N_4120', 1)[0]})

    ModelDict.update({'PQ_Limits_VDE_AR_N_4105': DERModel.GetContents('PQ_Limits_VDE_AR_N_4105', 1)[0],
                      'PQ_Limits_VDE_AR_N_4110': DERModel.GetContents('PQ_Limits_VDE_AR_N_4110', 1)[0],
                      'PQ_Limits_VDE_AR_N_4120_Var1': DERModel.GetContents('PQ_Limits_VDE_AR_N_4120_Var1', 1)[0],
                      'PQ_Limits_VDE_AR_N_4120_Var2': DERModel.GetContents('PQ_Limits_VDE_AR_N_4120_Var2', 1)[0],
                      'PQ_Limits_VDE_AR_N_4120_Var3': DERModel.GetContents('PQ_Limits_VDE_AR_N_4120_Var3', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_220kV_Var1': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_220kV_Var1', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_220kV_Var2': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_220kV_Var2', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_220kV_Var3': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_220kV_Var3', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_380kV_Var1': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_380kV_Var1', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_380kV_Var2': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_380kV_Var2', 1)[0],
                      'PQ_Limits_VDE_AR_N_4130_380kV_Var3': DERModel.GetContents('PQ_Limits_VDE_AR_N_4130_380kV_Var3', 1)[0]
                      })

    if copy_VS == True:
        ModelDict.update({'VoltageSource': VS_modell.GetContents('CtrlVSource_v2.BlkDef')[0]})

    # DER model parameters for each voltage level
    DERModel_params = {
        'Input_Layer': ('Input',
                              {'T_Vmeas': 0.01, 'T_Curmeas': 0.01, 'T_femeas': 0,
                               'T_qm': 0.01, 'T_pm': 0.01, 'u_tol': 0.1, 't_after': 0.2}),
        'Q_Control': ('QV_Control',
                      {'T_qstat_delay': 0.7, 'ddroop': 5, 'Vdead': 0, 'Vref': 1}),
        'Iq_Dyn': ('iq_Injection',
                   {'T_iq_inj': 0, 'K_iq_inj': 2, 'SDLWindV': 0}),
        'PCR': ('PCR',
                {'Tdelay_PCR': 2, 'droop_PCR': 5, 'fmin_PCR': 1.004, 'fmax_PCR': 0.996}),
        'LFSM': ('LFSM',
                 {'Tdelay_LFSM': 13, 'droop_LFSM': 5,
                  'df_LFSM_U': 0.004, 'df_LFSM_O': 0.004}),
        'Current_Controller': ('Current_Controller',
                            {'K_q': 0.5, 'T_q': 0.01, 'K_p': 0.5, 'T_p': 0.04,
                             'K_iq': 1, 'T_iq': 0.01, 'K_id': 1, 'T_id': 0.01}),
        'fprot': ('f-Prot_MinMax',
                  {'f_el_min': 0.95, 'f_el_max': 1.03, 't_sw_delay': 0, 't_uf_delay': 0.01, 't_of_delay': 0.01}),
        'vprot': ('V-Prot_4Step',
                  {'Umin_norm': 0.8, 't_UV_norm': 1.5, 'Umin_fast': 0.3, 't_UV_fast': 0.8, 'Umax_norm': 1.1
                      , 't_OV_norm': 600, 'Umax_fast': 1.25, 't_OV_fast': 0.1, 't_sw_delay': 0}),
        'FRT_LV': ('FRT_VDE_AR_N_4105',
                {'V_FRT_HV': 0.1, 'V_FRT_LV': 0.1, 't_delay': 0}),
        'FRT_MV': ('FRT_VDE_AR_N_4110',
                   {'V_FRT_HV': 0.1, 'V_FRT_LV': 0.1, 't_delay': 0}),
        'FRT_HV': ('FRT_VDE_AR_N_4120',
                   {'V_FRT_HV': 0.1, 'V_FRT_LV': 0.1, 't_delay': 0})
    }

    return ModelDict, DERModel_params