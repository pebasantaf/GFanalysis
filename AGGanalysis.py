import sys
import os
import time
import statistics
import matplotlib.pyplot as plt

import functions.aggregationstudy.aggfunctions as AGF
import functions.convcompar.datamanager as DTM
import numpy as np
import functions.convcompar.PFmanager as PFM
import datetime
import pandas as pd

if os.getlogin() != 'Pedro':
    sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8')

    import powerfactory as pf


    start = time.asctime()
    print(start)

    #Initiate powerfactory. In case there is an error, will give the error code number
    try:
        app = pf.GetApplicationExt(None, None, r"/ini C:\Users\ge25bod\config.ini")
        # ... some      calculations ...
    except pf.ExitError as error:
        print(error)
        print('error.code = %d' % error.code)
        print('error.code = %d' % error.code)

# mode selection

mode = 'multievent' # RXanaylsis, RMSE, multievent, enveloping
escens = ['hPV', 'hW', 'lPV', 'lW']

#available analysis to be made

if mode == 'RXanalysis':

    submode = 'boxplot'
    setup = 'GFMdepth'

    if submode == 'barplot':

        vlevel = 'MV'
        controller = 'constphi'

        Rdf, Xdf, Cdf = AGF.RXanalysis(app, vlevel, controller, escens)


        DTM.DFplot([Rdf,Xdf, Cdf], 'aggRX',
                   savefigures=True,
                   figurefolder='grid aggregation\line_parameter/' + vlevel + controller)

    elif submode == 'boxplot':

        vlevel = 'MV'
        controllers = ['constphi', 'constv']

        R = [None]*2
        X = [None]*2
        C = [None]*2

        for controller in controllers:

            i = controllers.index(controller)

            Rdf, Xdf, Cdf = AGF.RXanalysis(app, vlevel, controller, escens)

            R[i] = Rdf
            X[i] = Xdf
            C[i] = Cdf

        parameters = {'R': R, 'X': X, 'C': C}

        if setup ==  'controller':

            p = 'C'

            Rconcv = list(parameters[p][0].loc[:, 'hPV']) + list(parameters[p][0].loc[:, 'hW']) + list(
                parameters[p][0].loc[:, 'lPV']) + list(parameters[p][0].loc[:, 'lW'])
            Rconcphi = list(parameters[p][1].loc[:, 'hPV']) + list(parameters[p][1].loc[:, 'hW']) + list(
                parameters[p][1].loc[:, 'lPV']) + list(parameters[p][1].loc[:, 'lW'])

            plt.figure(0)

            plt.boxplot([Rconcv,Rconcphi])

            plt.show()


        #elif setup == 'GFMdepth':






elif mode == 'RMSE':

    submode = 'boxplot'

    print('Obtaining all csv for RMSE calculation')

    # selecte noramlization mode and var to plot
    var2plot = 'I1'
    normmode = 'commonscenario' # 'commonscenario' or 'individual'

    #select controller folder
    basepath = r"C:\Users\Usuario\Documents\Universidad\TUM\Subjects\5th semester\Masterarbeit\Code_and_data\aggresults/"
    controller = ['constv', 'constphi']
    cont = controller[0]

    if submode == 'barplot':

        normRMSE = AGF.RMSEanalysis(var2plot, normmode, basepath, cont, escens)
        DTM.DFplot([normRMSE], 'aggeval',
                    savefigures=True,
                    figurefolder= 'grid aggregation\Vdip_0,5_RMSE/' + mode + '_' + var2plot + '_' + cont + '.pdf')

    elif submode == 'boxplot':

        normRMSEv = AGF.RMSEanalysis(var2plot, normmode, basepath, controller[0], escens)
        normRMSEphi = AGF.RMSEanalysis(var2plot, normmode, basepath, controller[1], escens)

        setup = 'converter' #constvconstphi, GFMdepth, scenario

        plt.figure(0)

        fontsize = 9
        params = {'legend.fontsize': fontsize,
                  'figure.titlesize': fontsize,
                  'axes.labelsize': fontsize,
                  'axes.titlesize': fontsize,
                  'xtick.labelsize': fontsize,
                  'ytick.labelsize': fontsize,
                  'font.family': 'Times New Roman',
                  'legend.frameon': '0',
                  'legend.columnspacing': '1'}
        for attr, val in params.items():
            plt.rcParams[attr] = val

        if setup == 'constvconstphi':

            concv = list(normRMSEv.loc[:, 'hPV']) + list(normRMSEv.loc[:, 'hW']) + list(
                normRMSEv.loc[:, 'lPV']) + list(normRMSEv.loc[:, 'lW'])
            concphi = list(normRMSEphi.loc[:, 'hPV']) + list(normRMSEphi.loc[:, 'hW']) + list(
                normRMSEphi.loc[:, 'lPV']) + list(normRMSEphi.loc[:, 'lW'])

            box = plt.boxplot([concv, concphi], patch_artist=True)
            plt.violinplot([concv, concphi], )

            colors = ['c', 'c']
            for patch, color in zip(box['boxes'], colors):
                patch.set_facecolor(color)

            plt.grid(axis='y')
            plt.xticks([1, 2], controller)
            plt.xlabel('Local controller')
            plt.ylabel('RMSE/%')
            plt.show()

        elif setup == 'GFMdepth':

            alllist = []
            onelist = []
            mvlist = []
            lvlist = []

            dtfs = [normRMSEv, normRMSEphi]

            for dtf in dtfs:

                for index,row in dtf.iterrows():

                    if index == 'GFLAll':

                        continue

                    elif ('All' in index) and (index != 'GFLAll'):

                        alllist = alllist + row.to_list()

                    elif 'MV' in index:

                        mvlist = mvlist + row.to_list()

                    elif 'LV' in index:

                        lvlist = lvlist + row.to_list()

                    else:

                        onelist = onelist + row.to_list()

            plt.boxplot([alllist,  lvlist,mvlist, onelist])
            plt.violinplot([alllist,  lvlist,mvlist, onelist])
            plt.xticks([1, 2, 3, 4], ['GFMAll', 'GFMLV', 'GMFMV', 'GFM1'])
            plt.xlabel('Depth of GFM integration')
            plt.ylabel('RMSE/%')
            plt.grid(axis='y')
            plt.show()

        elif setup == 'scenario':

            dtfs = [normRMSEv, normRMSEphi]
            hPV = []
            hW = []
            lPV = []
            lW = []

            for dtf in dtfs:

                hPV = hPV + list(dtf.loc[:,'hPV'])
                hW = hW + list(dtf.loc[:, 'hW'])
                lPV = lPV + list(dtf.loc[:,'lPV'])
                lW = lW + list(dtf.loc[:, 'lW'])

            plt.boxplot([hPV,hW,lPV,lW])
            violing = plt.violinplot([hPV, hW, lPV, lW])

            for pc in violing['bodies']:
                pc.set_facecolor('cyan')
                pc.set_edgecolor('red')

            plt.xticks([1,2,3,4], escens)
            plt.xlabel('Scenario')
            plt.ylabel('RMSE/%')
            plt.grid(axis='y')
            plt.show()

        elif setup == 'converter':

            dtfs = [normRMSEv, normRMSEphi]
            dro = []
            syn = []
            vsm = []

            for dtf in dtfs:

                for index,row in dtf.iterrows():

                    if index == 'GFLAll':

                        continue

                    elif 'Dro' in index:

                        dro = dro + row.to_list()

                    elif 'Syn' in index:

                        syn = syn + row.to_list()

                    elif 'VSM' in index:

                        vsm = vsm + row.to_list()

            plt.boxplot([dro,syn,vsm])
            plt.xticks([1,2,3], ['Dro', 'Syn', 'VSM'])
            plt.show()


elif mode == 'multievent':

    eventype = 'vdip' # 'vdip' or 'freqramp'
    controller = 'constphi'
    pathmode = 'Read'

    print('Creating or reading results folder')

    if pathmode == 'Create':

        basepath = DTM.ReadorCreatePath('Create', folder=datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_' + eventype + '_' + controller + '_AG\\')

    elif pathmode == 'Read':

        basepath = DTM.ReadorCreatePath('Read', readmode='userfile', folder='/14.06.2021_07-38-00_vdip_constphi_AG/', filename='')

    else:

        sys.exit('Wrong path mode')

    if eventype == 'vdip':

        values = list(np.arange(-0.9, -0.05, 0.1))
        values.extend((-0.05, 0.05, 0.1))
        values = [round(val, 2) for val in values]

    else:

        sys.exit('Wrong event type name')


    print('Running ' + str(len(values)) + ' events for ' + eventype)

    #  get all projects
    allprj = app.GetCurrentUser().GetContents('*.IntPrj')

    # get relevant projects
    prjlist = [s for s in allprj if '_eq_' + controller in s.GetFullName()]
    prjlist = prjlist
    escens = escens
    # for every project of the specified controller
    for prj in prjlist:

        print('PROJECT: Activating ' + prj.GetAttribute('loc_name') + ' project')
        prj.Activate() # activate

        # for each of the for scenarios
        for scen in escens:

            # get active study case
            actcase = app.GetProjectFolder('study').GetContents('MV_'+scen+'.IntCase')[0]
            print('STUDY CASE: Activating MV_' + scen + ' study case')
            actcase.Activate()

            #get VS from detailed and equivalent grids for given scenario
            detVS = app.GetProjectFolder('netdat').GetContents('*rural*.ElmNet')[0].GetContents('*.ElmVac')[0]
            eqVS = app.GetProjectFolder('netdat').GetContents('MV_'+scen+'.ElmNet')[0].GetContents('*.ElmVac')[0]

            # for each of the values of the event selected
            for val in values:

                print('Applying values ' + str(val) + ' to VS signal')
                #if this event is a voltage dip
                if eventype == 'vdip':
                    detVS.GetAttribute('c_pmod').GetAttribute('sig_conditioning').SetAttribute('Vldip', val)
                    eqVS.GetAttribute('c_pmod').GetAttribute('sig_conditioning').SetAttribute('Vldip', val)

                # run simulation and save csv in Results folder

                print('Running simulation and saving results')
                PFM.RunNSave(app, True, tstop=1, path=basepath + '{project}_{scenario}_{eventype}_{value}.csv'.format(project=prj.GetAttribute('loc_name'), scenario=scen, eventype=eventype, value=val))

elif mode == 'enveloping':


    submode = None
    grid = '1-MVLV-rural-all-0-sw_SynAll_constv'
    var2plot = ['u1', 'I1P', 'I1Q']
    units = ['pu', 'kA', 'kA']
    var = var2plot[1]
    label = var + '/' + units[var2plot.index(var)]

    if submode == 'Createcsv':

        project = app.GetCurrentUser().GetContents(grid)[0]
        project.Activate()

        app.GetProjectFolder('scen').GetContents('hW')[0].Activate()

        basepath = DTM.ReadorCreatePath('Create', folder=datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_EV\\')

        PFM.RunNSave(app, True, tstop=1, path=basepath + 'enveloping_ubus.csv')

    else:

        df = DTM.importData(r"C:\Users\Usuario\Documents\Universidad\TUM\Subjects\5th semester\Masterarbeit\Code_and_data\aggresults\07.06.2021_11-19-02_EV\enveloping_ubus1-MVLV-rural-all-0-sw_VSMAll_constvhPV.csv")
        enveldf = pd.DataFrame(columns=['max', 'min'])

        col2plot = [s for s in list(df.columns) if var in s]
        allvaluesdf = df.loc[:, col2plot]
        allvaluesdf.insert(0, df.columns[0], df.loc[:, df.columns[0]])

        for i in range(df.shape[0]):

            enveldf.loc[i, 'max'] = max(df.loc[i, col2plot])
            enveldf.loc[i, 'min'] = min(df.loc[i, col2plot])

        dflist = [enveldf, allvaluesdf]

        DTM.DFplot(dflist,'enveloping', ylabel=label)