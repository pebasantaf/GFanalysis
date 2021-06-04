import sys
import os
import time
import functions.aggregationstudy.aggfunctions as AGF
import functions.convcompar.datamanager as DTM
import numpy as np
import functions.convcompar.PFmanager as PFM
import datetime

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

mode = 'RMSE' # RXanaylsis, RMSE, multievent
escens = ['hPV', 'hW', 'lPV', 'lW']

#available analysis to be made

if mode == 'RXanalysis':

    vlevel = 'LV4'
    controller = 'constphi'

    Rdf, Xdf, Cdf = AGF.RXanalysis(app, vlevel, controller, escens)


    DTM.DFplot([Rdf,Xdf, Cdf], 'aggRX',
               savefigures=True,
               figurefolder='grid aggregation\line_parameter/' + vlevel + controller)

elif mode == 'RMSE':

    print('Obtaining all csv for RMSE calculation')

    # selecte noramlization mode and var to plot
    var2plot = 'I1'
    normmode = 'commonscenario' # 'commonscenario' or 'individual'

    #select controller folder
    basepath = r"C:\Users\Usuario\Documents\Universidad\TUM\Subjects\5th semester\Masterarbeit\Code_and_data\aggresults/"
    controller = 'constv'

    normRMSE = AGF.RMSEanalysis(var2plot,normmode,basepath,controller,escens)

    DTM.DFplot([normRMSE], 'aggeval',
                savefigures=True,
                figurefolder= 'grid aggregation\Vdip_0,5_RMSE/' + mode + '_' + var2plot + '_' + controller + '.pdf')


elif mode == 'multievent':

    eventype = 'vdip' # 'vdip' or 'freqramp'
    controller = 'constv'
    pathmode = 'Read'

    print('Creating or reading results folder')

    if pathmode == 'Create':

        basepath = DTM.ReadorCreatePath('Create', folder=datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_' + eventype + '_' + controller + '_AG\\')

    elif pathmode == 'Read':

        basepath = DTM.ReadorCreatePath('Read', readmode='lastfile')

    else:

        sys.exit('Wrong path mode')

    if eventype == 'vdip':

        values = list(np.arange(-0.9, -0.05, 0.1))
        values.extend((-0.05, 0.05, 0.1, 0.15, 0.2))
        values = [round(val, 2) for val in values]

    else:

        sys.exit('Wrong event type name')


    print('Running ' + str(len(values)) + ' events for ' + eventype)

    #  get all projects
    allprj = app.GetCurrentUser().GetContents('*.IntPrj')

    # get relevant projects
    prjlist = [s for s in allprj if '_eq_' + controller in s.GetFullName()]
    prjlist = prjlist[2:]

    # for every project of the specified controller
    for prj in prjlist:

        print('Activating ' + prj.GetAttribute('loc_name') + ' project')
        prj.Activate() # activate

        # for each of the for scenarios
        for scen in escens:

            # get active study case
            actcase = app.GetProjectFolder('study').GetContents('MV_'+scen+'.IntCase')[0]
            print('Activating MV_' + scen + ' study case')
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
                    detVS.GetAttribute('c_pmod').GetAttribute('sig_conditioning').SetAttribute('Vldip', val)

                # run simulation and save csv in Results folder

                print('Running simulation and saving results')
                PFM.RunNSave(app, True, tstop=1, path=basepath + '{project}_{scenario}_{eventype}_{value}.csv'.format(project=prj.GetAttribute('loc_name'), scenario=scen, eventype=eventype, value=val))
