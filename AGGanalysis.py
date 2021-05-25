import sys
import time
import functions.aggregationstudy.aggfunctions as AGF
import functions.convcompar.datamanager as DTM
import numpy as np
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

mode = 'multievent' # RXanaylsis, RMSE, multievent
escens = ['hPV', 'hW', 'lPV', 'lW']

#available analysis to be made

if mode == 'RXanalysis':

    vlevel = 'MV'
    controller = 'constv'

    Rdf, Xdf, Cdf = AGF.RXanalysis(app, vlevel, controller, escens)


    DTM.DFplot([Rdf,Xdf, Cdf], 'aggRX',
               savefigures=True,
               figurefolder='grid aggregation\line_parameter/' + vlevel + controller)

elif mode == 'RMSE':

    print('Obtaining all csv for RMSE calculation')

    # selecte noramlization mode and var to plot
    var2plot = 'I1Q'
    normmode = 'commonscenario' # 'commonscenario' or 'individual'

    #select controller folder
    basepath = r"I:\05_Basanta Franco\Exchange\MasterThesis\Aggregation/"
    controller = 'constv'

    normRMSE = AGF.RMSEanalysis(var2plot,normmode,basepath,controller,escens)

    DTM.DFplot([normRMSE], 'aggeval',
                savefigures=True,
                figurefolder= 'grid aggregation\Vdip_0,5_RMSE/' + mode + '_' + var2plot + '_' + controller + '.png')


elif mode == 'multievent':

    eventype = 'vdip' # 'vdip' or 'freqramp'

    if eventype == 'vdip':

        values = list(np.arange(-0.1, -0.95, 0.05))
        values.extend((0.05, 0.1, 0.12))
        values = [round(val, 2) for val in values]

    controller = 'constphi'

    #  get all projects
    allprj = app.GetCurrentUser().GetContents('*.IntPrj')

    # get relevant projects
    prjlist = [s for s in allprj if '_eq_' + controller in s.GetFullName()]

    for prj in prjlist:

        prj.Activate()

        for scen in escens:

            # get active study case
            actcase = app.GetProjectFolder('study').GetContents('MV_'+scen+'.IntCase')[0]
            actcase.Activate()

            #get VS from detailed and equivalent grids for given scenario
            detVS = app.GetProjectFolder('netdat').GetContents('*rural*.ElmNet')[0].GetContents('*.ElmVac')
            eqVS = app.GetProjectFolder('netdat').GetContents('MV_'+scen+'.ElmNet')[0].GetContents('*.ElmVac')









