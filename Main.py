import sys
import os
import datetime
import functions.convcompar.PFmanager as PFM
import functions.convcompar.datamanager as DM
from conf.configs import confselect
import functions.dynamisation
#from conf.configs import confselect
from pprint import pprint


sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8')

import powerfactory as pf

#Initiate powerfactory. In case there is an error, will give the error code number
try:
    app = pf.GetApplicationExt(None, None, r"/ini C:\Users\ge25bod\config.ini")
    # ... some      calculations ...
except pf.ExitError as error:
    print(error)
    print('error.code = %d' % error.code)
    print('error.code = %d' % error.code)

# Activate the project
projname = str("DEAModel_gridforming_20210301")

app.ActivateProject(projname)

# Get project that is active as and application
project = app.GetActiveProject()


# Get Calculation object by getting active study case. This active case is related to a grid or part of a grid that is active.
#Then, getting the initial conditions and the simulation object to being executed at the end

#  function CreateSimpleStabilityStudy

PFM.CreateSimpleStabilityStudy(app, 0)

# manage variations

ResultsList = list()
faultname = 'allvoldip'
conf = confselect(faultname)

# Frequency Ramp, Voltage Step, Voltage Ramp
Modes = [1]  # Modes: 0-Run; 1-Plot

# activate study case

SelCase = app.GetProjectFolder('study').GetContents(conf.get('StudyCase'))[0]
SelCase.Activate()

NetData = app.GetProjectFolder('netdat')
Net = NetData.GetContents('110KV.ElmNet')

newfolder = datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_CC\\'
filenum = 1

for Mode in Modes:

    if Mode == 0:
        for varname in conf.get('Variation Name'):

            Variation = PFM.ActivateVariation(varname, app)

            Converter = Net[0].GetContents('*.ElmGenStat')[0]
            Frame = Converter.GetAttribute('c_pmod')

            for val in conf.get('faultvalues'):

                # apply desired fault values
                PFM.SetFaulEvent(app, conf.get('StudyCase'), conf.get('tinit'), conf.get('tend'), val)

                # apply desired inertia values
                for inval in conf.get('inertiavalues'):

                    if varname == 'Synchronverter':

                        print('Applying inertia value ' + str(inval))
                        Frame.GetAttribute('Synchronverter control').SetAttribute('Ta', inval)

                    elif varname == 'VSM':

                        print('Applying inertia value ' + str(inval))
                        Frame.GetAttribute('Grid-forming control').SetAttribute('Ta', inval)

                    else:

                        print('Converter has no inertia')

                    path = DM.ReadorCreatePath('Create', folder=newfolder, filename=str(filenum) + '-' +
                                                                                    "{controller}-{faultname}{faults}-inertia{inertia}.csv".format(controller=varname, faultname= faultname, faults=val,inertia=inval))

                    filenum += 1

                    # EXECUTING SIMULATION
                    print('Executing simulation')
                    PFM.RunNSave(app, True, tstop=2, path=path)

    elif Mode == 1:

        path = DM.ReadorCreatePath('Read', readmode='lastfile')
        # import results
        for direc in os.listdir(path):
            Results = DM.importData(path + direc).astype(float)

            ResultsList.append(Results)

        # print results

        DM.DFplot(ResultsList, [1, 1], 'convcompar',
                  xaxis=conf.get('xaxis'),
                  xlabel='Time (s)',
                  seriesnames=conf.get('seriesnames'),
                  savefigures=conf.get('savefigures'),
                  fixplot=conf.get('fixplot'),
                  figurefolder=conf.get('figurefolder'))

app.PostCommand("exit")
