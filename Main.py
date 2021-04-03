import sys
import os
import datetime
import functions.convcompar.PFmanager as PFM
import functions.convcompar.datamanager as DM
# from conf.configs import *
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


#get project folder

# manage variations

ResultsList = list()

#conf = confselect('freqramp')

StudyCase = 'Frequency Ramp' # Frequency Ramp, Voltage Step, Voltage Ramp
VariationName = ["Synchronverter"]

Eventname = 'fslope' # Frequency ramp:fslope; Voltage step:vdip; Voltage ramp:
faultvalues = ["0,001"]
inertiavalues = [3,5,7,10]
Modes = [0, 1, 2] # Modes: 0-Run; 1-Import; 2-Plot

# activate study case

SelCase = app.GetProjectFolder('study').GetContents(StudyCase)[0]
SelCase.Activate()

NetData = app.GetProjectFolder('netdat')
Net = NetData.GetContents('110KV.ElmNet')

newfolder = datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_CC\\'
for Mode in Modes:

    if Mode == 0:
        for varname in VariationName:

            Variation = PFM.ActivateVariation(varname, app)

            Converter = Net[0].GetContents('*.ElmGenStat')[0]
            Frame = Converter.GetAttribute('c_pmod')

            for val in faultvalues:

                # apply desired fault values
                PFM.SetAttributesforFaultEvent(app, Eventname, value=val)

                # apply desired inertia values
                for inval in inertiavalues:

                    if varname == 'Synchronverter':

                        print('Applying inertia value ' + str(inval))
                        Frame.GetAttribute('Synchronverter control').SetAttribute('Ta', inval)

                    elif varname == 'VSM':

                        print('Applying inertia value ' + str(inval))
                        Frame.GetAttribute('Grid-forming control').SetAttribute('Ta', inval)

                    else:

                        print('Converter has no inertia')

                    path1 = DM.ReadorCreatePath('Create', folder=newfolder, filename= "results{controller}fault{faults}inertia{inertia}.csv".format(controller=varname, faults=faultvalues,inertia=inval))

                    # EXECUTING SIMULATION
                    PFM.RunNSave(app, True, tstop=10, path=path1)

    elif Mode == 1:

        path2 = DM.ReadorCreatePath('Read', readmode='lastfile')
        # import results
        for direc in os.listdir(path):
            Results = DM.importData(path + direc).astype(float)

            ResultsList.append(Results)



    elif Mode == 2:

        # print results
        seriesnames = list(map(str, inertiavalues))
        #seriesnames = VariationName
        #seriesnames.append('Voltage Source')
        #columns = [[1,7],[18, 24],[18, 24],[1, 7]]
        columns = [[18,24],[18, 24],[18, 24]]
        DM.DFplot(ResultsList, [1, 1],
                  xaxis=0,
                  xlabel='Time (s)',
                  ylabel=['Voltage (p.u.)', 'Frequency (Hz)'],
                  fixplot=[25 ,31],
                  seriesnames=seriesnames,
                  savefigures=True,
                  figurefolder='converters_comparison/'+StudyCase+'/')

app.PostCommand("exit")
