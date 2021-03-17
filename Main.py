import sys
import os
import datetime
import functions.PFmanager as PFM
import functions.datamanager as DM
# from conf.configs import *
import functions.dynamisation
#from conf.configs import confselect
from pprint import pprint


sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8')

import powerfactory as pf

#Initiate powerfactory. In case there is an error, will give the error code number
try:
    app = pf.GetApplicationExt(None, None, "/ini {path}".format(path= "I:\'05_Basanta Franco'\config.ini"))
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

conf = confselect('freqramp')

StudyCase = 'Frequency Ramp'
VariationName = ["Droop", "VSM", "Synchronverter", "GridFollowing"]

Eventname = 'fslope'
faultvalues = ["0,001"]
inertiavalues = [3]

# activate study case

SelCase = app.GetProjectFolder('study').GetContents(StudyCase)[0]
SelCase.Activate()

for varname in VariationName:

    if VariationName.index(varname) == 0:
        Resultspath = os.getcwd() + "\\Results"
        newfolder = datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S\\")
        os.mkdir(Resultspath + newfolder)

    Variation = PFM.ActivateVariation(varname, app)

    NetData = app.GetProjectFolder('netdat')
    Net = NetData.GetContents('110KV.ElmNet')

    Converter = Net[0].GetContents(varname + 'Converter')[0]
    Frame = Converter.GetAttribute('c_pmod')

    if varname == "VSM":
        DSLobj = Frame.GetContents('Virtual Synchronous Machine')[0]

    elif varname == "Synchronverter":
        DSLobj = Frame.GetContents('Synchronverter')[0]

    if 'DSLobj' in locals():
        id = DSLobj.parnam[0].split(',').index('Ta')
        ElmParams = DSLobj.params



    for val in faultvalues:

        # apply desired fault values
        PFM.SetAttributesforFaultEvent(app, Eventname, value=val)

        # apply desired inertia values
        for inval in inertiavalues:

            if 'DSLobj' in locals():
                ElmParams[id] = inval
                DSLobj.SetAttribute('params', ElmParams)

            # EXECUTING OF SIMULATION

            ComInc = app.GetActiveStudyCase().GetContents('*.ComInc')[0]
            ComSim = app.GetActiveStudyCase().GetContents('*.ComSim')[0]


            # these functions optimize performnance
            app.WriteChangesToDb()
            app.SetWriteCacheEnabled(0)

            # excecute simulation

            ComInc.Execute() # execute initial conditions
            ComSim.Execute() # execute simulation


            # execute the object to export results


            f_name = Resultspath + newfolder + "/results{controller}fault{faults}inertia{inertia}.csv".format(controller=varname, faults=faultvalues,inertia=inval)

            ElmRes = app.GetFromStudyCase("ElmRes") #results object

            freqdrop = app.GetFromStudyCase('IntCase')  # study case

            # execute results and save as csv
            PFM.execComRes(freqdrop,ElmRes,f_name,
                           iopt_exp=6,
                           iopt_sep=0,
                           dec_Sep='.')

            # import results
            Results = DM.importData(f_name).astype(float)

            ResultsList.append(Results)

# print results
#seriesnames = list(map(str,inertiavalues))
seriesnames = VariationName
seriesnames.append('Voltage Source')
columns = [[1,7],[18, 24],[18, 24],[1, 7]]
#columns = [[18,24],[18, 24],[18, 24]]
DM.DFplot(ResultsList, 1, [2, 1], columns,
          xaxis=0,
          xlabel='Time (s)',
          ylabel=['Voltage (p.u.)', 'Frequency (Hz)'],
          fixplot=[25 ,31],
          seriesnames=seriesnames)

app.PostCommand("exit")
