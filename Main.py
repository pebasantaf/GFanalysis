import sys
import os
import datetime
import functions.PFmanager as PFM
import functions.datamanager as DM
from pprint import pprint


sys.path.append(r'C:\Program Files\DIgSILENT\PowerFactory 2021 SP1\Python\3.8')

import powerfactory as pf

#Initiate powerfactory. In case there is an error, will give the error code number
try:
    app = pf.GetApplicationExt()
    # ... some      calculations ...
except pf.ExitError as error:
    print(error)
    print('error.code = %d' % error.code)

# Activate the project
projname = str("DEAModel_gridforming_20210204")

app.ActivateProject(projname)

# Get project that is active as and application
project = app.GetActiveProject()


# Get Calculation object by getting active study case. This active case is related to a grid or part of a grid that is active.
#Then, getting the initial conditions and the simulation object to being executed at the end

#  function CreateSimpleStabilityStudy

PFM.CreateSimpleStabilityStudy(app)


#get project folder

# manage variations

ResultsList = list()
VariationName = ["Droop", "VSM", "Synchronverter"]

for varname in VariationName:


    Variation = PFM.ActivateVariation(varname, app)


    NetData = app.GetProjectFolder('netdat')
    Net = NetData.GetContents('110KV.ElmNet')
    # ConvComp = Net[0].GetContents('Test-gridformingConverter.ElmComp')

    '''
    #current control frame
    ControlFrameName = ConvComp[0].typ_id.GetFullName().split('\\')[-1][:-7]
    
    # select control frame to apply
    
    # get folder with user models
    folder = app.GetProjectFolder('blk')
    pprint(folder.GetContents())
    nfold = int(input("Select which control frame you want to use: "))
    frame = folder.GetContents()[nfold] #select 2 or three regarding the folder we want
    
    # control frame options
    
    frameopts = ["Grid_forming_VSM0H", "Grid_forming_VSM"]
    
    
    if ControlFrameName != frameopts[nfold]: #if our selected frame is not applied yet
    
        ConvComp[0].typ_id = frame.GetContents('*.BLkDef')[0]
    '''
    ComInc = app.GetActiveStudyCase().GetContents('*.ComInc')[0]
    ComSim = app.GetActiveStudyCase().GetContents('*.ComSim')[0]


    # these functions optimize performnance
    app.WriteChangesToDb()
    app.SetWriteCacheEnabled(0)

    # excecute simulation

    ComInc.Execute() # execute initial conditions
    ComSim.Execute() # execute simulation


    # execute the object to export results

    if VariationName.index(varname) == 0:
        Resultspath = os.getcwd()+"\\Results"
        newfolder = datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S\\")
        os.mkdir(Resultspath + newfolder)

    f_name = Resultspath + newfolder + "/results{controller}.csv".format(controller=varname)

    ElmRes = app.GetFromStudyCase("ElmRes") #results object

    freqdrop = app.GetFromStudyCase('IntCase')  # study case
    # Events = freqdrop.GetContents('*.EvtParam') contents of event

    # execute results
    PFM.execComRes(freqdrop,ElmRes,f_name,
                   iopt_exp=6,
                   iopt_sep=0,
                   dec_Sep='.')

    # import results
    Results = DM.importData(f_name).astype(float)

    ResultsList.append(Results)

seriesnames = VariationName
seriesnames.append('Node')
columns = [1, 7]
DM.DFplot(ResultsList, 1, [2, 1], columns,
          xaxis=0,
          fixplot=[8, 9],
          xlabel='Time (s)',
          ylabel=['Voltage (p.u.)', 'Frequency (Hz)'],
          seriesnames=seriesnames)

app.PostCommand("exit")
