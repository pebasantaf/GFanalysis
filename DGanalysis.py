import sys
import os
import datetime
import functions.convcompar.datamanager as DM
import functions.convcompar.PFmanager as PFM
from pathlib import Path
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

# project and simulation inputs

proj_name = '1-MVLV-rural-all-0-sw'
GFol_model = r"I:\05_Basanta Franco\Masterarbeit_local\Models\DEAModel_20210209_PQ_Limits.pfd"

app.ActivateProject(proj_name)
proj = app.GetActiveProject()

gentype = 'Wind'
convtype = 'GridFollowing'

# Mode: 0-Change Converters; 1-Execute simulation; 2-Import data; 3-Plot

Modes = 1
filemode = 'Create' #filemodes: Create or Read

# results input

if filemode == 'Create':

    folder = os.getcwd() + "\\Results" + datetime.datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S\\")
    os.mkdir(folder)
    path = folder + proj_name + "-GF-wind.csv"

elif filemode == 'Read':

    path = sorted(Path(os.getcwd()+'/Results').iterdir(), key=os.path.getmtime)[-1].__str__() #reads last file. This can be modified

# apply converters and select what to execute

for Mode in Modes:

    if Mode == 0:

        PFM.ApplyConverters(app, proj, GFol_model, gentype, convtype)

    elif Mode == 1:

        # excute and export
        PFM.RunNSave(app, path, tstop=1)

    elif Mode == 2:

        # import results
        Results = DM.importData(path).astype(float)

    elif Mode == 3:

        # select plots
        voltage_plot = [s for s in list(Results.columns) if 'u1' in s]
        voltage_id = [i for i in range(len(list(Results.columns))) if list(Results.columns)[i] in voltage_plot]

        frequency_plot = [s for s in list(Results.columns) if 'Electrical Frequency' in s]
        frequency_id = [i for i in range(len(list(Results.columns))) if list(Results.columns)[i] in frequency_plot]

        activepower_plot = [s for s in list(Results.columns) if 'Total Active Power' in s]
        activepower_id = [i for i in range(len(list(Results.columns))) if list(Results.columns)[i] in activepower_plot]

        # plot

        for id in frequency_id:
            DM.DFplot(Results, 1, [1, 1], columns,
                      xaxis=0,
                      xlabel='Time (s)',
                      ylabel=['Voltage (p.u.)', 'Frequency (Hz)'],
                      fixplot=[25, 31],
                      seriesnames=seriesnames,
                      savefigures=True)

    else:
        raise Exception('Wrong value')






