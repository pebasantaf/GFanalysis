import sys
import os
import datetime
import functions.convcompar.datamanager as DM
import functions.convcompar.PFmanager as PFM
from prettytable import PrettyTable


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

# RESET

Reset = False

if Reset:

    gentype = ['Wind', 'Photovoltaic']
    convtype = 'GridFollowing'
    print('Reseting Power Factory grid to initial state')
    PFM.ApplyConverters(app, proj, GFol_model, gentype, convtype, convname='LV2')
    quit()

# INPUT

gentype = [['Wind'], ['Wind', 'Photovoltaic']]
convtype = 'Synchronverter'

# Mode: 0-Change Converters; 1-Execute simulation; 2-Import data; 3-Plot

Modes = [1, 0, 1, 0, 1]
filemode = 'Create' #filemodes: Create or Read
folder = datetime.now().strftime("\\%d.%m.%Y_%H-%M-%S") + '_DG\\'
fileending = [".csv", "-GF-wind.csv", "-GF-wind-solar.csv"]


g = 0  # generation type counter
c = 0  # converter type counter
f = 0  # file counter

ResultsList = list()

for i in range(len(Modes)):

    if Modes[i] == 0:

        print('Applying ' + convtype +' for '+ str(gentype[g]) + ' generators')

        PFM.ApplyConverters(app, proj, GFol_model, gentype[g], convtype, convname='LV2')

        print(convtype + ' applied')

        if g + 1 < len(gentype):
            g += 1

        if c + 1 < len(convtype):
            c += 1

    elif Modes[i] == 1:

        path = DM.ReadorCreatePath('Create', folder=folder, filename=proj_name+fileending[f])
        f += 1

        print('Executing simulation and storing results')

        # excute and export
        PFM.RunNSave(app, 1, tstop=1, path=path)
        print('Simulation finished')

    elif Modes[i] == 2:

        path = DM.ReadorCreatePath('Read', readmode='lastfile')

        # import results
        for direc in os.listdir(path):

            Results = DM.importData(path+direc).astype(float)

            ResultsList.append(Results)

    elif Modes[i] == 3:
        # plot

        DM.DFplot(ResultsList, len(columns), [1, 1], columns,
                  xaxis=0,
                  xlabel='Time (s)',
                  savefigures=False,
                  ylabel=list(map(list(ResultsList[0].columns).__getitem__, columns)),
                  seriesnames=['0GF', 'WindGF', 'WindLV2SolarGF']
                  )

    else:
        raise Exception('Wrong value')






