import sys
import time
import os
import pandas as pd
import math
import functions.aggregationstudy.aggfunctions as AGF
import functions.convcompar.datamanager as DTM
import matplotlib.pyplot as plt
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

mode = 'RMSE' # RXanaylsis, RMSE
escens = ['hPV', 'hW', 'lPV', 'lW']

if mode == 'RXanalysis':

    #  get all projects
    allprj = app.GetCurrentUser().GetContents('*.IntPrj')

    # get relevant projects
    prjlist = [s for s in allprj if '_eq_constphi' in s.GetFullName()]
    columns = escens
    rows = [r.GetFullName().split('\\')[-1].split('_')[1] for r in prjlist]

    Rdf = pd.DataFrame(columns=columns, index=rows)
    Xdf = pd.DataFrame(columns=columns, index=rows)

    for prj in prjlist:

        prj.Activate()
        networkdata = app.GetProjectFolder('netdat')

        R = []
        X = []
        for scen in columns:

            R[columns.index(scen)] = networkdata.GetContents('MV_' + scen)[0].GetContents('*.ElmLne')[0].GetAttribute('R1')
            X[columns.index(scen)] = networkdata.GetContents('MV_' + scen)[0].GetContents('*.ElmLne')[0].GetAttribute('X1')

        Rdf.loc[rows[prjlist.index(prj)]] = R
        Xdf.loc[rows[prjlist.index(prj)]] = X

elif mode == 'RMSE':

    print('Obtaining all csv for RMSE calculation')

    var2plot = 'I1P'
    scendict = {p: 0 for p in escens}

    basepath = r"I:\05_Basanta Franco\Exchange\MasterThesis\Aggregation/"
    controller = 'constphi'
    thedir = basepath + controller

    print('Data in folder: ' + thedir)

    subdirs = [name for name in os.listdir(thedir) if os.path.isdir(os.path.join(thedir, name))] #get all directory names in thedir
    grids = [m.split('_')[0] for m in subdirs]
    RMSEdf = pd.DataFrame(columns=escens, index=grids)

    for folder in subdirs: #for every grid setup

        datadir = thedir + '/' + folder + '/' + '1-MVLV-rural-all-0-sw/'

        targetcsv = [file for file in os.listdir(datadir) if file.endswith(".csv") and 'MV' in file] # get the MV csvs

        for file in targetcsv: # for every csv

            df = pd.read_csv(os.path.join(datadir,file), delimiter=';',skiprows=[0]) # import the df

            col2calc = [n for n in list(df.columns) if var2plot in n] # localize columns to be plotted

            RMSE = AGF.CalculateRMSE(df[col2calc[0]],df[col2calc[1]] ) # calculate rmse
            normRMSE = RMSE/abs(df[col2calc[0]][0]) # normalize rmse

            RMSEdf.loc[folder.split('_')[0], escens[targetcsv.index(file)]] = normRMSE * 100

    print('RMSE calculated and stored in dataframe')

x = np.arange(len(grids))
color = ['b', 'g', 'r', 'y']
width = 0.15
plt.figure(1)


for esc in escens:

    plt.bar(x + (escens.index(esc) * width)-width*2, RMSEdf[esc], color = color[escens.index(esc)], width = width, label=esc, edgecolor='white', align='center', zorder=3)

plt.xticks(x, grids)
plt.xlabel('Grid variations')
plt.ylabel('RMSE (%)')

plt.grid(axis='y')
plt.legend()

plt.show()


#DTM.DFplot(grids, aggcompar)

