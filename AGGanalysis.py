import sys
import time
import os
import pandas as pd
import math
import functions.aggregationstudy.aggfunctions as AGF
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

    var2plot = 'I1P'
    scendict = {p: 0 for p in escens}

    basepath = r"I:\05_Basanta Franco\Exchange\MasterThesis\Aggregation/"
    controller = 'constphi'
    thedir = basepath + controller

    subdirs = [name for name in os.listdir(thedir) if os.path.isdir(os.path.join(thedir, name))]
    grids = {m.split('_')[0]:0 for m in subdirs}

    for folder in subdirs: #for every grid setup

        datadir = thedir + '/' + folder + '/' + '1-MVLV-rural-all-0-sw/'

        targetcsv = [file for file in os.listdir(datadir) if file.endswith(".csv") and 'MV' in file] # get the MV csvs

        for file in targetcsv: # for every csv

            df = pd.read_csv(os.path.join(datadir,file), delimiter=';',skiprows=[0]) # import the df

            col2calc = [n for n in list(df.columns) if var2plot in n] # localize columns to be plotted

            RMSE = AGF.CalculateRMSE(df[col2calc[0]],df[col2calc[1]] ) # calculate rmse
            normRMSE = RMSE/df[col2calc[0]][0] # normalize rmse

            scendict[file.split('_')[2].split('.')[0]] = normRMSE #store rmse information in dictionary

        grids[folder.split('_')[0]] = scendict # nest dictionary with scenario info in grid dictionary


