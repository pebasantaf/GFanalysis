import sys
import time
import os
import pandas as pd
import math
import functions.aggregationstudy.aggfunctions as AGF
import functions.convcompar.datamanager as DTM
import matplotlib.pyplot as plt
import numpy as np
from scipy.stats import pearsonr
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

mode = 'RXanalysis' # RXanaylsis, RMSE
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

        R = [None] * len(columns)
        X = [None] * len(columns)
        for scen in columns:

            R[columns.index(scen)] = networkdata.GetContents('MV_' + scen)[0].GetContents('*.ElmLne')[0].GetAttribute('R1')
            X[columns.index(scen)] = networkdata.GetContents('MV_' + scen)[0].GetContents('*.ElmLne')[0].GetAttribute('X1')

        Rdf.loc[rows[prjlist.index(prj)]] = R
        Xdf.loc[rows[prjlist.index(prj)]] = X

    DTM.DFplot([Rdf,Xdf], 'aggRX',
               savefigures=False,
               figurefolder='grid aggregation\line_parameter/x.png')

elif mode == 'RMSE':

    print('Obtaining all csv for RMSE calculation')

    # selecte noramlization mode and var to plot
    var2plot = 'I1'
    normmode = 'commonscenario' # 'commonscenario' or 'individual'

    #select controller folder
    basepath = r"I:\05_Basanta Franco\Exchange\MasterThesis\Aggregation/"
    controller = 'constv'
    thedir = basepath + controller

    print('Data in folder: ' + thedir)

    #select csv files in cntroller folder
    subdirs = [name for name in os.listdir(thedir) if os.path.isdir(os.path.join(thedir, name))] #get all directory names in thedir
    grids = [m.split('_')[0] for m in subdirs]

    #initialize dataframes
    RMSEdf = pd.DataFrame(columns=escens, index=grids)
    normRMSE = pd.DataFrame(columns=escens, index=grids)

    # initialize df depending on mode
    if normmode == 'commonscenario':

        normdf = pd.DataFrame(columns=escens, index=['max', 'min','dif'])

    elif normmode == 'individual':

        normdf = pd.DataFrame(columns=escens, index=grids)

    # for every grid setup

    for folder in subdirs:

        datadir = thedir + '/' + folder + '/' + '1-MVLV-rural-all-0-sw/'

        targetcsv = [file for file in os.listdir(datadir) if file.endswith(".csv") and 'MV' in file] # get the MV csvs

        for file in targetcsv: # for every csv

            df = pd.read_csv(os.path.join(datadir,file), delimiter=';',skiprows=[0]) # import the df

            # common scenario normalization: find max and min of a certain scenario for all grids

            # what to plot
            if var2plot == 'I1': # if we want absolute current, first we must calculated it

                activecol = [n for n in list(df.columns) if 'I1P' in n]
                reactivecol = [n for n in list(df.columns) if 'I1Q' in n]

                detcurr = np.sqrt(df[activecol[0]] ** 2 + df[reactivecol[0]] ** 2)
                eqcurr = np.sqrt(df[activecol[1]] ** 2 + df[reactivecol[1]] ** 2)

            else: # for the active and reactive current cases

                col2calc = [n for n in list(df.columns) if var2plot in n] # localize columns to be plotted
                detcurr = df[col2calc[0]]
                eqcurr = df[col2calc[1]]

            # calculate rmse

            RMSE = AGF.CalculateRMSE(detcurr, eqcurr)

            # way of calculating normalization min-max range: for each grid and scenario a value or for each scenario a value

            if normmode == 'commonscenario':

                actscen = file.split('_')[2].split('.')[0]

                actmax = max(detcurr)
                actmin = min(detcurr)

                storemax = normdf.loc['max', actscen]
                storemin = normdf.loc['min', actscen]

                if (actmax > storemax) or (subdirs.index(folder) == 0): # if present current value is bigger than stored for scenario or if it is the first grid

                    normdf.loc['max',actscen] = actmax

                else:
                    pass

                if actmin < storemin or (subdirs.index(folder) == 0): # if present current value is smaller than stored for scenario or if it is the first grid

                    normdf.loc['min', actscen] = actmin

                else:
                    pass

                normdf.loc['dif',actscen] = abs(normdf.loc['max',actscen] - normdf.loc['min',actscen])


            elif normmode == 'individual':

                normvalue = abs(max(detcurr)-min(detcurr))
                normdf.loc[folder.split('_')[0], escens[targetcsv.index(file)]] = normvalue


            RMSEdf.loc[folder.split('_')[0], escens[targetcsv.index(file)]] = RMSE

    # normalize RMSE
    if normmode == 'commonscenario':

        for index,row in RMSEdf.iterrows():

            normRMSE.loc[index] = (row / normdf.loc['dif'])*100

    elif normmode == 'individual':

        normRMSE = (RMSEdf/normdf)*100

    print('RMSE calculated and stored in dataframe')



    DTM.DFplot([normRMSE], 'aggeval',
                savefigures=True,
                figurefolder= 'grid aggregation\Vdip_0,5_RMSE/' + mode + '_' + var2plot + '_' + controller + '.png')

