import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from prettytable import PrettyTable
from pathlib import Path


def importData(path):

    # Import results from files
    EV_raw_df = pd.read_csv(path, sep=';', header=[0, 1],encoding='cp1252')

    # Reset headers ISO-8859-1
    EV_raw_df.columns = pd.MultiIndex.from_tuples(EV_raw_df.columns.tolist())
    Head_Obj = EV_raw_df.columns.get_level_values(0).tolist()
    Head_Var = EV_raw_df.columns.get_level_values(1).tolist()
    Head_Var = [s.split(' in ')[0] for s in Head_Var]  # take Varname, excluding the units (Original: Varname in MW)
    new_Head = [Head_Obj[i] + '|' + Head_Var[i] for i in range(len(Head_Obj))]
    EV_raw_df.columns = new_Head

    return EV_raw_df


def DFplot(DFlist, nsplts, **kwargs):

    # This function plots a variable number of figures with variable number of supplots that all have the same x axis (normally time)
    # DFlist: list with the dataframes imported from csvs
    # nfigs: how many figures we want
    # nsplts: list with two integers indicating rows and colums of subplots
    # columns: columns to plot

    table = PrettyTable()
    table.add_column('id', list(range(0, len(list(DFlist[0].columns)))))
    table.add_column('Variable', list(DFlist[0].columns))
    print(table)

    columns = input('Enter column id (format: x, x, x): ')
    columns = columns.split(',')
    columns = list(map(int, columns))

    if 'nfigs' not in kwargs.items():
        nfigs = len(columns)

    else:
        nfigs = kwargs.get('nfigs')

    for n in range(nfigs):
        plt.figure(n+1)

        for m in range(np.prod(nsplts)):
            plt.subplot(nsplts[0], nsplts[1], m+1)
            c = 0

            if 'fixplot' in kwargs:
                plt.figure(n + 1).axes[m].plot(DFlist[0].iloc[:, [kwargs.get('xaxis')]], DFlist[0].iloc[:, [kwargs.get('fixplot')[m]]],
                                               label=kwargs.get('seriesnames')[-1])


            #dataframe-dependent plots
            for DF in DFlist:

                plt.figure(n+1).axes[m].plot(DF.iloc[:, [kwargs.get('xaxis')]], DF.iloc[:, [columns[n]]], label=kwargs.get('seriesnames')[c])

                c += 1

            plt.figure(n+1).axes[m].set_xlabel(kwargs.get('xlabel'))
            plt.figure(n+1).axes[m].set_ylabel(kwargs.get('ylabel')[n].split('|')[1])
            plt.figure(n+1).axes[m].ticklabel_format(useOffset=False)
            plt.figure(n+1).axes[m].legend()



            plt.figure(n+1).axes[m].yaxis.grid()

        plt.show()

        if 'savefigures' in kwargs:

            if kwargs.get('savefigures') == True:

                plt.savefig(r'I:\05_Basanta Franco\Masterarbeit_local\images\grid_comparison/' + kwargs.get('ylabel')[n] + '.png')

            elif kwargs.get('savefigures') == False:

                pass


def ReadorCreatePath(filemode, **kwargs):

    if filemode == 'Create':

        directory = os.getcwd() + "\\Results" + kwargs.get('folder')

        if not os.path.isdir(directory):
            os.mkdir(directory)

        path = directory + kwargs.get('filename')

    elif filemode == 'Read':
        if 'readmode' in kwargs == 'lastfile':

            path = sorted(Path(os.getcwd() + '/Results').iterdir(), key=os.path.getmtime)[
                   -1].__str__() + '\\'  # reads last file. This can be modified

        elif 'readmode' in kwargs == 'userfile':

            path = os.getcwd() + '/Results' + kwargs.get('folder') + kwargs.get('filename')


    return path