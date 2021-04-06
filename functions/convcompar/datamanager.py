import os
import sys
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

    moreplots = True

    while moreplots:

        table = PrettyTable()
        table.add_column('id', list(range(0, len(list(DFlist[0].columns)))))
        table.add_column('Variable', list(DFlist[0].columns))
        print(table)

        if 'columns' not in kwargs.items() or kwargs.get('columns') == None: #if variable columns is not a kwarg, ask for input

            columns = input('Enter column id (format: x, x, x): ')
            columns = columns.split(',')
            columns = list(map(int, columns))

        else:
            columns = kwargs.get('columns')

        # if amount of figures is not given, just take the number of column
        if 'nfigs' not in kwargs.items():
            nfigs = len(columns)

        else:
            nfigs = kwargs.get('nfigs')

        variables = DFlist[0].columns

        # create reference strings. This is useful as sometimes equivalent columns of data from different converters will be in different positions in the dataframes

        reflist = list()

        for col in columns:

            reference = 'Converter' + '|' + DFlist[0].columns[col].split('|')[1]
            reflist.append(reference)

        # generate figurs, plots and subplots
        for n in range(nfigs):

            plt.figure(n+1)

            for m in range(np.prod(nsplts)):
                plt.subplot(nsplts[0], nsplts[1], m+1)

                # a plot that is common for all DF that only needs to be plotted once
                if kwargs.get('fixplot') == True:

                    fixplotcol = input('Enter column id for reference variable (format: x, x, x): ')
                    fixplotcol = fixplotcol.split(',')
                    fixplotcol = list(map(int, fixplotcol))

                    plt.figure(n + 1).axes[m].plot(DFlist[0].iloc[:, [kwargs.get('xaxis')]], DFlist[0].iloc[:, fixplotcol[m]],
                                                   label=kwargs.get('seriesnames')[-1], linewidth=0.5)

                #dataframe-dependent plots
                for c in range(len(DFlist)):

                    targetcol = [s for s in DFlist[c].columns if reflist[n] in s]
                    plt.figure(n+1).axes[m].plot(DFlist[c].iloc[:, [kwargs.get('xaxis')]], DFlist[c].loc[:, targetcol], label=kwargs.get('seriesnames')[c], linewidth=0.8)


                # set up axis labels and legend
                plt.figure(n+1).axes[m].set_xlabel(kwargs.get('xlabel'))

                if 'ylabel' in kwargs.items():

                    plt.figure(n+1).axes[m].set_ylabel(kwargs.get('ylabel')[n].split('|')[1])

                else:

                    plt.figure(n + 1).axes[m].set_ylabel(variables[columns[n]].split('|')[1])

                plt.figure(n+1).axes[m].ticklabel_format(useOffset=False)
                plt.figure(n+1).axes[m].legend()
                plt.figure(n+1).axes[m].yaxis.grid()

            # save images in folder

            if 'savefigures' in kwargs:

                if kwargs.get('savefigures') == True:
                    figure = plt.gcf()
                    figure.set_size_inches(14, 8)
                    plt.savefig(r'I:\05_Basanta Franco\Masterarbeit_local\images/'+ kwargs.get('figurefolder') + variables[columns[n]].replace('|', '-') + '.png', dpi=300)

                elif kwargs.get('savefigures') == False:

                    pass

            plt.show()

        moreplots = input('Plot more variables?')
        moreplots = moreplots == 'True'

        if moreplots != True and moreplots != False:

            sys.exit('Answer can only be boolean')




def ReadorCreatePath(filemode, **kwargs):

    if filemode == 'Create':

        directory = os.getcwd() + "\\Results" + kwargs.get('folder')

        if not os.path.isdir(directory):
            os.mkdir(directory)

        path = directory + kwargs.get('filename')

        return path

    elif filemode == 'Read':
        if kwargs.get('readmode') == 'lastfile':

            path = sorted(Path(os.getcwd() + '/Results').iterdir(), key=os.path.getmtime)[
                   -1].__str__() + '\\'  # reads last file. This can be modified
            return path

        elif kwargs.get('readmode') == 'userfile':

            path = os.getcwd() + '/Results' + kwargs.get('folder') + kwargs.get('filename')

            return path


