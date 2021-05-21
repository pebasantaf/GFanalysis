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

    for s in range(len(Head_Var)):
        if 'in' in Head_Var[s]:
            Head_Var[s] = Head_Var[s].split(' in ')[0] + ' (' + Head_Var[s].split('in ')[1] + ')'

        else:

            Head_Var[s] = Head_Var[s]

    new_Head = [Head_Obj[i] + '|' + Head_Var[i] for i in range(len(Head_Obj))]
    EV_raw_df.columns = new_Head

    return EV_raw_df


def DFplot(DFlist, mode, **kwargs):

    # This function plots a variable number of figures with variable number of supplots that all have the same x axis (normally time)
    # DFlist: list with the dataframes imported from csvs
    # mode: what do we wanna plot (grid results, aggregation results, etc.). convcompar, dncompar, aggcompar
    # kwargs:
        # columns: integer list with which columns to plot
        # nfgis: number of figures i want to plot
        # fixplot: in convcompar and dgcompar mode, we plot results from different DF. If we want to plot a variable that is equal for all and should not be repited (a reference value, f.e. ), it should go here
        # xaxis: which column should be the xaxis
        # seriesname: names for legend
        # xlabel: name for x axis
        # ylable: name for y axis
        # savefigures: save plotted figures in folder
        # figurefolder

    moreplots = True

    while moreplots:

        if mode == 'convcompar' or mode == 'dncompar':  # if the selected mode is comparison of models in test grid or of distribution grids
            variables = DFlist[0].columns

            # show available columns

            table = PrettyTable()
            table.add_column('id', list(range(0, len(list(variables)))))
            table.add_column('Variable', list(variables))
            print(table)

            # if amount of figures is not given, just take the number of columns selected (meaning, having as many figures as variables selected for plotting)
            if 'nfigs' not in kwargs.items():
                nfigs = len(columns)

            else:
                nfigs = kwargs.get('nfigs')


            # create reference strings. This is useful as sometimes equivalent columns of data from different converters will be in different positions in the dataframes

            reflist = list()
            if mode == 'convcompar':
                for col in columns:

                    reference = 'Converter' + '|' + variables[col].split('|')[1]
                    reflist.append(reference)

            elif mode == 'dncompar':

                for col in columns:

                    reference = variables[col]
                    reflist.append(reference)


            # generate figurs, plots and subplots
            for n in range(nfigs):

                plt.figure(n+1)

                # a plot that is common for all DF that only needs to be plotted once
                if kwargs.get('fixplot') == True:

                    fixplotcol = input('Enter column id for reference variable (format: x, x, x): ')
                    fixplotcol = fixplotcol.split(',')
                    fixplotcol = list(map(int, fixplotcol))

                    plt.figure(n + 1).axes.plot(DFlist[0].iloc[:, [kwargs.get('xaxis')]], DFlist[0].iloc[:, fixplotcol],
                                                   label=kwargs.get('seriesnames')[-1], linewidth=0.5)

                #dataframe-dependent plots
                for c in range(len(DFlist)):

                    targetcol = [s for s in DFlist[c].columns if reflist[n] in s]
                    plt.figure(n+1).axes.plot(DFlist[c].iloc[:, [kwargs.get('xaxis')]], DFlist[c].loc[:, targetcol], label=kwargs.get('seriesnames')[c], linewidth=1.2)


                # set up axis labels and legend
                plt.figure(n+1).axes.set_xlabel(kwargs.get('xlabel'))

                if 'ylabel' in kwargs.items():

                    plt.figure(n+1).axes.set_ylabel(kwargs.get('ylabel')[n].split('|')[1])

                else:

                    plt.figure(n + 1).axes.set_ylabel(variables[columns[n]].split('|')[1])

                plt.figure(n+1).axes.ticklabel_format(useOffset=False)
                plt.figure(n+1).axes.legend()
                plt.figure(n+1).axes.yaxis.grid()


        elif mode == 'aggeval':

            DF = DFlist[0]

            escens = list(DF.columns)
            x = np.arange(DF.shape[0])
            color = ['b', 'g', 'r', 'y']
            width = 0.15
            plt.figure(1)

            for esc in escens:
                plt.bar(x + (escens.index(esc) * width) - width * 2, DF[esc], color=color[escens.index(esc)],
                        width=width, label=esc, edgecolor='white', align='center', zorder=3)

            plt.xticks(x, list(DF.index))
            plt.xlabel('Grid variations')
            plt.ylabel('RMSE (%)')

            plt.grid(axis='y')
            plt.legend()

        elif mode == 'aggRX':

            labels = ['R', 'X']
            for n in range(len(DFlist)):

                escens = list(DFlist[n].columns)
                x = np.arange(DFlist[n].shape[0])
                color = ['b', 'g', 'r', 'y']
                marker = ['o', '+', 'v', 'd']

                plt.figure(n)

                for esc in escens:
                    plt.plot(x, DFlist[n][esc], color=color[escens.index(esc)], marker=marker[escens.index(esc)],
                            label=esc, linestyle='None', zorder=3)

                plt.xticks(x, list(DFlist[n].index))
                plt.xlabel('Grid variations')
                plt.ylabel(labels[n])

                plt.grid(axis='y')
                plt.legend()


        # save images in folder

        if 'savefigures' in kwargs:

            if kwargs.get('savefigures') == True:

                if mode == 'convcompar' or mode == 'dncompar':

                    figure = plt.gcf()
                    figure.set_size_inches(14, 8)
                    plt.savefig(r'I:\05_Basanta Franco\Masterarbeit_local\images/'+ kwargs.get('figurefolder') + reflist[n].replace('|', '-').replace('/',' ').replace('\\', ' ') + '.png', dpi=600)

                elif mode == 'aggeval':

                    figure = plt.gcf()
                    figure.set_size_inches(14, 8)
                    plt.savefig(r'I:\05_Basanta Franco\Masterarbeit_local\images/' + kwargs.get('figurefolder'))

            elif kwargs.get('savefigures') == False:

                pass

        plt.show()

        moreplots = input('Plot more variables? ')
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


