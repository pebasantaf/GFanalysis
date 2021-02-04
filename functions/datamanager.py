import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

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

def DFplot(DFlist, nfigs,nsplts,columns, **kwargs):

    # This function plots a variable number of figures with variable number of supplots that all have the same x axis (normally time)
    # DFlist: list with the dataframes imported from csvs
    # nfigs: how many figures we want
    # nsplts: list with two integers indicating rows and colums of subplots
    # columns: columns to plot

    for n in range(nfigs):
        plt.figure(n+1)

        for m in range(np.prod(nsplts)):
            plt.subplot(nsplts[0], nsplts[1], m+1)
            c = 0

            for DF in DFlist:

                xdata = kwargs.get('xaxis')

                plt.figure(n+1).axes[m].plot(DF.iloc[:, [xdata]], DF.iloc[:, [columns[m]]], label=kwargs.get('seriesnames')[c])

                plt.figure(n+1).axes[m].set_xlabel(DF.columns[xdata])
                plt.figure(n+1).axes[m].set_ylabel('cosa')
                plt.figure(n + 1).axes[m].legend()


                c += 1

            plt.figure(n+1).axes[m].yaxis.grid()

    plt.show()