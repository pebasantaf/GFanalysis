from __future__ import absolute_import, division, print_function, unicode_literals
from sklearn.metrics import mean_squared_error
import numpy as np
import matplotlib.pyplot as plt
from math import sqrt

# functions for plotting and testing

def TestPlot(model, tests):
    t = 1
    fig1 = plt.figure()
    fig2 = plt.figure()
    RMSE = []

    for te in tests:
        NN_pred = model.predict(te[0], batch_size=1)

        # ANALYSIS
        # reshape the prediction for plotting

        NN_pred = np.reshape(NN_pred, (te[1].shape[0], te[1].shape[2]))
        te[0] = np.reshape(te[0], (te[0].shape[0], te[0].shape[2]))
        te[1] = np.reshape(te[1], (te[1].shape[0], te[1].shape[2]))

        # plotspredicted and desired test output

        # plot currents
        graph1 = fig1.add_subplot(len(tests), 1, t)
        graph1.set_title('Test0' + np.str(t))
        graph1.plot(NN_pred[:, -2:])
        graph1.plot(te[1][:, -2:])
        graph1.legend(['I_P_pred', 'I_Q_pred', 'I_P', 'I_Q'])

        # plot voltages
        graph2 = fig2.add_subplot(len(tests), 1, t)
        graph2.set_title('Test0' + np.str(t))
        graph2.plot(NN_pred[:, 0])
        graph2.plot(te[1][:, 0])
        graph2.legend(['V_0_pred', 'V_0'])

        # mean squared error
        rmse = sqrt(mean_squared_error(te[1], NN_pred))
        RMSE.append(rmse)
        print('Test RMSE: %.3f' % rmse)

        t = t + 1
    '''
    #  print inputs yes or no
    printin = input('Print inputs as well? [y/n]: ')

    m = True
    while m == True:

        if printin == 'y':

            t = 1
            fig4 = plt.figure()

            for te in tests:
                graph4 = fig4.add_subplot(len(tests), 1, t)
                graph4.set_title('Inputs: V, P, Q')
                graph4.plot(te[0][:, 10])
                t = t + 1
            m = False
        elif printin == 'n':

            m = False
        else:

            printin = input('Answer not valid. Print inputs? [y/n]: ')
        '''
    return fig1, fig2, RMSE

def evaluateTraining(history):
    fig3 = plt.figure()

    graph3 = fig3.add_subplot(211)
    graph3.set_title('Loss')
    graph3.plot(history.history['loss'], label='train')
    graph3.plot(history.history['val_loss'], label='test')
    graph3.legend()
    # plot mse during training
    graph3 = fig3.add_subplot(212)
    graph3.set_title('Mean Squared Error')
    graph3.plot(history.history['mse'], label='train')
    graph3.plot(history.history['val_mse'], label='test')

    return fig3
