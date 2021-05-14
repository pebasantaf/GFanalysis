import numpy as np

def CalculateRMSE(I_det, I_eq):
    return np.sqrt((1 /len(I_det)) * np.sum((I_det-I_eq)**2))