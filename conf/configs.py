
def confselect(conf):

    if conf == 'freqramp':

        config = {
            'StudyCase': 'Frequency Ramp',
            'Eventname': 'fslope',
            'Variation Name': ["Droop", "VSM", "Synchronverter", "GridFollowing"],
            'faultvalues': 0.001,
            'inertiavalues': [3],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'columns': [[1,7],[18, 24],[18, 24],[1, 7]],
            '['Voltage (p.u.)', 'Frequency (Hz)']'
        }

    return config