
def confselect(conf):

    if conf == 'freqramp':

        studcase = 'Frequency Ramp'

        config = {
            'StudyCase': studcase,
            'tinit' : 3,
            'tend' : 7,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "Grid Following"],
            'faultvalues': ["0.001"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'fixplot': False,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config
    elif conf == 'voldip':

        studcase = 'Voltage Step'

        config = {
            'StudyCase': studcase,
            'tinit' : 3,
            'tend' : 7,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "GridFollowing"],
            'faultvalues': ["0.6"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'columns': [[1,7],[18, 24],[18, 24],[1, 7]],
            'fixplot': True,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config

    elif conf == 'volramp':

        studcase = 'Voltage Ramp'

        config = {
            'StudyCase': studcase,
            'tinit' : 3,
            'tend' : 7,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "GridFollowing"],
            'faultvalues': ["-0.4"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'columns': [1],
            'fixplot': True,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config

    elif conf == 'Inertia':

        inertiavalues = [3,5,7,10]
        studcase = 'Frequency Ramp'

        config = {
            'StudyCase': studcase,
            'tinit': 3,
            'tend': 7,
            'Variation Name': ["VSM"],
            'faultvalues': ["0.001"],
            'inertiavalues': inertiavalues,
            'seriesnames': ['Inertia: ' + s for s in list(map(str, inertiavalues))],
            'xaxis' : 0,
            'ylabel': ['Voltage (p.u.)', 'Frequency (Hz)'],
            'savefigures': True,
            'figurefolder': 'converters_comparison/' + studcase + '/inertia/'
        }
        return config