
def confselect(conf):

    if conf == 'allfreqramp':

        studcase = 'Frequency Ramp'

        config = {
            'StudyCase': studcase,
            'tinit' : 0.25,
            'tend' : 0.27,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "Grid Following"],
            'faultvalues': ["0.25"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'fixplot': False,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config
    elif conf == 'allvoldip':

        studcase = 'Voltage Step'

        config = {
            'StudyCase': studcase,
            'tinit' : 0.25,
            'tend' : 1,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "Grid Following"],
            'faultvalues': ["-0.1"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'fixplot': False,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config

    elif conf == 'allvolramp':

        studcase = 'Voltage Ramp'

        config = {
            'StudyCase': studcase,
            'tinit' : 3,
            'tend' : 7,
            'Variation Name': ["Droop", "VSM", "Synchronverter", "Grid Following"],
            'faultvalues': ["-0.4"],
            'inertiavalues': [5],
            'seriesnames': ["Droop", "VSM", "Synchronverter", "GridFollowing", "Voltage Source"],
            'xaxis': 0,
            'columns': [1],
            'fixplot': False,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config

    elif conf == 'Inertia':

        inertiavalues = [1,3,5,7,10]
        studcase = 'Frequency Ramp'

        config = {
            'StudyCase': studcase,
            'tinit': 3,
            'tend': 7,
            'Variation Name': ["VSM"],
            'faultvalues': ["0.01"],
            'inertiavalues': inertiavalues,
            'seriesnames': ['Inertia: ' + s for s in list(map(str, inertiavalues))],
            'xaxis' : 0,
            'savefigures': True,
            'figurefolder': 'converters_comparison/' + studcase + '/inertia/'
        }
        return config
    elif conf == 'multifault':

        studcase = 'Voltage Step'
        faultvalues = ["-0.2","-0.4","-0.6","-0.8" ]
        config = {
            'StudyCase': studcase,
            'tinit' : 3,
            'tend' : 7,
            'Variation Name': ["Synchronverter"], #["Droop", "VSM", "Synchronverter", "Grid Following"],
            'faultvalues': faultvalues,
            'inertiavalues': [5],
            'seriesnames': ['Fault value: ' + s for s in faultvalues],
            'xaxis': 0,
            'fixplot': False,
            'savefigures': True,
            'figurefolder': 'converters_comparison/'+studcase+'/totalcompar/'
        }

        return config