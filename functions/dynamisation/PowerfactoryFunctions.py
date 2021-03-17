import pandas as pd
import time
import pickle
# ----------------------------------------------------------------------------------------------------------------------
# Importiert eine pfd-Datei
# ----------------------------------------------------------------------------------------------------------------------
def ImportPfdFile(app, path):
    # Nutzer des Programmes auslesen
    user = app.GetCurrentUser()

    # Referenzprojektliste erstellen zum Abgleich mit der Liste danach
    projectsref = user.GetContents('*.IntPrj')

    # Script für den Import der pfd-Datei
    script = user.CreateObject('ComDpl', 'PFD-Import')
    obj = script.CreateObject('ComPfdimport', 'PFD-Import')

    # Pfad für den Aufruf der pfd-Datei setzen
    obj.SetAttribute('g_file', path)

    # Zielnutzer setzen
    obj.SetAttribute('g_target', user)

    # Script ausführen und anschließend löschen
    obj.Execute()
    script.Delete()

    # Abgleich mit alter Projektliste um neue Projekte zu identifizieren
    projects = user.GetContents('*.IntPrj')
    prjlist = [prj for prj in projects if prj not in projectsref]

    # Fehlerbehandlung -> Muss noch ausgeweitet werden
    if len(prjlist) == 1:
        prj = prjlist[0]
    else:
        print('ERR: Mehr als ein Projekt wurde beim Datenimport in einer pfd-Datei geladen')

    return prj


'''
------------------------------------------------------------------------------------------------------------------------
ExportToPfd(): Exports a project to a pfd file
------------------------------------------------------------------------------------------------------------------------
Input:
    - app: PowerFactory application
    - path: Output path for the pfd file
    - ExpObj: List of objects, which should be exported to the file
------------------------------------------------------------------------------------------------------------------------
'''
def ExportToPfd(app, path, ExpObj): # ToDo Überarbeiten

    # Get current user
    user = app.GetCurrentUser()

    # Create Script for Pfd-Export
    script = user.CreateObject('ComDpl', 'PFD-Export')
    obj = script.CreateObject('ComPfdexport', 'PFD-Export')

    # Path with file name for the export
    obj.SetAttribute('g_file', path)

    # Add object for export to selection
    obj.SetAttribute('g_objects', ExpObj)
    obj.SetAttribute('exportExternalFiles', 1)

    # Script ausführen und anschließend löschen
    obj.Execute()
    script.Delete()

# Get any PowerFactory object or create new one if none was found
# Additionally, any parameters can be set
def GetPowerFactoryObject(ElmType, location, loc_name='*', **kwargs):
    # Check, if object already exists (if new element is not required/forced)
    try:
        obj = location.GetContents(loc_name + '.' + ElmType)[0]
    except:
        if '*' not in loc_name:
            obj = location.CreateObject(ElmType, loc_name)
        else:
            obj = location.CreateObject(ElmType)

    # Apply settings
    for key, var in kwargs.items():
        obj.SetAttribute(key, var)

    return obj

# Execute Result Export
def execComRes(IntCase, ElmRes, f_name, **kwargs):

    try:
        ComRes = IntCase.GetContents('*.ComRes', 1)[0]
    except:
        ComRes = IntCase.CreateObject('ComRes', 'ASCII - Ergebnis Export')

    ComRes.SetAttribute('pResult', ElmRes)

    ComRes.SetAttribute('f_name', f_name)

    for key, var in kwargs.items():
        ComRes.SetAttribute(key, var)

    ComRes.Execute()


# Function to prepare a ComSim calculation
def PrepComSim(app, ElmRes, ElmVarObservation, optComLdf={}, optComInc={}, optComSim={}, ComTasks=None):

    # Add variables for each object in input data
    app.WriteChangesToDb()
    app.SetWriteCacheEnabled(0)
    for Elm, VarList in ElmVarObservation.items():
        for var in VarList:
            ElmRes.AddVariable(Elm, var)
    app.SetWriteCacheEnabled(1)

    # Parameterize calculations
    ComLdf = GetPowerFactoryObject('ComLdf', app.GetActiveStudyCase(), loc_name='*', **optComLdf)
    ComInc = GetPowerFactoryObject('ComInc', app.GetActiveStudyCase(), loc_name='*', **optComInc)
    ComSim = GetPowerFactoryObject('ComSim', app.GetActiveStudyCase(), loc_name='*', **optComSim)

    # Add to task automation if available
    if ComTasks is not None:
        ComTasks.AppendCommand(ComInc)
        ComTasks.AppendCommand(ComSim)

    return [ComInc, ComSim]

# ----------------------------------------------------------------------------------------------------------------------
# Run RMS Simulation and saves the results file
# ----------------------------------------------------------------------------------------------------------------------
def Run_RMS(ComInc, ComSim):

    start = time.time()
    ComInc.Execute()
    ComSim.Execute()
    end = time.time()
    elapsed_time = (end - start)

    return elapsed_time


def RetrieveResults(ElmRes, Variables, Objects,resolution):
    ElmRes.Load()
    NumCol = ElmRes.GetNumberOfColumns()
    NumRow = ElmRes.GetNumberOfRows()
    name = []

    df = pd.DataFrame(columns=range(len(Objects)))

    for i in range(NumRow):
        L = []
        for n in range(len(Variables)):
            ColIndex = ElmRes.FindColumn(Objects[n], Variables[n])
            try:
                Value = ElmRes.GetValue(i*resolution, ColIndex)[1]
                L.append(Value)
            except:
                continue


        df.loc[len(df)] = L

    return df


# Import csv data
def importData(path, combine_headers=True, sep=';'):

    # Import results from files
    EV_raw_df = pd.read_csv(path, sep=sep, header=[0, 1])

    # Reset headers
    EV_raw_df.columns = pd.MultiIndex.from_tuples(EV_raw_df.columns.tolist())
    Head_Obj = EV_raw_df.columns.get_level_values(0).tolist()
    Head_Var = EV_raw_df.columns.get_level_values(1).tolist()
    Head_Var = [s.split(' in ')[0] for s in Head_Var]  # take Varname, excluding the units (Original: Varname in MW)

    if combine_headers == True:
        new_Head = [Head_Obj[i] + '|' + Head_Var[i] for i in range(len(Head_Obj))]
    else:
        new_Head = [Head_Var[i] for i in range(len(Head_Obj))]

    EV_raw_df.columns = new_Head

    return EV_raw_df