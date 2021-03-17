#-----------------------------------------------------------------------------------------------------------------------
# Erstellt ein ElmNet-Objekt und die zugehörige graphische Darstellung
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmNet(app, ElmNetFolder, gridname):

    # Erstellung des ElmNet-Objekts
    ElmNet = ElmNetFolder.CreateObject("ElmNet", gridname)

    return ElmNet
#-----------------------------------------------------------------------------------------------------------------------
# Kopiert Benutzerdefinierte Modelle
#-----------------------------------------------------------------------------------------------------------------------
def get_usermodell(app, prjname):

    app.ActivateProject(prjname)
    prj = app.GetActiveProject()

    # Einlesen der Bibliothek-Ordnerstruktur
    bibfolder = prj.GetContents('Bibliothek.IntPrjfolder')[0]
    modfolder = bibfolder.GetContents('Benutzerdefinierte Modelle.IntPrjfolder')[0]

    usermods = []
    for mod in modfolder.GetContents():
        usermods.append(mod)

    return usermods

#-----------------------------------------------------------------------------------------------------------------------
# Speichert die kopierten benutzerdefinierten Modelle
#-----------------------------------------------------------------------------------------------------------------------
def importMods(app, ModList):

    prj = app.GetActiveProject()

    # Einlesen der Bibliothek-Ordnerstruktur
    bibfolder = prj.GetContents('Bibliothek.IntPrjfolder')[0]
    modfolder = bibfolder.GetContents('Benutzerdefinierte Modelle.IntPrjfolder')[0]

    for mod in ModList:

        modname = mod.GetAttribute('loc_name')

        # Überprüft, ob Modell mit gleichem Namen bereits vorhanden
        if modfolder.GetContents(modname) == []:
            modfolder.PasteCopy(mod)
            app.PrintInfo("Benutzerdefiniertes Modell " + modname + " wurde importiert")
        else:
            app.PrintInfo("Benutzerdefiniertes Modell " + modname + " ist bereits vorhanden und wurde nicht importiert")

#-----------------------------------------------------------------------------------------------------------------------
# Erstellung neuer Lastelemente
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmLod(name, target, ElmTerm, plini, qlini, TypLod = None, for_name=None, relay_model=None, relay_param=None):

    # Create element
    ElmLod = create_1end_element(name, 'ElmLod', target, ElmTerm, relay_model, relay_param)

    # Set active and reactive power
    ElmLod.SetAttribute('plini', plini)
    ElmLod.SetAttribute('qlini', qlini)

    # Check if load type is given
    if TypLod != None:
        ElmLod.SetAttribute('typ_id', TypLod)

    # Set foreign key, if available
    if for_name is not None:
        ElmLod.SetAttribute('for_name', for_name)

    # Space dedicated to the update of parameter
    ElmLod.SwitchOn()
    return ElmLod

#-----------------------------------------------------------------------------------------------------------------------
# Create new load model <ElmLod>
#-----------------------------------------------------------------------------------------------------------------------
def create_TypLod(name, LibFolder, params):
    TypLod = LibFolder.CreateObject('TypLod', name)

    for param, val in params.items():
        TypLod.SetAttribute(param, val)

    # TypLod.SetAttribute('loddy', 100)
    # TypLod.SetAttribute('udmax', 1.5)
    # TypLod.SetAttribute('udmin', 0.5)
    # TypLod.SetAttribute('i_nln', 0)
    # TypLod.SetAttribute('aP', params['Pp'])
    # TypLod.SetAttribute('bP', params['Ip'])
    # TypLod.SetAttribute('aQ', params['Pq'])
    # TypLod.SetAttribute('bQ', params['Iq'])
    # TypLod.SetAttribute('kpf', params['kpf'])
    # TypLod.SetAttribute('tpf', params['tpf'])
    # TypLod.SetAttribute('tpu', params['tpu'])
    # TypLod.SetAttribute('kqf', params['kqf'])
    # TypLod.SetAttribute('tqf', params['tqf'])
    # TypLod.SetAttribute('tqu', params['tqu'])
    # TypLod.SetAttribute('t1', params['t1'])

    return TypLod

#-----------------------------------------------------------------------------------------------------------------------
# Create new frequency relay <ElmFrl>
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmFrl(loc_name, ElmFolder, TypFrl, ElmTerm, ElmLod):

    # Create object
    ElmFrl = ElmFolder.CreateObject('ElmFrl', loc_name)

    # Set relay type
    ElmFrl.SetAttribute('typ_id', TypFrl)

    # Set measurement terminal
    ElmFrl.SetAttribute('plocbr', ElmTerm)

    # Set Load
    ElmFrl.SetAttribute('pload', ElmLod)

    # Return ElmFrl
    return ElmFrl

#-----------------------------------------------------------------------------------------------------------------------
# Create new frequency relay type <TypFrl>
#-----------------------------------------------------------------------------------------------------------------------
def create_TypFrl(loc_name, LibFolder, params):

    # Create object
    TypFrl = LibFolder.CreateObject('TypFrl', loc_name)

    # Set parameters
    for key, param in params.items():
        TypFrl.SetAttribute(key, param)

    return TypFrl

#-----------------------------------------------------------------------------------------------------------------------
# Create new relay model <ElmRelay>
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmRelay(name, target, relay_model, relay_para):
    '''This method creates a desired relay model for shunt elements (1 end elements). This method is developed only to
    model a frequency relay.

    Keyword arguments:
    name --> str name of the element to be created
    target --> str name of the class (example .ElmTerm)
    relay_model --> model of the relay to be added into the equipment library folder
    relay_para --> relay parameters

    Output:
    Returns a relay element object
'''

    # Create relay object
    ElmRelay = target.CreateObject('ElmRelay', name)
    ElmRelay.SetAttribute('typ_id', relay_model)
    ElmRelay_objects = ElmRelay.GetContents()
    for object_num in range(len(ElmRelay_objects)):
        if object_num == 0:
            continue
        elif object_num % 2 != 0:
            ElmRelay_object = ElmRelay.GetContents()[object_num]
            ElmRelay_object.SetAttribute('Fset', relay_para[object_num][0])
            ElmRelay_object.SetAttribute('Tdel', relay_para[object_num][1])
        else:
            ElmRelay_object = ElmRelay.GetContents()[object_num]
            ElmRelay_object.SetAttribute('shed', [relay_para[object_num][0]])
            ElmRelay_object.SetAttribute('Tbreak', [relay_para[object_num][1]])

    return ElmRelay

#-----------------------------------------------------------------------------------------------------------------------
# Erstellt ein Element mit einem Anschlusspunkt
#-----------------------------------------------------------------------------------------------------------------------
def create_1end_element(name, classname, target, terminal, relay_model=None, relay_para=None):
    """
    The function creates a 1end element (load, grid, machines), including the
    the graphical symbols and the graphical connections to it.

    Keyword arguments:
    name --> str name of the element to be created
    classname --> str name of the class (example .ElmTerm)
    Netzdaten_netz --> object associated to the grid where the element will be created
    Terminal --> object associated to the terminal where the element will be created

    Output:
    Object of the created element

    """
    # Creation of the element
    element = target.CreateObject(classname, name)

    if element is None: print("The element" + " " + name + " " + "could not be created")
    # else: print("The element"+" "+name+" "+"was created")

    # Creation of the associated field and swith in the terminal
    field = create_field_switch(terminal, element)

    # Create relay if model passed
    if relay_model != None:
        name = name + '_relaymodel'
        create_ElmRelay(name, field, relay_model, relay_para)  # Creates a relay model in the corresponding switch

    return element

#-----------------------------------------------------------------------------------------------------------------------
# Generic function to create element with no switch
#-----------------------------------------------------------------------------------------------------------------------
def create_Elm_NoStaCubic(target, classname, loc_name, **kwargs):

    # Create element
    Elm = target.CreateObject(classname, loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in kwargs.items():
        Elm.SetAttribute(key, item)

    return Elm

#-----------------------------------------------------------------------------------------------------------------------
# Erstellt ein Element mit zwei Anschlusspunkten
#-----------------------------------------------------------------------------------------------------------------------
def create_2end_element(name, classname, terminal1, terminal2, ElmNet):
    #check_repetition(name, ElmNet)

    element = ElmNet.CreateObject(classname, name)

    if element is None: print("The element" + " " + name + " " + "could not be created")

    create_field_switch(terminal1, element)
    create_field_switch(terminal2, element)

    return element

#-----------------------------------------------------------------------------------------------------------------------
# Überprüft Überschneidungen bei Namen von Elementen
#-----------------------------------------------------------------------------------------------------------------------
def check_repetition(name, Netzdaten_netz):

    if (len(Netzdaten_netz.GetContents(name)) >= 1):
        raise AttributeError("Be careful, the name " + " " + name + " " + "for the desired symbol/element in" +
                         " " + Netzdaten_netz.GetAttribute('loc_name') + " " + "network" + " " + "is already in use")


#-----------------------------------------------------------------------------------------------------------------------
# Erstellt ein Schaltfeld
#-----------------------------------------------------------------------------------------------------------------------
def create_field_switch(terminal, element, name=''):
    """
    When an element of the power system is connected to a terminal, a field
    (feld) has to be created in the terminal, and that field will be bond to the
    corresponding element. Additionaly, each field need an switch to operate the
    connection

    Keyword arguments:
    terminal --> object associated to the terminal where the element will be connected
    element --> object associated to the element to be connected

    Output:
    none

    """
    if name == '':
        name = 'Feld_1'

    Feld = terminal.CreateObject('StaCubic', name)
    Feld.SetAttribute('obj_id', element)
    Feld.CreateObject('StaSwitch', 'Schalter')

    return Feld

#-----------------------------------------------------------------------------------------------------------------------
# Erstellt ein Externes Netz 'ElmXnet'
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmXnet(name, ElmNet, ElmTerm):
    grid = create_1end_element(name, 'ElmXnet', ElmNet, ElmTerm)
    grid.SwitchOn()
    return grid

def create_ElmVac(name, target, ElmTerm):
    ElmVac = create_1end_element(name, 'ElmVac', target, ElmTerm)
    ElmVac.SwitchOn()
    return ElmVac

#-----------------------------------------------------------------------------------------------------------------------
# Funktion zur Erstellung eines Knoten
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmTerm(name, target, uknom, lat=None, lon=None, ElmZone=None, ElmZoneName=None, ElmAreaName=None, app=None, for_name=None):
    """
    Creates a terminal object (node)

    Keyword arguments:
    name --> str - name of the node to be created
    Netzdaten_netz --> object associated to the grid where the terminal will be created
    x & y --> position of the terminal
    List --> List to accumulate the created terminals

    Output:
    none

    """
    # Creation of the terminal in the corresponding grid
    ElmTerm = target.CreateObject('ElmTerm', name)
    ElmTerm.SetAttribute('uknom', uknom)

    # Set geo-coordinates
    if lat is not None and lon is not None and lat != 'NA' and lon != 'NA':
        ElmTerm.SetAttribute('GPSlat', lat)
        ElmTerm.SetAttribute('GPSlon', lon)
    elif lat is not None and lon is not None:
        print('No coordinates found for ElmSite ' + name)

    # Set foreign key if available
    if for_name is not None:
        ElmTerm.SetAttribute('for_name', for_name)

    # Set Zone
    if ElmZone is not None:
        ElmTerm.SetAttribute('pZone', ElmZone)

    # Set Zone by name
    elif ElmZoneName is not None:
        if app is not None:

            # Get folder:
            IntZone = app.GetDataFolder('ElmZone', 1)

            # Check, if ElmZone already exists
            if IntZone.GetContents(ElmZoneName + '.ElmZone') != []:
                ElmTerm.SetAttribute('pZone', IntZone.GetContents(ElmZoneName + '.ElmZone')[0])
            else:
                ElmZone = IntZone.CreateObject('ElmZone', ElmZoneName)
                ElmTerm.SetAttribute('pZone', ElmZone)
        else:
            print('The PowerFactory app was not passed. Could not set Zone ' + ElmZoneName + ' for busbar ' + name)


    # Set Area
    if ElmAreaName is not None:
        if app is not None:

            # Get folder:
            IntArea = app.GetDataFolder('ElmArea', 1)

            # Check, if ElmZone already exists
            if IntArea.GetContents(ElmZoneName + '.ElmArea') != []:
                ElmTerm.SetAttribute('pArea', IntArea.GetContents(ElmAreaName + '.ElmArea')[0])
            else:
                ElmArea = IntArea.CreateObject('ElmArea', ElmAreaName)
                ElmTerm.SetAttribute('pArea', ElmArea)
        else:
            print('The PowerFactory app was not passed. Could not set Area ' + ElmAreaName + ' for busbar ' + name)

    if ElmTerm is None: print("The Terminal" + " " + name + " " + "could not be created")
    # else: print("The Terminal"+" "+name+" "+"was created")

    ElmTerm.SwitchOn()
    return ElmTerm

#-----------------------------------------------------------------------------------------------------------------------
# Erstellt einen Standort
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmSite(name, target, lat=None, lon=None):

    ElmSite = target.CreateObject('ElmSite', name)

    if lat is not None and lon is not None and lat != 'NA' and lon != 'NA':
        ElmSite.SetAttribute('GPSlat', lat)
        ElmSite.SetAttribute('GPSlon', lon)
    else:
        print('No coordinates found for ElmTerm ' + name)


    if ElmSite is None: print("The Site" + " " + name + " " + "could not be created")
    # else: print("The Terminal"+" "+name+" "+"was created")

    ElmSite.SwitchOn()
    return ElmSite


#-----------------------------------------------------------------------------------------------------------------------
# Erstellt 2W-Trafo
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmTr2(name, target, terminal1, terminal2, TypTrafo):
    # Be careful with the order, we need a check here for the voltage of 1 and 2

    trafo = create_2end_element(name, 'ElmTr2', terminal1, terminal2, target)
    trafo.SetAttribute('typ_id', TypTrafo)
    trafo.SwitchOn()
    return trafo

# ----------------------------------------------------------------------------------------------------------------------
# Erstellt einen Impedanzzweig 'ElmZpu'
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmZpu(name, terminal1, terminal2, ElmNet):
    impedance = create_2end_element(name, 'ElmZpu', terminal1, terminal2, ElmNet)
    impedance.SwitchOn()
    return impedance

# ----------------------------------------------------------------------------------------------------------------------
# Erstellt eine Leitung ElmLne
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmLne(name, ElmTermA, ElmTermB, target, TypLne=None,length=1):
    ElmLne = create_2end_element(name, 'ElmLne', ElmTermA, ElmTermB, target)

    # Setzen der Leitungsdaten
    ElmLne.SetAttribute('dline', length)
    ElmLne.SetAttribute('typ_id', TypLne)
    ElmLne.SwitchOn()

    return ElmLne

# ----------------------------------------------------------------------------------------------------------------------
# Erstellt eine DEA 'ElmGenstat' mit einem hinterlegtem Modell
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmGenstat(name, ElmNet, terminal, sgn, cosn, pgini, qgini=0, modelfolder=None, default=True, for_name=None):

    # Erzeugung der DEA-Anlage
    ElmGenstat = create_1end_element(name, 'ElmGenstat', ElmNet, terminal)

    if ElmGenstat == None:
        print('stop')

    # Set nominal parameters
    ElmGenstat.SetAttribute('sgn', sgn)
    ElmGenstat.SetAttribute('cosn', cosn)

    # Set load flow parameters
    ElmGenstat.SetAttribute('pgini', pgini)
    ElmGenstat.SetAttribute('qgini', qgini)
    
    ElmGenstat.SetAttribute('iShcModel',1)
    ElmGenstat.SetAttribute('Skss',sgn)
    ElmGenstat.SetAttribute('Kfactor',3)
    ElmGenstat.SetAttribute('imax',1.1)
    

    # Set dynamic model for ElmGenstat
    if default == False:
        # Erzeugung des zugehörigen Modell-Elements
        DEA_comp = ElmNet.CreateObject('ElmComp', name + '_comp')
        DEA_comp.SetAttribute('typ_id', modelfolder.GetContents('DEA_Control_idq.BlkDef')[0])

        DEA_contr = DEA_comp.CreateObject('ElmDsl', name + '_contr')
        DEA_contr.SetAttribute('typ_id', modelfolder.GetContents('VAR-Idq.BlkDef')[0])

        DEA_comp.SetAttribute('pelm', [ElmGenstat, DEA_contr, terminal])

    else:
        DEA_comp = None

    # Set foreign key
    if for_name != None:
        ElmGenstat.SetAttribute('for_name', for_name)

    # Element einschalten
    ElmGenstat.SwitchOn()

    return ElmGenstat

# ----------------------------------------------------------------------------------------------------------------------
# Create a composite model
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmComp(loc_name, target, typ_id, pelm_dict={}):

    # Create element
    ElmComp = target.CreateObject('ElmComp', loc_name)

    # Set type from library
    ElmComp.SetAttribute('typ_id', typ_id)

    # Sets all blocks which are passed in the dictionary
    for key, item in pelm_dict.items():
        ElmComp.SetAttribute(key, item)

    return ElmComp

# ----------------------------------------------------------------------------------------------------------------------
# Create Coupler
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmCoup(name, target, ElmTermA, ElmTermB):

    # Create coupler object
    ElmCoup = create_2end_element(name, 'ElmCoup', ElmTermA, ElmTermB, target)

    # Close coupler
    ElmCoup.SetAttribute('on_off', 1)

# ----------------------------------------------------------------------------------------------------------------------
# Create a ElmDsl for a ElmComp
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmDsl(loc_name, target, typ_id, params_dict={}):

    # Create element
    ElmDsl = target.CreateObject('ElmDsl', loc_name)

    # Set type from library
    ElmDsl.SetAttribute('typ_id', typ_id)

    # Sets all blocks which are passed in the dictionary
    for key, item in params_dict.items():
        ElmDsl.SetAttribute(key, item)

    return ElmDsl

# ----------------------------------------------------------------------------------------------------------------------
# Create a ElmPLL 
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmPLL(loc_name, target, params_dict={}):

    # Create element
    ElmDsl = target.CreateObject('ElmPhi__pll', loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in params_dict.items():
        ElmDsl.SetAttribute(key, item)

    return ElmDsl

# ----------------------------------------------------------------------------------------------------------------------
# Create a Measurements
# ----------------------------------------------------------------------------------------------------------------------

def create_PQMeas(loc_name, target, params_dict={}):

    # Create element
    ElmDsl = target.CreateObject('StaPqMea', loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in params_dict.items():
        ElmDsl.SetAttribute(key, item)

    return ElmDsl

def create_VMeas(loc_name, target, params_dict={}):

    # Create element
    ElmDsl = target.CreateObject('StaVMea', loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in params_dict.items():
        ElmDsl.SetAttribute(key, item)

    return ElmDsl

def create_IMeas(loc_name, target, params_dict={}):

    # Create element
    ElmDsl = target.CreateObject('StaIMea', loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in params_dict.items():
        ElmDsl.SetAttribute(key, item)

    return ElmDsl


# ----------------------------------------------------------------------------------------------------------------------
# Create a EvtOutage for a protection scheme
# ----------------------------------------------------------------------------------------------------------------------
def create_EvtOutage(loc_name, ElmFolder, i_what, p_target):

    # Create element
    EvtOutage = ElmFolder.CreateObject('EvtOutage', loc_name)

    # Set event type
    EvtOutage.SetAttribute('i_what', i_what)

    # Set reference for element
    EvtOutage.SetAttribute('p_target', p_target)

    # Return object
    return EvtOutage

# ----------------------------------------------------------------------------------------------------------------------
# Create a EvtSwitch for a switching event
# ----------------------------------------------------------------------------------------------------------------------
def create_EvtSwitch(loc_name, ElmFolder, time, p_target, i_switch, i_allph):

    # Create element
    EvtSwitch = ElmFolder.CreateObject('EvtSwitch', loc_name)

    # Set time in seconds
    EvtSwitch.SetAttribute('time', time)

    # Set if switch is opened or closed
    EvtSwitch.SetAttribute('i_switch', i_switch)

    # Set three phase selection
    EvtSwitch.SetAttribute('i_allph', i_allph)

    # Set reference for element
    EvtSwitch.SetAttribute('p_target', p_target)

    # Return object
    return EvtSwitch

'''
------------------------------------------------------------------------------------------------------------------------
Create Synchonous Machine (ElmSym)
------------------------------------------------------------------------------------------------------------------------
Input parameters:
    - name: Unique name (loc_name)
    - folder: Location, where the element is created in PowerFactory
    - ElmTerm: Terminal, where the element is connected to the grid
    - TypSym: Type of the element
    - pgini: Active power for load flow calculations
    - ip_ctrl: Binary flag, sets machine as reference machine
    - av_mode: Stationary reactive power control
    - qgini: Reactive power for stationary behaviour
    - usetp: Voltage set point, if av_mode = constv
    - phiini: Voltage angle set point, if av_mode = constv and ip_ctrl = 1
'''
def create_ElmSym(name, folder, ElmTerm, TypSym, pgini, ip_ctrl=0, 
                  av_mode = 'constq', qgini=0, usetp=1, phiini=0, for_name=None):

    # Create Object ElmSym
    ElmSym = create_1end_element(name, 'ElmSym', folder, ElmTerm)

    # Set type TypSym
    ElmSym.SetAttribute('typ_id', TypSym)

    # Set active power of load flow
    ElmSym.SetAttribute('pgini', pgini)

    # Set reference machine flag
    ElmSym.SetAttribute('ip_ctrl', ip_ctrl)

    # Set stationary reactive power behaviour
    ElmSym.SetAttribute('av_mode', av_mode)

    # Constant reactive power control
    if av_mode == 'constq':
        ElmSym.SetAttribute('qgini', qgini)

    # Constant voltage control
    elif av_mode == 'constv':
        ElmSym.SetAttribute('qgini', qgini)
        ElmSym.SetAttribute('usetp', usetp)
        if ip_ctrl == 1: ElmSym.SetAttribute('phiini', phiini)

    # Set foreign key
    if for_name != None:
        ElmSym.SetAttribute('for_name', for_name)

    ElmSym.SwitchOn()

    return ElmSym

'''
------------------------------------------------------------------------------------------------------------------------
Create Type for Synchonous Machine (TypSym)
------------------------------------------------------------------------------------------------------------------------
Default parameters taken from entsoe controller test report
Input parameters:
    - name: Unique name (loc_name)
    - folder: Location, where the element is created in PowerFactory
    - sgn: Nominal apparent power in MVA
    - ugn: Nominal voltage in kV
    - cosn: Nominal power factor
    - iner_inp: Input mode for inertia constant (currently possible: HSC)
    - h: Inertia constant referenced to sgn in s
    - rstr: armature resistance in pu
    - xl: armature leakage reactance in pu
    - xd: synchronous reactance in d-axis in pu
    - xq: synchronous reactance in q-axis in pu
    - tds: transient short circuit time constant in d-axis in s
    - tqs: transient short circuit time constant in q-axis in s
    - xds: transient reactance in d-axis in pu
    - xqs: transient reactance in q-axis in pu
    - tdss: subtransient short circuit time constant in d-axis in s
    - tqss: subtransient short circuit time constant in q-axis in s
    - xdss: subtransient reactance in d-axis in pu
    - xqss: subtransient reactance in q-axis in pu
'''
def create_TypSym(name, folder, sgn, ugn, cosn, iner_inp='HSC', h=4, rstr=0, xl=0.15, xd=2, xq=1.8,
                  tds=0.9, tqs=0.6, xds=0.35, xqs=0.5,
                  tdss=0.03, tqss=0.05, xdss=0.25, xqss=0.3):

    # Create Object ElmSym
    TypSym = folder.CreateObject('TypSym', name)

    # Set base data
    TypSym.SetAttribute('sgn', sgn)
    TypSym.SetAttribute('ugn', ugn)
    TypSym.SetAttribute('cosn', cosn)

    # RMS parameters
    # Inertia parameters
    TypSym.SetAttribute('iner_inp', iner_inp)
    TypSym.SetAttribute('h', h)

    # Armature parameters
    TypSym.SetAttribute('rstr', rstr)
    TypSym.SetAttribute('xl', xl)

    # Synchronous reactance
    TypSym.SetAttribute('xd', xd)
    TypSym.SetAttribute('xq', xq)

    # Transient time constants
    TypSym.SetAttribute('tds', tds)
    TypSym.SetAttribute('tqs', tqs)

    # Transient reactance
    TypSym.SetAttribute('xds', xds)
    TypSym.SetAttribute('xqs', xqs)

    # Subtransient time constants
    TypSym.SetAttribute('tdss', tdss)
    TypSym.SetAttribute('tqss', tqss)

    # Subtransient reactance
    TypSym.SetAttribute('xdss', xdss)
    TypSym.SetAttribute('xqss', xqss)

    return TypSym

# ----------------------------------------------------------------------------------------------------------------------
# Erstellt einen Typen für einen 2-Wicklungstrafo
# ----------------------------------------------------------------------------------------------------------------------
def create_TypTr2(name, ElmNet, app, strn, utrn_h, utrn_l, iopt_uk='pcu', iopt_uk0='r0x0', pcutr=0, uktr=0,
                  uktrr=0, xtor=0, r1pu=0, x1pu=0, pfe=0, curmg=0, nt2ag=0, tr2cn_h='YN' ,tr2cn_l='YN'):

    TrafoTyp = ElmNet.CreateObject('TypTr2', name)

    TrafoTyp.SetAttribute('strn', strn)
    TrafoTyp.SetAttribute('utrn_h', utrn_h)
    TrafoTyp.SetAttribute('utrn_l', utrn_l)

    # Get Settings Folder
    SetFolder = app.GetActiveProject().GetContents('Einstellungen.SetFold')[0]
    if SetFolder.GetContents('*.IntOpt') == []:
        IntOpt = SetFolder.CreateObject("IntOpt", 'Eingabeoptionen')
    else:
        IntOpt = SetFolder.GetContents('*.IntOpt')[0]

    if IntOpt.GetContents('*.OptTyptr2') == []:
        OptTyptr2 = IntOpt.CreateObject("OptTyptr2", '2-Wicklungstransformator')
    else:
        OptTyptr2 = IntOpt.GetContents('*.OptTyptr2')[0]

    OptTyptr2.SetAttribute('iopt_uk', iopt_uk)
    OptTyptr2.SetAttribute('iopt_uk0', iopt_uk0)

    if iopt_uk == 'pcu':
        TrafoTyp.SetAttribute('pcutr', pcutr)
        TrafoTyp.SetAttribute('uktr', uktr)
    elif iopt_uk == 'ukr':
        TrafoTyp.SetAttribute('uktr', uktr)
        TrafoTyp.SetAttribute('uktrr', uktrr)
    elif iopt_uk == 'xtr':
        TrafoTyp.SetAttribute('uktr', uktr)
        TrafoTyp.SetAttribute('xtor', xtor)
    elif iopt_uk == 'rx':
        TrafoTyp.SetAttribute('r1pu', r1pu)
        TrafoTyp.SetAttribute('x1pu', x1pu)
    else:
        raise AttributeError('Invalid attribute <' + iopt_uk + '> for iopt_uk while creating a TypTr2')

    TrafoTyp.SetAttribute('pfe', pfe)
    TrafoTyp.SetAttribute('curmg', curmg)
    TrafoTyp.SetAttribute('tr2cn_h', tr2cn_h)
    TrafoTyp.SetAttribute('tr2cn_h', tr2cn_l)
    TrafoTyp.SetAttribute('nt2ag', nt2ag)


    return TrafoTyp

# ----------------------------------------------------------------------------------------------------------------------
# Erstellt einen Leitungstyp TypLne
# ----------------------------------------------------------------------------------------------------------------------
def create_TypLne(name, target,  U_n, rline=1, xline=1, cline=None, bline=None, sline=1, InomAir=1, cohl_=1, gline=0):

    TypLne = target.CreateObject('TypLne', name)
    TypLne.SetAttribute('uline', U_n)
    TypLne.SetAttribute('rline', rline)
    TypLne.SetAttribute('xline', xline)
    TypLne.SetAttribute('cohl_', cohl_)
    TypLne.SetAttribute('sline', sline)
    TypLne.SetAttribute('InomAir', InomAir)

    if cline == None and bline == None:
        TypLne.SetAttribute('cline', 0)
    elif cline != None and bline == None:
        TypLne.SetAttribute('cline', cline)
    elif cline == None and bline != None:
        TypLne.SetAttribute('bline', bline)
    elif cline != None and bline != None:
        print('Two values for C and B were given for TypLne ' + name + '. Using C')
        TypLne.SetAttribute('cline', cline)


    return TypLne

# ----------------------------------------------------------------------------------------------------------------------
# Create a Zone element (ElmZone)
# ----------------------------------------------------------------------------------------------------------------------
def create_ElmZone(name, target, icolor=1, for_name=None):

    ElmZone = target.CreateObject('ElmZone', name)
    ElmZone.SetAttribute('icolor', icolor)

    if for_name != None:
        ElmZone.SetAttribute('for_name', for_name)

#-----------------------------------------------------------------------------------------------------------------------
# Create a compensation element ElmShnt (inductor or capacitor)
#-----------------------------------------------------------------------------------------------------------------------
def create_ElmShnt(name, target, terminal, ushnm, shtype, qrean=0, qcapn=0, lat=None, lon=None, ElmZoneName=None, ElmAreaName=None, app=None, for_name=None):

    # Check if type is in allowed values ToDo Test
    valid_shtype = {0, 1, 2, 3, 4}
    if shtype not in valid_shtype:
        raise ValueError('Error: Shunt element ' + name + ' type must be one of %r.' % valid_shtype)

    # Create element
    ElmShnt = create_1end_element(name, 'ElmShnt', target, terminal)

    # Set nominal voltage
    ElmShnt.SetAttribute('ushnm', ushnm)

    # Set shunt type
    ElmShnt.SetAttribute('shtype', shtype)
    # Capacitor (C)
    if shtype == 2:
        ElmShnt.SetAttribute('qcapn', qcapn)
    # Inductor (R,L)
    elif shtype == 1:
        ElmShnt.SetAttribute('qrean', qrean)

    # Set foreign identification key
    if for_name != None:
        ElmShnt.SetAttribute('for_name', for_name)

    # Switch element on
    ElmShnt.SwitchOn()

    return ElmShnt

# ----------------------------------------------------------------------------------------------------------------------
# Create a calculated result variable (IntCalcres)
# ----------------------------------------------------------------------------------------------------------------------
def create_IntCalcres(loc_name, ElmRes, resdesc, resunit, ElmList, VarName):

    # Define formula
    formula = ""
    i = 1
    ElmListInServ = []

    for Elm in ElmList:

        # Check if element is in service
        if Elm.GetAttribute('outserv') == 0:

            # Add variable for object
            ElmRes.AddVariable(Elm, VarName)

            # Add to list of elements in service
            ElmListInServ.append(Elm)

            # Add to formula
            formula += 'in' + str(i) + '+'

            i+=1

    # Create IntCalcres
    IntCalcres = ElmRes.CreateObject('IntCalcres', 'Calc_' + loc_name)

    # Set attributes for IntCalcres
    IntCalcres.SetAttribute('resname', loc_name)
    IntCalcres.SetAttribute('resdesc', resdesc)
    IntCalcres.SetAttribute('resunit', resunit)
    IntCalcres.SetAttribute('element', ElmListInServ)
    IntCalcres.SetAttribute('inpvar', len(ElmListInServ) * [VarName])

    # Delete last "+" sign
    formula = formula[:-1]

    # Set formula
    IntCalcres.SetAttribute('formula', [formula])

