# -----------------------------------------------------------------------------------------------------------------------
# Generic function to create node element with single switch
# -----------------------------------------------------------------------------------------------------------------------
def create_Elm_NoStaCubic(target, classname, loc_name, **kwargs):

    # Create element
    Elm = target.CreateObject(classname, loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in kwargs.items():
        Elm.SetAttribute(key, item)

    return Elm

# -----------------------------------------------------------------------------------------------------------------------
# Generic function to create element with no switch
# -----------------------------------------------------------------------------------------------------------------------
def create_Elm_StaCubic(target, classname, loc_name, ElmTermA=None, ElmTermB=None, **kwargs):

    # Create element
    Elm = target.CreateObject(classname, loc_name)

    # Sets all blocks which are passed in the dictionary
    for key, item in kwargs.items():
        Elm.SetAttribute(key, item)

    # Create connections to busbars ElmTerm
    if ElmTermA is not None: create_StaCubic(ElmTermA, Elm, f'Cub_{loc_name}')
    if ElmTermB is not None: create_StaCubic(ElmTermB, Elm, f'Cub_{loc_name}')

    # Activate all switches
    Elm.SwitchOn()

    return Elm

#-----------------------------------------------------------------------------------------------------------------------
# Create a switch field
#-----------------------------------------------------------------------------------------------------------------------
def create_StaCubic(ElmTerm, Elm, loc_name='', StaSwitch=True):

    # Create field at target node
    StaCubic = ElmTerm.CreateObject('StaCubic', loc_name)

    # Connect element
    StaCubic.SetAttribute('obj_id', Elm)

    # Create a switch if required
    if StaSwitch: StaCubic.CreateObject('StaSwitch')

    return StaCubic