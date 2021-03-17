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


def ActivateVariation(SelVariation, app):

    Scheme = app.GetProjectFolder('scheme')
    Variations = Scheme.GetContents('*.IntScheme')

    ActVar = app.GetActiveNetworkVariations()

    # make sure there is not active variation

    if not ActVar:

        print("No variation activated. Proceeding to activate selected variation")

    else:

        ActVar[0].Deactivate()

    for varis in Variations:

        VarName = varis.GetFullName().split('\\')[-1][:-10]

        if VarName == SelVariation:

            varis.Activate()


    return varis

def CreateSimpleStabilityStudy(app, CreateScens):

    MyCases = app.GetProjectFolder('study')
    MyScenarios = app.GetProjectFolder('scen')
    AktCase = app.GetActiveStudyCase()
    CaseNames = ['Frequency Ramp', 'Voltage Step', 'Voltage Ramp']

    if len(MyCases.GetContents()) == 1:

        if CaseNames[0] not in AktCase.GetFullName():
            AktCase.SetAttribute('loc_name', CaseNames[0])
            print('Changing name of active study case')
        print('Procceeding to create the rest of cases')

        for n in range(3):
            MyCases.AddCopy(AktCase, CaseNames[n + 1])

    if CreateScens == 1:

        if not MyScenarios.GetContents():

            print('Scenarios object empty. Creating SM and VS scenarios')
            MyScenarios.CreateObject('IntScenario', 'SM')
            MyScenarios.CreateObject('IntScenario', 'VS')

        else:

            s = 0
            t = 0
            for scen in MyScenarios.GetContents():

                if 'SM' in scen.GetFullName():

                    s += 1

                    if s == 0:
                        print("Creating SM scenario")
                        MyScenarios.CreateObject('IntScenario', 'SM')

                elif 'VS' in scen.GetFullName():

                    t += 1

                    if t == 0:

                        print("Creating VS scenario")
                        MyScenarios.CreateObject('IntScenario', 'VS')
    elif CreateScens == 0:

        print("No scenarios created")

    else:

        raise ValueError("Flag value not valid. Choose 1 or 0")

def SetAttributesforFaultEvent(app, faultname, **kwargs):

    faultcase = app.GetFromStudyCase('IntEvt')
    event = faultcase.GetContents(faultname)[0]

    for key, var in kwargs.items():
        event.SetAttribute(key, var)

