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