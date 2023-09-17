Attribute VB_Name = "Formules_Geometrie"
Public Const JTOOLS_Pi As Double = 3.141592654

Function sec(X)
Attribute sec.VB_Description = "Secans"
Attribute sec.VB_ProcData.VB_Invoke_Func = " \n29"
    sec = 1 / Cos(X)
End Function

Function Cosec(X)
Attribute Cosec.VB_Description = "Cosecans"
Attribute Cosec.VB_ProcData.VB_Invoke_Func = " \n29"
    Cosec = 1 / Sin(X)
End Function

Function Cotan(X)
Attribute Cotan.VB_Description = "Cotangens"
Attribute Cotan.VB_ProcData.VB_Invoke_Func = " \n29"
    Cotan = 1 / Tan(X)
End Function

Function Arcsin(X)
Attribute Arcsin.VB_ProcData.VB_Invoke_Func = " \n29"
    Arcsin = Atn(X / Sqr(-X * X + 1))
End Function

Function Arccos(X)
Attribute Arccos.VB_ProcData.VB_Invoke_Func = " \n29"
    Arccos = Atn(-X / Sqr(-X * X + 1)) + 2 * Atn(1)
End Function

Function Arcsec(X)
Attribute Arcsec.VB_ProcData.VB_Invoke_Func = " \n29"
    Arcsec = Atn(X / Sqr(X * X - 1)) + Sgn((X) - 1) * (2 * Atn(1))
End Function

Function Arccosec(X)
Attribute Arccosec.VB_ProcData.VB_Invoke_Func = " \n29"
    Arccosec = Atn(X / Sqr(X * X - 1)) + (Sgn(X) - 1) * (2 * Atn(1))
End Function

Function Arccotan(X)
Attribute Arccotan.VB_ProcData.VB_Invoke_Func = " \n29"
    Arccotan = Atn(X) + 2 * Atn(1)
End Function

Function HSin(X)
Attribute HSin.VB_ProcData.VB_Invoke_Func = " \n29"
    HSin = (Exp(X) - Exp(-X)) / 2
End Function

Function HCos(X)
Attribute HCos.VB_ProcData.VB_Invoke_Func = " \n29"
    HCos = (Exp(X) + Exp(-X)) / 2
End Function

Function HTan(X)
Attribute HTan.VB_ProcData.VB_Invoke_Func = " \n29"
    HTan = (Exp(X) - Exp(-X)) / (Exp(X) + Exp(-X))
End Function

Function HSec(X)
Attribute HSec.VB_ProcData.VB_Invoke_Func = " \n29"
    HSec = 2 / (Exp(X) + Exp(-X))
End Function

Function HCosec(X)
Attribute HCosec.VB_ProcData.VB_Invoke_Func = " \n29"
    HCosec = 2 / (Exp(X) - Exp(-X))
End Function

Function HCotan(X)
    HCotan = (Exp(X) + Exp(-X)) / (Exp(X) - Exp(-X))
End Function

Function HArcsin(X)
Attribute HArcsin.VB_ProcData.VB_Invoke_Func = " \n29"
    HArcsin = Log(X + Sqr(X * X + 1))
End Function

Function HArccos(X)
Attribute HArccos.VB_ProcData.VB_Invoke_Func = " \n29"
    HArccos = Log(X + Sqr(X * X - 1))
End Function

Function HArctan(X)
Attribute HArctan.VB_ProcData.VB_Invoke_Func = " \n29"
    HArctan = Log((1 + X) / (1 - X)) / 2
End Function

Function HArcsec(X)
    HArcsec = Log((Sqr(-X * X + 1) + 1) / X)
End Function

Function HArccosec(X)
Attribute HArccosec.VB_ProcData.VB_Invoke_Func = " \n29"
    HArccosec = Log((Sgn(X) * Sqr(X * X + 1) + 1) / X)
End Function

Function HArccotan(X)
Attribute HArccotan.VB_ProcData.VB_Invoke_Func = " \n29"
    HArccotan = Log((X + 1) / (X - 1)) / 2
End Function

Function LogN(X, n)
Attribute LogN.VB_ProcData.VB_Invoke_Func = " \n29"
    LogN = Log(X) / Log(n)
End Function

Function Pi()
    Pi = JTOOLS_Pi
  
End Function

Function Rad2Deg(X)
Attribute Rad2Deg.VB_ProcData.VB_Invoke_Func = " \n29"
    Rad2Deg = X * (360 / (2 * JTOOLS_Pi))
End Function

Function Deg2Rad(X)
Attribute Deg2Rad.VB_ProcData.VB_Invoke_Func = " \n29"
    Deg2Rad = X / (360 / (2 * JTOOLS_Pi))
End Function
