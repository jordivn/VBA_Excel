Attribute VB_Name = "Formules_Inhoud"



Function TankInhoudRondLiggend(StraalTank, LengteTank, Optional PeilglasStand = 1)
Attribute TankInhoudRondLiggend.VB_ProcData.VB_Invoke_Func = " \n30"
If PeilglasStand < 1 Then
Hoogte = StraalTank * 2 * PeilglasStand
If Hoogte = 0 Then
TankInhoudRondLiggend = 0
Else
TankInhoudRondLiggend = ((Rad2Deg(Arccos((StraalTank - Hoogte) / StraalTank)) / 360) * 2 * ((StraalTank ^ 2) * JTOOLS_Pi) - (Sqr(StraalTank ^ 2 - (StraalTank - Hoogte) ^ 2) * (StraalTank - Hoogte))) * LengteTank
End If
Else
TankInhoudRondLiggend = StraalTank ^ 2 * Pi() * LengteTank
End If
End Function

Function TankInhoudRondStaand(StraalTank, Optional PeilglasStand = 1)
Attribute TankInhoudRondStaand.VB_ProcData.VB_Invoke_Func = " \n30"
Hoogte = LengteTank * PeilglasStand
TankInhoudRondStaand = (StraalTank ^ 2 * JTOOLS_Pi) * Hoogte
End Function

Function StraalUitbollingTank(StraalTank, UitbollingDiepte)
Attribute StraalUitbollingTank.VB_ProcData.VB_Invoke_Func = " \n30"
HoogteTank = StraalTank * 2
If HoogteTank <> 0 And UitbollingDiepte <> 0 Then
StraalUitbollingTank = (Tan(Deg2Rad((Rad2Deg(Atn(((HoogteTank / 2) / UitbollingDiepte)))) - (Rad2Deg(Atn((UitbollingDiepte / (HoogteTank / 2))))))) * (HoogteTank / 2)) + UitbollingDiepte
Else
StraalUitbollingTank = 0
End If
End Function

Function InhoudUitbollingTank(StraalTank, UitbollingDiepte, Optional PeilglasStand = 1)
Attribute InhoudUitbollingTank.VB_ProcData.VB_Invoke_Func = " \n30"
If PeilglasStand < 1 Then
Hoogte = StraalTank * 2 * PeilglasStand
TotaleBolInhoud = InhoudUitbollingTank(StraalTank, UitbollingDiepte)
StraalBol = StraalUitbollingTank(StraalTank, UitbollingDiepte)
AangepasteStraal = Sqr(StraalBol ^ 2 - (StraalTank - Hoogte) ^ 2)
AangepasteBollingDiepte = (StraalBol - AangepasteStraal)
CilindrischeInhoud = ((StraalTank - Hoogte) ^ 2 * Pi()) * (UitbollingDiepte - AangepasteBollingDiepte)
NieuweBolInhoud = InhoudUitbollingTank((StraalTank - Hoogte), AangepasteBollingDiepte)
If (StraalTank - Hoogte) > 0 Then
InhoudUitbollingTank = (TotaleBolInhoud - CilindrischeInhoud - NieuweBolInhoud) / 2

Else
InhoudUitbollingTank = CilindrischeInhoud + NieuweBolInhoud + (TotaleBolInhoud - CilindrischeInhoud - NieuweBolInhoud) / 2

End If
Else
If StraalTank <> 0 And UitbollingDiepte <> 0 Then
HoogteTank = StraalTank * 2
InhoudUitbollingTank = ((Pi() * UitbollingDiepte) / 6) * (3 * (Sqr((StraalUitbollingTank(StraalTank, UitbollingDiepte)) ^ 2 - ((StraalUitbollingTank(StraalTank, UitbollingDiepte)) - UitbollingDiepte) ^ 2)) ^ 2 + UitbollingDiepte ^ 2)
Else
InhoudUitbollingTank = 0
End If
End If
End Function


Function InhoudUitbollingTankX(StraalTank, UitbollingDiepte, Optional PeilglasStand = 1)
Attribute InhoudUitbollingTankX.VB_ProcData.VB_Invoke_Func = " \n30"

Hoogte = StraalTank * 2 * PeilglasStand
TotaleBolInhoud = InhoudUitbollingTank(StraalTank, UitbollingDiepte)
StraalBol = StraalUitbollingTank(StraalTank, UitbollingDiepte)
AangepasteStraal = Sqr(StraalBol ^ 2 - (StraalTank - Hoogte) ^ 2)
AangepasteBollingDiepte = (StraalBol - AangepasteStraal)
InhoudUitbollingTankX = ((StraalTank - Hoogte) ^ 2 * Pi()) * (UitbollingDiepte - AangepasteBollingDiepte)

End Function

Function InhoudUitbollingTankY(StraalTank, UitbollingDiepte, Optional PeilglasStand = 1)
Attribute InhoudUitbollingTankY.VB_ProcData.VB_Invoke_Func = " \n30"

Hoogte = StraalTank * 2 * PeilglasStand
TotaleBolInhoud = InhoudUitbollingTank(StraalTank, UitbollingDiepte)
StraalBol = StraalUitbollingTank(StraalTank, UitbollingDiepte)
AangepasteStraal = Sqr(StraalBol ^ 2 - (StraalTank - Hoogte) ^ 2)
AangepasteBollingDiepte = (StraalBol - AangepasteStraal)
CilindrischeInhoud = ((StraalTank - Hoogte) ^ 2 * Pi()) * (UitbollingDiepte - AangepasteBollingDiepte)
InhoudUitbollingTankY = InhoudUitbollingTank((StraalTank - Hoogte), AangepasteBollingDiepte)

End Function
