Attribute VB_Name = "Formules_Electrotechniek"
Function I_UR(Spanning, Weerstand)
Attribute I_UR.VB_Description = "Berekend Stroom met spanning en weerstand"
Attribute I_UR.VB_ProcData.VB_Invoke_Func = " \n22"
    I_UR = Spanning / Weerstand
End Function

Function U_IR(Stroom, Weerstand)
Attribute U_IR.VB_Description = "Berekend spanning met weerstand en stroom"
Attribute U_IR.VB_ProcData.VB_Invoke_Func = " \n22"
    U_IR = Stroom * Weerstand
End Function

Function R_UI(Spanning, Stroom)
Attribute R_UI.VB_Description = "Berekend weerstand met spanning en stroom"
Attribute R_UI.VB_ProcData.VB_Invoke_Func = " \n22"
    R_UI = Spanning / Stroom
End Function

Function P_UI(Spanning, Stroom, Optional Arbeidsfactor = 0)
Attribute P_UI.VB_Description = "Berekend vermogen met spanning en stroom"
Attribute P_UI.VB_ProcData.VB_Invoke_Func = " \n22"
    P_UI = (((Spanning * Stroom * Cos(Arbeidsfactor)) ^ 2) + ((Spanning * Stroom * Sin(Arbeidsfactor)) ^ 2)) ^ 0.5
End Function

Function Pr_UI(Spanning, Stroom, Optional Arbeidsfactor = 0)
Attribute Pr_UI.VB_Description = "Berekend blindvermogen met spanning, stroom en cos phi"
Attribute Pr_UI.VB_ProcData.VB_Invoke_Func = " \n22"
    Pr_UI = Spanning * Stroom * Sin(Arbeidsfactor)
End Function

Function Pa_UI(Spanning, Stroom, Optional Arbeidsfactor = 0)
Attribute Pa_UI.VB_Description = "Berekend actiefvermogen met spanning, stroom en cos phi"
Attribute Pa_UI.VB_ProcData.VB_Invoke_Func = " \n22"
    Pa_UI = Spanning * Stroom * Cos(Arbeidsfactor)
End Function

Function Rv_Ser(R1, R2, Optional R3 = 0, Optional R4 = 0, Optional R5 = 0, Optional R6 = 0, Optional R7 = 0, Optional R8 = 0, Optional R9 = 0, Optional R10 = 0)
Attribute Rv_Ser.VB_Description = "Berekend vervangingsweerstand van serie geschakelde weerstanden"
Attribute Rv_Ser.VB_ProcData.VB_Invoke_Func = " \n22"
Rv = R1 + R2

If R3 <> 0 Then
 Rv = Rv + R3
End If

If R4 <> 0 Then
 Rv = Rv + R4
End If
If R5 <> 0 Then
 Rv = Rv + R5
End If
If R6 <> 0 Then
 Rv = Rv + R6
End If
If R7 <> 0 Then
 Rv = Rv + R7
End If
If R8 <> 0 Then
 Rv = Rv + R8
End If
If R9 <> 0 Then
 Rv = Rv + R9
End If
If R10 <> 0 Then
 Rv = Rv + R10
End If

Rv_Ser = Rv


End Function

Function Rv_Par(R1, R2, Optional R3 = 0, Optional R4 = 0, Optional R5 = 0, Optional R6 = 0, Optional R7 = 0, Optional R8 = 0, Optional R9 = 0, Optional R10 = 0)
Attribute Rv_Par.VB_Description = "Berekend vervangingsweerstand van parallel geschakelde weerstanden"
Attribute Rv_Par.VB_ProcData.VB_Invoke_Func = " \n22"
Rv = (1 / R1) + (1 / R2)

If R3 <> 0 Then
 Rv = Rv + (1 / R3)
End If

If R4 <> 0 Then
 Rv = Rv + 1 / R4
End If
If R5 <> 0 Then
 Rv = Rv + 1 / R5
End If
If R6 <> 0 Then
 Rv = Rv + 1 / R6
End If
If R7 <> 0 Then
 Rv = Rv + 1 / R7
End If
If R8 <> 0 Then
 Rv = Rv + 1 / R8
End If
If R9 <> 0 Then
 Rv = Rv + 1 / R9
End If
If R10 <> 0 Then
 Rv = Rv + 1 / R10
End If

Rv_Par = 1 / Rv


End Function
