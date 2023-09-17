Attribute VB_Name = "Formules_Stromingsleer"
Function Re(rho, v, Di, n)
Attribute Re.VB_Description = "Reynolds getal"
Attribute Re.VB_ProcData.VB_Invoke_Func = " \n32"

Re = (rho * v * Di) / n
' < 2300 Laminair
' > 3500 Turbulent
' daartussen overgangsgebied
End Function

Function Dpwl(rho, v, l, Di, la)
Attribute Dpwl.VB_Description = "Leiding druk verlies"
Attribute Dpwl.VB_ProcData.VB_Invoke_Func = " \n32"
Dpwl = 0.5 * rho * (v ^ 2) * (l / Di) * la
'Dpwl is in Pa
End Function

Function la_laminair(Re)
Attribute la_laminair.VB_Description = "Labda laminar"
Attribute la_laminair.VB_ProcData.VB_Invoke_Func = " \n32"
If Re < 2300 Then
la = 64 / Re
Else
la = "n/b"
End If
End Function

Function la_swanee_jain(e, Di, Re)
Attribute la_swanee_jain.VB_Description = "Labda volgens swanee jain"
Attribute la_swanee_jain.VB_ProcData.VB_Invoke_Func = " \n32"
la_swanee_jain = (2 * (Log((((e / Di) / 3.7) + (5.74 * (Re ^ (-0.9))))) / Log(10))) ^ (-2)
End Function

Function relative_ruwheid(materiaal)
Attribute relative_ruwheid.VB_ProcData.VB_Invoke_Func = " \n32"
Select Case materiaal
    Case "getrokken metalen buis", "koper", "messing", "brons", "tin", "aluminium"
        relative_ruwheid = 0.0014
    Case "kunststof", "glas", "plexiglas", "rubber"
        relative_ruwheid = 0.0015
    Case "staal gewalst naadloos"
        relative_ruwheid = 0.04
    Case "staal gewalst lasnaad"
        relative_ruwheid = 0.07
    Case "staal verzinkt naadloos"
        relative_ruwheid = 0.12
    Case "staal gegalvaniseerd lasnaad"
        relative_ruwheid = 0.008
    Case "staal matig verroest"
        relative_ruwheid = 0.0175
    Case "staal sterk verroest"
        relative_ruwheid = 2.5
    Case "gietijzer nieuw"
        relative_ruwheid = 0.4
    Case "gietijzer nieuw bitumen"
        relative_ruwheid = 0.115
    Case "gietijzer licht verroest"
        relative_ruwheid = 1
    Case "gietijzer sterk verroest"
        relative_ruwheid = 2.5
End Select
End Function
