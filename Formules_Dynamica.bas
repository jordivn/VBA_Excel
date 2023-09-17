Attribute VB_Name = "Formules_Dynamica"
Function v_opt(v0, a, t)
Attribute v_opt.VB_Description = "Snelheid op gegeven tijd"
Attribute v_opt.VB_ProcData.VB_Invoke_Func = " \n28"
v_opt = v0 + (a * t)
End Function

Function s_opt(s0, v0, a, t)
Attribute s_opt.VB_Description = "Afstand op gegeven tijd"
Attribute s_opt.VB_ProcData.VB_Invoke_Func = " \n28"
s_opt = s0 + (v0 * t) + (0.5 * a * (t ^ 2))
End Function

Function v_ops(v0, a, s, s0)
Attribute v_ops.VB_Description = "Snelheid op gegeven afstand"
Attribute v_ops.VB_ProcData.VB_Invoke_Func = " \n28"
v_ops = ((v0 ^ 2) + (2 * a * (s - s0))) ^ 0.5
End Function

Function an(v, r)
Attribute an.VB_Description = "Versnelling in N"
Attribute an.VB_ProcData.VB_Invoke_Func = " \n28"
an = (v ^ 2) / r
End Function

Function a_atan(at, an)
Attribute a_atan.VB_ProcData.VB_Invoke_Func = " \n28"
a = ((at) ^ 2 + (an) ^ 2) ^ 0.5
End Function

Function a_atvr(at, v, r)
Attribute a_atvr.VB_ProcData.VB_Invoke_Func = " \n28"
a = ((at) ^ 2 + (an(v, r)) ^ 2) ^ 0.5
End Function

Function sy_opt(s0y, v0y, t)
Attribute sy_opt.VB_Description = "Afstand Y op ggeven tijd"
Attribute sy_opt.VB_ProcData.VB_Invoke_Func = " \n28"
sy_opt = s0y + ((v0y) * t) - (0.5 * 9.81 * (t ^ 2))
End Function

Function t_opsy(sy, s0y, v0y)
Attribute t_opsy.VB_Description = "Tijd tot gegeven afstand"
Attribute t_opsy.VB_ProcData.VB_Invoke_Func = " \n28"
t_opsy = (-v0y - ((((2500 * (v0y ^ 2)) + (49010 * s0y) - (49010 * sy)) ^ 0.5) / 50)) / -9.802
End Function

Function vy_opt(v0y, t)
Attribute vy_opt.VB_Description = "Snelheid Y op gegeven tijd"
Attribute vy_opt.VB_ProcData.VB_Invoke_Func = " \n28"
vy_opt = v0y - (9.81 * (t ^ 2))
End Function

Function vy_opsy(v0y, sy, sy0)
Attribute vy_opsy.VB_Description = "Stelheid T op gegeven afstand"
Attribute vy_opsy.VB_ProcData.VB_Invoke_Func = " \n28"
vy_opsy = ((v0y) ^ 2 - (2 * 9.81 * (sy - sy0)))
End Function
