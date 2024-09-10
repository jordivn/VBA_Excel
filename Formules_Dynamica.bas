'===============================================
'@details       Function for Dynamics
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

Function v_opt(v0, a, t)
v_opt = v0 + (a * t)
End Function

Function s_opt(s0, v0, a, t)
s_opt = s0 + (v0 * t) + (0.5 * a * (t ^ 2))
End Function

Function v_ops(v0, a, s, s0)
v_ops = ((v0 ^ 2) + (2 * a * (s - s0))) ^ 0.5
End Function

Function an(v, r)
an = (v ^ 2) / r
End Function

Function a_atan(at, an)
a = ((at) ^ 2 + (an) ^ 2) ^ 0.5
End Function

Function a_atvr(at, v, r)
a = ((at) ^ 2 + (an(v, r)) ^ 2) ^ 0.5
End Function

Function sy_opt(s0y, v0y, t)
sy_opt = s0y + ((v0y) * t) - (0.5 * 9.81 * (t ^ 2))
End Function

Function t_opsy(sy, s0y, v0y)
t_opsy = (-v0y - ((((2500 * (v0y ^ 2)) + (49010 * s0y) - (49010 * sy)) ^ 0.5) / 50)) / -9.802
End Function

Function vy_opt(v0y, t)
vy_opt = v0y - (9.81 * (t ^ 2))
End Function

Function vy_opsy(v0y, sy, sy0)
vy_opsy = ((v0y) ^ 2 - (2 * 9.81 * (sy - sy0)))
End Function
