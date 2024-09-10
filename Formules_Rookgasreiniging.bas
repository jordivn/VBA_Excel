Attribute VB_Name = "Formules_Rookgasreiniging"
'===============================================
'@details       Formula's for calculating on flue gas cleaning based on dutch Martek studies
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

Function Quench_temp(PercentageWaterDamp As Double, QuenchDruk As Double)
Attribute Quench_temp.VB_ProcData.VB_Invoke_Func = " \n31"
'Waterpercentage bv 0.25 (%/100)
'QuenchDruk bv 0,985 (Bara)
Quench_temp = Tsat_p(PercentageWaterDamp * QuenchDruk)
End Function


Function Quench_druk(PercentageWaterDamp As Double, QuenchTemp As Double)
Attribute Quench_druk.VB_ProcData.VB_Invoke_Func = " \n31"
'Waterpercentage bv 0.25 (%/100)
'QuenchTemp bv 65 (C)
Quench_druk = psat_T(QuenchTemp) / PercentageWaterDamp
End Function

Function Massa_suppletie(MVerdamping As Double, DichtheidSup As Double, DichtheidSpui As Double)
Attribute Massa_suppletie.VB_ProcData.VB_Invoke_Func = " \n31"
'MVerdamping BV 1.2 (kg/s)
'DichtheidSup BV 1025 (kg/m3)
'DichtheidSpui BV 1025 (kg/m3)

Indikking = DichtheidSpui / DichtheidSup
Massa_suppletie = MVerdamping * (Indikking / (Indikking - 1))

End Function
