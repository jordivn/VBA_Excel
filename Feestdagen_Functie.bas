Attribute VB_Name = "Feestdagen_Functie"
Function Pasen(Optional Jaar = 1, Optional dag = 1)
Attribute Pasen.VB_Description = "Geeft de datum van Pasen weer"
Attribute Pasen.VB_ProcData.VB_Invoke_Func = " \n21"
If Jaar = 1 Then Jaar = Year(Now)
a = DateSerial(Jaar, 4, 1) / 7
If Jaar Mod 19 = 0 Then b = 19
c = (Jaar Mod 19 + b) * 19 - 7
D = (c Mod 30) / 7
Pasen = (Round(a + D, 0) * 7 - 6) + (dag - 1)
End Function

Function Carnaval(Optional Jaar = 1, Optional dag = 1)
Attribute Carnaval.VB_Description = "Geeft de datum van carnaval weer"
Attribute Carnaval.VB_ProcData.VB_Invoke_Func = " \n21"
Carnaval = Pasen(Jaar) - 50 + (dag - 1)
End Function

Function GoedeVrijdag(Optional Jaar = 1)
Attribute GoedeVrijdag.VB_Description = "Geeft de datum van goede vrijdag weer"
Attribute GoedeVrijdag.VB_ProcData.VB_Invoke_Func = " \n21"
GoedeVrijdag = Pasen(Jaar) - 2
End Function

Function Hemelvaart(Optional Jaar = 1)
Attribute Hemelvaart.VB_Description = "Geeft de datum van Hemelvaart weer"
Attribute Hemelvaart.VB_ProcData.VB_Invoke_Func = " \n21"
Hemelvaart = Pasen(Jaar) + 39
End Function

Function Pinksteren(Optional Jaar = 1, Optional dag = 1)
Attribute Pinksteren.VB_ProcData.VB_Invoke_Func = " \n29"
Pinksteren = Pasen(Jaar) + 49 + (dag - 1)
End Function

Function Vierdaagse(Optional Jaar = 1, Optional dag = 1)
Attribute Vierdaagse.VB_Description = "Geeft de datum van de vierdaagse weer"
Attribute Vierdaagse.VB_ProcData.VB_Invoke_Func = " \n21"
If Jaar = 1 Then Jaar = Year(Now)
startDate = CDate("1-7-" & Jaar)
I = 0
While I <> 3

If Weekday(startDate, vbMonday) = 2 Then I = I + 1
startDate = startDate + 1
Wend
Vierdaagse = startDate - 1 + (dag - 1)
End Function


