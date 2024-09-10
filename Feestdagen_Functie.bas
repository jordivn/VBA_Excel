'===============================================
'@details       Function for calculate dutch special dates (easter, christmas and so on)
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================

Function Pasen(Optional Jaar = 1, Optional dag = 1)
If Jaar = 1 Then Jaar = Year(Now)
a = DateSerial(Jaar, 4, 1) / 7
If Jaar Mod 19 = 0 Then b = 19
c = (Jaar Mod 19 + b) * 19 - 7
D = (c Mod 30) / 7
Pasen = (Round(a + D, 0) * 7 - 6) + (dag - 1)
End Function

Function Carnaval(Optional Jaar = 1, Optional dag = 1)
Carnaval = Pasen(Jaar) - 50 + (dag - 1)
End Function

Function GoedeVrijdag(Optional Jaar = 1)
GoedeVrijdag = Pasen(Jaar) - 2
End Function

Function Hemelvaart(Optional Jaar = 1)
Hemelvaart = Pasen(Jaar) + 39
End Function

Function Pinksteren(Optional Jaar = 1, Optional dag = 1)
Pinksteren = Pasen(Jaar) + 49 + (dag - 1)
End Function

Function Vierdaagse(Optional Jaar = 1, Optional dag = 1)
If Jaar = 1 Then Jaar = Year(Now)
startDate = CDate("1-7-" & Jaar)
I = 0
While I <> 3

If Weekday(startDate, vbMonday) = 2 Then I = I + 1
startDate = startDate + 1
Wend
Vierdaagse = startDate - 1 + (dag - 1)
End Function


