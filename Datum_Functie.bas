'===============================================
'@details       Date functions, calculate age and days until specified date
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================
Function Leeftijd(GeboorteDatum)
Leeftijd = (Now - CDate(GeboorteDatum)) / 365.25
End Function

Function dagenTotDatum(datum)
dagenTotDatum = CDate(datum) - Now
End Function
