Attribute VB_Name = "Datum_Functie"
'===============================================
'@details       Date functions, calculate age and days until specified date
'@author        Jordi van Nistelrooij @ Webs en Systems
'@email         info@websensystems.nl
'@version       1.0.0
'@date          2024-09-10
'@copyright     Non of these scripts maybe copied or modified without permission of the author
'===============================================
Function Leeftijd(GeboorteDatum)
Attribute Leeftijd.VB_Description = "Functie voor het berekenen van de leeftijd"
Attribute Leeftijd.VB_ProcData.VB_Invoke_Func = " \n21"
Leeftijd = (Now - CDate(GeboorteDatum)) / 365.25
End Function

Function dagenTotDatum(datum)
Attribute dagenTotDatum.VB_Description = "Functie voor het berekenen van het aantal dagen tot een bepaalde datum"
Attribute dagenTotDatum.VB_ProcData.VB_Invoke_Func = " \n21"
dagenTotDatum = CDate(datum) - Now
End Function
