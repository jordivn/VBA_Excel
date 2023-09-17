Attribute VB_Name = "Datum_Functie"
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
