Attribute VB_Name = "AddFunctions"

Sub CreateFunctionsDiscriptions()
    On Error Resume Next
    Dim catagories(1 To 9) As Variant
    
    catagories(1) = "JTools - Rooster"
    catagories(2) = "JTools - Database"
    catagories(3) = "JTools - Omrekenen"
    catagories(4) = "JTools - Feestdagen"
    catagories(5) = "JTools - Tijd&Datum functies"
    catagories(6) = "JTools - KKS functies"
    catagories(7) = "JTools - Weer functies"
    catagories(8) = "JTools - Tennet functies"
    'catagories(6) = "JTools - Baily blokken"
    
    Dim argdesc() As Variant
    
    Dim helpFile As String
    helpFile = "https://websensystems.nl"
    
    
    '================================
    ' Rooster
    '================================
    ReDim argdesc(0 To 1)
    argdesc(0) = "Ploeg. Mogelijkheden zijn A,B,C,D,E,F."
    argdesc(1) = "Optioneel: Datum. Mag een cell verwijzing zijn. Bij geen opgave wordt de huidige dag gebruikt."
    Application.MacroOptions "Dienst_PD", "Functie voor het weer geven van een dienst behorende bij een ploeg en datum.", Category:=catagories(1), ArgumentDescriptions:=argdesc, StatusBar:=catagories(1), helpFile:=helpFile & "#Dienst_PD"


    ReDim argdesc(0 To 1)
    argdesc(0) = "Shift. Een dienst. Mag Enkele letter zijn maar ook uitgescheven voorbeelden: V,v,vroege. "
    argdesc(1) = "Optioneel: Datum. Mag een cell verwijzing zijn. Bij geen opgave wordt de huidige dag gebruikt."
    Application.MacroOptions "Dienst_SD", "Functie voor het weer geven van een ploeg behorende bij een dienst en datum.", Category:=catagories(1), ArgumentDescriptions:=argdesc, StatusBar:=catagories(1), helpFile:=helpFile & "#Dienst_SD"
    
    '================================
    ' Database
    '================================
    
    ReDim argdesc(0 To 6)
    argdesc(0) = "KKS nummer (bv 10hbk01ct044q01 of _1hbk01ct044q01)"
    argdesc(1) = "Optioneel: Bewerking. Kan zijn: SUM (totaal), AVG (Gemiddelde), MAX (Grootste), MIN (Kleinste)"
    argdesc(2) = "Start tijd en datum. Kan eventueel verwijzing zijn naar een cel met =nu() of =vandaag()"
    argdesc(3) = "Stop tijd en datum. Kan eventueel verwijzing zijn naar een cel met =nu() of =vandaag()"
    argdesc(4) = "Het rij nummer van het gegeven. 2 is bijvoorbeeld het 3de gegeven. Standaart is 0"
    argdesc(5) = "Een tabel. Bijvoorbeeld HOUR/DAY/WEEK. Bij geen opgegegeven bewerking wordt avg (gemiddelde) aangehouden. Bij leeg laten wordt de tabel RAW (minuut) gebruikt"
    argdesc(6) = "Comment toevoegen met gegevens van de query True/False Default=true"
    Application.MacroOptions "Get_DB_value", "Functie voor het weergeven van een enkele gegeven uit de database", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "KKS nummer. "
    argdesc(1) = "Optioneel: Gegevens bewerking (AVG/SUM/MIN/MAX)."
    argdesc(2) = "Optioneel: Tabel (RAW/HOUR/DAY)."
    Application.MacroOptions "get_DB_table", "Functie voor het weergeven van de mogelijke tabel uit de database", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "KKS nummer. "
    argdesc(1) = "Optioneel: Gegevens bewerking (AVG/SUM/MIN/MAX)."
    argdesc(2) = "Optioneel: Tabel (RAW/HOUR/DAY)."
    Application.MacroOptions "get_DB_kks", "Functie voor het weergeven van de juite database kks", Category:=catagories(2), ArgumentDescriptions:=argdesc, StatusBar:=catagories(2)
    
    
    
    '================================
    ' Omrekenen
    '================================
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenEnergie", "Functie voor het omrekenen van energie.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenDruk", "Functie voor het omrekenen van drukken.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenGewicht", "Functie voor het omrekenen van gewichten.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenInhoud", "Functie voor het omrekenen van inhouden.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenLengte", "Functie voor het omrekenen van lengtes.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
  
    ReDim argdesc(0 To 2)
    argdesc(0) = "Rekenwaarde. "
    argdesc(1) = "Eenheid van het eindproduct."
    argdesc(2) = "Optioneel: Eenheid van het beginproduct."
    Application.MacroOptions "OmrekenenSnelheid", "Functie voor het omrekenen van snelheid.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Flow in m3."
    argdesc(1) = "Druk in Bar"
    argdesc(2) = "Temperatuur in C"
    Application.MacroOptions "OmrekenennaarNm3", "Functie voor het omrekenen van flow van m3 naar Nm3.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Flow in Nm3."
    argdesc(1) = "Druk in Bar"
    argdesc(2) = "Temperatuur in C"
    Application.MacroOptions "Omrekenennaarm3", "Functie voor het omrekenen van flow van Nm3 naar m3.", Category:=catagories(3), ArgumentDescriptions:=argdesc, StatusBar:=catagories(3)
    
    '================================
    ' Feestdagen
    '================================
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 of 2"
    Application.MacroOptions "Pasen", "Functie voor het weergeven van de datum van Pasen.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 t/m 4"
    Application.MacroOptions "Carnaval", "Functie voor het weergeven van de datum van Carnaval.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Optioneel: Van welk jaar"
    Application.MacroOptions "GoedeVrijdag", "Functie voor het weergeven van de datum van Goede Vrijdag.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Optioneel: Van welk jaar"
    Application.MacroOptions "Hemelvaart", "Functie voor het weergeven van de datum van Hemelvaart.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 of 2"
    Application.MacroOptions "Pinksteren", "Functie voor het weergeven van de datum van Pinksteren.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Van welk jaar"
    argdesc(1) = "Optioneel: 1 t/m 4"
    Application.MacroOptions "Vierdaagse", "Functie voor het weergeven van de datum van de Vierdaagse.", Category:=catagories(4), ArgumentDescriptions:=argdesc, StatusBar:=catagories(4)
    
    '================================
    ' Datum en Tijd
    '================================

    ReDim argdesc(0 To 0)
    argdesc(0) = "Geboortedatum tussen aanhalingstekens "
    Application.MacroOptions "Leeftijd", "Functie voor het berekenen van een leeftijd.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)

    ReDim argdesc(0 To 0)
    argdesc(0) = "Datum tussen aanhalingstekens"
    Application.MacroOptions "dagenTotDatum", "Functie voor het berekenen van het aantal dagen tot een bepaalde datum.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "VolgendeHeleUur", "Functie voor het berekenen van het aankomend hele uur.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 2)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    argdesc(2) = "Optioneel: Verschuiving in de tijd"
    Application.MacroOptions "LaatsteHeleUur", "Functie voor het berekenen van het laatste hele uur.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "TijdEindeWacht", "Functie voor het berekenen van het einde van de wacht.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "TijdStartWacht", "Functie voor het berekenen van het begin van de wacht.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
   
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "StartWeek", "Functie voor het berekenen van het begin van de week.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "EindWeek", "Functie voor het berekenen van het einde van de week.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "StartCyclus", "Functie voor het berekenen van het begin van de cycles.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Datum en tijd (Default is nu)"
    argdesc(1) = "Optioneel: Format (bv dd-mm-yyyy)"
    Application.MacroOptions "EindeCyclus", "Functie voor het berekenen van het einde van de cyclus.", Category:=catagories(5), ArgumentDescriptions:=argdesc, StatusBar:=catagories(5)
    
    '================================
    ' KKS functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Kks nummer"
    argdesc(1) = "Optioneel: Lange over korte versie (waar/onwaar)"
    Application.MacroOptions "getDBKKS", "Haalt de bekende db kks format op.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoDiscr", "Geeft de ultimo omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoKostenplaatsNum", "Geeft het ultimo kostenplaatsnummer weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoPIDNum", "Geeft het ultimo pid weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoElectrischeVerdelerNum", "Geeft het in ultimo bekende verdeler weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoWSinHoofdstroom", "Geeft aan of de werkschakelaar volgens ultimo in de hoofdstroom zit.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoZone", "Geeft het in ultimo bekende flitslicht zone weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kostenplaatscode"
    Application.MacroOptions "getUltimoKostenplaats", "Geeft de kostenplaats omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDDiscr", "Geeft P&ID omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDVersion", "Geeft P&ID versie weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDLastChange", "Geeft P&ID versie datum weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "PID nummer"
    Application.MacroOptions "getUltimoPIDResponsible", "Geeft P&ID verantwoordelijke weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Electrische verdeler"
    Application.MacroOptions "getUltimoElectrischeVerdeler", "Geeft de omschrijving van de electrische verdeler weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoZoneDiscr", "Geeft de flitslicht omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoZoneLevel", "Geeft de flitslicht zone hoogte/verdieping weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Kks nummer"
    Application.MacroOptions "getUltimoFlitslicht", "Geeft het in ultimo bekende flitslicht nummer weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)
    
    ReDim argdesc(0 To 0)
    argdesc(0) = "Flitslicht Zone"
    Application.MacroOptions "getUltimoFlitslichtDiscr", "Geeft de flitslicht omschrijving weer.", Category:=catagories(6), ArgumentDescriptions:=argdesc, StatusBar:=catagories(6)

    '================================
    ' Weer functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_StationNaam", "Haalt de naam van het station op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Moment", "Haalt de datum en tijd van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Temperatuur", "Haalt de temperatuur van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Vochtigheid", "Haalt de vochtigheid van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windsnelheid", "Haalt de Windsnelheid in m/s van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_WindrichtingGR", "Haalt de windrichting in graden van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windrichting", "Haalt de windrichting in kompassrichtingen van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Luchtdruk", "Haalt de luchtdruk van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Windstoten", "Haalt de snelheid van windstoten in m/s van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Regen", "Haalt de de hoeveelheid regen in mm/h van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_Zicht", "Haalt de zichtafstand in meters van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_ZonIntensiteit", "Haalt de zonintensiteit in W/m2 van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: Stationsnummer, Default = Arnhem"
    Application.MacroOptions "getWeerData_TemperatuurOp10cm", "Haalt de temperatuur op 10cm vanaf maaiveld van de huidige set op.", Category:=catagories(7), ArgumentDescriptions:=argdesc, StatusBar:=catagories(7)
    
    '================================
    ' Tennet functies
    '================================
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_Tijd", "Haalt de tijd van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_OpregelVermogen", "Haalt het opgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_AfregelVermogen", "Haalt het afgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_OpregelVermogenReserve", "Haalt het reserve opgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_AfregelVermogenReserve", "Haalt het reserve afgeregelvermogen van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_Noodvermogen", "Haalt op of er noodvermogen is geactiveerd van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_PrijsMin", "Haalt de minimale prijs van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
    ReDim argdesc(0 To 1)
    argdesc(0) = "Optioneel: ItemNum, aantal blokken eerder"
    Application.MacroOptions "getTennetData_PrijsMax", "Haalt de maximale prijs van de huidige gegevensset op.", Category:=catagories(8), ArgumentDescriptions:=argdesc, StatusBar:=catagories(8)
    
End Sub
