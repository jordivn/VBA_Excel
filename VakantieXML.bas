Attribute VB_Name = "Module1"
Sub test()
rijnum = 1
Set oXMLFile = CreateObject("Microsoft.XMLDOM")
    oXMLFile.async = False
    oXMLFile.Load ("https://opendata.rijksoverheid.nl/v1/sources/rijksoverheid/infotypes/schoolholidays")
    
    Set years = oXMLFile.SelectNodes("/documents/document")
    TotalNumYears = years.Length - 1
    
    For i = 0 To TotalNumYears
        Set Vacanties = oXMLFile.SelectNodes("/documents/document[" & i & "]/content/contentblock/vacations/vacation")
        TotalNumVac = Vacanties.Length - 1
        For x = 0 To TotalNumVac
            Range("a" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/schoolyear/text()").NodeValue), " ", "")
            Range("b" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/type/text()").NodeValue), " ", "")
            If oXMLFile.SelectNodes("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions").Length > 1 Then
            Range("c" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions[2]/region/text()").NodeValue), " ", "")
            Range("d" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions[2]/startdate/text()").NodeValue), " ", "")
Range("e" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions[2]/enddate/text()").NodeValue), " ", "")
            Else
             Range("c" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions/region/text()").NodeValue), " ", "")
            Range("d" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions/startdate/text()").NodeValue), " ", "")
Range("e" & rijnum).Value = Replace(Trim(oXMLFile.SelectSingleNode("/documents/document[" & i & "]/content/contentblock/vacations/vacation[" & x & "]/regions/enddate/text()").NodeValue), " ", "")
            End If
            
            
            Range("f" & rijnum).Value = CDate(Format(Split(Range("d" & rijnum).Value, "T")(0), "dd-mm-yyyy"))
            Range("g" & rijnum).Value = CDate(Format(Split(Range("e" & rijnum).Value, "T")(0), "dd-mm-yyyy"))
            
            
            
            
            
            rijnum = rijnum + 1
        Next
    
    Next
    
'Range("a1").Value = oXMLFile.SelectSingleNode("/documents/document[2]/content/contentblock/title/text()").NodeValue

End Sub
