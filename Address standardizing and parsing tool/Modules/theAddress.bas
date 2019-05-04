Attribute VB_Name = "theAddress"
Const Col_Status = 2
Const Col_formatted_address = 3
Const Col_street_number = 4
Const Col_street_address = 5
Const Col_Route = 6
Const Col_premise = 7
Const Col_subpremise = 8
Const Col_sublocality = 9
Const Col_locality = 10
Const Col_administrative_area_level_1 = 11
Const Col_administrative_area_level_2 = 12
Const Col_administrative_area_level_3 = 13
Const Col_administrative_area_level_4 = 14
Const Col_administrative_area_level_5 = 15
Const Col_Country = 16
Const Col_postal_code = 17

Function Tuning(ByVal b As String, ByVal a As String) As String
Dim c, d As String
c = "</" & a & ">"
b = Replace(b, c, "")

d = "<" & a & "/>"
b = Replace(b, d, "")


a = "<" & a & ">"
b = Replace(b, a, "")
Tuning = b
End Function

Sub Address(ByVal API As Boolean)
Dim Address, Web, Status As String
Dim address_count, i, j, z As Long
Dim IE, No_result_Check As Object
Dim EachResult, EachLevel, EachType, Eachformatted_address, Eachstreet_number, Eachstreet_address, EachRoute, Eachpremise, Eachsubpremise, Eachsublocality, Eachlocality, Eachadministrative_area_level_1, Eachadministrative_area_level_2, Eachadministrative_area_level_3, Eachadministrative_area_level_4, Eachadministrative_area_level_5, EachCountry, Eachpostal_code As Object

With Worksheets("Workplace")
    .Activate
    .Cells.Font.Name = "Calibri"
    .Cells.Font.Size = 11
End With
Worksheets("Workplace").Range("B:Z").Clear

ActiveWindow.Zoom = 100


address_count = Cells(1, 1).CurrentRegion.Rows.Count
Cells(1, Col_Status).Value = "Status"
Cells(1, Col_formatted_address).Value = "formatted_address"
Cells(1, Col_street_number).Value = "street_number"
Cells(1, Col_street_address).Value = "street_address"
Cells(1, Col_Route).Value = "Route"
Cells(1, Col_premise).Value = "premise"
Cells(1, Col_subpremise).Value = "subpremise"
Cells(1, Col_sublocality).Value = "sublocality"
Cells(1, Col_locality).Value = "locality"
Cells(1, Col_administrative_area_level_1).Value = "administrative_area_level_1"
Cells(1, Col_administrative_area_level_2).Value = "administrative_area_level_2"
Cells(1, Col_administrative_area_level_3).Value = "administrative_area_level_3"
Cells(1, Col_administrative_area_level_4).Value = "administrative_area_level_4"
Cells(1, Col_administrative_area_level_5).Value = "administrative_area_level_5"
Cells(1, Col_Country).Value = "Country"
Cells(1, Col_postal_code).Value = "postal_code"

Set IE = CreateObject("InternetExplorer.Application")
IE.Visible = False

For i = 2 To address_count
Address = Worksheets("Workplace").Cells(i, 1).Value
If API Then
Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & Address & "&sensor=false&language=en"
Else
Web = "https://maps.googleapis.com/maps/api/geocode/xml?address=" & Address & "&sensor=false&language=en"
End If
IE.Navigate (Web)

Do While IE.Busy Or IE.ReadyState <> 4: DoEvents: Loop
Set No_result_Check = IE.document.getElementsBytagName("status")
On Error Resume Next
Status = Tuning(No_result_Check(0).innertext, "status")

If StrComp(Status, "OK") = 0 Then
Set EachResult = IE.document.getElementsBytagName("result")
Eachformatted_address = IE.document.getElementsBytagName("formatted_address")
Worksheets("Workplace").Cells(i, Col_formatted_address).Value = Tuning(Eachformatted_address.innertext, "formatted_address")

If EachResult.Length > 1 Then
Worksheets("Workplace").Cells(i, Col_Status).Value = "Return multiple results"
Else
Worksheets("Workplace").Cells(i, Col_Status).Value = "Single result"
End If
'For j = 0 To EachResult.Length - 1
j = 0
Set EachLevel = IE.document.getElementsBytagName("result")(0).getElementsBytagName("address_component")
For k = 0 To EachLevel.Length - 1
Set EachType = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("type")

For z = 0 To EachType.Length - 1
If StrComp(EachType(z).innertext, "<type>street_number</type>") = 0 Then
Set Eachstreet_number = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_street_number).Value = Tuning(Eachstreet_number(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>street_address</type>") = 0 Then
Set Eachstreet_address = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_street_address).Value = Tuning(Eachstreet_address(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>route</type>") = 0 Then
Set EachRoute = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_Route).Value = Tuning(EachRoute(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>premise</type>") = 0 Then
Set Eachpremise = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_premise).Value = Tuning(Eachpremise(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>subpremise</type>") = 0 Then
Set Eachsubpremise = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_subpremise).Value = Tuning(Eachsubpremise(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>sublocality</type>") = 0 Then
Set Eachsublocality = IE.document.getElementsBytagName("result")(j).getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_sublocality).Value = Tuning(Eachsublocality(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>locality</type>") = 0 Then
Set Eachlocality = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_locality).Value = Tuning(Eachlocality(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>administrative_area_level_1</type>") = 0 Then
Set Eachadministrative_area_level_1 = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_administrative_area_level_1).Value = Tuning(Eachadministrative_area_level_1(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>administrative_area_level_2</type>") = 0 Then
Set Eachadministrative_area_level_2 = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_administrative_area_level_2).Value = Tuning(Eachadministrative_area_level_2(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>administrative_area_level_3</type>") = 0 Then
Set Eachadministrative_area_level_3 = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_administrative_area_level_3).Value = Tuning(Eachadministrative_area_level_3(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>administrative_area_level_4</type>") = 0 Then
Set Eachadministrative_area_level_4 = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_administrative_area_level_4).Value = Tuning(Eachadministrative_area_level_4(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>administrative_area_level_5</type>") = 0 Then
Set Eachadministrative_area_level_5 = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_administrative_area_level_5).Value = Tuning(Eachadministrative_area_level_5(0).innertext, "long_name")
End If


If StrComp(EachType(z).innertext, "<type>country</type>") = 0 Then
Set EachCountry = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_Country).Value = Tuning(EachCountry(0).innertext, "long_name")
End If

If StrComp(EachType(z).innertext, "<type>postal_code</type>") = 0 Then
Set Eachpostal_code = IE.document.getElementsBytagName("result")(j).document.getElementsBytagName("address_component")(k).getElementsBytagName("long_name")
Worksheets("Workplace").Cells(i, Col_postal_code).Value = Tuning(Eachpostal_code(0).innertext, "long_name")
End If
Next z
Next k
'Next j

Set No_result_Check = Nothing
Set Eachformatted_address = Nothing
Set Eachstreet_number = Nothing
Set Eachstreet_address = Nothing
Set EachRoute = Nothing
Set Eachpremise = Nothing
Set Eachsubpremise = Nothing
Set Eachsublocality = Nothing
Set Eachlocality = Nothing
Set Eachadministrative_area_level_1 = Nothing
Set Eachadministrative_area_level_2 = Nothing
Set Eachadministrative_area_level_3 = Nothing
Set Eachadministrative_area_level_4 = Nothing
Set Eachadministrative_area_level_5 = Nothing
Set EachCountry = Nothing
Set Eachpostal_code = Nothing
Set EachResult = Nothing
Set EachLevel = Nothing
Set EachType = Nothing
End If

If Worksheets("Workplace").Cells(i, Col_formatted_address).Value = "" Then
Worksheets("Workplace").Cells(i, Col_Status).Value = "Zero Result"
End If

Next i

IE.Quit
Set IE = Nothing
End Sub
