# geotagging-xslx
Geotagging VBA Macro

```
Sub LoopAddresses()
    Dim c As Object
    Dim columnTarget As Integer
    Dim columnSource As String
    Dim columnError As Integer
    Dim isNew As Boolean
    Dim sError As String
    columnSource = "D"
    columnTarget = 7
    columnError = columnTarget + 3
    For Each c In ActiveSheet.UsedRange.Columns(columnSource).Cells
        isNew = False
        If Cells(c.Row, columnTarget).Value = vbNullString Then
            If Cells(c.Row, columnError).Value <> "ZERO_RESULTS" Then
                Cells(c.Row, columnTarget).Value = getGoogleMapsGeocode(c.Value, sError)
                Cells(c.Row, columnError).Value = sError
                isNew = True
            End If
        End If
    Next c
End Sub



Rem http://stackoverflow.com/questions/4158492/code-to-get-gps-coordinates-from-address-vb6-vba-vbscript
Function getGoogleMapsGeocode(sAddr As String, ByRef sError As String) As String

Dim xhrRequest As XMLHTTP60
Dim sQuery As String
Dim domResponse As DOMDocument60
Dim ixnStatus As IXMLDOMNode
Dim ixnLat As IXMLDOMNode
Dim ixnLng As IXMLDOMNode


' Use the empty string to indicate failure
getGoogleMapsGeocode = ""
sError = ""

Set xhrRequest = New XMLHTTP60
sQuery = "http://maps.googleapis.com/maps/api/geocode/xml?sensor=false&address="
sQuery = sQuery & WorksheetFunction.EncodeURL(sAddr)
xhrRequest.Open "GET", sQuery, False
xhrRequest.send

Set domResponse = New DOMDocument60
domResponse.LoadXML xhrRequest.responseText
Set ixnStatus = domResponse.SelectSingleNode("//status")

If (ixnStatus.Text <> "OK") Then
    sError = ixnStatus.Text
    Exit Function
End If

Set ixnLat = domResponse.SelectSingleNode("/GeocodeResponse/result/geometry/location/lat")
Set ixnLng = domResponse.SelectSingleNode("/GeocodeResponse/result/geometry/location/lng")

getGoogleMapsGeocode = ixnLat.Text & ", " & ixnLng.Text

End Function
```
