Sub downloadfile()

Dim oStream As Object
Dim myURL As String

myURL = "https://www.gstatic.com/covid19/mobility/Region_Mobility_Report_CSVs.zip"

Dim WinHttpReq As Object
Set WinHttpReq = CreateObject("Microsoft.XMLHTTP")


WinHttpReq.Open "GET", myURL, False, "username", "password"

WinHttpReq.setRequestHeader "Accept", "*/*"
WinHttpReq.setRequestHeader "User-Agent", "Mozilla/4.0 (compatible; MSIE 6.0; Windows NT 5.0)"
WinHttpReq.setRequestHeader "Proxy-Connection", "Keep-Alive"

WinHttpReq.send

myURL = WinHttpReq.responseBody

If WinHttpReq.Status = 200 Then
    Set oStream = CreateObject("ADODB.Stream")
    oStream.Open
    oStream.Type = 1
    oStream.Write WinHttpReq.responseBody
    oStream.SaveToFile Environ("Userprofile") & "\Downloads\TempFiles\Region_Mobility_Report_CSVs.zip", 2
    oStream.Close
Else
    MsgBox "Returncode:" & WinHttpReq.Status & " Unable to download file."
End If

End Sub
