Sub Check_Date()

Dim request As Object
Dim response As String
Dim html As New HTMLDocument
Dim website As String
Dim checktime As Variant

website = "https://www.google.com/covid19/mobility/"
Set request = CreateObject("MSXML2.XMLHTTP")
request.Open "GET", website, False
request.setRequestHeader "If-Modified-Since", "Sat, 1 Jan 2000 00:00:00 GMT"
request.send
response = StrConv(request.responseBody, vbUnicode)
html.body.innerHTML = response
checktime = html.getElementsByClassName("report-info-text").item(0).innerText

Dim WB1 As Workbook
Set WB1 = Workbooks.Add
ActiveWorkbook.SaveAs FileName:=Environ("Userprofile") & "\Downloads\TempFiles\PasteLinksHere.xlsx"
Set WB = Workbooks.Open(FileName:=Environ("Userprofile") & "\Downloads\TempFiles\PasteLinksHere.xlsx", local:=True)
WB.Sheets("Sheet1").Range("A1").Value = checktime
WB.Sheets("Sheet1").Range("A2").Formula = "=RIGHT(LEFT(RIGHT(A1,LEN(A1)-15),LEN(RIGHT(A1,LEN(A1)-15))-1),LEN(LEFT(RIGHT(A1,LEN(A1)-15),LEN(RIGHT(A1,LEN(A1)-15))-1))-1)"
WB.Sheets("Sheet1").Range("A2").Copy
WB.Sheets("Sheet1").Range("A2").PasteSpecial xlPasteValues
If Format(Now, "dd-mm-yyyy") = WB.Sheets("Sheet1").Range("A2").Value _
    Or Format(Now, "d-m-yyyy") = WB.Sheets("Sheet1").Range("A2").Value _
    Or Format(Now, "dd-m-yyyy") = WB.Sheets("Sheet1").Range("A2").Value _
    Or Format(Now, "d-mm-yyyy") = WB.Sheets("Sheet1").Range("A2").Value Then
    Exit Sub
    Else
    WB.Close
    MsgBox "Wait for today's data to be Issued"
    End
    Dim FSO
    Set FSO = CreateObject("Scripting.FileSystemObject")
    FSO.DeleteFolder Environ("Userprofile") & "\Downloads\TempFiles", False
    End
End If

End Sub
