Sub GetLinks()

Dim path1 As String
Dim i As Integer
Dim xmlHttp As Object
Dim URL As String
Dim WB As Workbook
Dim html As New HTMLDocument
Dim collection As MSHTML.IHTMLElementCollection
Dim element As MSHTML.HTMLInputElement, subElement As MSHTML.HTMLInputElement

Dim WB1 As Workbook
Set WB1 = Workbooks.Add
ActiveWorkbook.SaveAs FileName:=Environ("Userprofile") & "\Downloads\TempFiles\PasteLinksHere.xlsx"
Set WB = Workbooks.Open(FileName:=Environ("Userprofile") & "\Downloads\TempFiles\PasteLinksHere.xlsx", local:=True)

URL = "https://www.google.com/covid19/mobility/"

Set xmlHttp = CreateObject("MSXML2.XMLHTTP")
xmlHttp.Open "GET", URL, False
xmlHttp.setRequestHeader "Content-Type", "text/xml"
xmlHttp.send

Set html = CreateObject("htmlfile")
html.body.innerHTML = xmlHttp.responseText

Set collection = html.getElementsByTagName("a")

i = 1
For Each element In collection
    'Debug.Print element.href
    WB.Sheets("Sheet1").Range("A" & i).Value = element.href
    i = i + 1
    Next

myURL1 = WB.Sheets("Sheet1").Range("A16").Value
myURL2 = WB.Sheets("Sheet1").Range("A19").Value

WB.Close

End Sub
