
'Lets suppose we have an array made of values from cells A2 To A20 of the first sheet of our Template Workbook (TEM)
RangeArray = TEM.Worksheets(1).Range("A2:A20").Value

'We need this array to filter column C of TEM.Worksheets(2)
'For the filter to work we need each value of the RangeArray to be turned into a string

Dim StringArray() As String
ReDim StringArray(UBound(RangeArray)) As String
For i = 1 To UBound(RangeArray)
    StringArray(i) = CStr(RangeArray(i, 1))
Debug.Print StringArray(i)
Next i

'Now that we have a new StringArray, we can use it to filter column C

TEM.Worksheets(2).Range("A1:E1").AutoFilter Field:=3, Criteria1:=sArray, Operator:=xlFilterValues
