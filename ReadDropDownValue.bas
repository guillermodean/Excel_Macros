'Function revisar todo

Dim selector as Object
'define the range where the drop down is
Set selector = Range("C3")
  ' Write the list 
ThisWorkbook.Sheets("DL").Range("A1").Value = sentto(selector)

  'this is a function called from a sub
  
Function sentto(selector As Object)
Dim rng As Range
  'selec the sheet where you wan to read 
    Sheets("DL").Select
    Select Case selector.Value
        Case Is = "ST"
            ActiveSheet.ListObjects("Table8").AutoFilter.ShowAllData
            ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=5, Criteria1:="STFP", Operator:=xlFilterValues
            Set rng = Selection.SpecialCells(xlCellTypeVisible)
            sentto = Join(Application.Transpose(rng.Value), ";")

        Case Is = "WF"
            ActiveSheet.ListObjects("Table8").AutoFilter.ShowAllData
            ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=5, Criteria1:="WFFP", Operator:=xlFilterValues
            sentto = Join(Application.Transpose(ActiveSheet.ListObjects("Table8").ListColumns(4).DataBodyRange.Value), ";")

        Case Is = "GL"
            ActiveSheet.ListObjects("Table8").AutoFilter.ShowAllData
            ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=5, Criteria1:="GAFP", Operator:=xlFilterValues
            sentto = Join(Application.Transpose(ActiveSheet.ListObjects("Table8").ListColumns(4).DataBodyRange.Value), ";")

     End Select
    ' sheet DL Table8

End Function
