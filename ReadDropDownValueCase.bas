'Function revisar todo

Dim selector as Object
'define the range where the drop down is
Set selector = Range("C3")
  ' call the function and assing to a range value 
ThisWorkbook.Sheets("DL").Range("A1").Value = sentto(selector)

  'this is a function called from a sub
  
Function sentto(selector As Object)
Dim rng As Range
    'selec the sheet where you wan to perform the action
    Sheets("DL").Select
      'Case method receive input from selector, we read the value of it
    Select Case selector.Value
        Case Is = "ST"
        'In case the value of the selector is ST the do things, in this case i will explain
            ActiveSheet.ListObjects("Table8").AutoFilter.ShowAllData
        'We are filtering a table where the values are STFP
            ActiveSheet.ListObjects("Table8").Range.AutoFilter Field:=5, Criteria1:="STFP", Operator:=xlFilterValues
        'Now we define this range to operate later
            Set rng = Selection.SpecialCells(xlCellTypeVisible)
          'we assing the return of the function "sentto" a string wich joins all the columns transposed into a row on the previous range (rng) wit mettho join and delimiter ;
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
