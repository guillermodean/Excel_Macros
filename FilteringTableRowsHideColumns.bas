Sub TN_WF_New()
'
  ' Filtering rows by color 
  ActiveSheet.ListObjects("Table13").Range.AutoFilter Field:=3, Criteria1:= _
        RGB(255, 255, 0), Operator:=xlFilterCellColor
  ' Filtering rows by multiple values

    ActiveSheet.ListObjects("Table13").Range.AutoFilter Field:=16, Criteria1:= _
        Array("NA/WF", "NA/WF/SERV", "ST/WF", "WF", "WF/SERV"), Operator:=xlFilterValues
  
  'hide and unhide columns
    Range("A:B").EntireColumn.Hidden = True
  Range("T:V").EntireColumn.Hidden = False
  
End Sub
