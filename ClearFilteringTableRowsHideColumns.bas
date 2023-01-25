Sub clear()
'
' clear Macro
'
    Application.ScreenUpdating = False
  ' Define Table name (in this case Table13) and select a field from it
    Range("Table13[[#Headers],[Change description]]").Select
    ' Clear filters
    ActiveSheet.ShowAllData
    'unhide columns
    Range("A:AI").EntireColumn.Hidden = False
    
End Sub
