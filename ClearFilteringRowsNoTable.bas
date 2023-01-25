Sub ClearUpdate()
'
' Clear ClearUpdate
'

Dim Sht As Worksheet
Dim i As Long
Application.ScreenUpdating = False
  'Select the active sheet
Set Sht = ActiveSheet
   With Sht.AutoFilter
      'Iterate through the filters
      For i = 1 To .Filters.Count
         If .Filters(i).On Then
            Sht.ShowAllData
            Exit For
         End If
      Next i
   End With

  'un hide rows from 5 to 500
    Range("5:500").EntireRow.Hidden = False
  'unhide columns from A to AI
    Range("A:AI").EntireColumn.Hidden = False

'
End Sub
