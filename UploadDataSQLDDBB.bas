Sub uploadData()

    Dim lr As Integer
  'define de variables that are going to take the values of our range of cells
    Dim eco, id_rev_old, id_rev_new, denomin_old, denomin_new, denomin_eng_old, denomin_eng_new, ref_prov_old, ref_prov_new, prov_old, prov_new, cod_purch_old, cod_purch_new, ch_descriip, origin, resp, rel_date, applicab, prior As String
  'Define a variable String for the string query
    
    Dim strQuery As String
  'identify lenght of the range we are going to uload
    lr = Cells(Rows.Count, 1).End(xlUp).Row
    
    If connectDB() Then
'See connectDB() function on this repository as ConnectionToDDBBFunc.bas
        For i = 9 To lr
' iterate al the rows and assign to de variables
            
            eco = ActiveSheet.Cells(i, 1).Value
            id_rev_old = ActiveSheet.Cells(i, 2).Value
            id_rev_new = ActiveSheet.Cells(i, 3).Value
            
            
            
            
            rev = ActiveSheet.Cells(i, 3).Value
            applicability = ActiveSheet.Cells(i, 10).Value
            projects = ActiveSheet.Cells(i, 16).Value
            impCom = ActiveSheet.Cells(i, 15).Value
            impDecProd = ActiveSheet.Cells(i, 19).Value
            impDecWind = ActiveSheet.Cells(i, 21).Value
            impDecServ = ActiveSheet.Cells(i, 23).Value
            permanentSolution = ActiveSheet.Cells(i, 25).Value
            DueDate = ActiveSheet.Cells(i, 26).Value
            ECR = ActiveSheet.Cells(i, 27).Value
            eco = ActiveSheet.Cells(i, 28).Value
            agenda = ActiveSheet.Cells(i, 29).Value
            openPoints = ActiveSheet.Cells(i, 30).Value
            escalation = ActiveSheet.Cells(i, 31).Value
            feedback = ActiveSheet.Cells(i, 32).Value
            Comments = ActiveSheet.Cells(i, 33).Value
            TN = ActiveSheet.Cells(i, 2).Value
' check the values with the function checkValue defined in this repository with name CheckEmptyCell.bas
            rev = checkValue(rev)
            applicability = checkValueInt(applicability)
            projects = checkValue(projects)
            impCom = checkValue(impCom)
            impDecProd = checkValue(impDecProd)
            impDecWind = checkValue(impDecWind)
            impDecServ = checkValue(impDecServ)
            permanentSolution = checkValue(permanentSolution)
            DueDate = checkValue(DueDate)
            ECR = checkValue(ECR)
            eco = checkValue(eco)
            agenda = checkValue(agenda)
            openPoints = checkValue(openPoints)
            escalation = checkValueInt(escalation)
            feedback = checkValue(feedback)
            Comments = checkValue(Comments)
            
            
'Form the query string in SQLnote that the variable fileds have  ' to identify the values in SQL  and " + & to assing the variables from the code
            strQuery = "UPDATE [PM_EPC].[dbo].[TNComittee$] SET [Rev] = " & rev & ", [Applicability] = " & applicability & _
            ",[Affected Projects]= '" & projects & "' ,[Implementation Committee] = '" & impCom & "' , [Implementation decision production] = '" & impDecProd & _
            "' , [Implementation decision windfarm] = '" & impDecWind & "' , [Implementation decision service] = '" & impDecServ & "' , [Permanent solution needed] = '" & permanentSolution & _
            "' , [Due date for permanent solution] = '" & DueDate & "' , [ECR(AST)/ETO/ECO/DECO (…)] = '" & ECR & "' , [ECO Release date] = '" & eco & "' , [Agenda follow up] = '" & agenda & _
            "' , [Open points] = '" & openPoints & "' , [Escalation] = " & escalation & ", [Feedback needed from] = '" & feedback & "' WHERE [Change Nr#] = '" & TN & "'"
            
            Set rsData = New ADODB.Recordset
  'execute command
            rsData.Open strQuery, Conn
        
        Next i
        
        
    End If

End Sub
