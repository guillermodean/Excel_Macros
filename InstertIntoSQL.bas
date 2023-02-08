Sub LoadAll()

    Dim lr As Integer
    Dim index As Integer
    Dim proj As String
    
    
    Dim strQuery As String
  ' table starts in column 2
    lr = Cells(Rows.Count, 2).End(xlUp).Row
' table starts in row 4
    If connectDB() Then
        For i = 4 To lr
'assing the value of the cell in row i and column 2 to the variable
            index = ActiveSheet.Cells(i, 2).Value
            TNtype = ActiveSheet.Cells(i, 3).Value
            ECO = ActiveSheet.Cells(i, 4).Value
            change = ActiveSheet.Cells(i, 5).Value
            rev = ActiveSheet.Cells(i, 6).Value
            reldate = ActiveSheet.Cells(i, 7).Value
            desc = ActiveSheet.Cells(i, 8).Value
            impc = ActiveSheet.Cells(i, 9).Value
            oritypeimpl = ActiveSheet.Cells(i, 10).Value
            origin = ActiveSheet.Cells(i, 11).Value
            eng = ActiveSheet.Cells(i, 12).Value
            app = ActiveSheet.Cells(i, 13).Value
            ecr = ActiveSheet.Cells(i, 14).Value
            firstag = ActiveSheet.Cells(i, 15).Value
            firstagday = ActiveSheet.Cells(i, 16).Value
            assesc = ActiveSheet.Cells(i, 17).Value
            implcom = ActiveSheet.Cells(i, 18).Value
            affProj = ActiveSheet.Cells(i, 19).Value
            nojobs = ActiveSheet.Cells(i, 20).Value
            mat = ActiveSheet.Cells(i, 21).Value
            cost = ActiveSheet.Cells(i, 22).Value
            costall = ActiveSheet.Cells(i, 23).Value
            leadt = ActiveSheet.Cells(i, 24).Value
            impldecp = ActiveSheet.Cells(i, 25).Value
            impldecpdate = ActiveSheet.Cells(i, 26).Value
            impldecw = ActiveSheet.Cells(i, 27).Value
            impldecwdate = ActiveSheet.Cells(i, 28).Value
            impldecs = ActiveSheet.Cells(i, 29).Value
            impltype = ActiveSheet.Cells(i, 30).Value
            perm = ActiveSheet.Cells(i, 31).Value
            permduedate = ActiveSheet.Cells(i, 32).Value
            ECOdate = ActiveSheet.Cells(i, 33).Value
            agenda = ActiveSheet.Cells(i, 34).Value
            openp = ActiveSheet.Cells(i, 35).Value
            escalate = ActiveSheet.Cells(i, 36).Value
            feedback = ActiveSheet.Cells(i, 37).Value
            comments = ActiveSheet.Cells(i, 38).Value
            status = ActiveSheet.Cells(i, 39).Value
'date format accepted by MSSQL database is yyyy-mm-dd with the format function we get it.
            impldecwdate = Format(checkValueDate(impldecwdate), "yyyy-mm-dd")
            impldecpdate = Format(checkValueDate(impldecpdate), "yyyy-mm-dd")
            firstagday = Format(checkValueDate(firstagday), "yyyy-mm-dd")
            reldate = Format(checkValueDate(reldate), "yyyy-mm-dd")

'define query including the variables 
            strQuery = "INSERT INTO [NORDEXAG\DeanG].[TN_Committee] ([Type of change],[ECO/DECO] ,[Change Nr#] ,[Rev] ,[Release date] , [Change description]  ,[Impact] ,[Origin type] ,[Origin] ,[Engineering Responsible],[Applicability] ,[ECR/AST],[First agenda] ,[First agenda day] ,[Assessment escalation],[Implementation Committee],[Affected Projects],[Number of Jobs] ,[Material Needed],[Costs],[Cost allocation] ,[Lead time] ,[Implementation decision production],[Prod implementation decision Date] ,[Implementation decision windfarm] ,[WF implementation decision Date] ,[Implementation decision service] ,[Implementation Type] ,[Permanent solution needed] ,[Due date for permanent solution] ,[ECO Release date] ,[Agenda follow up] ,[Open points] ,[Escalation], [Feedback needed from] ,[Comments] ,[Status]) VALUES ('" & TNtype & "', '" & ECO & "'  , '" & change & "', '" & rev & "', '" & reldate & _
            "', '" & desc & "' , '" & impc & "' , '" & oritypeimpl & "' , '" & origin & _
            "', '" & eng & "' , '" & app & "' , '" & ecr & _
            "' , '" & firstag & "' ,'" & firstagday & "' , '" & assesc & _
            "'  ,'" & implcom & "' ,'" & affProj & "'      , '" & nojobs & _
            "' , '" & mat & "' , '" & cost & "'      , '" & costall & "' , '" & leadt & _
            "' , '" & impldecp & "' , '" & impldecpdate & _
            "' , '" & impldecw & "' , '" & impldecwdate & _
            "' , '" & impldecs & "' , '" & impltype & _
            "', '" & perm & "' , '" & permduedate & "' , '" & ECOdate & _
            "' ,'" & agenda & "' , '" & openp & "' , '" & escalate & _
            "' , '" & feedback & "' , '" & comments & "', '" & status & "' )"
            
            
'below I have write a console log of the query => CTRL+R to launch the inmediate screen and see the result. it is used to copy the query with the variables updated and paste it in SQL query to find the errors faster
            
            Debug.Print strQuery
    
            Set rsData = New ADODB.Recordset
            rsData.Open strQuery, Conn
        
        Next i
        
    End If

End Sub
