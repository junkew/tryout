Sub testRunner()
    Dim wb As Workbook
   
    Set wb = ActiveWorkbook
    runWorksheets wb
    
End Sub

Sub runWorksheets(wb As Workbook)
    Dim ws As Worksheet
    For Each ws In wb.Worksheets
        If ws.Range("A1").Value = "Testset sheet" Then
            runTests ws
        End If
    Next
End Sub

Sub runTests(ws As Worksheet)
    Dim c As Range
    Dim testStartRow As Long
    Dim testEndRow As Long
    
    Set c = ws.Columns(1)
    'Debug.Print ws.UsedRange.Address

    Set foundcell = c.Find(what:="Testcase")
    
    If Not foundcell Is Nothing Then
        FirstFound = foundcell.Address
       
        Do Until foundcell Is Nothing
            testStartRow = foundcell.Row
            
            Set foundcell = c.FindNext(after:=foundcell)
    
            If foundcell.Address = FirstFound Then
                testEndRow = ws.UsedRange.Rows.Count
                Set foundcell = Nothing
            Else
                testEndRow = foundcell.Row
            End If
            
            runTest Application.Range(ws.Cells(testStartRow, 1), ws.Cells(testEndRow, ws.UsedRange.Column))
            
        Loop
    End If
         
End Sub
Sub runTest(testRange As Range)
    Dim iRow As Long
    Dim parCount As Long
    ReDim parData(1, 0) As String
    
    Debug.Print "Running " & testRange.Address & testRange.Rows.Count & testRange.Columns.Count
    
    For iRow = 3 To testRange.Rows.Count Step 2
        ReDim parData(1, 0) As String
        parCount = 0
        
        While testRange.Cells(iRow, 3 + parCount) <> ""
            ReDim Preserve parData(1, parCount)
            parData(0, parCount) = testRange.Cells(iRow, 3 + parCount).Value
            parData(1, parCount) = testRange.Cells(iRow + 1, 3 + parCount).Value
            parCount = parCount + 1
        Wend
        
        runKeyword testRange.Cells(iRow + 1, 2), parData
        
    Next
End Sub
Sub runKeyword(kwName As String, kwData)
   ' Debug.Print kwName & " " & UBound(kwData, 2) & " parameters"
    On Error Resume Next
        Application.Run Replace(kwName, " ", "_"), kwData
    On Error GoTo 0
End Sub
Sub keyword_1(kwData)
    Debug.Print "keyword 1 with " & UBound(kwData, 2) & " parameters"
End Sub
