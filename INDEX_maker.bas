Attribute VB_Name = "INDEX_maker"
Sub INDEX_maker_main()
    'Summary
    'This program creates a table of contents sheet.
   
    Dim sheetCount As Integer
    Dim j As Integer
    Dim ASCII As Integer
    Dim worksheetName As String
    Dim subAddress_ As String
    Dim worksheet As worksheet

    ASCII = 65
    j = 2
       
    Application.ScreenUpdating = False
    
    If ExistsWorksheet("INDEX") Then
    
        Worksheets("INDEX").Select
        Application.DisplayAlerts = False
        ActiveSheet.Delete
        Application.DisplayAlerts = True
        
    End If
    
    Worksheets.Add
    ActiveSheet.name = "INDEX"
    Worksheets("INDEX").Activate
    
    Rows("1:30").Select
    Selection.Delete Shift:=xlUp
    Range("A2").Select
    sheetCount = Worksheets.Count
    Range("A1").Activate
    ActiveCell.Value = sheetCount & "Sheet"
    

    'Get sheet information and write to INDEX sheet.
    For i = 2 To sheetCount
    
        worksheetName = Worksheets(i).name
        
        If Mid(worksheetName, 1, 1) = "-" Then
            j = 1
            ASCII = ASCII + 1
        End If
        
        Worksheets(i).Activate
        Dim b_row As String
        b_row = Range("A1").Value
        subAddress_ = "'" & worksheetName & "'" & "!" & ActiveCell.Address

        Worksheets("INDEX").Activate
        Range(Chr(ASCII) & j).Activate
        
        If ASCII < 122 Then
        
            ActiveSheet.Hyperlinks.Add Anchor:=Selection, Address:=ActiveWorkbook.FullName, SubAddress:=subAddress_
            ActiveCell.Value = worksheetName
            
            With Range(Chr(ASCII) & j)
                .Value = worksheetName
                .Value = Worksheets(i).name
                .Interior.colorIndex = Worksheets(i).Tab.colorIndex
                .Borders.LineStyle = xlContinuous
                .Font.Size = 12
                .EntireColumn.ColumnWidth = 30
            End With
            
        End If
        
        j = j + 1
        
    Next i
    
    Columns("B:BB").ColumnWidth = 10
    Columns("A:A").Select
    Columns("A:A").EntireColumn.AutoFit
    Range("A1").Select
    Cells.Select
    Cells.EntireColumn.AutoFit
    Range("A1").Select
    FreezePanes
    
End Sub

Sub FreezePanes()

    With ActiveWindow
        .SplitColumn = 0
        .SplitRow = 1
    End With
    
    ActiveWindow.FreezePanes = True
    
End Sub
Public Function ExistsWorksheet(ByVal name As String)

    Dim ws As worksheet
    
    For Each ws In sheets
    
        If ws.name = name Then
            ExistsWorksheet = True
            Exit Function
        End If
        
    Next
    
    ExistsWorksheet = False
    
End Function



