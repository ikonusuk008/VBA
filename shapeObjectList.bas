Attribute VB_Name = "shapeObjectList"
Sub accelerate()

    With Application
        .ScreenUpdating = False
        .DisplayAlerts = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
    End With

End Sub
Sub clearAccelerate()

    With Application
        .ScreenUpdating = True
        .DisplayAlerts = True
        .EnableEvents = True
        .Calculation = xlCalculationAutomatic
    End With

End Sub
Sub shapeObjectListMain()

On Error GoTo Err

    Dim sheets As Worksheet, i As Integer, m As Integer, O As Shape, thisSheet As Worksheet
    Set thisSheet = ActiveSheet
   
    Dim rng As Range
    Set rng = Application.InputBox(prompt:="Choose the cell to open the table", Type:=8)
    
    Dim n As Integer

    i = rng.Row
    m = rng.Column

    Call accelerate

Done:

    With thisSheet
        .Cells(i, m).Value = "Sheet Name"
        .Cells(i, m + 1).Value = "Object Name"
        .Cells(i, m + 2).Value = "Object Texts"
    End With
    
    With thisSheet.Range(Cells(i, m), Cells(i, m + 2))
        .Interior.colorIndex = 48
        .Font.colorIndex = 2
        .Font.Bold = True
    End With
    
    For Each sheets In ActiveWorkbook.Worksheets
        If sheets.Shapes.Count > 0 Then
            For Each O In sheets.Shapes
                thisSheet.Cells(i + 1, m).Value = sheets.name
                thisSheet.Cells(i + 1, m + 1).Value = O.name
                If O.name Like "Comment*" Then
                    thisSheet.Cells(i + 1, m + 2).Value = "-"
                Else
                    If O.TextFrame2.HasText Then
                        thisSheet.Cells(i + 1, m + 2).Value = O.TextFrame2.TextRange.Text
                    Else
                        thisSheet.Cells(i + 1, m + 2).Value = "-"
                    End If
                End If
                i = i + 1
            Next O
        End If
    Next sheets

    For n = 1 To 3
        thisSheet.Columns(m).EntireColumn.AutoFit
        m = m + 1
    Next n
    
    Call clearAccelerate
    Exit Sub

Err:
End Sub
