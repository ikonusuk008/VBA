Attribute VB_Name = "zoom"
Sub zoom_main()

    Dim buf As String
    buf = InputBox("imput color", "80", "80")

    For i = 1 To Worksheets.Count
    
        Worksheets(i).Select
        ActiveWindow.zoom = buf
        Range("a1").Select
        previewFlag = 1
        
        If previewFlag = 1 Then
            ActiveWindow.View = xlNormalView
        Else
            ActiveWindow.View = xlPageBreakPreview
        End If
         
        Cells.Select
    
        Cells.Font.name = "ÉÅÉCÉäÉI"
        
        
    Next i
    
    Worksheets(1).Select
    
End Sub
