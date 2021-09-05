Attribute VB_Name = "Image_resizing"
Sub Image_resizing_100()
Attribute Image_resizing_100.VB_ProcData.VB_Invoke_Func = "i\n14"
    Dim s As Long
    s = resizing(100)
    
    Debug.Print s

End Sub
Sub Image_resizing_450()
Attribute Image_resizing_450.VB_ProcData.VB_Invoke_Func = "u\n14"
    Dim s As Long
    s = resizing(450)
    
    Debug.Print s

End Sub

Function resizing(ByVal a As Long) As Long

'
    With ActiveSheet.Shapes(Selection.ShapeRange.ZOrderPosition)
        .LockAspectRatio = True
        .Width = a
    End With
    
    Selection.ShapeRange.ZOrder msoBringToFront
    
    resizing = a

End Function


