Attribute VB_Name = "ChangeColor"
Sub ChangeColor_main()

    Dim rng As Range
    
    Dim ptr As Integer
    
    Const dataType As Long = 23 '23:When "Numeric value", "Character", "Logical value", and "Error value" are all selected
    
    Dim colorChangeText As String
    colorChangeText = InputBox("text")
    
    Dim colorIndex As String
    colorIndex = InputBox("imput color", "3", "3")
    
    For Each rng In ActiveSheet.Cells.SpecialCells(xlCellTypeConstants, dataType)
    
        ptr = InStr(rng.Value, colorChangeText)
        
        If ptr > 0 Then
            rng.Characters(Start:=ptr, Length:=Len(colorChangeText)).Font.colorIndex = CInt(colorIndex)
        End If
    
    Next rng

End Sub

