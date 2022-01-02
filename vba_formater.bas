Attribute VB_Name = "vba_formater"
Sub vba_formater_()

    'http://thom.hateblo.jp/entry/2015/11/21/114328
    'C:\WINNT(または Windows)\system32\FM20.DLL
    
    
    Dim CB As New DataObject
    CB.GetFromClipboard
        
    If Not CB.GetFormat(1) Then
        MsgBox "クリップボードが空です。", vbExclamation
        Exit Sub
    End If
        
    Dim Lines() As String: Lines = split(CB.GetText, vbNewLine)
        
    For Each x In Lines
        
        x2 = ConvertString(x)
                
        If InStr(1, x2, "End Sub") > 0 Then
            t = 0
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Exit Sub") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Go Sub") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Sub ") > 0 Then
            Debug.Print
            Debug.Print String(t, vbTab) & Trim(x)
            t = 1
        ElseIf InStr(1, x2, "End Function") > 0 Then
            t = 0
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Function ") > 0 Then
            Debug.Print
            Debug.Print String(t, vbTab) & Trim(x)
            t = 1
        ElseIf InStr(1, x2, "For ") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
            t = t + 1
        ElseIf InStr(1, x2, "Next") > 0 Then
            t = t - 1
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Loop") > 0 Then
            t = t - 1
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Do") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
            t = t + 1
        ElseIf InStr(1, x2, "End With") > 0 Then
            t = t - 1
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "With") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
            t = t + 1
        ElseIf InStr(1, x2, "End If") > 0 Then
            t = t - 1
            Debug.Print String(t, vbTab) & Trim(x)
        ElseIf InStr(1, x2, "Else") > 0 Then
            t = t - 1
            Debug.Print String(t, vbTab) & Trim(x)
            t = t + 1
        ElseIf InStr(1, x2, "If ") > 0 Then
            Debug.Print String(t, vbTab) & Trim(x)
            t = t + 1
        ElseIf Trim(x) <> "" Then
            Debug.Print String(t, vbTab) & Trim(x)
        End If
    Next
        
End Sub

Function ConvertString(ByVal x As String) As String

    ConvertString = x
    
    If InStr(1, x, """") > 0 Then
    
        If InStr(1, x, "'") > 0 Then
            If InStr(1, x, """") < InStr(1, x, "'") Then
                ConvertString = InnerConvertString(x)
            End If
        Else
            ConvertString = InnerConvertString(x)
        End If
        
    End If
    
End Function

Private Function InnerConvertString(x As String) As String

    Dim newstr As String
    Dim IsString As Boolean
    
    For i = 1 To Len(x)
    
        IsString = (Mid(x, i, 1) = """") Xor IsString
        
        If Not IsString Then
            newstr = newstr & Mid(x, i, 1)
        End If
        
    Next
    
    InnerConvertString = newstr
    
End Function
