Attribute VB_Name = "DeleteOther13DAY"
Sub ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����()
        
        Dim �ŏI�s As Long
        �ŏI�s = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ����̎��Ԑ��J�E���g As Long
        ����̎��Ԑ��J�E���g = 0
        
        Dim BREAK_FLAG As String
        BREAK_FLAG = 0
        
        Dim i As Long
        
        For EVERYDAY = 0 To 100
        
                If BREAK_FLAG = 1 Then
                        Exit For
                End If
             
                For i = 1 To �ŏI�s
                
                        If Range("a" & i).Value = "" Then
                                BREAK_FLAG = 1
                                Exit For
                        End If
               
                         ����̎��Ԑ��J�E���g = ����̎��Ԑ��J�E���g + 1
        
                        If Range("a" & i).Value <> Range("a" & i + 1).Value Then
                               
                               If ����̎��Ԑ��J�E���g = 13 Then
                                        ����̎��Ԑ��J�E���g = 0
                                        
                               Else
                                        '��̍s���J�E���g���������폜����B
                                        Rows(i + 1 - ����̎��Ԑ��J�E���g & ":" & CStr(i)).Delete
                                        
                                        ����̎��Ԑ��J�E���g = 0
                                        
                                        Debug.Print i
                                        
                                        Exit For
                               End If
                        End If
                Next i
       Next EVERYDAY
      
End Sub

