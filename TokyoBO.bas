Attribute VB_Name = "TokyoBO"
Sub main()
        '����24�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����
        
        For n = 0 To 5000
                 ���� (n)
                ���� (n)
        Next n
End Sub
Function ����(ByVal n As Long)

        Dim �����s�ꍂ�l As Double '
        Dim ���B1���Ԓl As Double   '
        Dim ���B�I�l As Double    '
        Dim �����C���f�b�N�X As Long '
        Dim ���ώ����� As Long '���ώ�����
        Dim ��24���� As Long
        ��24���� = 24
        
        ���ώ����� = 22
        
        �����C���f�b�N�X = (n * ��24����)
        
        ���B�I�l = Range("f" & (���ώ����� + �����C���f�b�N�X)).Value
        
        Dim buyFlag As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        buyFlag = 1
       
        �����s�ꍂ�l = WorksheetFunction.Max(Range("d" & 10 + �����C���f�b�N�X & ":d" & 15 + �����C���f�b�N�X))     '�����s��@���l
        
        For i = 16 + �����C���f�b�N�X To ���ώ����� + �����C���f�b�N�X    '15������Q�Q���܂�
            
                ���B1���Ԓl = CDbl(Range("f" & i).Value)
        
                If ���B1���Ԓl > �����s�ꍂ�l Then  '�����s��@���l�@�u���[�N
                        buyFlag = 2
                        Exit For
                End If
        Next i
        
        For i2 = i To ���ώ����� + �����C���f�b�N�X   '�u���[�N��̑��؂蔻��
                ���B1���Ԓl = CDbl(Range("f" & i2).Value)
        
                If ((���B1���Ԓl - �����s�ꍂ�l) * 100) < -30 Then
                    
                        buyFlag = 3
                        
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If buyFlag = 1 Then
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = 0
        ElseIf buyFlag = 2 Then
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = (���B�I�l - �����s�ꍂ�l) * 100
        ElseIf buyFlag = 3 Then
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = -30
        End If

End Function
Function ����(ByVal n As Long)

        Dim min9_15 As Double '�����s��@min
        Dim ���B1���Ԓl As Double   '���B1���Ԓl
        Dim ���B�I�l As Double    '���B�I�l
        Dim �����C���f�b�N�X As Long '�����C���f�b�N�X
        Dim ���ώ����� As Long '���ώ�����
        Dim ��24���� As Long
        
        ��24���� = 24
        ���ώ����� = 22
        
        �����C���f�b�N�X = (n * ��24����)
        
        ���B�I�l = Range("f" & (���ώ����� + �����C���f�b�N�X)).Value
        
        Dim buyFlag As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        
        buyFlag = 1
       
        min9_15 = WorksheetFunction.Min(Range("e" & 10 + �����C���f�b�N�X & ":e" & 15 + �����C���f�b�N�X))     '�����s��@���l
        
        For i = 16 + �����C���f�b�N�X To ���ώ����� + �����C���f�b�N�X    '15������Q�Q���܂�
            
                ���B1���Ԓl = CDbl(Range("f" & i).Value)
        
                If ���B1���Ԓl < min9_15 Then  '�����s��@���l�@�u���[�N
                        buyFlag = 2
                        Exit For
                End If
        Next i
        
        For i2 = i To ���ώ����� + �����C���f�b�N�X   '�u���[�N��̑��؂蔻��
                ���B1���Ԓl = CDbl(Range("f" & i2).Value)
        
                If ((min9_15 - ���B1���Ԓl) * 100) < -30 Then
                        buyFlag = 3
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If buyFlag = 1 Then
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = 0
        ElseIf buyFlag = 2 Then
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = (min9_15 - ���B�I�l) * 100
        ElseIf buyFlag = 3 Then
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = -30
        End If

End Function
Sub ����24�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����()
        
        Dim �ŏI�s As Long
        
        �ŏI�s = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim ����̎��Ԑ��J�E���g As Long
        ����̎��Ԑ��J�E���g = 0
        Dim f As String
        f = 0
        
        Dim i As Long
        For h = 0 To 2000
        
                If f = 1 Then
                        Exit For
                End If
             
                For i = 1 To �ŏI�s
                
                        If Range("a" & i).Value = "" Then
                                f = 1
                                Exit For
                        End If
               
                         ����̎��Ԑ��J�E���g = ����̎��Ԑ��J�E���g + 1
        
                        If Range("a" & i).Value <> Range("a" & i + 1).Value Then
                               
                               If ����̎��Ԑ��J�E���g = 24 Then
                                        ����̎��Ԑ��J�E���g = 0
                               Else
                                        '��̍s���J�E���g���������폜����B
                                        Rows(i + 1 - ����̎��Ԑ��J�E���g & ":" & CStr(i)).Delete
                                        ����̎��Ԑ��J�E���g = 0
                                        Exit For
                               End If
                        End If
                Next i
       Next h
                
                
End Sub




Sub ����14���Ԃ��܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����()
        
        Dim n As Long
        
        n = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 1 To n
        
                If (Format(Range("b" & i).Value, "hh:nn") = "09:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "10:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "12:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "11:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "13:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "12:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "14:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "13:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "15:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "14:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "16:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "15:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "17:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "16:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "18:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "17:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "19:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "18:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "20:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "19:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "21:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "22:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "22:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "22:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Then
                       
                       
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                        (Format(Range("b" & i).Value, "hh:nn") = "12:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "15:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Then
                       
                       
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                         Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                 
                 ElseIf (Format(Range("b" & i).Value, "hh:nn") = "19:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                        Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                         Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                         Range("b" & i + 1).EntireRow.Insert '�Z�������w�肷��ꍇ
                

                End If

                Debug.Print Format(Range("b" & i).Value, "hh:nn")
        Next i

End Sub

