Attribute VB_Name = "FX_tool"
Sub main()

        extract_2200_of_13H

        ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����
        
        Debug.Print "Start Input BUY and SELL"

        '�R�U�T�~�T���P�W�Q�T��
        For n = 0 To 1829
        
                ���� (n)
                ���� (n)
                
        Next n
        
        Debug.Print "End Input BUY and SELL"
        
End Sub
Function ����(ByVal n As Long)

        Dim �����s�ꍂ�l As Double
        Dim ���B1���Ԓl As Double
        Dim ���B�I�l As Double
        
                        
        Dim ��24���� As Long
        ��24���� = 13
        
        Dim �����C���f�b�N�X As Long
        �����C���f�b�N�X = (n * ��24����)
        
        Dim ���ώ����� As Long
        ���ώ����� = 13

       
        
        ���B�I�l = Range("f" & (���ώ����� + �����C���f�b�N�X)).Value
        
        Dim �u���[�N����l As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        �u���[�N����l = 1  '���Z�b�g
       
        �����s�ꍂ�l = WorksheetFunction.Max(Range("d" & 1 + �����C���f�b�N�X & ":d" & 6 + �����C���f�b�N�X))
        
        
        '�����s��@���l�@�u���[�N
        For i = 1 + �����C���f�b�N�X To ���ώ����� + �����C���f�b�N�X    '15������Q�Q���܂�
            
                ���B1���Ԓl = CDbl(Range("f" & i).Value)
        
                If ���B1���Ԓl > �����s�ꍂ�l Then
                        �u���[�N����l = 2
                        Exit For
                End If
        Next i
        
        '�u���[�N��̑��؂蔻��
        For i2 = i To ���ώ����� + �����C���f�b�N�X
                ���B1���Ԓl = CDbl(Range("f" & i2).Value)
        
                If ((���B1���Ԓl - �����s�ꍂ�l) * 100) < -30 Then
                    
                        �u���[�N����l = 3
                        
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If �u���[�N����l = 1 Then
                '�g���[�h����
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = 0
        ElseIf �u���[�N����l = 2 Then
                '�u���[�N��̌���
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = (���B�I�l - �����s�ꍂ�l) * 100
        ElseIf �u���[�N����l = 3 Then
                '�u���[�N��̑��؂�
                Range("g" & ���ώ����� + �����C���f�b�N�X).Value = -30
        End If

End Function
Function ����(ByVal n As Long)

        Dim �����s����l As Double
        Dim ���B1���Ԓl As Double
        Dim ���B�I�l As Double
        
        
        Dim ��24���� As Long
        ��24���� = 13
        
        Dim �����C���f�b�N�X As Long
        �����C���f�b�N�X = (n * ��24����)
        
        Dim ���ώ����� As Long
        ���ώ����� = 13

        
        
        ���B�I�l = Range("f" & (���ώ����� + �����C���f�b�N�X)).Value
        
        Dim �u���[�N����l As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        
        �u���[�N����l = 1  '���Z�b�g
       
        �����s����l = WorksheetFunction.Min(Range("e" & 1 + �����C���f�b�N�X & ":e" & 6 + �����C���f�b�N�X))
        
        '�����s��@���l�@�u���[�N
        For i = 1 + �����C���f�b�N�X To ���ώ����� + �����C���f�b�N�X
            
                ���B1���Ԓl = CDbl(Range("f" & i).Value)
        
                
                If ���B1���Ԓl < �����s����l Then
                        �u���[�N����l = 2
                        Exit For
                End If
        Next i
        
        '�����s��@���l�@���؂�
        For i2 = i To ���ώ����� + �����C���f�b�N�X   '�u���[�N��̑��؂蔻��
                ���B1���Ԓl = CDbl(Range("f" & i2).Value)
        
                If ((�����s����l - ���B1���Ԓl) * 100) < -30 Then
                        �u���[�N����l = 3
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If �u���[�N����l = 1 Then
                '�g���[�h����
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = 0
        ElseIf �u���[�N����l = 2 Then
                '�u���[�N��̌���
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = (�����s����l - ���B�I�l) * 100
        ElseIf �u���[�N����l = 3 Then
                '�u���[�N��̑��؂�
                Range("h" & ���ώ����� + �����C���f�b�N�X).Value = -30
        End If

End Function
Sub ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����()

        Debug.Print "END ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����"

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
       
       Debug.Print "END ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����"
       
      
End Sub

Sub extract_2200_of_13H()
Attribute extract_2200_of_13H.VB_ProcData.VB_Invoke_Func = " \n14"
'
' extract_13_of Macro
'

'
Macro1

    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$H$44000").AutoFilter Field:=2, Criteria1:="15:00"
    Range("G104").Select
    Selection.End(xlDown).Select
    ActiveWindow.SmallScroll Down:=9
    Range("H23751").Select
    ActiveWindow.SmallScroll Down:=-129
 
End Sub
Sub Macro1()
'
' Macro1 Macro
'

'
    Cells.Select
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Select
    Selection.AutoFilter
    ActiveSheet.Range("$A$1:$G$44000").AutoFilter Field:=2, Criteria1:=Array( _
        "0:00", "1:00", "16:00", "17:00", "18:00", "19:00", "2:00", "20:00", "21:00", "22:00", _
        "23:00"), Operator:=xlFilterValues
    ActiveWindow.SmallScroll Down:=-18
    Rows("1:1").Select
    Range(Selection, Selection.End(xlDown)).Select
    Selection.Delete Shift:=xlUp
    ActiveWindow.SmallScroll Down:=-210
    Range("A1").Select
End Sub

