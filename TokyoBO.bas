Attribute VB_Name = "TokyoBO"
Sub main()
        '�R�U�T�~�T���P�W�Q�T��
        For n = 0 To 1825
                 ���� (n)
                ���� (n)
        Next n
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

