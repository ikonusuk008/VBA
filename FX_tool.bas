Attribute VB_Name = "FX_tool"
Sub FX_tool_main()

        extract_2200_of_13H

        ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����
        
        Debug.Print "Start Input BUY and SELL"

        '�R�U�T�~�T���P�W�Q�T��
        '���ōs�r�������A�s���������Ώې��ƂȂ�B
        For n = 0 To 1829
        
                ���� (n)
                ���� (n)
        Next n
        
        Debug.Print "End Input BUY and SELL"
        
End Sub
Function ����(ByVal n As Long)

        Dim Tokyo_market_high_price As Double
        Dim Europe_1_hour_value As Double
        Dim European_closing_price As Double
        
                        
        Dim t_13_hours_a_day As Long
        t_13_hours_a_day = 13
        
        Dim Days_index As Long
        Days_index = (n * t_13_hours_a_day)
        
        Dim Settlement_time_sequence As Long
        Settlement_time_sequence = 13

       
        
        European_closing_price = Range("f" & (Settlement_time_sequence + Days_index)).Value
        
        Dim Break_judgment_value As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        Break_judgment_value = 1  '���Z�b�g
       
        Tokyo_market_high_price = WorksheetFunction.Max(Range("d" & 1 + Days_index & ":d" & 6 + Days_index))
        
        
        '�����s��@���l�@�u���[�N
        For i = 1 + Days_index To Settlement_time_sequence + Days_index    '15������Q�Q���܂�
            
                Europe_1_hour_value = CDbl(Range("f" & i).Value)
        
                If Europe_1_hour_value > Tokyo_market_high_price Then
                        Break_judgment_value = 2
                        Exit For
                End If
        Next i
        
        '�u���[�N��̑��؂蔻��
        For i2 = i To Settlement_time_sequence + Days_index
                Europe_1_hour_value = CDbl(Range("f" & i2).Value)
        
                If ((Europe_1_hour_value - Tokyo_market_high_price) * 100) < -30 Then
                    
                        Break_judgment_value = 3
                        
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If Break_judgment_value = 1 Then
                '�g���[�h����
                Range("g" & Settlement_time_sequence + Days_index).Value = 0
        ElseIf Break_judgment_value = 2 Then
                '�u���[�N��̌���
                Range("g" & Settlement_time_sequence + Days_index).Value = (European_closing_price - Tokyo_market_high_price) * 100
        ElseIf Break_judgment_value = 3 Then
                '�u���[�N��̑��؂�
                Range("g" & Settlement_time_sequence + Days_index).Value = -30
        End If

End Function
Function ����(ByVal n As Long)

        Dim Tokyo_market_low As Double
        Dim Europe_1_hour_value As Double
        Dim European_closing_price As Double
        
        
        Dim t_13_hours_a_day As Long
        t_13_hours_a_day = 13
        
        Dim Days_index As Long
        Days_index = (n * t_13_hours_a_day)
        
        Dim Settlement_time_sequence As Long
        Settlement_time_sequence = 13

        
        
        European_closing_price = Range("f" & (Settlement_time_sequence + Days_index)).Value
        
        Dim Break_judgment_value As Integer  '�i�P�F�u���[�N�Ȃ��B�Q�F�u���[�N����B�R�F�u���[�N���؁j
        
        Break_judgment_value = 1  '���Z�b�g
       
        Tokyo_market_low = WorksheetFunction.Min(Range("e" & 1 + Days_index & ":e" & 6 + Days_index))
        
        '�����s��@���l�@�u���[�N
        For i = 1 + Days_index To Settlement_time_sequence + Days_index
            
                Europe_1_hour_value = CDbl(Range("f" & i).Value)
        
                
                If Europe_1_hour_value < Tokyo_market_low Then
                        Break_judgment_value = 2
                        Exit For
                End If
        Next i
        
        '�����s��@���l�@���؂�
        For i2 = i To Settlement_time_sequence + Days_index   '�u���[�N��̑��؂蔻��
                Europe_1_hour_value = CDbl(Range("f" & i2).Value)
        
                If ((Tokyo_market_low - Europe_1_hour_value) * 100) < -30 Then
                        Break_judgment_value = 3
                        Exit For
                End If
        Next i2
        
        '���茋�ʂ��Q�Q����G��ɏ�������
        If Break_judgment_value = 1 Then
                '�g���[�h����
                Range("h" & Settlement_time_sequence + Days_index).Value = 0
        ElseIf Break_judgment_value = 2 Then
                '�u���[�N��̌���
                Range("h" & Settlement_time_sequence + Days_index).Value = (Tokyo_market_low - European_closing_price) * 100
        ElseIf Break_judgment_value = 3 Then
                '�u���[�N��̑��؂�
                Range("h" & Settlement_time_sequence + Days_index).Value = -30
        End If

End Function
Sub ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����()

        Debug.Print "END ����13�s���܂܂�Ă��Ȃ���Τ���̓��͑ΏۊO�Ƃ��č폜����"

        Dim Last_line As Long
        Last_line = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim Count_the_number_of_hours_of_the_day As Long
        Count_the_number_of_hours_of_the_day = 0
        
        Dim BREAK_FLAG As String
        BREAK_FLAG = 0
        
        Dim i As Long
        
        For EVERYDAY = 0 To 100
        
                If BREAK_FLAG = 1 Then
                        Exit For
                End If
             
                For i = 1 To Last_line
                
                        If Range("a" & i).Value = "" Then
                                BREAK_FLAG = 1
                                Exit For
                        End If
               
                         Count_the_number_of_hours_of_the_day = Count_the_number_of_hours_of_the_day + 1
        
                        If Range("a" & i).Value <> Range("a" & i + 1).Value Then
                               
                               If Count_the_number_of_hours_of_the_day = 13 Then
                                        Count_the_number_of_hours_of_the_day = 0
                                        
                               Else
                                        '��̍s���J�E���g���������폜����B
                                        Rows(i + 1 - Count_the_number_of_hours_of_the_day & ":" & CStr(i)).Delete
                                        
                                        Count_the_number_of_hours_of_the_day = 0
                                        
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
Macro2

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


   Columns("A:A").Select
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
Sub Macro2()
'
' Macro2 Macro
'

'
    Columns("G:G").Select
    Selection.ClearContents
    Range("G1").Select
End Sub
