Attribute VB_Name = "Tools"
Sub insert_row_on_active_cell()
Attribute insert_row_on_active_cell.VB_ProcData.VB_Invoke_Func = "y\n14"
    '���t��t�����s��ǉ�����B
    '�@������Ƃ̃V�[�g��Ώ�
    
    For i = 1 To 10
        
        If Range("a" & i).Value = "���t" Or _
        Range("a" & i).Value = "���t" _
 _
        Then
        
            Range("a" & i + 1).Select
            '�A�N�e�B�u�Z���̏�ɍs��ǉ�����B
            Rows(ActiveCell.Row).Insert
            '�ǉ������s��I������
            ActiveCell.EntireRow.Select
            
            
            ActiveCell.EntireRow.Font.Color = vbBlack '��
            
            ActiveCell.EntireRow.HorizontalAlignment = xlLeft
            ActiveCell.EntireRow.VerticalAlignment = xlTop
            ActiveCell.EntireRow.WrapText = True
            ActiveCell.EntireRow.Font.Bold = False
            
            
            
            
            '�I�������s�̐F���N���A�B
            ActiveCell.EntireRow.Interior.colorIndex = 0
            '�I�������Z���Ɂ@�{����
            ActiveCell.Value = Date
            ActiveCell.End(xlToLeft).Select
            
            ActiveCell.EntireRow.RowHeight = 80
            
            Columns("A:A").EntireColumn.AutoFit
            
            Exit For
        End If
        
    Next i
    
End Sub



Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    
End Sub
