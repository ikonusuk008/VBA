Attribute VB_Name = "Tools1"
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
Sub One_row_delete()
Attribute One_row_delete.VB_ProcData.VB_Invoke_Func = "p\n14"
'
ActiveCell.EntireRow.Select

Rows(ActiveCell.Row).Delete
  
End Sub
Sub One_row_insert()
Attribute One_row_insert.VB_ProcData.VB_Invoke_Func = "i\n14"
'
ActiveCell.EntireRow.Select

Rows(ActiveCell.Row).Insert
  
End Sub
Sub One_Columun_delete()
Attribute One_Columun_delete.VB_ProcData.VB_Invoke_Func = "P\n14"
'
ActiveCell.EntireColumn.Select

Columns(ActiveCell.Column).Delete
  
End Sub
Sub One_Columun_insert()
Attribute One_Columun_insert.VB_ProcData.VB_Invoke_Func = "I\n14"
'
ActiveCell.EntireColumn.Select

Columns(ActiveCell.Column).Insert
  
End Sub
Sub �s�̍�����������_Automatic_row_height_adjustment()
Attribute �s�̍�����������_Automatic_row_height_adjustment.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' Automatic_row_height_adjustment
'
    Cells.Select
    Cells.EntireRow.AutoFit
    Range("A1").Select
End Sub
Sub �A�N�e�B�u�Z���̉��s�Ƌ󔒂�r������_Eliminate_line_breaks_and_blanks_in_the_active_Cell()
Attribute �A�N�e�B�u�Z���̉��s�Ƌ󔒂�r������_Eliminate_line_breaks_and_blanks_in_the_active_Cell.VB_ProcData.VB_Invoke_Func = "Q\n14"
  
    a = ActiveCell
    b = Replace(a, vbLf, "")
    b = Replace(b, " ", "")
    
    ActiveCell = b
      
End Sub

'Private Sub Worksheet_Activate()
'
'    Dim last_row As Long
'    last_row = ThisWorkbook.Worksheets("TIME").Cells(Rows.Count, 1).End(xlUp).Row
'    ThisWorkbook.Worksheets("TIME").Range("A" & last_row + 1).Value = ActiveSheet.name
'    ThisWorkbook.Worksheets("TIME").Range("B" & last_row + 1).Value = Format(Time, "hh:mm")
'    ThisWorkbook.Worksheets("TIME").Range("C" & last_row + 1).Value = DateDiff("n", Format(ThisWorkbook.Worksheets("TIME").Range("B" & last_row).Value, "hh:mm"), Format(Time, "hh:mm"))
'
'End Sub





