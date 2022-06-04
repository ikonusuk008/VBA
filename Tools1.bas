Attribute VB_Name = "Tools1"
Sub Insert_todays_row_in_the_date_column()
Attribute Insert_todays_row_in_the_date_column.VB_ProcData.VB_Invoke_Func = "y\n14"

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
            '�I�������Z���Ɂ@�{���̓��t�����
            ActiveCell.Value = Date
            ActiveCell.End(xlToLeft).Select
            
            ActiveCell.EntireRow.RowHeight = 80
            
            Columns("A:A").EntireColumn.AutoFit
            
            ActiveSheet.Range("b2").Select
            
            
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
Sub �V�[�g�̌o�ߎ��Ԃ�0���Ȃ�΍s���폜����()
Attribute �V�[�g�̌o�ߎ��Ԃ�0���Ȃ�΍s���폜����.VB_ProcData.VB_Invoke_Func = " \n14"
    '��ƒ�
    
    
End Sub
Sub �A�N�e�B�u�Z���̉��̃Z�����R�s�[����()
Attribute �A�N�e�B�u�Z���̉��̃Z�����R�s�[����.VB_ProcData.VB_Invoke_Func = "q\n14"
    
    ActiveCell.Cells(1, 1).Value = ActiveCell.Cells(2, 1).Value
   
    
End Sub
'-------------------------------------------
Sub �V���[�g�J�b�g�ꗗ�̎擾_GetShortCutKeys()
'���݂̐ݒ�́A���u�b�N�Ɍ���

'INDEX_maker_main() : Ctrl +�@u
'seve_Time_series_main() : Ctrl +�@b
'ToIndex_main() : Ctrl +�@l
'extract_13H_between_0300to1500() : Ctrl +
'Get_final_balance() : Ctrl +�@q
'Scraiping2() : Ctrl +�@j
'insert_row_on_active_cell() : Ctrl +�@y
'One_row_delete() : Ctrl +�@p
'One_row_insert() : Ctrl +�@i
'One_Columun_delete() : Ctrl +�@P
'One_Columun_insert() : Ctrl +�@I
'�s�̍�����������_Automatic_row_height_adjustment() : Ctrl +�@t
'�A�N�e�B�u�Z���̉��s�Ƌ󔒂�r������_Eliminate_line_breaks_and_blanks_in_the_active_Cell() : Ctrl +�@Q
'�V�[�g�̌o�ߎ��Ԃ�0���Ȃ�΍s���폜����() : Ctrl +
'�A�N�e�B�u�Z���̉��̃Z�����R�s�[����() : Ctrl +�@q

Dim DefPath As String
Dim FNo As Integer
Dim LineBuf As String
Dim i As Integer
Dim buf() As String
Dim bufName As String
Dim bufKeyName As String
Dim vbc As Object
Const AT1 As String = "Attribute "
Const AT2 As String = "VB_Invoke_Func ="
Const TMPF As String = "Temp1.bas"

DefPath = ThisWorkbook.path & "\"
  With ThisWorkbook.VBProject
  For Each vbc In .VBComponents
  .VBComponents(vbc.name).Export filename:=DefPath & TMPF
  FNo = FreeFile()
  Open DefPath & TMPF For Input As #FNo
  While Not EOF(FNo)
    Line Input #FNo, LineBuf
    If InStr(1, LineBuf, "Sub", vbTextCompare) = 1 Then
      bufName = Mid$(LineBuf, InStr(LineBuf, "Sub") + 4)
    End If
    If InStr(LineBuf, AT1) = 1 And InStr(LineBuf, AT2) > 0 Then
     ReDim Preserve buf(i)
      bufKeyName = " : Ctrl +�@" & Mid$(LineBuf, InStrRev(LineBuf, "=") + 3, 1)
      buf(i) = bufName & bufKeyName  '�z��o��
   
      'Debug.Print��"
      'Debug.Print bufName; bufKeyName
      i = i + 1
      bufName = ""
    End If
    LineBuf = ""
  Wend
  Close #FNo
  Kill DefPath & TMPF
  Next
  End With
  Debug.Print Join(buf, vbCrLf)
End Sub






















