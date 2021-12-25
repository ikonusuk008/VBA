Attribute VB_Name = "Tools"
Sub insert_row_on_active_cell()
Attribute insert_row_on_active_cell.VB_ProcData.VB_Invoke_Func = "y\n14"
    '日付を付けた行を追加する。
    '　毎日作業のシートを対象
    
    For i = 1 To 10
        
        If Range("a" & i).Value = "日付" Or _
        Range("a" & i).Value = "日付" _
 _
        Then
        
            Range("a" & i + 1).Select
            'アクティブセルの上に行を追加する。
            Rows(ActiveCell.Row).Insert
            '追加した行を選択する
            ActiveCell.EntireRow.Select
            
            
            ActiveCell.EntireRow.Font.Color = vbBlack '黒
            
            ActiveCell.EntireRow.HorizontalAlignment = xlLeft
            ActiveCell.EntireRow.VerticalAlignment = xlTop
            ActiveCell.EntireRow.WrapText = True
            ActiveCell.EntireRow.Font.Bold = False
            
            
            
            
            '選択した行の色をクリア。
            ActiveCell.EntireRow.Interior.colorIndex = 0
            '選択したセルに　本日の
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
Sub 行の高さ自動調整_Automatic_row_height_adjustment()
Attribute 行の高さ自動調整_Automatic_row_height_adjustment.VB_ProcData.VB_Invoke_Func = "t\n14"
'
' Automatic_row_height_adjustment
'
    Cells.Select
    Cells.EntireRow.AutoFit
    Range("A1").Select
End Sub
Sub アクティブセルの改行と空白を排除する_Eliminate_line_breaks_and_blanks_in_the_active_Cell()
Attribute アクティブセルの改行と空白を排除する_Eliminate_line_breaks_and_blanks_in_the_active_Cell.VB_ProcData.VB_Invoke_Func = "Q\n14"
  
    a = ActiveCell
    b = Replace(a, vbLf, "")
    b = Replace(b, " ", "")
    
    ActiveCell = b
      
End Sub
Sub テンプレートシートをコピーして当日のシートを作成する()
'
' テンプレートシートをコピーして当日のシートを作成する Macro
' Copy the template sheet to create the sheet for the day.

    sheets("T").Select
    sheets("T").Copy Before:=sheets(4)
    sheets("T (2)").Select
    sheets("T (2)").name = Format(Date, "（mmdd")
    
End Sub



