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



Sub Macro7()
Attribute Macro7.VB_ProcData.VB_Invoke_Func = " \n14"
'
' Macro7 Macro
'

'
    
End Sub
