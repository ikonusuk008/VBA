Attribute VB_Name = "FX_tool"
Dim STOP_LOSS As Long
Sub FX_tool_main()

        'TODO　自動で開いて書き込む。開いている場合は、閉じて実行する。
        'Workbooks.Open "C:\Users\User\Google ドライブ\00-share\MT4\GBPJPY60.csv"

        extract_2200_of_13H '対象外の時間の行を削除する。

        日に13行が含まれていなければ､その日は対象外として削除する '対象時間が欠けている日の行は削除する｡
        
        Debug.Print "Start Input BUY and SELL"
            
        '最終行--------------------------------
        Dim xlLastRow As Long
        Dim end_r As Long
        xlLastRow = Cells(Rows.Count, 1).Row
        end_r = Cells(xlLastRow, 1).End(xlUp).Row
        '--------------------------------最終

        For n = 0 To end_r \ 13
                買い (n)
                売り (n)
        Next n
        
        Debug.Print "End Input BUY and SELL"
        
End Sub
Function 買い(ByVal n As Long)

        STOP_LOSS = -5

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
        
        Dim Break_judgment_value As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        Break_judgment_value = 1  'リセット
       
        Tokyo_market_high_price = WorksheetFunction.Max(Range("d" & 1 + Days_index & ":d" & 6 + Days_index))
        
        '東京市場の高値をブレークブレークする行を検索し、グラグ２（ブレークあり）を設定する。
        For i = 1 + Days_index To Settlement_time_sequence + Days_index    '15時から２２時まで
            
                Europe_1_hour_value = CDbl(Range("f" & i).Value)
        
                If Europe_1_hour_value > Tokyo_market_high_price Then
                
            
                        'ここで更に、東京市場高値で買いができた場合　という条件が必要。・・・①
                        'または、東京市場高値より、Xpips（最適解が必要）下がった時、という条件が必要。・・・②
                        'または、そのまま、東京市場高値を超えた終値から計算する。・・・③
                        
                        '調査
                        '１H終値のブレークが必要か。東京市場Xpipsブレークでいいのではないか。
                
                        Break_judgment_value = 2
                        Exit For
                End If
        Next i
        
        'ブレーク後、損切り判定を行い、Xpipsでフラグ３（ブレーク損切）を設定する。
        For i2 = i To Settlement_time_sequence + Days_index
                Europe_1_hour_value = CDbl(Range("f" & i2).Value)

                If ((Europe_1_hour_value - Tokyo_market_high_price) * 100) < STOP_LOSS Then
                    
                        Break_judgment_value = 3
                        
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If Break_judgment_value = 1 Then
                'トレード無し
                Range("g" & Settlement_time_sequence + Days_index).Value = 0
        ElseIf Break_judgment_value = 2 Then
        
                Dim a As Long
                
                a = Europe_1_hour_value - European_closing_price
        
        
                'ブレーク後の決済
                '
               'Range("g" & Settlement_time_sequence + Days_index).Value = (European_closing_price - Tokyo_market_high_price) * 100
                
                Range("g" & Settlement_time_sequence + Days_index).Value = (European_closing_price - Tokyo_market_high_price - a) * 100
                
        ElseIf Break_judgment_value = 3 Then
                'ブレーク後の損切り
                Range("g" & Settlement_time_sequence + Days_index).Value = STOP_LOSS
        End If

End Function
Function 売り(ByVal n As Long)

        STOP_LOSS = -5

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
        
        Dim Break_judgment_value As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        
        Break_judgment_value = 1  'リセット
       
        Tokyo_market_low = WorksheetFunction.Min(Range("e" & 1 + Days_index & ":e" & 6 + Days_index))
        
        '東京市場　安値　ブレーク
        For i = 1 + Days_index To Settlement_time_sequence + Days_index
            
                Europe_1_hour_value = CDbl(Range("f" & i).Value)
        
                
                If Europe_1_hour_value < Tokyo_market_low Then
                        Break_judgment_value = 2
                        Exit For
                End If
        Next i
        
        '東京市場　安値　損切り
        For i2 = i To Settlement_time_sequence + Days_index   'ブレーク後の損切り判定
                Europe_1_hour_value = CDbl(Range("f" & i2).Value)
        
                If ((Tokyo_market_low - Europe_1_hour_value) * 100) < STOP_LOSS Then
                        Break_judgment_value = 3
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If Break_judgment_value = 1 Then
                'トレード無し
                Range("h" & Settlement_time_sequence + Days_index).Value = 0
        ElseIf Break_judgment_value = 2 Then
                'ブレーク後の決済
                Range("h" & Settlement_time_sequence + Days_index).Value = (Tokyo_market_low - European_closing_price) * 100
        ElseIf Break_judgment_value = 3 Then
                'ブレーク後の損切り
                Range("h" & Settlement_time_sequence + Days_index).Value = STOP_LOSS
        End If

End Function
Sub 日に13行が含まれていなければ､その日は対象外として削除する()

        Debug.Print "END 日に13行が含まれていなければ､その日は対象外として削除する"

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
                                        '上の行をカウント数分だけ削除する。
                                        Rows(i + 1 - Count_the_number_of_hours_of_the_day & ":" & CStr(i)).Delete
                                        
                                        Count_the_number_of_hours_of_the_day = 0
                                        
                                        Debug.Print i
                                        
                                        Exit For
                               End If
                        End If
                Next i
       Next EVERYDAY
       
       Debug.Print "END 日に13行が含まれていなければ､その日は対象外として削除する"
      
End Sub
Sub extract_2200_of_13H()
Attribute extract_2200_of_13H.VB_ProcData.VB_Invoke_Func = " \n14"

Macro1
Macro2

    'TODO　この当たり、動的に数値を変更できるようにする。

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
        '日付の区切りを.から／に置換
    Columns("A:A").Select
    Selection.Replace What:=".", Replacement:="/", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, _
        ReplaceFormat:=False
    Cells.Select
    
    'フィルタ設定
    Selection.AutoFilter
    
    '削除対象を抽出する。
    ActiveSheet.Range("$A$1:$G$200000").AutoFilter Field:=2, Criteria1:=Array( _
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

    Columns("G:G").Select
    Selection.ClearContents
    Range("G1").Select
    
End Sub
Sub prepare_pibot_table()
'ル－ル設定
' 東京市場ブレークアウトのデータ取得後のピボットの準備
'

    Rows("1:1").Select
    Selection.Insert Shift:=xlDown, CopyOrigin:=xlFormatFromLeftOrAbove '先頭行を挿入
    
    Range("A1").Select
    ActiveCell.FormulaR1C1 = "A1"
    Range("A1").Select

    'A1からA８まで列名を作成する。
    Selection.AutoFill Destination:=Range("A1:H1"), Type:=xlFillDefault
    Range("A1:H1").Select
    
    Selection.AutoFilter 'フィルタ設定
    
    'G列の空列を排除
    ActiveSheet.Range("$A$1:$H$200000").AutoFilter Field:=7, Criteria1:="<>"
    
    Range(Selection, Selection.End(xlDown)).Select '最終行まで選択状態にする｡
    Selection.Copy 'コピーする。
    
    sheets.Add After:=ActiveSheet '横にシートを作成する。
    
    Selection.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks _
        :=False, Transpose:=False 'テキストで貼り付ける。
    
    '日付と時間のセル設定
    Columns("A:A").Select
    Application.CutCopyMode = False
    Selection.NumberFormatLocal = "yyyy/m/d"
    Columns("B:B").Select
    Selection.NumberFormatLocal = "[$-x-systime]h:mm:ss AM/PM"
    
    '表の選択
    Range("A1:H1").Select
    Range(Selection, Selection.End(xlDown)).Select
    
End Sub
