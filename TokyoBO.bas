Attribute VB_Name = "TokyoBO"
Sub main()
        '３６５×５＝１８２５日
        For n = 0 To 1825
                 買い (n)
                売り (n)
        Next n
End Sub
Function 買い(ByVal n As Long)

        Dim 東京市場高値 As Double
        Dim 欧州1時間値 As Double
        Dim 欧州終値 As Double
        
                        
        Dim 日24時間 As Long
        日24時間 = 24
        
        Dim 日数インデックス As Long
        日数インデックス = (n * 日24時間)
        
        Dim 決済時刻列 As Long
        決済時刻列 = 22

       
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim ブレーク判定値 As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        ブレーク判定値 = 1  'リセット
       
        東京市場高値 = WorksheetFunction.Max(Range("d" & 10 + 日数インデックス & ":d" & 15 + 日数インデックス))
        
        
        '東京市場　高値　ブレーク
        For i = 16 + 日数インデックス To 決済時刻列 + 日数インデックス    '15時から２２時まで
            
                欧州1時間値 = CDbl(Range("f" & i).Value)
        
                If 欧州1時間値 > 東京市場高値 Then
                        ブレーク判定値 = 2
                        Exit For
                End If
        Next i
        
        'ブレーク後の損切り判定
        For i2 = i To 決済時刻列 + 日数インデックス
                欧州1時間値 = CDbl(Range("f" & i2).Value)
        
                If ((欧州1時間値 - 東京市場高値) * 100) < -30 Then
                    
                        ブレーク判定値 = 3
                        
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If ブレーク判定値 = 1 Then
                'トレード無し
                Range("g" & 決済時刻列 + 日数インデックス).Value = 0
        ElseIf ブレーク判定値 = 2 Then
                'ブレーク後の決済
                Range("g" & 決済時刻列 + 日数インデックス).Value = (欧州終値 - 東京市場高値) * 100
        ElseIf ブレーク判定値 = 3 Then
                'ブレーク後の損切り
                Range("g" & 決済時刻列 + 日数インデックス).Value = -30
        End If

End Function
Function 売り(ByVal n As Long)

        Dim 東京市場安値 As Double
        Dim 欧州1時間値 As Double
        Dim 欧州終値 As Double
        
        
        Dim 日24時間 As Long
        日24時間 = 24
        
        Dim 日数インデックス As Long
        日数インデックス = (n * 日24時間)
        
        Dim 決済時刻列 As Long
        決済時刻列 = 22

        
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim ブレーク判定値 As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        
        ブレーク判定値 = 1  'リセット
       
        東京市場安値 = WorksheetFunction.Min(Range("e" & 10 + 日数インデックス & ":e" & 15 + 日数インデックス))
        
        '東京市場　安値　ブレーク
        For i = 16 + 日数インデックス To 決済時刻列 + 日数インデックス
            
                欧州1時間値 = CDbl(Range("f" & i).Value)
        
                
                If 欧州1時間値 < 東京市場安値 Then
                        ブレーク判定値 = 2
                        Exit For
                End If
        Next i
        
        '東京市場　安値　損切り
        For i2 = i To 決済時刻列 + 日数インデックス   'ブレーク後の損切り判定
                欧州1時間値 = CDbl(Range("f" & i2).Value)
        
                If ((東京市場安値 - 欧州1時間値) * 100) < -30 Then
                        ブレーク判定値 = 3
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If ブレーク判定値 = 1 Then
                'トレード無し
                Range("h" & 決済時刻列 + 日数インデックス).Value = 0
        ElseIf ブレーク判定値 = 2 Then
                'ブレーク後の決済
                Range("h" & 決済時刻列 + 日数インデックス).Value = (東京市場安値 - 欧州終値) * 100
        ElseIf ブレーク判定値 = 3 Then
                'ブレーク後の損切り
                Range("h" & 決済時刻列 + 日数インデックス).Value = -30
        End If

End Function
Sub 日に24行が含まれていなければ､その日は対象外として削除する()
        
        Dim 最終行 As Long
        最終行 = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim 一日の時間数カウント As Long
        一日の時間数カウント = 0
        
        Dim BREAK_FLAG As String
        BREAK_FLAG = 0
        
        Dim i As Long
        
        For EVERYDAY = 0 To 30
        
                If BREAK_FLAG = 1 Then
                        Exit For
                End If
             
                For i = 1 To 最終行
                
                        If Range("a" & i).Value = "" Then
                                BREAK_FLAG = 1
                                Exit For
                        End If
               
                         一日の時間数カウント = 一日の時間数カウント + 1
        
                        If Range("a" & i).Value <> Range("a" & i + 1).Value Then
                               
                               If 一日の時間数カウント = 24 Then
                                        一日の時間数カウント = 0
                                        
                               Else
                                        '上の行をカウント数分だけ削除する。
                                        Rows(i + 1 - 一日の時間数カウント & ":" & CStr(i)).Delete
                                        
                                        一日の時間数カウント = 0
                                        
                                        Debug.Print i
                                        
                                        Exit For
                               End If
                        End If
                Next i
       Next EVERYDAY
      
End Sub

