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
        日24時間 = 13
        
        Dim 日数インデックス As Long
        日数インデックス = (n * 日24時間)
        
        Dim 決済時刻列 As Long
        決済時刻列 = 13

       
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim ブレーク判定値 As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        ブレーク判定値 = 1  'リセット
       
        東京市場高値 = WorksheetFunction.Max(Range("d" & 1 + 日数インデックス & ":d" & 6 + 日数インデックス))
        
        
        '東京市場　高値　ブレーク
        For i = 1 + 日数インデックス To 決済時刻列 + 日数インデックス    '15時から２２時まで
            
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
        日24時間 = 13
        
        Dim 日数インデックス As Long
        日数インデックス = (n * 日24時間)
        
        Dim 決済時刻列 As Long
        決済時刻列 = 13

        
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim ブレーク判定値 As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        
        ブレーク判定値 = 1  'リセット
       
        東京市場安値 = WorksheetFunction.Min(Range("e" & 1 + 日数インデックス & ":e" & 6 + 日数インデックス))
        
        '東京市場　安値　ブレーク
        For i = 1 + 日数インデックス To 決済時刻列 + 日数インデックス
            
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

