Attribute VB_Name = "TokyoBO"
Sub main()
        '日に24行が含まれていなければ､その日は対象外として削除する
        
        For n = 0 To 5000
                 買い (n)
                売り (n)
        Next n
End Sub
Function 買い(ByVal n As Long)

        Dim 東京市場高値 As Double '
        Dim 欧州1時間値 As Double   '
        Dim 欧州終値 As Double    '
        Dim 日数インデックス As Long '
        Dim 決済時刻列 As Long '決済時刻列
        Dim 日24時間 As Long
        日24時間 = 24
        
        決済時刻列 = 22
        
        日数インデックス = (n * 日24時間)
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim buyFlag As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        buyFlag = 1
       
        東京市場高値 = WorksheetFunction.Max(Range("d" & 10 + 日数インデックス & ":d" & 15 + 日数インデックス))     '東京市場　高値
        
        For i = 16 + 日数インデックス To 決済時刻列 + 日数インデックス    '15時から２２時まで
            
                欧州1時間値 = CDbl(Range("f" & i).Value)
        
                If 欧州1時間値 > 東京市場高値 Then  '東京市場　高値　ブレーク
                        buyFlag = 2
                        Exit For
                End If
        Next i
        
        For i2 = i To 決済時刻列 + 日数インデックス   'ブレーク後の損切り判定
                欧州1時間値 = CDbl(Range("f" & i2).Value)
        
                If ((欧州1時間値 - 東京市場高値) * 100) < -30 Then
                    
                        buyFlag = 3
                        
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If buyFlag = 1 Then
                Range("g" & 決済時刻列 + 日数インデックス).Value = 0
        ElseIf buyFlag = 2 Then
                Range("g" & 決済時刻列 + 日数インデックス).Value = (欧州終値 - 東京市場高値) * 100
        ElseIf buyFlag = 3 Then
                Range("g" & 決済時刻列 + 日数インデックス).Value = -30
        End If

End Function
Function 売り(ByVal n As Long)

        Dim min9_15 As Double '東京市場　min
        Dim 欧州1時間値 As Double   '欧州1時間値
        Dim 欧州終値 As Double    '欧州終値
        Dim 日数インデックス As Long '日数インデックス
        Dim 決済時刻列 As Long '決済時刻列
        Dim 日24時間 As Long
        
        日24時間 = 24
        決済時刻列 = 22
        
        日数インデックス = (n * 日24時間)
        
        欧州終値 = Range("f" & (決済時刻列 + 日数インデックス)).Value
        
        Dim buyFlag As Integer  '（１：ブレークなし。２：ブレークあり。３：ブレーク損切）
        
        buyFlag = 1
       
        min9_15 = WorksheetFunction.Min(Range("e" & 10 + 日数インデックス & ":e" & 15 + 日数インデックス))     '東京市場　高値
        
        For i = 16 + 日数インデックス To 決済時刻列 + 日数インデックス    '15時から２２時まで
            
                欧州1時間値 = CDbl(Range("f" & i).Value)
        
                If 欧州1時間値 < min9_15 Then  '東京市場　高値　ブレーク
                        buyFlag = 2
                        Exit For
                End If
        Next i
        
        For i2 = i To 決済時刻列 + 日数インデックス   'ブレーク後の損切り判定
                欧州1時間値 = CDbl(Range("f" & i2).Value)
        
                If ((min9_15 - 欧州1時間値) * 100) < -30 Then
                        buyFlag = 3
                        Exit For
                End If
        Next i2
        
        '判定結果を２２時のG列に書き込む
        If buyFlag = 1 Then
                Range("h" & 決済時刻列 + 日数インデックス).Value = 0
        ElseIf buyFlag = 2 Then
                Range("h" & 決済時刻列 + 日数インデックス).Value = (min9_15 - 欧州終値) * 100
        ElseIf buyFlag = 3 Then
                Range("h" & 決済時刻列 + 日数インデックス).Value = -30
        End If

End Function
Sub 日に24行が含まれていなければ､その日は対象外として削除する()
        
        Dim 最終行 As Long
        
        最終行 = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim 一日の時間数カウント As Long
        一日の時間数カウント = 0
        Dim f As String
        f = 0
        
        Dim i As Long
        For h = 0 To 2000
        
                If f = 1 Then
                        Exit For
                End If
             
                For i = 1 To 最終行
                
                        If Range("a" & i).Value = "" Then
                                f = 1
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
                                        Exit For
                               End If
                        End If
                Next i
       Next h
                
                
End Sub




Sub 日に14時間が含まれていなければ､その日は対象外として削除する()
        
        Dim n As Long
        
        n = Cells(Rows.Count, 1).End(xlUp).Row
        
        
        For i = 1 To n
        
                If (Format(Range("b" & i).Value, "hh:nn") = "09:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "10:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "12:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "11:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "13:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "12:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "14:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "13:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "15:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "14:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "16:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "15:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "17:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "16:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "18:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "17:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "19:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "18:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "20:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "19:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "21:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "22:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "22:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "22:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Then
                       
                       
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "09:00") Or _
                        (Format(Range("b" & i).Value, "hh:nn") = "12:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "15:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Then
                       
                       
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "10:00") Or _
                       (Format(Range("b" & i).Value, "hh:nn") = "21:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                
                ElseIf (Format(Range("b" & i).Value, "hh:nn") = "20:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                         Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                 
                 ElseIf (Format(Range("b" & i).Value, "hh:nn") = "19:00" And Format(Range("b" & i + 1).Value, "hh:nn") = "11:00") Then
                       
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                        Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                         Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                         Range("b" & i + 1).EntireRow.Insert 'セル名を指定する場合
                

                End If

                Debug.Print Format(Range("b" & i).Value, "hh:nn")
        Next i

End Sub

