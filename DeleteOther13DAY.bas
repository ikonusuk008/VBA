Attribute VB_Name = "DeleteOther13DAY"
Sub 日に13行が含まれていなければ､その日は対象外として削除する()
        
        Dim 最終行 As Long
        最終行 = Cells(Rows.Count, 1).End(xlUp).Row
        
        Dim 一日の時間数カウント As Long
        一日の時間数カウント = 0
        
        Dim BREAK_FLAG As String
        BREAK_FLAG = 0
        
        Dim i As Long
        
        For EVERYDAY = 0 To 100
        
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
                               
                               If 一日の時間数カウント = 13 Then
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

