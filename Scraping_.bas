Attribute VB_Name = "Scraping_"
Sub GetTable4()

    Dim IE As InternetExplorer
    Dim Doc As HTMLDocument
    Dim objTAG As Object
    Dim i As Long
    Dim r As Long
    
    Set objIE = CreateObject("InternetExplorer.Application")
    
    objIE.Visible = False
    objIE.Navigate "C:\mydrive\00-share\05-Finance\スーパークリーナー医薬品\楽天証券株式会社_467ec72c.htm"
    
    Do While objIE.Busy Or objIE.ReadyState < READYSTATE_COMPLETE
        DoEvents
    Loop
    
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"
    
    r = 1

    Cells(r, 1).Select
    
    Debug.Print objIE.document.all.Length - 1
    
        'ページ上部を走査対象外とすることで、お店に限定する
        For i = 1 To objIE.document.all.Length - 1
        
        Debug.Print objIE.document.all(i).outerHTML
        
        'DIV,DIV,DIV,P,Aという並びが出現する箇所を探す
'            If objIE.document.all(i).tagName = "LI" Then
'             If objIE.document.all(i + 1).tagName = "H4" Then
'              If objIE.document.all(i + 2).tagName = "PRE" Then
               If objIE.document.all(i + 3).tagName = "DIV" Then
'                If objIE.document.all(i + 4).tagName = "DL" Then
             
            
                '店名があると判断
                r = r + 1
                
                Debug.Print "r: " & r
                Debug.Print objIE.document.all(i).tagName
                Debug.Print objIE.document.all(i).innerHTML
                
                Cells(r, 1) = objIE.document.all(i).tagName
                Cells(r, 2) = objIE.document.all(i).innerHTML
                Cells(r, 3) = r
                Cells(r, 4) = i
                
                '以降のタグから、目印のSPANタグを走査
                For i2 = i To objIE.document.all.Length - 1

                    '夜の点数があれば、そのタグを起点に値を取得
                    Debug.Print "i2: " & i2

                    If InStr(objIE.document.all(i2 + 2).outerHTML, "<tr class=") > 0 Then

                        Cells(r, 1) = objIE.document.all(i2 + 1).innerText
                        Cells(r, 2) = objIE.document.all(i2 + 2).innerText
                        Cells(r, 3) = objIE.document.all(i2 + 3).innerText
                        Cells(r, 4) = objIE.document.all(i2 + 4).innerText
                        Cells(r, 5) = objIE.document.all(i2 + 5).innerText
                        Cells(r, 6) = objIE.document.all(i2 + 6).innerText
                        Cells(r, 7) = objIE.document.all(i2 + 7).innerText
                        Cells(r, 8) = objIE.document.all(i2 + 8).innerText
                        Cells(r, 9) = objIE.document.all(i2 + 9).innerText
                        Cells(r, 10) = objIE.document.all(i2 + 10).innerText
                        Cells(r, 11) = objIE.document.all(i2 + 11).innerText
                        Cells(r, 12) = objIE.document.all(i2 + 12).innerText

'                        Cells(r, 1) = objIE.document.all(i2 + 1).outerHTML
'                        Cells(r, 2) = objIE.document.all(i2 + 2).outerHTML
'                        Cells(r, 3) = objIE.document.all(i2 + 3).outerHTML
'                        Cells(r, 4) = objIE.document.all(i2 + 4).outerHTML
'                        Cells(r, 5) = objIE.document.all(i2 + 5).outerHTML
'                        Cells(r, 6) = objIE.document.all(i2 + 6).outerHTML
'                        Cells(r, 7) = objIE.document.all(i2 + 7).outerHTML
'                        Cells(r, 8) = objIE.document.all(i2 + 8).outerHTML
'                        Cells(r, 9) = objIE.document.all(i2 + 9).outerHTML
'                        Cells(r, 10) = objIE.document.all(i2 + 10).outerHTML
'                        Cells(r, 11) = objIE.document.all(i2 + 11).outerHTML
'                        Cells(r, 12) = objIE.document.all(i2 + 12).outerHTML
'
'                        Cells(r, 1) = objIE.document.all(i2 + 1).tagName
'                        Cells(r, 2) = objIE.document.all(i2 + 2).tagName
'                        Cells(r, 3) = objIE.document.all(i2 + 3).tagName
'                        Cells(r, 4) = objIE.document.all(i2 + 4).tagName
'                        Cells(r, 5) = objIE.document.all(i2 + 5).tagName
'                        Cells(r, 6) = objIE.document.all(i2 + 6).tagName
'                        Cells(r, 7) = objIE.document.all(i2 + 7).tagName
'                        Cells(r, 8) = objIE.document.all(i2 + 8).tagName
'                        Cells(r, 9) = objIE.document.all(i2 + 9).tagName
'                        Cells(r, 10) = objIE.document.all(i2 + 10).tagName
'                        Cells(r, 11) = objIE.document.all(i2 + 11).tagName
'                        Cells(r, 12) = objIE.document.all(i2 + 12).tagName


                        Exit For
                    
                    End If

                Next i2
                End If
'               End If
'              End If
'             End If
'            End If
            
        Next i

End Sub
Sub GetTable5()
    Dim IE As InternetExplorer
    Dim Doc As HTMLDocument
    Dim objTAG As Object
    Dim i As Long
    Dim n As Long
    
    Set objIE = CreateObject("InternetExplorer.Application")
    
    objIE.Visible = False
    objIE.Navigate "https://member.rakuten-sec.co.jp/app/info_jp_quants_research.do;BV_SessionID=635AAF413A3E6A375BA41739AF9DDA3E.40c4d9b7?eventType=init&gmn=J&smn=01&lmn=01&fmn=01"
    
    Do While objIE.Busy Or objIE.ReadyState < READYSTATE_COMPLETE
        DoEvents
    Loop
   
    Cells.ClearContents
    Cells.NumberFormatLocal = "G/標準"

    Dim flag_1 As Boolean
    flag_1 = False
    Dim s_1 As String
        
    For Each objTAG In objIE.document.all

        s_1 = Replace(objTAG.outerHTML, Chr(10), "")
        's_1 = objTAG.outerHTML

'        If s_1 = "メソッドの詳細" Then
'            flag_1 = True
'            Debug.Print ""
'        End If

            'If flag_1 = True Then
                n = n + 1
                'Cells(n + 2, 1) = "'" & TypeName(objTAG) 'TypeNameでオブジェクトのタイプを表示
                Cells(n + 2, 1) = "'" & objTAG.tagName   'タグの名前
                Cells(n + 2, 2) = s_1
            'End If

    Next

End Sub

