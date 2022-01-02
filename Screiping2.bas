Attribute VB_Name = "Screiping2"
Sub Scraiping1()
    'Dim driver As New Selenium.ChromeDriver
    Dim driver As New Selenium.PhantomJSDriver
    Dim elmDoc     As WebElement
        Dim OutputTarget As Range
        Dim sKeyWord As String
        Range("4:999").Clear                'サンプルプログラムなので手抜き
        sKeyWord = Range("検索")
        If sKeyWord = "" Then
            Exit Sub
        End If
        Set OutputTarget = Range("OutputArea")
        With driver
            .Start
            '.Window.SetSize 1920, 1080
            .Get "https://www.library.toyota.aichi.jp/" '豊田市図書館のHPにアクセス
            '検索するキーワードを投入
            .FindElementById("kensaku_keyword").SendKeys Range("検索") & vbCrLf
            'スクレイピング開始
            'doclistの中に1冊ごとにdoc,doc,doc… という繰り返しで本の情報が入っている
            For Each elmDoc In .FindElementByClass("doclist").FindElementsByClass("doc")
                '各CSS名にアクセス
                OutputTarget.Cells(, 1) = elmDoc.FindElementByClass("doc-title").Text       '本のタイトル
                    Set OutputTarget = OutputTarget.Offset(1)
                    OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-writer").Text      '著者
                        Set OutputTarget = OutputTarget.Offset(1)
                        OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-recap").Text       '本の概略
                            Set OutputTarget = OutputTarget.Offset(1)
                            OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-available").Text   '貸出可否
                                Set OutputTarget = OutputTarget.Offset(2)
                            Next
                        End With
                        
                        
                        
End Sub
Sub Scraiping2()
Attribute Scraiping2.VB_ProcData.VB_Invoke_Func = "j\n14"

    'https://kawattawatta.com/it/vba-web-scraping-with-google-chrome-and-selenium/

    ActiveSheet.Hyperlinks.Add Anchor:=ActiveCell, Address:=ActiveCell.Value
    
    Dim driver As New Selenium.PhantomJSDriver
    
    With driver
        .Start
        .Get ActiveCell.Value
        
        Cells(ActiveCell.Row, ActiveCell.Column + 1) = driver.Window.Title



    End With
    
End Sub
Sub ScraipingLIST()
    'https://kawattawatta.com/it/vba-web-scraping-with-google-chrome-and-selenium/
    
    
    Dim driver As New Selenium.PhantomJSDriver
    
    
    For i = 2 To Cells(Cells(Rows.Count, 2).Row, 2).End(xlUp).Row
    
    
        With driver
            .Start
            .Get Range("B" & i).Value
            
            Range("A" & i).Value = driver.Window.Title
    
    
        End With
        
        driver.Wait 3000
    
    
    Next i
    
End Sub


