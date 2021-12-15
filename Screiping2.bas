Attribute VB_Name = "Screiping2"
Sub Scraiping1()
    'Dim driver As New Selenium.ChromeDriver
    Dim driver As New Selenium.PhantomJSDriver
    Dim elmDoc     As WebElement
        Dim OutputTarget As Range
        Dim sKeyWord As String
        Range("4:999").Clear                '�T���v���v���O�����Ȃ̂Ŏ蔲��
        sKeyWord = Range("����")
        If sKeyWord = "" Then
            Exit Sub
        End If
        Set OutputTarget = Range("OutputArea")
        With driver
            .Start
            '.Window.SetSize 1920, 1080
            .Get "https://www.library.toyota.aichi.jp/" '�L�c�s�}���ق�HP�ɃA�N�Z�X
            '��������L�[���[�h�𓊓�
            .FindElementById("kensaku_keyword").SendKeys Range("����") & vbCrLf
            '�X�N���C�s���O�J�n
            'doclist�̒���1�����Ƃ�doc,doc,doc�c �Ƃ����J��Ԃ��Ŗ{�̏�񂪓����Ă���
            For Each elmDoc In .FindElementByClass("doclist").FindElementsByClass("doc")
                '�eCSS���ɃA�N�Z�X
                OutputTarget.Cells(, 1) = elmDoc.FindElementByClass("doc-title").Text       '�{�̃^�C�g��
                    Set OutputTarget = OutputTarget.Offset(1)
                    OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-writer").Text      '����
                        Set OutputTarget = OutputTarget.Offset(1)
                        OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-recap").Text       '�{�̊T��
                            Set OutputTarget = OutputTarget.Offset(1)
                            OutputTarget.Cells(, 2) = elmDoc.FindElementByClass("doc-available").Text   '�ݏo��
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


