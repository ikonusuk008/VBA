Attribute VB_Name = "seve_Time_series"
Sub seve_Time_series_main()
Attribute seve_Time_series_main.VB_ProcData.VB_Invoke_Func = "b\n14"

    ActiveWorkbook.Save
    Dim objFSO As Object, txtSource As String, txtDestination
    Set objFSO = CreateObject("Scripting.FileSystemObject")
    Dim backupFolderPath As String
    backupFolderPath = "C:\Time_series\"
    Dim TargetFilePath As String
    TargetFilePath = ActiveWorkbook.FullName
    Dim save_destination As String
    save_destination = backupFolderPath & objFSO.GetBaseName(TargetFilePath) + Format(Date, "_eemmdd-") + Format(Time, "hhmmss.") + objFSO.GetExtensionName(TargetFilePath)
    Dim GetExtensionName As String
    GetExtensionName = objFSO.GetExtensionName(TargetFilePath)
    objFSO.CopyFile TargetFilePath, save_destination
    
   

    For i = 1 To Worksheets.Count
    
        Worksheets(i).Select
        ActiveWindow.zoom = 90
        Range("a1").Select
        previewFlag = 1
        
        If previewFlag = 1 Then
            ActiveWindow.View = xlNormalView
        Else
            ActiveWindow.View = xlPageBreakPreview
        End If
        
    Next i
    
    Worksheets(1).Select


End Sub
