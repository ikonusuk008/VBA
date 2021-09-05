Attribute VB_Name = "file_analyzer"
 Sub file_analyzer_main()
  
    Range("b1").Value = "Path"
    Range("c1").Value = "macro mame"
    Range("d1").Value = "date"
    Range("e1").Value = "size"
    
    Call setFileList(Cells(1, 1))
    
End Sub

Sub setFileList(searchPath)

    Dim startCell As Range
    Dim maxRow As Long
    Dim maxCol As Long

    Set startCell = Cells(2, 2)
    startCell.Select
    
    maxRow = startCell.SpecialCells(xlLastCell).Row
    maxCol = startCell.SpecialCells(xlLastCell).Column
    Range(startCell, Cells(maxRow + 1, maxCol)).ClearContents
    
    Call getFileList(searchPath)
    startCell.Select
    
End Sub

Sub getFileList(searchPath)

    Dim FSO As New FileSystemObject
    Dim objFiles As File
    Dim objFolders As Folder
    Dim separateNum As Long

    For Each objFolders In FSO.GetFolder(searchPath).SubFolders
    
        Call getFileList(objFolders.Path)
        
    Next
    
    For Each objFiles In FSO.GetFolder(searchPath).Files
    
        separateNum = InStrRev(objFiles.Path, "\")
        
        ActiveCell.Value = Left(objFiles.Path, separateNum - 1)
        ActiveCell.Offset(0, 1).Value = Right(objFiles.Path, Len(objFiles.Path) - separateNum)
        ActiveCell.Offset(0, 2).Value = FileDateTime(objFiles)
        ActiveCell.Offset(0, 3).Value = Format((FileLen(objFiles) / 1024), "#.0")
        ActiveCell.Offset(1, 0).Select
        
    Next
    
End Sub

