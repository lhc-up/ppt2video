Sub EXPORTNotes()
    On Error Resume Next
    Dim str As String
    str = ""
    For i = 1 To ActivePresentation.Slides.Count
        str = str & "P" & VBA.str(i)& ":" 
        str = str & ActivePresentation.Slides.Item(i).NotesPage.Shapes.Placeholders(2).TextFrame.TextRange.Text & Chr(13) & Chr(10)
    Next
    FileName = Left(Application.ActivePresentation.Name, 29) & ".txt"
    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set sFile = FSO.CreateTextFile(FileName, True, True)
        sFile.WriteLine (str)
        sFile.Close
    Set sFile = Nothing
    Set FSO = Nothing
    MsgBox "备注文件已保存至：" & FileName
End Sub