Sub PdfToText()

    Dim answer
    Dim source As String
    Dim dest As String
    Dim exe As String
    Dim wsh As Object
    
    Set wsh = VBA.CreateObject("WScript.Shell")
    answer = Application.GetOpenFilename(Title:="Please choose a file to import", FileFilter:="PDF Files *.pdf (*.pdf),")
    If answer = False Then
        MsgBox "No file specified.", vbExclamation, "Warning"
        Exit Sub
    End If

    
    source = answer
    exe = "C:\Users\Bruce\Documents\pdftotext.exe"
    dest = Replace(source, ".pdf", ".txt")
    
    wsh.Run exe & " " & source & " -layout", vbHide, True
    
    Workbooks.OpenText Filename:=dest, _
        StartRow:=1, _
        DataType:=xlFixedWidth, _
        TrailingMinusNumbers:=True
    
End Sub
