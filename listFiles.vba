' List all files on informed folder, printing each file in one cell on column A, starting on A2
' Folder path shold be on cell H2
' File extension filter should be on cell H4. If left blank will return all files

Sub listFiles()

    Application.ScreenUpdating = False
    
    ' Clear previous results
    Range("A2", Range("A2").End(xlDown)).Clear
    
    Dim filePath As String
    filePath = Range("H2").Value
    
    If Right(filePath, 1) <> "\" Then
        filePath = filePath + "\"
    End If
    
    Dim extension As String, extensionSize As Integer
    
    extension = Range("H4").Value
    extensionSize = Len(extension)
    
    Dim row As Integer
    row = 2
    
    
    Dim strfile As String, filenum As String
    strfile = Dir(filePath)
    
    If extensionSize = 0 Then
        Do While strfile <> ""
            Range("A" & row).Value = strfile
            row = row + 1
            strfile = Dir
        Loop
    Else
        Do While strfile <> ""
            If Right(strfile, extensionSize) = extension Then
                Range("A" & row).Value = strfile
                row = row + 1
            End If
            strfile = Dir
        Loop
    End If
    
    Application.ScreenUpdating = True
End Sub

Sub Clear()
    Application.ScreenUpdating = False
    Range("H2").ClearContents
    Range("H4").Value = ""
    Range("A2", Range("A2").End(xlDown)).Clear
    Application.ScreenUpdating = True
End Sub
