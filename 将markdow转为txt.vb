Sub CommandButton1_Click()
    On Error Resume Next
    Dim MyName, Dic, Did, i, F, MyFileName, SheetSize, Cell
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "chose", 0, 0)
    Set Dic = CreateObject("Scripting.Dictionary")
    If Not objFolder Is Nothing Then
        lj = objFolder.self.Path & "\"
    Else
        Exit Sub
    End If
    Dic.Add (lj), ""
    Set Did = CreateObject("Scripting.Dictionary")
    i = 0
    Do While i < Dic.Count
        Ke = Dic.keys   '?????鯊晙???

        MyName = Dir(Ke(i), vbDirectory)
        If Err.Number <> 0 Then
        End If
        On Error GoTo 0
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If (GetAttr(Ke(i) & MyName) And vbDirectory) = vbDirectory Then
                    Dic.Add (Ke(i) & MyName & "\"), ""
                End If
            End If
            MyName = Dir
        Loop
        i = i + 1
    Loop
    For Each Ke In Dic.keys
        MyFileName = Dir(Ke & "*.md")
        Do While MyFileName <> ""
            Did.Add (Ke & MyFileName), ""
            Call ChangeFile(Ke & MyFileName)
            MyFileName = Dir
        Loop
    Next
    MsgBox ("converted files : " & Did.Count)
End Sub

Function ChangeFile(fileToRead)
    On Error GoTo Err_Handle
    Dim fileToWrite
    'Declare ALL of your variables :)
    Const ForReading = 1 '
    fileToWrite = Replace(fileToRead, ".md", ".txt") ' the path of a new file
    Dim FSO As Object
    Dim readFile As Object 'the file you will READ
    Dim writeFile As Object 'the file you will CREATE
    Dim repLine As Variant 'the array of lines you will WRITE
    Dim ln As Variant
    Dim l As Long

    Set FSO = CreateObject("Scripting.FileSystemObject")
    Set readFile = FSO.OpenTextFile(fileToRead, ForReading, False)
    Set writeFile = FSO.CreateTextFile(fileToWrite, True, False)

    '# Read entire file into an array & close it
    repLine = readFile.ReadAll
    readFile.Close
    repLine = "<markdown>" & repLine & "</markdown>"

    '# Write to the array items to the file
    writeFile.Write repLine
    writeFile.Close

    '# clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
Err_Handle:
    '# clean up
    Set readFile = Nothing
    Set writeFile = Nothing
    Set FSO = Nothing
End Function




