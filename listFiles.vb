
Function ListFiles(DirPath, PathType)
    Dim MyName, Dic, Did, i, F, MyFileName, SheetName, SheetSize, Cell 
    Set objFolder = Nothing
    Set objShell = Nothing
    Set Dic = CreateObject("Scripting.Dictionary")    '创建一个字典对象
    Set Did = CreateObject("Scripting.Dictionary")
    Dic.Add (DirPath), ""
    i = 0
    If PathType <> "Root" Then
    Do While i < Dic.Count
        Ke = Dic.keys   '开始遍历字典
        MyName = Dir(Ke(i), vbDirectory)    '查找目录
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If (GetAttr(Ke(i) & MyName) And vbDirectory) = vbDirectory Then    '如果是次级目录
                    Dic.Add (Ke(i) & MyName & "\"), ""  '就往字典中添加这个次级目录名作为一个条目
                End If
            End If
            MyName = Dir    '继续遍历寻找
        Loop
        i = i + 1
    Loop
    End If
    SheetSize = Split(DirPath, "\")
    SheetName = SheetSize(UBound(SheetSize) - 1)
    Did.Add (SheetName), ""
    For Each Ke In Dic.keys
        MyFileName = Dir(Ke & "*.*")
        Do While MyFileName <> ""
            Did.Add (Ke & MyFileName), ""
            MyFileName = Dir
        Loop
    Next
    For Each Sh In ThisWorkbook.Worksheets
        If Sh.Name = SheetName Then
            Sheets(SheetName).Cells.Delete
            F = True
            Exit For
        Else
            F = False
        End If
    Next
    If Not F Then
        Sheets.Add.Name = SheetName
    End If
    Sheets(SheetName).[A1].Resize(Did.Count, 1) = WorksheetFunction.Transpose(Did.keys)
    Sheets(SheetName).[B1].Resize(Did.Count, 1) = WorksheetFunction.Transpose(Did.keys)
    For Each Cell In Sheets(SheetName).Range("A2:A"& Did.Count)
    If Cell <> "" Then
        CellArray = Split(Cell.Value, "\")
        CellName = CellArray(UBound(CellArray))
        CellPath = Replace(Cell.Value,CellName,"")
        Cell.Value = CellPath
        Sheets(SheetName).Hyperlinks.Add Cell, Cell.Value
    End If
    Next
    For Each Cell In Sheets(SheetName).Range("B2:B"& Did.Count)
    If Cell <> "" Then
        Sheets(SheetName).Hyperlinks.Add Cell, Cell.Value
        CellArray = Split(Cell.Value, "\")
        CellName = CellArray(UBound(CellArray))
        Cell.Value = CellName
    End If
    Next
    Call MergeCells(SheetName,Did.Count)
    Sheets(SheetName).Range("A1:B1").EntireColumn.AutoFit
    Sheets(SheetName).Range("A1:A"&Did.Count).HorizontalAlignment = xlCenter '水平居中
    Sheets(SheetName).Range("A1:A"&Did.Count).VerticalAlignment = xlCenter '垂直居中
End Function


Function MergeCells(SheetName,CellNumber)
    'set your data rows here 
    Dim Rows As Integer: Rows = CellNumber 

    Dim First As Integer: First = 2 
    Dim Last As Integer: Last = 0 
    Dim Rng As Range 

    Application.DisplayAlerts = False 
    With ActiveSheet 
     For i = 1 To Rows + 1 
      If Sheets(SheetName).Range("A" & i).Value <> Sheets(SheetName).Range("A" & First).Value Then 
       If i - 1 > First Then 
        Last = i - 1 

        Set Rng = Sheets(SheetName).Range("A" & First, "A" & Last) 
        Rng.MergeCells = True
       End If 

       First = i 
       Last = 0 
      End If 
     Next i 
    End With 
    Application.DisplayAlerts = True 
End Function

Private Sub CommandButton1_Click()
    Dim MyName, Dirs, RootNum, Sheets, t, TT
    t = Time
    Set objShell = CreateObject("Shell.Application")
    Set objFolder = objShell.BrowseForFolder(0, "chose", 0, 0)
    Set Dirs = CreateObject("Scripting.Dictionary")
    If Not objFolder Is Nothing Then lj = objFolder.self.Path & "\"
    Dirs.Add (lj), ""
    i = 0
    Do While i < Dirs.Count
        Ke = Dirs.keys   '开始遍历字典

        MyName = Dir(Ke(i), vbDirectory)    '查找目录
        Do While MyName <> ""
            If MyName <> "." And MyName <> ".." Then
                If (GetAttr(Ke(i) & MyName) And vbDirectory) = vbDirectory Then    '如果是次级目录
                    Dirs.Add (Ke(i) & MyName & "\"), ""  '就往字典中添加这个次级目录名作为一个条目
                End If
            End If
            MyName = Dir    '继续遍历寻找
        Loop
        i = i + 1
    Loop
    RootNum = Len(lj) - Len(Replace(lj, "\", ""))
    For Each keysss In Dirs.keys
    keysssNum = Len(keysss) - Len(Replace(keysss, "\", ""))
    If keysssNum = RootNum + 1 Then
    Call ListFiles(keysss, "")
    End If
    Next
    Call ListFiles(lj, "Root")
    TT = Time - t
    MsgBox Minute(TT) & " min" & Second(TT) & " sec"
End Sub

