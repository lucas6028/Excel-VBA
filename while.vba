Sub While1()
'
'while 巨集
'快速鍵 Ctrl+shift+W
'註記為一次20筆原始值
'參數更改Line: 10, 33, 39
'
    Dim x, Counter, n  As Integer
    x = 16
    n = 2 ' Modify n !! 19 = 20 -1
    Counter = 0 'Initialize counter
        ActiveCell.Range("A1:E14").Select 'Select A1 to E14
        Application.CutCopyMode = False 'Cancel all cut, copy previously
        Selection.Copy 
        ActiveCell.Offset(0,6).Range("A1").Select 'Move selection 0 rows down, 6 columns right
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True 'Only transpose it
        ' The first data is finish

    While Counter < n
        Counter = Counter + 1 'Increment counter
        'ActiveWindow.SmallScroll Down:=9 
        ActiveCell.Offset(x, -6).Range("A1:E14").Select
        Application.CutCopyMode = False
        Selection.Copy
        'ActiveWindow.SmallScroll Down:=-6
        ActiveCell.Offset((-x + 1), 6).Range("A1").Select
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True
        x = x + 15
    Wend 'End of while loop
        'ActiveWindow.SmallScroll Down:=-12
        ActiveCell.Offset(-n, -6).Range("A1:E46").Select ' Modify E !! 318 = 16(n + 1) - 2
        With Selection.Font
            .Color = -16776961 'Color red
            .TintAndShade = 0
        End With
        'ActiveWindow.SmallScroll Down:=9
        ActiveCell.Offset(0, 6).Range("A1:Q3").Select ' Modify Q !! 20 = n + 1
        Selection.Copy
        'ActiveWindow.SmallScroll Down:=16(n + 1)
End Sub
