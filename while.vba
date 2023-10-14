Sub while()
'
'while 巨集
'快速鍵 Ctrl+w
'
    Dim intx, inty, intz, Counter, Down As Integer
    x = 16
    y = 6
    Counter = 0 'Initialize counter
        ActiveCell.Range("A1:E14").Select 'Select A1 to E14
        Application.CutCopyMode = False 'Cancel all cut, copy previously
        Selection.Copy 
        ActiveCell.Offset(0,6).Range("A1").Select 'Move selection 0 rows down, 6 columns right
        Selection.PasteSpecial Paste:=xlPasteAll, Operation:=xlNone, SkipBlanks:= _
            False, Transpose:=True 'Only transpose it
        ' The first data is finish

    While Counter < 20
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
        ActiveWindow.SmallScroll Down:=-12
        ActiveCell.Offset(-19, -6).Range("A1:E318").Select
        With Selection.Font
            .Color = -16776961 'Color red
            .TintAndShade = 0
        End With
        ActiveWindow.SmallScroll Down:=9
        ActiveCell.Offset(0, 6).Range("A1:Q20").Select
        Selection.Copy
        ActiveWindow.SmallScroll Down:=310
End Sub