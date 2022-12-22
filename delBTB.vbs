Sub delBTB()
'
' delBTB Macro
'
' Keyboard Shortcut: Ctrl+d
'
    Columns("A:A").Select
    Selection.FormatConditions.AddUniqueValues
    Selection.FormatConditions(Selection.FormatConditions.Count).SetFirstPriority
    Selection.FormatConditions(1).DupeUnique = xlDuplicate
    With Selection.FormatConditions(1).Font
        .Color = -16383844
        .TintAndShade = 0
    End With
    With Selection.FormatConditions(1).Interior
        .PatternColorIndex = xlAutomatic
        .Color = 13551615
        .TintAndShade = 0
    End With
    Selection.FormatConditions(1).StopIfTrue = False
    Selection.AutoFilter
    Columns("A:B").Select
    Selection.AutoFilter
    Selection.AutoFilter
    Range("A1").Select
    ActiveSheet.Range("$A$1:$B$4328").AutoFilter Field:=1, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
    ActiveSheet.Range("$A$1:$B$4328").AutoFilter Field:=2, Criteria1:="BTB"
End Sub
