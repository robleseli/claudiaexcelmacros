Sub delDuplicates()
'
' delDuplicates Macro
'
' Keyboard Shortcut: Ctrl+e
'
    ActiveSheet.Range("$A$1:$B$8569").AutoFilter Field:=1, Criteria1:=RGB(255, _
        199, 206), Operator:=xlFilterCellColor
    Columns("A:A").Select
    ActiveSheet.Range("$A$1:$A$8570").RemoveDuplicates Columns:=1, Header:= _
        xlYes
End Sub
