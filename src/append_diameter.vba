Sub AppendDiameter()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "D\n14"

'
' Keyboard Shortcut: Ctrl+Shift+D
'
    For Each cell In selection
        If (InStr(1, cell.Value, ChrW(248)) = 0 And IsNumeric(cell.Value)) Then
            cell.Value = cell.Value & ChrW(248)
        End If
    Next

End Sub