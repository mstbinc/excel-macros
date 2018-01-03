Sub SelectBoxTotals()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "T\n14"
'
' SelectBoxTotals Macro
' Selects all totals in Column E of the Outlet Sheet for easy copying to the box sheets
' Created by Matt P on 3/13/2014
'
' Keyboard Shortcut: Ctrl+Shift+T
'
Dim sel As range

For Each c In ActiveSheet.UsedRange.Columns("E").cells
    If c.Value <> "" And c.Font.Bold Then
        If sel Is Nothing Then
            Set sel = c
        Else
            Set sel = Union(sel, c)
        End If
    End If
Next c

If Not sel Is Nothing Then
    sel.Select
End If

End Sub