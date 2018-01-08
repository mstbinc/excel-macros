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
Dim answer As Integer
Dim col As String

answer = MsgBox("Select Design Column?", vbYesNo + vbQuestion, "Select Box Totals")

If answer = vbYes Then
	col = "E"
Else
	col = "H"
End If

For Each c In ActiveSheet.UsedRange.Columns(col).cells
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