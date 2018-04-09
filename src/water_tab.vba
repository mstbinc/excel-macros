Sub GenerateWaterTab()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "W\n14"
'
' EasyWater Macro
' Makes a Water --> Tab because I'm lazy and shouldn't have to
'
' Keyboard Shortcut: Ctrl+Shift+W
'

    Dim after As Integer

    For i = 1 To Sheets.Count

        'MsgBox (Sheets(i).Name & " " & Sheets(i).Tab.color)
        after = i
        If Sheets(i).Tab.color = "False" Then Exit For

    Next i

    ActiveWorkbook.Sheets.Add Before:=Sheets(after)
    ActiveSheet.name = "WATER --->"
    ActiveSheet.Tab.color = RGB(0, 0, 0)
    range("A:XFD").EntireColumn.Hidden = True
    range("1:1048576").EntireRow.Hidden = True

End Sub