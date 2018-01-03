Sub ExhaustSetup()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "E\n14"
'
' EasyExhaust Macro
' Makes an Exhaust Tab with Inlets because I'm lazy and I shouldn't have to.
'
' Keyboard Shortcut: Ctrl+Shift+E
'

Dim after As Integer
Dim num_efs As Integer
Dim name_efs As Integer
Dim pages As Integer


    num_efs = InputBox("How many exhaust fans are on this job?", "Number of Exhaust Fans")
    name_efs = MsgBox("Autoname EF-1 thru EF-" & num_efs & "?", vbYesNo, "AutoNumber Exhaust?")
    pages = Application.WorksheetFunction.Ceiling(num_efs / 3, 1)


    For i = 1 To Sheets.Count
        after = i
        If Sheets(i).Tab.color = "False" Then Exit For
    Next i

    Sheets("FANTEST").Copy Before:=Sheets(after)
    ActiveSheet.name = "EFs"
    ActiveSheet.Tab.color = 2646607

    If (name_efs <> 6) Then
        range("B10:N10").Value = ""
    End If

    ' Loop through and add all the exhaust pages
    Dim pagetop As Integer
    Dim remainder As Integer
    Dim curr_ef As Integer

    curr_ef = 4
    pagetop = 6
    remainder = num_efs Mod 3

    For i = 1 To pages - 1

        Rows(pagetop & ":" & pagetop + 45).Select
        range("S" & pagetop + 45).Activate
        selection.Copy
        Rows(pagetop + 46 & ":" & pagetop + 46).Select
        ActiveSheet.Paste

        pagetop = pagetop + 46

        'name efs
        If (name_efs = 6) Then

            range("B" & pagetop + 4).Value = "EF-" & curr_ef
            range("H" & pagetop + 4).Value = "EF-" & curr_ef + 1
            range("N" & pagetop + 4).Value = "EF-" & curr_ef + 2
            curr_ef = curr_ef + 3

        End If

        If (i = pages - 1) Then

            If (remainder = 1) Then
                range("B" & pagetop + 4 & ":B" & pagetop + 36).Select
                selection.Value = ""
                selection.Interior.ColorIndex = xlNone

                If (name_efs = 6) Then
                    range("H" & pagetop + 4).Value = "EF-" & curr_ef - 3
                End If

            End If

            If (remainder = 2 Or remainder = 1) Then
                range("N" & pagetop + 4 & ":N" & pagetop + 36).Select
                selection.Value = ""
                selection.Interior.ColorIndex = xlNone
            End If

        End If

    Next i

    ActiveSheet.PageSetup.PrintArea = "$A$6:$S$" & pagetop + 45
    range("B10").Select


    Sheets("OUTLET TEST SHEET").Copy Before:=Sheets(after + 1)
    ActiveSheet.name = "EF_INs"
    ActiveSheet.range("G8").Value = "EXHAUST"
    ActiveSheet.Tab.color = 2646607


End Sub
