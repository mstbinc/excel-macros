Sub AHUSetup()

    Dim index As Integer ' place sheets after last colored tab
    Dim do_sp As Integer ' set up static profile or not?
    Dim ahu_name As String


    ahu_name = InputBox("What do you want to call the AHU?", "Number of Exhaust Fans", "AHU-1")
    do_sp = MsgBox("Include a generic static profile?", vbYesNo, "Use Generic Static Profile?")

    For i = 1 To Sheets.Count
        index = i
        If Sheets(i).Tab.color = "False" Then Exit For
    Next i

    ' Head sheet
    Sheets("AIRAPPTR DATA").Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name
    Range("B10").Value = ahu_name
    Range("B20").Formula = "=B19-B21"
    If InStr(ahu_name, "RTU") Then
        Range("B11").Value = "ROOF"
    End If
    set_tab_color

    index = index + 1

    ' Exhaust fan
    Sheets("FANTEST").Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name & " EF"
    Range("N10:S33").Select
    Selection.ClearContents
    Selection.FormatConditions.Delete
    Range("B10:G33").Select
    Selection.ClearContents
    Selection.FormatConditions.Delete

    Range("H10").Select
    set_tab_color

    index = index + 1

    ' Static Profile
    Dim sp_name As String
    sp_name = "STATIC PROFILE - BLANK"
    If do_sp = vbYes Then
        sp_name = "STATIC PROFILE - GENERIC"
    End If

    Sheets(sp_name).Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name & " SP"
    Worksheets(ahu_name).Range("B10").Copy
    Range("F45").Select
    ActiveSheet.Paste Link:=True
    set_tab_color

    index = index + 1

    ' FP Boxes
    Sheets("FP BOXES (CFM)").Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name & " FP BOXES"
    set_tab_color

    index = index + 1

    ' VAV Boxes
    Sheets("BOXES (CFM)").Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name & " BOXES"
    set_tab_color

    index = index + 1

    ' Outlets
    Sheets("OUTLET TEST SHEET").Copy Before:=Sheets(index)
    ActiveSheet.name = ahu_name & " OUTLETS"
    set_tab_color

    index = index + 1


    ' Go back to the beginning
    Sheets(ahu_name).Select
    Range("B10").Activate

End Sub

Function set_tab_color()
    ActiveSheet.Tab.color = RGB(0, 112, 192)
End Function
