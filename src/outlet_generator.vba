Sub OutletGenerator()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "O\n14"

' System Data
Dim systems As Variant
Dim name As String
Dim outlets As Integer

Dim line_total As Integer
Dim page_start As Integer
Dim first_page_end As Integer

systems = selection
page_start = 14 'Row 14
first_page_end = 43 'Row 43

For i = 1 To UBound(systems)
    outlets = systems(i, 2)
    ' add on additional rows for each system. e.g. "Total" and blank rows
    If (outlets > 0) Then
        If (outlets = 1) Then
            line_total = line_total + 2
        Else
            line_total = line_total + outlets + 2
        End If
    End If
Next i

' Trim off the added row at the beginning and end
line_total = line_total - 1

' Overwrite default "Total" line
range("A17:I17").Copy
range("A16:I16").Select
selection.FormatConditions.Delete
ActiveSheet.Paste
range("E15").Value = ""

' Add additional rows if necessary
If (page_start + line_total > first_page_end) Then
    range("43:43").Select
    For Row = page_start + line_total - first_page_end To 1 Step -1
        selection.EntireRow.Insert
    Next Row
End If

' Populate outlet sheet data
Dim current_row As Integer
Dim current_outlet As Integer

current_row = 15
current_outlet = 1

range("D15:I15").Copy
range("P8").Select
ActiveSheet.Paste

' loop through systems
For i = 1 To UBound(systems)

    name = systems(i, 1)
    outlets = systems(i, 2)

    ' loop through outlets
    For j = 1 To outlets

        If (j = 1) Then
            range("A" & current_row).Value = name
        End If

        range("B" & current_row).Value = current_outlet
        range("P8:U8").Copy
        range("D" & current_row).Select
        ActiveSheet.Paste

        current_row = current_row + 1
        current_outlet = current_outlet + 1

    Next j

    If (outlets > 1) Then

        range("E15:I15").Copy
        range("E" & current_row).Select
        ActiveSheet.Paste

        ' add total line
        range("C" & current_row & ":D" & current_row).Merge
        range("C" & current_row & ":E" & current_row).Font.Bold = True
        range("G" & current_row & ":I" & current_row).Font.Bold = True
        range("C" & current_row).Value = "TOTAL:"
        range("E" & current_row).Formula = "=SUM(E" & current_row - outlets & ":E" & current_row - 1 & ")"
        range("E" & current_row).Copy
        Union(range("F" & current_row), range("H" & current_row)).PasteSpecial xlPasteFormulas
        current_row = current_row + 1

    End If

    If (outlets = 1) Then
        current_row = current_row - 1
        range("E" & current_row).Font.Bold = True
        range("G" & current_row & ":I" & current_row).Font.Bold = True
        Union(range("E" & current_row & ":F" & current_row), range("H" & current_row)).Value = 0
        current_row = current_row + 1
    End If

    current_row = current_row + 1
Next i

range("P8:U8").Delete

' Scroll to top and start working
ActiveWindow.ScrollRow = 1
range("C15").Value = ""
range("C15").Select

End Sub