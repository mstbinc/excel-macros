Sub CopySelectedSystem()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "N\n14"
'
' nextTabGroup Macro
' Takes selected worksheets and duplicates the group while incrementing the unit number
'
' Keyboard Shortcut: Ctrl+Shift+N
'

Dim tabGroup As New Collection
Dim regExp As Object
Dim matches As Object
Dim tabIndex As Integer
Dim tabName As String
Dim newTabName As String
Dim tabArray() As String
Dim tabColors(1 To 8) As String
Dim prediction As String
Dim prediction_msg As String

tabColors(1) = "12611584"
tabColors(2) = "255"
tabColors(3) = "15773696"
tabColors(4) = "4626167"
tabColors(5) = "10498160"
tabColors(6) = "6684927"
tabColors(7) = "16776960"
tabColors(8) = "477335"

prediction = ""

Set regExp = CreateObject("vbscript.regexp")
regExp.Pattern = "^(.*?)(\d+)$"
regExp.ignorecase = True
regExp.Global = True

tabName = ActiveSheet.name

Set matches = regExp.Execute(tabName)

If (matches.Count > 0) Then
    prediction = matches(0).submatches(0) & (matches(0).submatches(1) + 1)
End If

' Loop through each worksheet in workbook and collect ones that belong to the selected tab name
For Each ws In Application.Worksheets
    If (InStr(ws.name, tabName)) Then
        tabGroup.Add ws ' No params means pass val by reference
    End If
Next ws

If (prediction = "") Then
    prediction_msg = "No prediction available"
    prediction = tabName
Else
    prediction_msg = prediction
End If

newTabName = InputBox("What would you like to name the new tab group? My prediction: " & prediction_msg, "Tab Group Name", prediction)
tabIndex = tabGroup(tabGroup.Count).Index

ReDim tabArray(1 To tabGroup.Count)
Dim i As Integer

i = 1

For Each ws In tabGroup
    tabArray(i) = ws.name
    i = i + 1
Next ws

Sheets(tabArray).Copy after:=Sheets(tabIndex)

i = 1

Dim color As String

For Each s In tabArray

    tabIndex = tabIndex + 1
    Sheets(tabIndex).name = Replace(tabArray(i), tabName, newTabName)

    Dim color_int As Integer
    color_int = in_array(tabColors, Sheets(tabIndex).Tab.color)

    If (color_int > 0) Then
        If (color_int + 1 <= 8) Then
            Sheets(tabIndex).Tab.color = tabColors(color_int + 1)
        Else
            Sheets(tabIndex).Tab.color = tabColors(1)
        End If
    Else
        Sheets(tabIndex).Tab.color = tabColors(1)
    End If

    i = i + 1
Next s

Worksheets(newTabName).range("B10").Value = newTabName
Worksheets(newTabName).range("B10").Activate

End Sub

Function in_array(ByRef arr() As String, ByVal test As String) As Integer

Dim i As Integer

i = 1

For Each s In arr

If test = s Then

in_array = i
Exit Function

End If

i = i + 1

Next s

in_array = -1
Exit Function


End Function


