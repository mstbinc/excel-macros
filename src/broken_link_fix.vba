Sub BrokenLinkFix()
Attribute MyFunction.VB_ProcData.VB_Invoke_Func = "L\n14"

    Dim current As String
    current = ActiveWorkbook.ActiveSheet.name

    For Each aSheet In ActiveWorkbook.Worksheets
         aSheet.Activate
         cells.Replace What:="'*[*]", Replacement:="'", LookAt:=xlPart, SearchOrder:= _
         xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
    Next aSheet

    Worksheets(current).Activate

End Sub