Attribute VB_Name = "SortToggle"
'==============================================================================
' SortToggle - click a column header to sort it, click again to flip direction.
'
' Works with:  "Search any withot error.xlsx" (Live Search + Final Calculation
'              Sheet). The workbook's formula sort engine does the real work;
'              this macro only writes your header click into the SORT BY (AJ2)
'              and ORDER (AL2) boxes on Live Search, or sorts the data block on
'              Final Calculation Sheet.
'
' INSTALL (one minute, Excel for Windows / Mac - not Excel Online):
'   1. Open the workbook, then File > Save As > "Excel Macro-Enabled Workbook
'      (*.xlsm)". (An .xlsx file cannot store macros.)
'   2. Developer tab > Visual Basic (if there is no Developer tab: File >
'      Options > Customize Ribbon > tick "Developer").
'   3. In the VBA editor: File > Import File > pick SortToggle.bas.
'   4. Double-click the "Live Search" sheet module, paste this at the top:
'
'        Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'            SortToggle.ToggleLiveSearchSort Target
'        End Sub
'
'        Private Sub Worksheet_BeforeDoubleClick(ByVal Target As Range, Cancel As Boolean)
'            If Not Intersect(Target, Me.Range("A5:AI5")) Is Nothing Then
'                SortToggle.ResetLiveSearchSort
'                Cancel = True
'            End If
'        End Sub
'
'   5. Double-click the "Final Calculation Sheet" sheet module, paste this:
'
'        Private Sub Worksheet_SelectionChange(ByVal Target As Range)
'            SortToggle.ToggleFinalSort Target
'        End Sub
'
'   6. File > Close and Return to Excel, save, and click "Enable Content" when
'      Excel asks about macros.
'
' USE:
'   Live Search: single-click any header in row 5 (A..AI) to sort by it;
'                click the same header again to flip ascending/descending.
'                Double-click a header to go back to the original order.
'                Rows reorder inside each group; every total row stays put.
'   Final Calculation Sheet: single-click any header in row 3 (A..AG) to sort
'                the 37 data rows; click again to flip. Total rows 41..43 are
'                never touched.
'==============================================================================
Option Explicit

' Arrow glyphs are built with ChrW so this file stays plain ASCII and
' imports cleanly under any system codepage.
Private Function SortAscText() As String
    SortAscText = ChrW(&H25B2) & " Ascending"    ' up-triangle + " Ascending"
End Function

Private Function SortDescText() As String
    SortDescText = ChrW(&H25BC) & " Descending"  ' down-triangle + " Descending"
End Function

Private lastFinalCol As Long
Private lastFinalDesc As Boolean

'--- Live Search ---------------------------------------------------------------
Public Sub ToggleLiveSearchSort(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Target.Worksheet
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, ws.Range("A5:AI5")) Is Nothing Then Exit Sub

    Dim letter As String
    letter = Split(ws.Cells(1, Target.Column).Address, "$")(1)

    Dim base As String
    base = SortBaseName(ws, letter)
    If base = vbNullString Then Exit Sub

    Dim entry As String
    entry = letter & " - " & base

    Application.EnableEvents = False
    On Error GoTo done
    If CStr(ws.Range("AJ2").Value) = entry Then
        If CStr(ws.Range("AL2").Value) = SortDescText() Then
            ws.Range("AL2").Value = SortAscText()
        Else
            ws.Range("AL2").Value = SortDescText()
        End If
    Else
        ws.Range("AJ2").Value = entry
        ws.Range("AL2").Value = SortAscText()
    End If
done:
    Application.EnableEvents = True
End Sub

Public Sub ResetLiveSearchSort()
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Live Search")
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub
    Application.EnableEvents = False
    ws.Range("AJ2").Value = "(none)"
    Application.EnableEvents = True
End Sub

' Finds the plain header name for a column letter in the AJ6:AJ40 sort list.
Private Function SortBaseName(ByVal ws As Worksheet, ByVal letter As String) As String
    Dim c As Range
    Dim prefix As String
    prefix = letter & " - "
    For Each c In ws.Range("AJ6:AJ40").Cells
        If Left$(CStr(c.Value), Len(prefix)) = prefix Then
            SortBaseName = Mid$(CStr(c.Value), Len(prefix) + 1)
            Exit Function
        End If
    Next c
    SortBaseName = vbNullString
End Function

'--- Final Calculation Sheet ----------------------------------------------------
Public Sub ToggleFinalSort(ByVal Target As Range)
    Dim ws As Worksheet
    Set ws = Target.Worksheet
    If Target.CountLarge > 1 Then Exit Sub
    If Intersect(Target, ws.Range("A3:AG3")) Is Nothing Then Exit Sub

    Dim col As Long
    col = Target.Column
    Dim desc As Boolean
    If col = lastFinalCol Then
        desc = Not lastFinalDesc
    Else
        desc = False ' first click on a new column always sorts ascending
    End If

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo done
    ws.Sort.SortFields.Clear
    ws.Sort.SortFields.Add _
        Key:=ws.Range(ws.Cells(4, col), ws.Cells(40, col)), _
        SortOn:=xlSortOnValues, _
        Order:=IIf(desc, xlDescending, xlAscending), _
        DataOption:=xlSortNormal
    With ws.Sort
        .SetRange ws.Range("A4:AG40") ' data only - totals in 41..43 stay put
        .Header = xlNo
        .MatchCase = False
        .Orientation = xlTopToBottom
        .SortMethod = xlPinYin
        .Apply
    End With
    lastFinalCol = col
    lastFinalDesc = desc
done:
    Application.ScreenUpdating = True
    Application.EnableEvents = True
End Sub
