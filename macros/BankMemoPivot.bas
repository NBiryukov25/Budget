Attribute VB_Name = "BankMemoPivot"
Option Explicit

'==============================================================================
' Bank Memo / Returned Payment summary pivot
'==============================================================================
' Rebuilds the Summary sheet's pivot table for whatever period the Data sheet
' currently holds, and re-applies the exact layout and formatting of the
' 8-6-26 report, so every month's report comes out looking the same.
'
' Three entry points (Developer > Macros, or assign one to a button):
'
'   RefreshSummaryPivot   Normal monthly run. Re-points the pivot at the
'                         current extent of the Data sheet, refreshes it, and
'                         re-applies the layout and design.
'   FormatSummaryPivot    Design only. Use when someone has hand-edited the
'                         sheet and you just want the look back.
'   RebuildSummaryPivot   Clears the Summary sheet and builds the pivot again
'                         from scratch. Use if the pivot was deleted or the
'                         field layout got scrambled.
'
' The design this reproduces:
'   Source     Data!A1:N<last row>, measured each run
'   Rows       Region > Customer > SAP No., tabular, labels repeated,
'              subtotals at the bottom of each group
'   Columns    Date grouped Years > Quarters > Months
'   Values     Count of SAP No., captioned
'              "Bank Memo Count, Returned Payment Count"
'   Totals     Row grand total on ("Totals"), column grand total off
'   Style      PivotStyleLight13, no stripes, last column emphasised
'   Type       Calibri 10 throughout, vertically centred; values centred;
'              Region and Customer left; corner block, top-right block and
'              the SAP No. header bold
'   Shading    Value cells banded by quarter. Odd quarters (Qtr1, Qtr3) are
'              light yellow, even quarters and the year/grand total columns
'              are white. A quarter-total column takes its quarter's shade.
'   Chrome     Gridlines off, 90% zoom, frozen under the header rows, value
'              columns 13.14 wide, header rows 20.1 / 20.1 / 13.5 tall
'==============================================================================


' ---- Sheet, object and field names -----------------------------------------
' Change these if the tabs or the Data header row are ever renamed.
Private Const DATA_SHEET       As String = "Data"
Private Const SUMMARY_SHEET    As String = "Summary"
Private Const PIVOT_NAME       As String = "PivotTable1"

Private Const FLD_SAP          As String = "SAP No."
Private Const FLD_REGION       As String = "Region"
Private Const FLD_CUSTOMER     As String = "Customer"
Private Const FLD_DATE         As String = "Date"

' ---- Captions ---------------------------------------------------------------
Private Const DATA_CAPTION     As String = "Bank Memo Count, Returned Payment Count"
Private Const GRAND_TOTAL_CAP  As String = "Totals"
Private Const COL_HEADER_CAP   As String = "Months"

' ---- Design -----------------------------------------------------------------
Private Const PIVOT_STYLE      As String = "PivotStyleLight13"
Private Const BODY_FONT        As String = "Calibri"
Private Const BODY_SIZE        As Double = 10

' Light yellow, RGB(255, 255, 153). Set BAND_QUARTERS to False for a plain
' white value area.
Private Const BAND_COLOR       As Long = &H99FFFF&
Private Const BAND_QUARTERS    As Boolean = True

' Excel writes quarter headers as "Qtr1" and "Qtr1 Total". A non-English Excel
' uses a different word; change this and the banding keeps working.
Private Const QUARTER_PREFIX   As String = "Qtr"

Private Const VALUE_COL_WIDTH  As Double = 13.14
Private Const ROW1_HEIGHT      As Double = 20.1
Private Const ROW2_HEIGHT      As Double = 20.1
Private Const ROW3_HEIGHT      As Double = 13.5
Private Const SHEET_ZOOM       As Long = 90


'==============================================================================
' Entry points
'==============================================================================

Public Sub RefreshSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    Set pt = FindPivot(ws)

    If pt Is Nothing Then
        Set pt = CreatePivot(ws)
    Else
        RepointPivotCache pt
    End If

    ApplyLayout pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not refresh the summary pivot." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Bank Memo Pivot"
End Sub


Public Sub FormatSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    Set pt = FindPivot(ws)

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "There is no pivot table on the " & SUMMARY_SHEET & " sheet. " & _
            "Run RebuildSummaryPivot instead."
    End If

    ApplyLayout pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not format the summary pivot." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Bank Memo Pivot"
End Sub


Public Sub RebuildSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    If MsgBox("This clears the " & SUMMARY_SHEET & " sheet and builds the " & _
              "pivot table again from the " & DATA_SHEET & " sheet." & _
              vbCrLf & vbCrLf & "Continue?", _
              vbOKCancel + vbQuestion, "Bank Memo Pivot") <> vbOK Then Exit Sub

    On Error GoTo Failed
    Application.ScreenUpdating = False

    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    ws.Cells.Clear

    Set pt = CreatePivot(ws)
    ApplyLayout pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not rebuild the summary pivot." & vbCrLf & vbCrLf & _
           Err.Description, vbExclamation, "Bank Memo Pivot"
End Sub


'==============================================================================
' Source data and pivot creation
'==============================================================================

' The Data sheet grows every month, so the source range is measured rather than
' remembered. Column A (SAP No.) is populated on every real row.
Private Function DataSourceRange() As Range
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long

    Set ws = ThisWorkbook.Worksheets(DATA_SHEET)
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        Err.Raise vbObjectError + 2, , _
            "The " & DATA_SHEET & " sheet has a header row but no data rows."
    End If

    Set DataSourceRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function


Private Function FindPivot(ws As Worksheet) As PivotTable
    On Error Resume Next
    Set FindPivot = ws.PivotTables(PIVOT_NAME)
    If FindPivot Is Nothing Then
        If ws.PivotTables.Count > 0 Then Set FindPivot = ws.PivotTables(1)
    End If
    On Error GoTo 0
End Function


' Always hands the pivot a cache built over the current extent of the Data
' sheet. That discards the date grouping, which ApplyLayout puts back, so the
' grouping is rebuilt from the same rules every run instead of being inherited.
Private Sub RepointPivotCache(pt As PivotTable)
    pt.ChangePivotCache ThisWorkbook.PivotCaches.Create( _
        SourceType:=xlDatabase, SourceData:=DataSourceRange())
    pt.RefreshTable
End Sub


Private Function CreatePivot(ws As Worksheet) As PivotTable
    Set CreatePivot = ThisWorkbook.PivotCaches.Create( _
            SourceType:=xlDatabase, SourceData:=DataSourceRange()) _
        .CreatePivotTable(TableDestination:=ws.Range("A1"), _
                          TableName:=PIVOT_NAME)
End Function


'==============================================================================
' Layout: which fields sit where, and how the pivot behaves
'==============================================================================

Private Sub ApplyLayout(pt As PivotTable)
    ' Rows: Region > Customer > SAP No.
    PlaceField pt, FLD_REGION, xlRowField, 1
    PlaceField pt, FLD_CUSTOMER, xlRowField, 2
    PlaceField pt, FLD_SAP, xlRowField, 3

    ' Region and Customer each get a subtotal row; the innermost field does not.
    SetAutoSubtotal pt.PivotFields(FLD_REGION), True
    SetAutoSubtotal pt.PivotFields(FLD_CUSTOMER), True
    SetAutoSubtotal pt.PivotFields(FLD_SAP), False

    pt.PivotFields(FLD_REGION).AutoSort xlAscending, FLD_REGION
    pt.PivotFields(FLD_CUSTOMER).AutoSort xlAscending, FLD_CUSTOMER

    ' Columns: Date grouped into Years > Quarters > Months
    EnsureDateGrouping pt

    ' Values: count of SAP No.
    SetValueField pt

    With pt
        .RowAxisLayout xlTabularRow
        .RepeatAllLabels xlRepeatLabels
        .SubtotalLocation xlAtBottom

        .RowGrand = True
        .ColumnGrand = False
        .GrandTotalName = GRAND_TOTAL_CAP
        .ColumnHeaderCaption = COL_HEADER_CAP

        .ShowDrillIndicators = False

        ' Keep our formatting and column widths across a plain Refresh Data.
        .PreserveFormatting = True
        .HasAutoFormat = False

        .TableStyle2 = PIVOT_STYLE
        .ShowTableStyleRowHeaders = True
        .ShowTableStyleColumnHeaders = True
        .ShowTableStyleRowStripes = False
        .ShowTableStyleColumnStripes = False
        .ShowTableStyleLastColumn = True
    End With
End Sub


' SAP No. sits in the row area and the value area at the same time. An existing
' count over it is reused rather than dropped and re-added, so a repeat run does
' not disturb the field it is already reporting on.
Private Sub SetValueField(pt As PivotTable)
    Dim i As Long
    Dim df As PivotField

    For i = pt.DataFields.Count To 1 Step -1
        If pt.DataFields(i).SourceName = FLD_SAP Then
            Set df = pt.DataFields(i)
        Else
            pt.DataFields(i).Orientation = xlHidden
        End If
    Next i

    If df Is Nothing Then
        Set df = pt.AddDataField(pt.PivotFields(FLD_SAP), DATA_CAPTION, xlCount)
    End If

    With df
        .Function = xlCount     ' resets the caption, so set the caption after
        .Caption = DATA_CAPTION
        .NumberFormat = "General"
    End With
End Sub


Private Sub PlaceField(pt As PivotTable, fieldName As String, _
                       axis As XlPivotFieldOrientation, slot As Long)
    With pt.PivotFields(fieldName)
        If .Orientation <> axis Then .Orientation = axis
        .Position = slot
    End With
End Sub


' Subtotals(1) is the Automatic slot. Setting it True clears the eleven
' explicit functions; turning subtotals off means clearing all twelve.
Private Sub SetAutoSubtotal(pf As PivotField, wanted As Boolean)
    Dim i As Long

    If wanted Then
        pf.Subtotals(1) = True
    Else
        For i = 1 To 12
            pf.Subtotals(i) = False
        Next i
    End If
End Sub


' Puts the Date field on the column axis grouped by Months, Quarters and Years,
' then orders those three levels Years > Quarters > Months.
'
' Excel names the generated fields "Months (Date)" / "Quarters (Date)" /
' "Years (Date)" in some versions and plain "Months" / "Quarters" / "Years" in
' others, and in older versions it renames the source field rather than adding
' a third, so they are looked up by prefix rather than by exact name.
Private Sub EnsureDateGrouping(pt As PivotTable)
    Dim pfYears As PivotField, pfQuarters As PivotField, pfMonths As PivotField

    Set pfYears = FieldStartingWith(pt, "Years")
    Set pfQuarters = FieldStartingWith(pt, "Quarters")
    Set pfMonths = FieldStartingWith(pt, "Months")

    If pfYears Is Nothing Or pfQuarters Is Nothing Or pfMonths Is Nothing Then
        GroupDateField pt
        Set pfYears = FieldStartingWith(pt, "Years")
        Set pfQuarters = FieldStartingWith(pt, "Quarters")
        Set pfMonths = FieldStartingWith(pt, "Months")
    End If

    If pfYears Is Nothing Or pfQuarters Is Nothing Or pfMonths Is Nothing Then
        Err.Raise vbObjectError + 3, , _
            "Could not group the " & FLD_DATE & " column into Years, " & _
            "Quarters and Months. Check that every value in the " & _
            FLD_DATE & " column of the " & DATA_SHEET & " sheet is a real " & _
            "date and not text."
    End If

    ' Where Excel kept a separate ungrouped Date field, it stays off the pivot.
    On Error Resume Next
    pt.PivotFields(FLD_DATE).Orientation = xlHidden
    On Error GoTo 0

    PlaceField pt, pfYears.Name, xlColumnField, 1
    PlaceField pt, pfQuarters.Name, xlColumnField, 2
    PlaceField pt, pfMonths.Name, xlColumnField, 3

    ' Quarter subtotals are the "Qtr n Total" columns and year subtotals the
    ' "<year> Total" column. Months, being innermost, get none.
    SetAutoSubtotal pfYears, True
    SetAutoSubtotal pfQuarters, True
    SetAutoSubtotal pfMonths, False
End Sub


Private Sub GroupDateField(pt As PivotTable)
    Dim pf As PivotField

    Set pf = pt.PivotFields(FLD_DATE)
    pf.Orientation = xlColumnField
    pf.Position = 1

    ' Excel 2016 and later auto-group a date field the moment it lands on an
    ' axis. If that already produced all three levels there is nothing to do,
    ' and calling Group again would fail.
    If HasAllDateLevels(pt) Then Exit Sub

    ' Auto-grouping may instead have produced a partial hierarchy, which Group
    ' also refuses to sit on top of. Flatten it back to raw dates first.
    On Error Resume Next
    DateAxisField(pt).DataRange.Cells(1, 1).Ungroup
    On Error GoTo 0

    Set pf = DateAxisField(pt)
    If pf Is Nothing Then Exit Sub      ' EnsureDateGrouping reports the failure
    If pf.Orientation <> xlColumnField Then pf.Orientation = xlColumnField

    ' Periods: Seconds, Minutes, Hours, Days, Months, Quarters, Years
    pf.DataRange.Cells(1, 1).Group _
        Start:=True, End:=True, _
        Periods:=Array(False, False, False, False, True, True, True)
End Sub


Private Function HasAllDateLevels(pt As PivotTable) As Boolean
    HasAllDateLevels = Not FieldStartingWith(pt, "Years") Is Nothing _
                  And Not FieldStartingWith(pt, "Quarters") Is Nothing _
                  And Not FieldStartingWith(pt, "Months") Is Nothing
End Function


' The field carrying the dates, whatever grouping has done to its name.
Private Function DateAxisField(pt As PivotTable) As PivotField
    On Error Resume Next
    Set DateAxisField = pt.PivotFields(FLD_DATE)
    On Error GoTo 0

    If DateAxisField Is Nothing Then Set DateAxisField = FieldStartingWith(pt, "Months")
    If DateAxisField Is Nothing Then Set DateAxisField = FieldStartingWith(pt, "Quarters")
    If DateAxisField Is Nothing Then Set DateAxisField = FieldStartingWith(pt, "Years")
End Function


Private Function FieldStartingWith(pt As PivotTable, prefix As String) As PivotField
    Dim pf As PivotField

    For Each pf In pt.PivotFields
        If InStr(1, pf.Name, prefix, vbTextCompare) = 1 Then
            Set FieldStartingWith = pf
            Exit Function
        End If
    Next pf
End Function


'==============================================================================
' Formatting: type, alignment and the quarter banding
'==============================================================================

Private Sub ApplyFormatting(pt As PivotTable)
    Dim ws As Worksheet
    Dim tr As Range, body As Range
    Dim headerRows As Long, labelCols As Long, buttonCols As Long
    Dim firstRow As Long, firstCol As Long, lastCol As Long, lastRow As Long
    Dim c As Long

    Set ws = pt.Parent
    Set tr = pt.TableRange1
    Set body = pt.DataBodyRange
    If body Is Nothing Then Exit Sub

    firstRow = tr.Row
    firstCol = tr.Column
    lastCol = firstCol + tr.Columns.Count - 1
    lastRow = body.Row + body.Rows.Count - 1

    headerRows = body.Row - tr.Row          ' 4: buttons, year, quarter, month
    labelCols = body.Column - tr.Column     ' 3: Region, Customer, SAP No.
    buttonCols = pt.ColumnFields.Count      ' 3: one button cell per column field

    ' Baseline for the whole table.
    With tr
        With .Font
            .Name = BODY_FONT
            .Size = BODY_SIZE
            .Bold = False
            .Italic = False
        End With
        .VerticalAlignment = xlCenter
        .HorizontalAlignment = xlGeneral
        .Interior.Pattern = xlNone
    End With

    ' Corner block above the row labels: bold and centred.
    With ws.Range(ws.Cells(firstRow, firstCol), _
                  ws.Cells(firstRow + headerRows - 2, firstCol + labelCols - 1))
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
    End With

    ' The row-label header row: Region left, SAP No. bold.
    ws.Cells(firstRow + headerRows - 1, firstCol).HorizontalAlignment = xlLeft
    ws.Cells(firstRow + headerRows - 1, firstCol + labelCols - 1).Font.Bold = True

    ' Top-right block: the column-field buttons stay plain, everything to the
    ' right of them is bold.
    ws.Range(ws.Cells(firstRow, firstCol + labelCols), _
             ws.Cells(firstRow, lastCol)).HorizontalAlignment = xlCenter

    If lastCol > firstCol + labelCols + buttonCols - 1 Then
        ws.Range(ws.Cells(firstRow, firstCol + labelCols + buttonCols), _
                 ws.Cells(firstRow, lastCol)).Font.Bold = True
    End If

    ' Year, quarter and month header rows.
    ws.Range(ws.Cells(firstRow + 1, firstCol + labelCols), _
             ws.Cells(firstRow + headerRows - 1, lastCol)) _
      .HorizontalAlignment = xlCenter

    ' Body: Region and Customer left, values centred.
    ws.Range(ws.Cells(body.Row, firstCol), _
             ws.Cells(lastRow, firstCol + labelCols - 2)) _
      .HorizontalAlignment = xlLeft

    body.HorizontalAlignment = xlCenter
    body.NumberFormat = "General"

    ' Quarter banding.
    For c = body.Column To body.Column + body.Columns.Count - 1
        With ws.Range(ws.Cells(body.Row, c), ws.Cells(lastRow, c)).Interior
            .Pattern = xlSolid
            If ShadeColumn(pt, c) Then
                .Color = BAND_COLOR
            Else
                .Color = vbWhite
            End If
        End With
    Next c
End Sub


' A value column is shaded when it belongs to an odd-numbered quarter. A
' quarter-total column inherits its quarter's shade because its header reads
' "Qtr n Total"; the year-total and grand-total columns carry no quarter header
' at all and so stay white.
Private Function ShadeColumn(pt As PivotTable, col As Long) As Boolean
    Dim q As Long

    If Not BAND_QUARTERS Then Exit Function

    q = QuarterOfColumn(pt, col)
    ShadeColumn = (q > 0) And ((q Mod 2) = 1)
End Function


Private Function QuarterOfColumn(pt As PivotTable, col As Long) As Long
    Dim ws As Worksheet
    Dim r As Long, hit As Long
    Dim caption As String

    Set ws = pt.Parent

    For r = pt.TableRange1.Row To pt.DataBodyRange.Row - 1
        caption = CStr(ws.Cells(r, col).Value)
        hit = InStr(1, caption, QUARTER_PREFIX, vbTextCompare)
        If hit > 0 Then
            QuarterOfColumn = Val(Mid$(caption, hit + Len(QUARTER_PREFIX)))
            Exit Function
        End If
    Next r
End Function


'==============================================================================
' Sheet chrome: widths, heights, freeze, gridlines, zoom
'==============================================================================

Private Sub ApplyChrome(ws As Worksheet, pt As PivotTable)
    Dim tr As Range, body As Range
    Dim labelCols As Long, firstCol As Long, lastCol As Long

    Set tr = pt.TableRange1
    Set body = pt.DataBodyRange
    If body Is Nothing Then Exit Sub

    firstCol = tr.Column
    lastCol = firstCol + tr.Columns.Count - 1
    labelCols = body.Column - tr.Column

    ' Row-label columns size to their contents. Value columns are fixed, so the
    ' table stays the same width whatever month names are on show.
    ws.Range(ws.Columns(firstCol), ws.Columns(firstCol + labelCols - 1)).AutoFit
    ws.Range(ws.Columns(firstCol + labelCols), _
             ws.Columns(lastCol)).ColumnWidth = VALUE_COL_WIDTH

    ws.Rows(tr.Row).RowHeight = ROW1_HEIGHT
    ws.Rows(tr.Row + 1).RowHeight = ROW2_HEIGHT
    ws.Rows(tr.Row + 2).RowHeight = ROW3_HEIGHT

    ' Gridlines, zoom and freeze panes are window properties, so the sheet has
    ' to be the active one while they are set.
    Application.Goto ws.Cells(1, 1), True
    With ActiveWindow
        .DisplayGridlines = False
        .Zoom = SHEET_ZOOM
        .FreezePanes = False
        ws.Cells(body.Row, firstCol).Select
        .FreezePanes = True
    End With

    ws.Cells(body.Row, body.Column).Select
End Sub
