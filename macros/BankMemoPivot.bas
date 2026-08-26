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
' DATA_SHEET is only the first place to look. If no sheet by that name has the
' right header row, every other sheet is searched, so the source tab can be
' called "Data" in one month's file and "Sheet1" in the next.
Private Const DATA_SHEET       As String = "Data"
Private Const SUMMARY_SHEET    As String = "Summary"
Private Const PIVOT_NAME       As String = "PivotTable1"

' Field names are matched loosely: case, spaces and punctuation are ignored, so
' "SAP No." and "SAP No" both resolve to the same column.
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

' Date grouping invents an item for every month of the year plus one either
' side of the range, so periods with no records would show as empty columns.
' The 8-6-26 report has those hidden. Set False to show every period.
Private Const HIDE_EMPTY_PERIODS As Boolean = True

Private Const VALUE_COL_WIDTH  As Double = 13.14
Private Const ROW1_HEIGHT      As Double = 20.1
Private Const ROW2_HEIGHT      As Double = 20.1
Private Const ROW3_HEIGHT      As Double = 13.5
Private Const SHEET_ZOOM       As Long = 90

' ---- Version ----------------------------------------------------------------
' Stamped into every message this module shows, so there is never any doubt
' about which copy of the code a workbook is actually running.
Private Const MACRO_VERSION    As String = "v5"

' ---- Run state ---------------------------------------------------------------
' gStep names the phase in progress so a failure can say where it happened.
' gSkipped collects settings this Excel would not accept, which are reported
' together at the end rather than being allowed to abort the run.
Private gStep                  As String
Private gSkipped               As String


'==============================================================================
' Entry points
'==============================================================================

Public Sub RefreshSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error GoTo Failed
    Application.ScreenUpdating = False
    BeginRun

    gStep = "finding the Summary sheet and its pivot"
    Set ws = SummarySheet()
    Set pt = FindPivot(ws)

    If pt Is Nothing Then
        gStep = "creating the pivot table"
        Set pt = CreatePivot(ws)
    Else
        gStep = "re-pointing the pivot at the source sheet"
        RepointPivotCache pt
    End If

    ApplyLayout pt
    gStep = "moving the pivot to cell A1"
    MoveToOrigin pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    ReportSkipped
    WarnIfRegionEmpty
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not refresh the summary pivot." & vbCrLf & vbCrLf & _
           "Step: " & gStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Bank Memo Pivot " & MACRO_VERSION
End Sub


Public Sub FormatSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    On Error GoTo Failed
    Application.ScreenUpdating = False
    BeginRun

    Set ws = SummarySheet()
    Set pt = FindPivot(ws)

    If pt Is Nothing Then
        Err.Raise vbObjectError + 1, , _
            "There is no pivot table on the " & SUMMARY_SHEET & " sheet. " & _
            "Run RebuildSummaryPivot instead."
    End If

    ApplyLayout pt
    gStep = "moving the pivot to cell A1"
    MoveToOrigin pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    ReportSkipped
    WarnIfRegionEmpty
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not format the summary pivot." & vbCrLf & vbCrLf & _
           "Step: " & gStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Bank Memo Pivot " & MACRO_VERSION
End Sub


Public Sub RebuildSummaryPivot()
    Dim ws As Worksheet
    Dim pt As PivotTable

    If MsgBox("This clears the " & SUMMARY_SHEET & " sheet and builds the " & _
              "pivot table again from the source sheet." & _
              vbCrLf & vbCrLf & "Continue?", _
              vbOKCancel + vbQuestion, "Bank Memo Pivot") <> vbOK Then Exit Sub

    On Error GoTo Failed
    Application.ScreenUpdating = False
    BeginRun

    Set ws = SummarySheet()
    ws.Cells.Clear

    Set pt = CreatePivot(ws)
    ApplyLayout pt
    gStep = "moving the pivot to cell A1"
    MoveToOrigin pt
    ApplyFormatting pt
    ApplyChrome ws, pt

    Application.ScreenUpdating = True
    ReportSkipped
    WarnIfRegionEmpty
    Exit Sub

Failed:
    Application.ScreenUpdating = True
    MsgBox "Could not rebuild the summary pivot." & vbCrLf & vbCrLf & _
           "Step: " & gStep & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "Bank Memo Pivot " & MACRO_VERSION
End Sub


'==============================================================================
' Source data and pivot creation
'==============================================================================

' The source sheet grows every month, so its extent is measured rather than
' remembered. The SAP column is populated on every real row, so it is the one
' that decides where the data ends.
Private Function DataSourceRange() As Range
    Dim ws As Worksheet
    Dim lastRow As Long, lastCol As Long, keyCol As Long

    Set ws = ResolveDataSheet()
    keyCol = HeaderColumn(ws, FLD_SAP)
    lastRow = ws.Cells(ws.Rows.Count, keyCol).End(xlUp).Row
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    If lastRow < 2 Then
        Err.Raise vbObjectError + 2, , _
            "The " & ws.Name & " sheet has a header row but no data rows."
    End If

    Set DataSourceRange = ws.Range(ws.Cells(1, 1), ws.Cells(lastRow, lastCol))
End Function


' The source tab is called "Data" in one month's file and "Sheet1" in another,
' so it is identified by its header row rather than by its name.
Private Function ResolveDataSheet() As Worksheet
    Dim ws As Worksheet

    On Error Resume Next
    Set ResolveDataSheet = ThisWorkbook.Worksheets(DATA_SHEET)
    On Error GoTo 0

    If Not ResolveDataSheet Is Nothing Then
        If SheetHasHeaders(ResolveDataSheet) Then Exit Function
        Set ResolveDataSheet = Nothing
    End If

    For Each ws In ThisWorkbook.Worksheets
        If ws.Name <> SUMMARY_SHEET Then
            If SheetHasHeaders(ws) Then
                Set ResolveDataSheet = ws
                Exit Function
            End If
        End If
    Next ws

    Err.Raise vbObjectError + 5, , _
        "Could not find the source sheet. One tab needs a header row in " & _
        "row 1 with columns called " & FLD_SAP & ", " & FLD_REGION & ", " & _
        FLD_CUSTOMER & " and " & FLD_DATE & "."
End Function


Private Function SheetHasHeaders(ws As Worksheet) As Boolean
    SheetHasHeaders = HeaderColumn(ws, FLD_SAP) > 0 _
                 And HeaderColumn(ws, FLD_REGION) > 0 _
                 And HeaderColumn(ws, FLD_CUSTOMER) > 0 _
                 And HeaderColumn(ws, FLD_DATE) > 0
End Function


Private Function HeaderColumn(ws As Worksheet, wanted As String) As Long
    Dim lastCol As Long, c As Long, target As String

    target = Canon(wanted)
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If lastCol > 256 Then lastCol = 256

    For c = 1 To lastCol
        If Canon(CStr(ws.Cells(1, c).Value)) = target Then
            HeaderColumn = c
            Exit Function
        End If
    Next c
End Function


' Strips everything but letters and digits and lower-cases the rest, so
' "SAP No." and "SAP No" compare equal.
Private Function Canon(text As String) As String
    Dim i As Long, ch As String, out As String

    For i = 1 To Len(text)
        ch = LCase$(Mid$(text, i, 1))
        If (ch >= "0" And ch <= "9") Or (ch >= "a" And ch <= "z") Then
            out = out & ch
        End If
    Next i

    Canon = out
End Function


' Finds a pivot field by loose name match, so a header row that says "SAP No"
' rather than "SAP No." still resolves.
Private Function PF(pt As PivotTable, wanted As String) As PivotField
    Dim pf2 As PivotField
    Dim target As String

    On Error Resume Next
    Set PF = pt.PivotFields(wanted)
    On Error GoTo 0
    If Not PF Is Nothing Then Exit Function

    target = Canon(wanted)
    For Each pf2 In pt.PivotFields
        If Canon(pf2.Name) = target Then
            Set PF = pf2
            Exit Function
        End If
    Next pf2

    Err.Raise vbObjectError + 4, , _
        "The pivot has no field matching '" & wanted & "'. The source " & _
        "sheet needs a column with that heading."
End Function


' "Subscript out of range" from a bare Worksheets("Summary") says nothing about
' what is missing, so the tab is looked up here and named in the failure.
Private Function SummarySheet() As Worksheet
    On Error Resume Next
    Set SummarySheet = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    On Error GoTo 0

    If SummarySheet Is Nothing Then
        Err.Raise vbObjectError + 6, , _
            "This workbook has no tab called '" & SUMMARY_SHEET & "'. " & _
            "Rename the tab the pivot belongs on to " & SUMMARY_SHEET & _
            ", or change SUMMARY_SHEET at the top of the module."
    End If
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
                          TableName:=FreePivotName())
End Function


' Pivot table names are workbook-wide. If PIVOT_NAME is already taken by a
' pivot on another sheet, creating a second one under that name fails, so a
' free variant is found instead.
Private Function FreePivotName() As String
    Dim candidate As String
    Dim n As Long

    candidate = PIVOT_NAME
    Do While PivotNameTaken(candidate)
        n = n + 1
        candidate = PIVOT_NAME & "_" & n
        If n > 100 Then Exit Do
    Loop

    FreePivotName = candidate
End Function


Private Function PivotNameTaken(candidate As String) As Boolean
    Dim ws As Worksheet
    Dim pt As PivotTable

    For Each ws In ThisWorkbook.Worksheets
        For Each pt In ws.PivotTables
            If StrComp(pt.Name, candidate, vbTextCompare) = 0 Then
                PivotNameTaken = True
                Exit Function
            End If
        Next pt
    Next ws
End Function


'==============================================================================
' Layout: which fields sit where, and how the pivot behaves
'==============================================================================

Private Sub ApplyLayout(pt As PivotTable)
    gStep = "placing the row fields"
    ' Rows: Region > Customer > SAP No.
    PlaceField PF(pt, FLD_REGION), xlRowField, 1
    PlaceField PF(pt, FLD_CUSTOMER), xlRowField, 2
    PlaceField PF(pt, FLD_SAP), xlRowField, 3

    ' Region and Customer each get a subtotal row; the innermost field does not.
    SetAutoSubtotal PF(pt, FLD_REGION), True
    SetAutoSubtotal PF(pt, FLD_CUSTOMER), True
    SetAutoSubtotal PF(pt, FLD_SAP), False

    OptSortAscending PF(pt, FLD_REGION)
    OptSortAscending PF(pt, FLD_CUSTOMER)

    ' Columns: Date grouped into Years > Quarters > Months
    gStep = "grouping the Date column into years, quarters and months"
    EnsureDateGrouping pt

    ' Values: count of SAP No.
    gStep = "adding the count of " & FLD_SAP
    SetValueField pt

    ' Every option below is applied on its own. Excel builds differ in which
    ' of these they expose, and one unsupported property should not cost the
    ' whole run, so anything that will not take is collected and reported at
    ' the end instead of raising.
    gStep = "applying the pivot layout options"
    OptCall pt, "RowAxisLayout", xlTabularRow
    OptCall pt, "RepeatAllLabels", xlRepeatLabels
    OptCall pt, "SubtotalLocation", xlAtBottom

    gStep = "applying the pivot total options"
    OptSet pt, "RowGrand", True
    OptSet pt, "ColumnGrand", False
    OptSet pt, "GrandTotalName", GRAND_TOTAL_CAP
    OptSet pt, "ColumnHeaderCaption", COL_HEADER_CAP
    OptSet pt, "ShowDrillIndicators", False

    ' Keep our formatting and column widths across a plain Refresh Data.
    OptSet pt, "PreserveFormatting", True
    OptSet pt, "HasAutoFormat", False

    gStep = "applying the pivot table style"
    OptSet pt, "TableStyle2", PIVOT_STYLE
    OptSet pt, "ShowTableStyleRowHeaders", True
    OptSet pt, "ShowTableStyleColumnHeaders", True
    OptSet pt, "ShowTableStyleRowStripes", False
    OptSet pt, "ShowTableStyleColumnStripes", False
    OptSet pt, "ShowTableStyleLastColumn", True

    gStep = "expanding the year, quarter and month levels"
    ExpandAllLevels pt
End Sub


'==============================================================================
' Expanding the hierarchy
'==============================================================================

' A pivot whose year or quarter items are collapsed shows that level's totals
' and nothing beneath, so the month columns never appear. The 8-6-26 report has
' every level expanded, and since the design turns drill indicators off there is
' no +/- left to open them by hand, so they are opened here.
Private Sub ExpandAllLevels(pt As PivotTable)
    Dim pfYears As PivotField, pfQuarters As PivotField, pfMonths As PivotField

    ExpandField PF(pt, FLD_REGION)
    ExpandField PF(pt, FLD_CUSTOMER)

    Set pfYears = FieldStartingWith(pt, "Years")
    Set pfQuarters = FieldStartingWith(pt, "Quarters")
    Set pfMonths = FieldStartingWith(pt, "Months")

    If Not pfYears Is Nothing Then ExpandField pfYears
    If Not pfQuarters Is Nothing Then ExpandField pfQuarters

    If HIDE_EMPTY_PERIODS Then
        gStep = "hiding periods with no records"
        If Not pfMonths Is Nothing Then HideEmptyItems pfMonths
        If Not pfQuarters Is Nothing Then HideEmptyItems pfQuarters
        If Not pfYears Is Nothing Then HideEmptyItems pfYears
    End If
End Sub


Private Sub ExpandField(pf As PivotField)
    Dim pi As PivotItem

    If pf Is Nothing Then Exit Sub

    On Error Resume Next
    pf.ShowDetail = True

    If Err.Number <> 0 Then
        ' Some builds only take it one item at a time.
        Err.Clear
        For Each pi In pf.PivotItems
            pi.ShowDetail = True
            Err.Clear
        Next pi
    End If
    On Error GoTo 0
End Sub


' Grouping a date field creates all twelve months whether or not the data
' reaches them, plus a "<" and a ">" item either side of the range. Anything
' with no records behind it is hidden so the table only spans real periods.
Private Sub HideEmptyItems(pf As PivotField)
    Dim pi As PivotItem
    Dim populated As Long

    On Error Resume Next

    ' Count first: hiding every item in a field is not allowed, so if nothing
    ' has records the field is left alone.
    For Each pi In pf.PivotItems
        If pi.RecordCount > 0 Then populated = populated + 1
        Err.Clear
    Next pi

    If populated = 0 Then
        On Error GoTo 0
        Exit Sub
    End If

    For Each pi In pf.PivotItems
        If pi.RecordCount > 0 Then
            If pi.Visible <> True Then pi.Visible = True
        Else
            If pi.Visible <> False Then pi.Visible = False
        End If
        Err.Clear
    Next pi

    On Error GoTo 0
End Sub


'==============================================================================
' Fail-soft setters
'==============================================================================

Private Sub BeginRun()
    gStep = "starting"
    gSkipped = ""
End Sub


' Sets a property by name. A property this Excel does not expose is noted and
' skipped rather than stopping the run.
Private Sub OptSet(target As Object, propertyName As String, value As Variant)
    On Error Resume Next
    CallByName target, propertyName, VbLet, value
    If Err.Number <> 0 Then
        Note propertyName, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


' Same, for the settings Excel exposes as one-argument methods rather than
' properties.
Private Sub OptCall(target As Object, methodName As String, argument As Variant)
    On Error Resume Next
    CallByName target, methodName, VbMethod, argument
    If Err.Number <> 0 Then
        Note methodName, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


Private Sub OptSortAscending(pf As PivotField)
    On Error Resume Next
    pf.AutoSort xlAscending, pf.Name
    If Err.Number <> 0 Then
        Note "sorting " & pf.Name, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


Private Sub Note(what As String, why As String)
    gSkipped = gSkipped & vbCrLf & "    " & what & " - " & why
End Sub


Private Sub ReportSkipped()
    If Len(gSkipped) = 0 Then Exit Sub

    MsgBox "The pivot was rebuilt and formatted, but this copy of Excel " & _
           "would not accept these settings:" & vbCrLf & gSkipped & _
           vbCrLf & vbCrLf & "Everything else was applied.", _
           vbInformation, "Bank Memo Pivot"
End Sub


' SAP No. sits in the row area and the value area at the same time. An existing
' count over it is reused rather than dropped and re-added, so a repeat run does
' not disturb the field it is already reporting on.
Private Sub SetValueField(pt As PivotTable)
    Dim i As Long
    Dim df As PivotField

    For i = pt.DataFields.Count To 1 Step -1
        If Canon(pt.DataFields(i).SourceName) = Canon(FLD_SAP) Then
            Set df = pt.DataFields(i)
        Else
            pt.DataFields(i).Orientation = xlHidden
        End If
    Next i

    If df Is Nothing Then
        Set df = pt.AddDataField(PF(pt, FLD_SAP), DATA_CAPTION, xlCount)
    End If

    df.Function = xlCount       ' resets the caption, so set the caption after
    OptSet df, "Caption", DATA_CAPTION
    OptSet df, "NumberFormat", "General"
End Sub


Private Sub PlaceField(pf As PivotField, axis As XlPivotFieldOrientation, _
                       slot As Long)
    With pf
        If .Orientation <> axis Then .Orientation = axis
        .Position = slot
    End With
End Sub


' Subtotals(1) is the Automatic slot. Setting it True clears the eleven
' explicit functions; turning subtotals off means clearing all twelve.
Private Sub SetAutoSubtotal(pf As PivotField, wanted As Boolean)
    Dim i As Long

    On Error Resume Next
    If wanted Then
        pf.Subtotals(1) = True
    Else
        For i = 1 To 12
            pf.Subtotals(i) = False
        Next i
    End If
    If Err.Number <> 0 Then
        Note "subtotals on " & pf.Name, Err.Description
        Err.Clear
    End If
    On Error GoTo 0
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
    PF(pt, FLD_DATE).Orientation = xlHidden
    On Error GoTo 0

    PlaceField pfYears, xlColumnField, 1
    PlaceField pfQuarters, xlColumnField, 2
    PlaceField pfMonths, xlColumnField, 3

    ' Quarter subtotals are the "Qtr n Total" columns and year subtotals the
    ' "<year> Total" column. Months, being innermost, get none.
    SetAutoSubtotal pfYears, True
    SetAutoSubtotal pfQuarters, True
    SetAutoSubtotal pfMonths, False
End Sub


Private Sub GroupDateField(pt As PivotTable)
    Dim pf As PivotField

    Set pf = PF(pt, FLD_DATE)
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
    Set DateAxisField = PF(pt, FLD_DATE)
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
    gStep = "formatting the pivot cells"

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
    gStep = "setting column widths, freeze panes and zoom"

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


'==============================================================================
' Housekeeping
'==============================================================================

' The 8-6-26 report has the pivot in the top-left corner. A pivot that was
' built further down is nudged up by deleting the empty rows above it, but only
' when there is nothing up there worth keeping.
Private Sub MoveToOrigin(pt As PivotTable)
    Dim ws As Worksheet
    Dim above As Long

    On Error Resume Next
    Set ws = pt.Parent
    above = pt.TableRange2.Row - 1
    If above >= 1 Then
        If Application.CountA(ws.Rows("1:" & above)) = 0 Then
            ws.Rows("1:" & above).Delete
        End If
    End If
    If Err.Number <> 0 Then
        Note "moving the pivot to A1", Err.Description
        Err.Clear
    End If
    On Error GoTo 0
End Sub


' A blank Region column collapses every customer into a single "(blank)" group,
' so the pivot comes out structurally unlike the report however good the
' formatting is. Worth saying out loud rather than leaving to be spotted.
Private Sub WarnIfRegionEmpty()
    Dim ws As Worksheet
    Dim regionCol As Long, lastRow As Long

    On Error Resume Next
    Set ws = ResolveDataSheet()
    On Error GoTo 0
    If ws Is Nothing Then Exit Sub

    regionCol = HeaderColumn(ws, FLD_REGION)
    lastRow = ws.Cells(ws.Rows.Count, HeaderColumn(ws, FLD_SAP)).End(xlUp).Row
    If regionCol = 0 Or lastRow < 2 Then Exit Sub

    If Application.CountA(ws.Range(ws.Cells(2, regionCol), _
                                   ws.Cells(lastRow, regionCol))) > 0 Then Exit Sub

    MsgBox "The pivot is formatted, but the " & FLD_REGION & " column on the " & _
           ws.Name & " sheet is empty, so every customer lands under one " & _
           "'(blank)' heading instead of being grouped by region." & _
           vbCrLf & vbCrLf & _
           "Fill that column in, then run this macro again.", _
           vbInformation, "Bank Memo Pivot"
End Sub


'==============================================================================
' Diagnostics
'==============================================================================

' Reports what this module can and cannot see, without changing anything.
' Run this first whenever something is not behaving: it settles which version
' of the code the workbook is running and what it resolves each name to.
Public Sub CheckSetup()
    Dim report As String
    Dim ws As Worksheet, src As Worksheet, pt As PivotTable
    Dim regionCol As Long, sapCol As Long, lastRow As Long, filled As Long

    report = "Module version: " & MACRO_VERSION & vbCrLf & _
             "Excel version: " & Application.Version & vbCrLf & _
             "Workbook: " & ThisWorkbook.Name & vbCrLf & vbCrLf

    ' --- tabs ----------------------------------------------------------------
    report = report & "Tabs in this workbook:" & vbCrLf
    For Each ws In ThisWorkbook.Worksheets
        report = report & "    " & ws.Name
        If ws.Visible <> xlSheetVisible Then report = report & "  (hidden)"
        report = report & vbCrLf
    Next ws
    report = report & vbCrLf

    ' --- summary sheet and pivot ---------------------------------------------
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SUMMARY_SHEET)
    On Error GoTo 0

    If ws Is Nothing Then
        report = report & "Summary tab '" & SUMMARY_SHEET & "': NOT FOUND" & vbCrLf
    Else
        report = report & "Summary tab: " & ws.Name & vbCrLf
        If ws.PivotTables.Count = 0 Then
            report = report & "    pivot tables: none - run RebuildSummaryPivot" & vbCrLf
        Else
            For Each pt In ws.PivotTables
                report = report & "    pivot: " & pt.Name & " at " & _
                         pt.TableRange2.Address(False, False) & vbCrLf
            Next pt
        End If
    End If
    report = report & vbCrLf

    ' --- source sheet ---------------------------------------------------------
    On Error Resume Next
    Set src = ResolveDataSheet()
    On Error GoTo 0

    If src Is Nothing Then
        report = report & "Source sheet: NOT FOUND" & vbCrLf & _
                 "    No tab has all four headings in row 1." & vbCrLf
    Else
        sapCol = HeaderColumn(src, FLD_SAP)
        lastRow = src.Cells(src.Rows.Count, sapCol).End(xlUp).Row
        report = report & "Source sheet: " & src.Name & _
                 "  (rows 2 to " & lastRow & ")" & vbCrLf
        report = report & "    " & FLD_SAP & " -> column " & sapCol & _
                 " headed """ & src.Cells(1, sapCol).Value & """" & vbCrLf
        report = report & ColumnLine(src, FLD_REGION)
        report = report & ColumnLine(src, FLD_CUSTOMER)
        report = report & ColumnLine(src, FLD_DATE)

        regionCol = HeaderColumn(src, FLD_REGION)
        If regionCol > 0 And lastRow > 1 Then
            filled = Application.CountA(src.Range(src.Cells(2, regionCol), _
                                                  src.Cells(lastRow, regionCol)))
            report = report & vbCrLf & FLD_REGION & " values filled in: " & _
                     filled & " of " & (lastRow - 1) & vbCrLf
            If filled = 0 Then
                report = report & "    Every customer will group under " & _
                         """(blank)"" until this column is filled." & vbCrLf
            End If
        End If
    End If

    MsgBox report, vbInformation, "Bank Memo Pivot " & MACRO_VERSION & " - setup"
End Sub


Private Function ColumnLine(ws As Worksheet, wanted As String) As String
    Dim c As Long

    c = HeaderColumn(ws, wanted)
    If c = 0 Then
        ColumnLine = "    " & wanted & " -> NOT FOUND" & vbCrLf
    Else
        ColumnLine = "    " & wanted & " -> column " & c & _
                     " headed """ & ws.Cells(1, c).Value & """" & vbCrLf
    End If
End Function
