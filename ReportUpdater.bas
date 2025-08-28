Attribute VB_Name = "ReportUpdater"
Option Explicit

Private Const CFG_SHEET As String = "Config"
Private Const NAME_CVM_PATH As String = "CVM_PATH"
Private Const NAME_SBT_PATH As String = "SBT_PATH"
Private Const NAME_ARCHER_PATH As String = "ARCHER_PATH"

Private Const Q_CVM As String = "CVMQuery"
Private Const Q_SBT As String = "SBTQuery"
Private Const Q_ARCHER As String = "ArcherQuery"

Private Const SHEET_CVM_RAW As String = "CVMRaw"
Private Const SHEET_SBT_RAW As String = "SBTRaw"
Private Const SHEET_ARCH_RAW As String = "ArcherRaw"

Private Const SHEET_CVM_RPT As String = "cvm_camp_rpt"
Private Const SHEET_SBT_RPT As String = "sbt_camp_rpt"
Private Const SHEET_ARCH_RPT As String = "archer_status_rpt"
Private Const SHEET_CVM_DEP As String = "cvm_open_in_camp"
Private Const SHEET_SBT_DEP As String = "sbt_open_in_camp"

Private Const ITEM_CAMP As String = "CAMP Excel Report"
Private Const ITEM_ARCHER As String = "Archer Search Report"

Public Sub UpdateReportsAutomation()
    Dim wb As Workbook
    Dim logWs As Worksheet
    Dim logRow As Long

    Set wb = ThisWorkbook

    PrepareAppState True
    On Error GoTo CleanUp

    EnsureConfig wb
    EnsureNamedPath wb, NAME_CVM_PATH, "C:\CVM_UPDATE\CAMP_CVM.xlsx"
    EnsureNamedPath wb, NAME_SBT_PATH, "C:\CVM_UPDATE\CAMP_SBT.xlsx"
    EnsureNamedPath wb, NAME_ARCHER_PATH, "C:\CVM_UPDATE\Archer.xlsx"

    If Not ValidateRequiredSheets(wb) Then GoTo CleanUp
    If Not ValidateFilesExist(wb) Then GoTo CleanUp

    EnsureQueryUsesNamedParam wb, Q_CVM, NAME_CVM_PATH, ITEM_CAMP
    EnsureQueryUsesNamedParam wb, Q_SBT, NAME_SBT_PATH, ITEM_CAMP
    EnsureQueryUsesNamedParam wb, Q_ARCHER, NAME_ARCHER_PATH, ITEM_ARCHER

    RefreshPowerQueriesSync wb

    Set logWs = GetOrCreateLogSheet(wb)
    logRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1

    UpdateSheetV2 wb.Sheets(SHEET_CVM_RAW), wb.Sheets(SHEET_CVM_RPT), logWs, logRow, SHEET_CVM_RPT: logRow = logRow + 1
    UpdateSheetV2 wb.Sheets(SHEET_SBT_RAW), wb.Sheets(SHEET_SBT_RPT), logWs, logRow, SHEET_SBT_RPT: logRow = logRow + 1
    UpdateSheetV2 wb.Sheets(SHEET_ARCH_RAW), wb.Sheets(SHEET_ARCH_RPT), logWs, logRow, SHEET_ARCH_RPT: logRow = logRow + 1

    ResizeDependentSheetV2 wb.Sheets(SHEET_CVM_RPT), wb.Sheets(SHEET_CVM_DEP), logWs, logRow, SHEET_CVM_DEP: logRow = logRow + 1
    ResizeDependentSheetV2 wb.Sheets(SHEET_SBT_RPT), wb.Sheets(SHEET_SBT_DEP), logWs, logRow, SHEET_SBT_DEP

    MsgBox "Reports and dependent sheets updated successfully! Check UpdateLog sheet for details.", vbInformation

CleanUp:
    PrepareAppState False
    If Err.Number <> 0 Then MsgBox "Update failed: " & Err.Description, vbExclamation
End Sub

Private Sub UpdateSheetV2(rawWs As Worksheet, tgtWs As Worksheet, logWs As Worksheet, logRow As Long, sheetName As String)
    Dim rawLastRow As Long, rawLastCol As Long
    Dim tgtLastRow As Long, tgtLastCol As Long
    Dim oldRecords As Long, newRecords As Long
    Dim status As String
    Dim clearRng As Range
    Dim lo As ListObject

    On Error GoTo ErrorHandler

    rawLastCol = rawWs.Cells(1, rawWs.Columns.Count).End(xlToLeft).Column
    rawLastRow = rawWs.Cells(rawWs.Rows.Count, "A").End(xlUp).Row
    newRecords = Application.Max(rawLastRow - 1, 0)

    tgtLastCol = tgtWs.Cells(1, tgtWs.Columns.Count).End(xlToLeft).Column
    tgtLastRow = tgtWs.Cells(tgtWs.Rows.Count, "A").End(xlUp).Row
    oldRecords = Application.Max(tgtLastRow - 1, 0)

    If newRecords = 0 Then
        status = "No new data; target preserved"
        GoTo LogAndExit
    End If

    If tgtLastRow >= 2 And tgtLastCol >= 1 Then
        Set clearRng = tgtWs.Range(tgtWs.Cells(2, 1), tgtWs.Cells(tgtLastRow, tgtLastCol))
        If tgtWs.ListObjects.Count > 0 Then
            For Each lo In tgtWs.ListObjects
                If Not Application.Intersect(clearRng, lo.Range) Is Nothing Then lo.Unlist
            Next lo
        End If
        clearRng.UnMerge
        clearRng.ClearContents
    End If

    tgtWs.Cells(2, 1).Resize(newRecords, rawLastCol).Value2 = _
        rawWs.Cells(2, 1).Resize(newRecords, rawLastCol).Value2

    status = "Success (" & newRecords & " rows)"

LogAndExit:
    logWs.Cells(logRow, "A").Value = Now()
    logWs.Cells(logRow, "B").Value = sheetName
    logWs.Cells(logRow, "C").Value = oldRecords
    logWs.Cells(logRow, "D").Value = newRecords
    logWs.Cells(logRow, "E").Value = status
    Exit Sub

ErrorHandler:
    status = "Error: " & Err.Description
    Resume LogAndExit
End Sub

Private Sub ResizeDependentSheetV2(srcWs As Worksheet, depWs As Worksheet, logWs As Worksheet, logRow As Long, sheetName As String)
    Dim srcLastRow As Long, srcLastCol As Long
    Dim depLastRow As Long, depLastCol As Long
    Dim rowsToAdd As Long, rowsToDelete As Long
    Dim oldRecords As Long, newRecords As Long
    Dim status As String
    Dim hasFormulas As Boolean
    Dim r As Range

    On Error GoTo ErrorHandler

    srcLastCol = srcWs.Cells(1, srcWs.Columns.Count).End(xlToLeft).Column
    srcLastRow = srcWs.Cells(srcWs.Rows.Count, "A").End(xlUp).Row
    newRecords = Application.Max(srcLastRow - 1, 0)

    depLastCol = depWs.Cells(1, depWs.Columns.Count).End(xlToLeft).Column
    depLastRow = depWs.Cells(depWs.Rows.Count, "A").End(xlUp).Row
    oldRecords = Application.Max(depLastRow - 1, 0)

    If depLastRow >= 2 And depLastCol >= 1 Then
        depWs.Range(depWs.Cells(2, 1), depWs.Cells(depLastRow, depLastCol)).ClearContents
    End If

    If srcLastCol > depLastCol Then
        depWs.Columns(depLastCol + 1).Resize(, srcLastCol - depLastCol).Insert Shift:=xlToRight, _
            CopyOrigin:=xlFormatFromLeftOrAbove
        depLastCol = srcLastCol
    ElseIf srcLastCol < depLastCol Then
        depWs.Columns(srcLastCol + 1 & ":" & depLastCol).Clear
        depLastCol = srcLastCol
    End If

    On Error Resume Next
    Set r = depWs.Rows(2).SpecialCells(xlCellTypeFormulas)
    hasFormulas = Not r Is Nothing
    On Error GoTo ErrorHandler

    rowsToAdd = newRecords - oldRecords
    If rowsToAdd > 0 Then
        depWs.Rows(depLastRow + 1 & ":" & depLastRow + rowsToAdd).Insert Shift:=xlDown, _
            CopyOrigin:=xlFormatFromLeftOrAbove
        If hasFormulas Then
            depWs.Rows(2).Copy
            depWs.Rows(depLastRow + 1 & ":" & depLastRow + rowsToAdd).PasteSpecial xlPasteFormulasAndNumberFormats
            Application.CutCopyMode = False
        Else
            depWs.Rows(1).Copy
            depWs.Rows(depLastRow + 1 & ":" & depLastRow + rowsToAdd).PasteSpecial xlPasteFormats
            Application.CutCopyMode = False
        End If
    ElseIf rowsToAdd < 0 Then
        rowsToDelete = Abs(rowsToAdd)
        depWs.Rows(depLastRow - rowsToDelete + 1 & ":" & depLastRow).Delete Shift:=xlUp
    End If

    If newRecords > 0 Then
        depWs.Rows(2).Resize(newRecords).FillDown
    End If

    depLastRow = 1 + newRecords
    status = "Updated to " & depLastRow & " rows, " & depLastCol & " cols"

    logWs.Cells(logRow, "A").Value = Now()
    logWs.Cells(logRow, "B").Value = sheetName
    logWs.Cells(logRow, "C").Value = oldRecords
    logWs.Cells(logRow, "D").Value = newRecords
    logWs.Cells(logRow, "E").Value = status
    Exit Sub

ErrorHandler:
    status = "Error: " & Err.Description
    logWs.Cells(logRow, "A").Value = Now()
    logWs.Cells(logRow, "B").Value = sheetName
    logWs.Cells(logRow, "C").Value = oldRecords
    logWs.Cells(logRow, "D").Value = 0
    logWs.Cells(logRow, "E").Value = status
End Sub

Private Sub RefreshPowerQueriesSync(wb As Workbook)
    Dim c As WorkbookConnection
    On Error Resume Next
    For Each c In wb.Connections
        Select Case c.Type
            Case xlConnectionTypeOLEDB: c.OLEDBConnection.BackgroundQuery = False: c.Refresh
            Case xlConnectionTypeODBC: c.ODBCConnection.BackgroundQuery = False: c.Refresh
            Case Else: c.Refresh
        End Select
    Next
    On Error GoTo 0
    Application.CalculateUntilAsyncQueriesDone
End Sub

Private Function GetOrCreateLogSheet(wb As Workbook) As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets("UpdateLog")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(After:=wb.Sheets(wb.Sheets.Count))
        ws.Name = "UpdateLog"
        ws.Range("A1:E1").Value = Array("Date/Time", "Sheet Name", "Old Records", "New Records", "Status")
    End If
    Set GetOrCreateLogSheet = ws
End Function

Private Function ValidateRequiredSheets(wb As Workbook) As Boolean
    Dim ok As Boolean
    ok = WorksheetExists(wb, SHEET_CVM_RAW) And _
         WorksheetExists(wb, SHEET_SBT_RAW) And _
         WorksheetExists(wb, SHEET_ARCH_RAW) And _
         WorksheetExists(wb, SHEET_CVM_RPT) And _
         WorksheetExists(wb, SHEET_SBT_RPT) And _
         WorksheetExists(wb, SHEET_ARCH_RPT) And _
         WorksheetExists(wb, SHEET_CVM_DEP) And _
         WorksheetExists(wb, SHEET_SBT_DEP)
    If Not ok Then MsgBox "One or more required worksheets are missing.", vbExclamation
    ValidateRequiredSheets = ok
End Function

Private Function ValidateFilesExist(wb As Workbook) As Boolean
    Dim p1 As String, p2 As String, p3 As String
    p1 = GetNamedPathValue(wb, NAME_CVM_PATH)
    p2 = GetNamedPathValue(wb, NAME_SBT_PATH)
    p3 = GetNamedPathValue(wb, NAME_ARCHER_PATH)
    If Dir(p1, vbNormal) = "" Or Dir(p2, vbNormal) = "" Or Dir(p3, vbNormal) = "" Then
        MsgBox "One or more source files not found. Please update named paths on Config sheet or place files at the specified locations.", vbExclamation
        ValidateFilesExist = False
    Else
        ValidateFilesExist = True
    End If
End Function

Private Function WorksheetExists(wb As Workbook, name As String) As Boolean
    On Error Resume Next
    WorksheetExists = Not wb.Worksheets(name) Is Nothing
    On Error GoTo 0
End Function

Private Sub EnsureConfig(wb As Workbook)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = wb.Sheets(CFG_SHEET)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = wb.Sheets.Add(Before:=wb.Sheets(1))
        ws.Name = CFG_SHEET
        ws.Range("A1").Value = "Name"
        ws.Range("B1").Value = "Path"
        ws.Visible = xlSheetHidden
    End If
End Sub

Private Sub EnsureNamedPath(wb As Workbook, name As String, defaultPath As String)
    Dim ws As Worksheet
    Dim targetCell As Range
    Dim exists As Boolean
    Dim nm As Name

    Set ws = wb.Sheets(CFG_SHEET)

    On Error Resume Next
    Set nm = wb.Names(name)
    exists = Not nm Is Nothing
    On Error GoTo 0

    If Not exists Then
        Set targetCell = NextEmptyConfigCell(ws)
        targetCell.Offset(0, -1).Value = name
        targetCell.Value = defaultPath
        wb.Names.Add Name:=name, RefersTo:=targetCell
    Else
        If Len(Trim$(Evaluate(nm.RefersTo))) = 0 Then
            Range(nm.RefersTo).Value = defaultPath
        End If
    End If
End Sub

Private Function NextEmptyConfigCell(ws As Worksheet) As Range
    Dim r As Long
    r = ws.Cells(ws.Rows.Count, "B").End(xlUp).Row + 1
    If r < 2 Then r = 2
    Set NextEmptyConfigCell = ws.Cells(r, "B")
End Function

Private Function GetNamedPathValue(wb As Workbook, name As String) As String
    Dim nm As Name
    On Error Resume Next
    Set nm = wb.Names(name)
    On Error GoTo 0
    If nm Is Nothing Then
        GetNamedPathValue = ""
    Else
        GetNamedPathValue = CStr(Evaluate(nm.RefersTo))
    End If
End Function

Private Sub EnsureQueryUsesNamedParam(wb As Workbook, qName As String, paramName As String, itemName As String)
    Dim q As WorkbookQuery
    Dim exists As Boolean
    Dim m As String

    exists = False
    For Each q In wb.Queries
        If StrComp(q.Name, qName, vbTextCompare) = 0 Then exists = True: Exit For
    Next q

    m = BuildParamQueryM(paramName, itemName)

    If exists Then
        wb.Queries(qName).Formula = m
    Else
        wb.Queries.Add Name:=qName, Formula:=m
    End If
End Sub

Private Function BuildParamQueryM(paramName As String, itemName As String) As String
    Dim s As String
    s = "let" & vbCrLf & _
        "    PathRow = Excel.CurrentWorkbook(){[Name=""" & paramName & """]}[Content]," & vbCrLf & _
        "    Path = if Table.RowCount(PathRow) > 0 then PathRow{0}[Column1] else """"""," & vbCrLf & _
        "    Source = Excel.Workbook(File.Contents(Path), null, true)," & vbCrLf & _
        "    Sheet = Source{[Item=""" & itemName & """,Kind=""Sheet""]}[Data]," & vbCrLf & _
        "    #""Promoted Headers"" = Table.PromoteHeaders(Sheet, [PromoteAllScalars=true])" & vbCrLf & _
        "in" & vbCrLf & _
        "    #""Promoted Headers"""
    BuildParamQueryM = s
End Function

Private Sub PrepareAppState(start As Boolean)
    If start Then
        Application.ScreenUpdating = False
        Application.EnableEvents = False
        Application.DisplayAlerts = False
        Application.Calculation = xlCalculationManual
    Else
        Application.ScreenUpdating = True
        Application.EnableEvents = True
        Application.DisplayAlerts = True
        Application.Calculation = xlCalculationAutomatic
    End If
End Sub

