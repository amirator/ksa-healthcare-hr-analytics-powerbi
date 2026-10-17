Attribute VB_Name = "modAlerts"
Option Explicit
'==============================================================================
' Compliance radar: scans the raw data against the thresholds in tblConfig
' and writes a prioritized alert register, plus a summary on the dashboard.
'==============================================================================

Private mRow As Long
Private mHigh As Long, mMed As Long, mInfo As Long

Public Sub BuildAlerts()
    Dim ws As Worksheet: Set ws = GetWS("Alerts")
    ws.Cells.Clear
    ws.Cells.Interior.Color = ClrLight()
    mRow = 4: mHigh = 0: mMed = 0: mInfo = 0

    With ws.Range("A1")
        .Value = "COMPLIANCE & RISK ALERTS  -  auto-generated " & Format(Now, "dd-mmm-yyyy hh:nn")
        .Font.Size = 14: .Font.Bold = True: .Font.Color = ClrNavy()
    End With
    Dim hdr As Variant: hdr = Array("Severity", "Area", "Item", "Detail", "Value")
    Dim j As Long
    For j = 0 To 4
        With ws.Cells(3, 1 + j)
            .Value = hdr(j)
            .Font.Bold = True: .Font.Color = vbWhite
            .Interior.Color = ClrNavy()
        End With
    Next j

    Dim asD As Double: asD = CDbl(AsOfDate)
    Dim cutoff As Double: cutoff = asD - 365
    Dim wsC As Worksheet: Set wsC = GetWS("_Calc")
    Dim tgtSaudi As Double: tgtSaudi = CfgVal("SaudizationTarget")
    Dim i As Long, v As Double

    ' ---- 1. facility saudization below Nitaqat target (High) ------------------
    Dim lastF As Long: lastF = wsC.Cells(wsC.Rows.Count, 19).End(xlUp).Row
    For i = 2 To lastF
        v = wsC.Cells(i, 21).Value
        If v < tgtSaudi Then
            AddAlert ws, "High", "Saudization", CStr(wsC.Cells(i, 19).Value), _
                "Below Nitaqat target of " & Format(tgtSaudi, "0%"), Format(v, "0.0%")
        End If
    Next i

    ' ---- 2. department attrition above ceiling (High) ----------------------------
    Dim tgtAttr As Double: tgtAttr = CfgVal("AttritionTargetMax")
    Dim lastD As Long: lastD = wsC.Cells(wsC.Rows.Count, 10).End(xlUp).Row
    For i = 2 To lastD
        v = wsC.Cells(i, 14).Value
        If v > tgtAttr Then
            AddAlert ws, "High", "Attrition", CStr(wsC.Cells(i, 10).Value), _
                "12-month attrition above ceiling of " & Format(tgtAttr, "0%"), Format(v, "0.0%")
        End If
    Next i

    ' ---- 3. ageing open requisitions (Medium / High) -------------------------------
    Dim loR As ListObject: Set loR = GetLO("tblRecruitment")
    Dim rc As Variant: rc = loR.DataBodyRange.Value2
    Dim cSt As Long, cPd As Long, cPo As Long, cIdR As Long, cFa As Long
    cSt = ColIx(loR, "Status"): cPd = ColIx(loR, "PostingDate")
    cPo = ColIx(loR, "Position"): cIdR = ColIx(loR, "RequisitionID")
    cFa = ColIx(loR, "Facility")
    Dim ageLimit As Double: ageLimit = CfgVal("OpenReqAgeAlertDays")
    Dim age As Double
    For i = 1 To UBound(rc, 1)
        If rc(i, cSt) = "Open" Then
            age = asD - rc(i, cPd)
            If age > ageLimit * 2 Then
                AddAlert ws, "High", "Recruitment", _
                    rc(i, cIdR) & " - " & rc(i, cPo), _
                    "Open requisition ageing at " & rc(i, cFa), Format(age, "0") & " days"
            ElseIf age > ageLimit Then
                AddAlert ws, "Medium", "Recruitment", _
                    rc(i, cIdR) & " - " & rc(i, cPo), _
                    "Open requisition ageing at " & rc(i, cFa), Format(age, "0") & " days"
            End If
        End If
    Next i

    ' ---- 4. performance watchlist (Medium) --------------------------------------------
    Dim loE As ListObject: Set loE = GetLO("tblEmployees")
    Dim a As Variant: a = loE.DataBodyRange.Value2
    Dim cH As Long, cT As Long, cIdE As Long, cNm As Long, cDp As Long, cRt As Long
    cH = ColIx(loE, "HireDate"): cT = ColIx(loE, "TerminationDate")
    cIdE = ColIx(loE, "EmployeeID"): cNm = ColIx(loE, "FullName")
    cDp = ColIx(loE, "Department"): cRt = ColIx(loE, "PerformanceRating")
    Dim lowRate As Double: lowRate = CfgVal("LowRatingThreshold")
    Dim h As Double, t As Double
    For i = 1 To UBound(a, 1)
        h = a(i, cH)
        t = 0
        If Not IsEmpty(a(i, cT)) Then
            If IsNumeric(a(i, cT)) Then t = a(i, cT)
        End If
        If h <= asD And (t = 0 Or t > asD) Then
            If a(i, cRt) < lowRate Then
                AddAlert ws, "Medium", "Performance", _
                    a(i, cIdE) & " - " & a(i, cNm), _
                    "Active employee in " & a(i, cDp) & " rated below " & lowRate, _
                    Format(a(i, cRt), "0.0")
            End If
        End If
    Next i

    ' ---- 5. heavy sick leave last 12m (Medium) -------------------------------------------
    Dim loL As ListObject: Set loL = GetLO("tblLeave")
    Dim lv As Variant: lv = loL.DataBodyRange.Value2
    Dim cLt As Long, cLs As Long, cLd As Long, cLid As Long
    cLt = ColIx(loL, "LeaveType"): cLs = ColIx(loL, "StartDate")
    cLd = ColIx(loL, "Days"): cLid = ColIx(loL, "EmployeeID")
    Dim sick As Object: Set sick = NewDict()
    For i = 1 To UBound(lv, 1)
        If lv(i, cLt) = "Sick" And lv(i, cLs) > cutoff And lv(i, cLs) <= asD Then
            sick(CStr(lv(i, cLid))) = sick(CStr(lv(i, cLid))) + lv(i, cLd)
        End If
    Next i
    Dim sickLimit As Double: sickLimit = CfgVal("SickDaysAlertThreshold")
    Dim k As Variant
    For Each k In sick.Keys
        If sick(k) > sickLimit Then
            AddAlert ws, "Medium", "Absence", CStr(k), _
                "Sick leave above " & sickLimit & " days in 12 months", _
                sick(k) & " days"
        End If
    Next k

    ' ---- 6. recent training no-shows (Info) ----------------------------------------------
    Dim loT As ListObject: Set loT = GetLO("tblTraining")
    Dim tr As Variant: tr = loT.DataBodyRange.Value2
    Dim cTs As Long, cTd As Long, cTe As Long, cTc As Long
    cTs = ColIx(loT, "CompletionStatus"): cTd = ColIx(loT, "TrainingDate")
    cTe = ColIx(loT, "EmployeeID"): cTc = ColIx(loT, "Course")
    For i = 1 To UBound(tr, 1)
        If tr(i, cTs) = "No Show" And tr(i, cTd) > asD - 90 Then
            AddAlert ws, "Info", "Training", CStr(tr(i, cTe)), _
                "No-show: " & tr(i, cTc), Format(CDate(tr(i, cTd)), "dd-mmm-yyyy")
        End If
    Next i

    ws.Columns("A:E").AutoFit
    If ws.Columns("D").ColumnWidth > 60 Then ws.Columns("D").ColumnWidth = 60
    ws.Range("A3").AutoFilter
    ws.Activate
    On Error Resume Next
    ActiveWindow.FreezePanes = False
    ws.Range("A4").Select
    ActiveWindow.FreezePanes = True
    On Error GoTo 0
    ws.Tab.Color = IIf(mHigh > 0, ClrRed(), ClrGreen())

    ' ---- dashboard mini-panel ------------------------------------------------------------
    Dim wsD As Worksheet: Set wsD = GetWS("Dashboard")
    wsD.Cells(CARD_R1 + 1, 14).Value = ChrW(9679) & " High: " & mHigh
    wsD.Cells(CARD_R1 + 1, 14).Font.Color = ClrRed()
    wsD.Cells(CARD_R1 + 1, 14).Font.Bold = True
    wsD.Cells(CARD_R1 + 2, 14).Value = ChrW(9679) & " Medium: " & mMed
    wsD.Cells(CARD_R1 + 2, 14).Font.Color = ClrAccent()
    wsD.Cells(CARD_R1 + 3, 14).Value = ChrW(9679) & " Info: " & mInfo
    wsD.Cells(CARD_R1 + 3, 14).Font.Color = ClrGrey()
    wsD.Cells(CARD_R1 + 4, 14).Value = "Top risk: " & TopRisk(wsC, tgtSaudi)
    wsD.Cells(CARD_R1 + 5, 14).Value = "Press ALERTS for the full register"
    wsD.Cells(CARD_R1 + 5, 14).Font.Italic = True
    wsD.Activate
End Sub

Private Function TopRisk(wsC As Worksheet, ByVal tgtSaudi As Double) As String
    Dim lastF As Long: lastF = wsC.Cells(wsC.Rows.Count, 19).End(xlUp).Row
    Dim i As Long, worst As Double, nm As String
    worst = 1
    For i = 2 To lastF
        If wsC.Cells(i, 21).Value < worst Then
            worst = wsC.Cells(i, 21).Value
            nm = wsC.Cells(i, 19).Value
        End If
    Next i
    If worst < tgtSaudi Then
        TopRisk = nm & " (" & Format(worst, "0%") & " Saudization)"
    Else
        TopRisk = "none - all facilities on target"
    End If
End Function

Private Sub AddAlert(ws As Worksheet, ByVal sev As String, ByVal area As String, _
                     ByVal item As String, ByVal detail As String, ByVal val As String)
    ws.Cells(mRow, 1).Value = sev
    ws.Cells(mRow, 2).Value = area
    ws.Cells(mRow, 3).Value = item
    ws.Cells(mRow, 4).Value = detail
    ws.Cells(mRow, 5).Value = val
    ws.Range(ws.Cells(mRow, 1), ws.Cells(mRow, 5)).Interior.Color = vbWhite
    With ws.Cells(mRow, 1)
        .Font.Bold = True
        Select Case sev
            Case "High":   .Interior.Color = ClrRed():    .Font.Color = vbWhite: mHigh = mHigh + 1
            Case "Medium": .Interior.Color = ClrAccent(): .Font.Color = vbWhite: mMed = mMed + 1
            Case Else:     .Interior.Color = ClrGrey():   .Font.Color = vbWhite: mInfo = mInfo + 1
        End Select
    End With
    mRow = mRow + 1
End Sub
