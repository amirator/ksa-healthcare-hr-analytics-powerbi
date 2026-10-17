Attribute VB_Name = "modKPI"
Option Explicit
'==============================================================================
' Computes the 12 headline KPIs (single in-memory pass per table) and writes
' them into the dashboard cards, plus the conditionally-formatted department
' scorecard sourced from the _Calc aggregates.
'==============================================================================

Public Sub ComputeKPIs()
    Dim ws As Worksheet: Set ws = GetWS("Dashboard")
    Dim asD As Double: asD = CDbl(AsOfDate)
    Dim cutoff As Double: cutoff = asD - 365

    ' ---- employees pass -------------------------------------------------------
    Dim loE As ListObject: Set loE = GetLO("tblEmployees")
    Dim a As Variant: a = loE.DataBodyRange.Value2
    Dim cH As Long, cT As Long, cG As Long, cN As Long, cS As Long
    cH = ColIx(loE, "HireDate"): cT = ColIx(loE, "TerminationDate")
    cG = ColIx(loE, "Gender"): cN = ColIx(loE, "NationalityGroup")
    cS = ColIx(loE, "MonthlySalarySAR")
    Dim i As Long, h As Double, t As Double
    Dim hc As Double, saudi As Double, fem As Double
    Dim tenureSum As Double, salSum As Double
    For i = 1 To UBound(a, 1)
        h = a(i, cH)
        t = 0
        If Not IsEmpty(a(i, cT)) Then
            If IsNumeric(a(i, cT)) Then t = a(i, cT)
        End If
        If h <= asD And (t = 0 Or t > asD) Then
            hc = hc + 1
            salSum = salSum + a(i, cS)
            tenureSum = tenureSum + (asD - h) / 365.25
            If a(i, cN) = "Saudi" Then saudi = saudi + 1
            If a(i, cG) = "Female" Then fem = fem + 1
        End If
    Next i

    ' ---- attrition (12M) from the _Calc monthly block ---------------------------
    Dim wsC As Worksheet: Set wsC = GetWS("_Calc")
    Dim lastR As Long: lastR = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    Dim terms12 As Double, hcSum As Double, mCount As Long
    For i = lastR To 2 Step -1
        If mCount >= 12 Then Exit For
        terms12 = terms12 + wsC.Cells(i, 4).Value
        hcSum = hcSum + wsC.Cells(i, 2).Value
        mCount = mCount + 1
    Next i
    Dim attr12 As Double
    If hcSum > 0 Then attr12 = terms12 / (hcSum / mCount)

    ' ---- recruitment pass --------------------------------------------------------
    Dim loR As ListObject: Set loR = GetLO("tblRecruitment")
    Dim rc As Variant: rc = loR.DataBodyRange.Value2
    Dim cSt As Long, cDf As Long
    cSt = ColIx(loR, "Status"): cDf = ColIx(loR, "DaysToFill")
    Dim openPos As Double, ttfSum As Double, ttfN As Double
    For i = 1 To UBound(rc, 1)
        If rc(i, cSt) = "Open" Then openPos = openPos + 1
        If rc(i, cSt) = "Filled" Then
            ttfSum = ttfSum + rc(i, cDf)
            ttfN = ttfN + 1
        End If
    Next i

    ' ---- training pass (last 12m) ---------------------------------------------------
    Dim loT As ListObject: Set loT = GetLO("tblTraining")
    Dim tr As Variant: tr = loT.DataBodyRange.Value2
    Dim cTd As Long, cThr As Long, cTcs As Long
    cTd = ColIx(loT, "TrainingDate"): cThr = ColIx(loT, "Hours")
    cTcs = ColIx(loT, "CompletionStatus")
    Dim trnHrs As Double, trnAll As Double, trnDone As Double
    For i = 1 To UBound(tr, 1)
        If tr(i, cTd) > cutoff And tr(i, cTd) <= asD Then
            trnHrs = trnHrs + tr(i, cThr)
            trnAll = trnAll + 1
            If tr(i, cTcs) = "Completed" Then trnDone = trnDone + 1
        End If
    Next i

    ' ---- leave pass (sick days last 12m) ----------------------------------------------
    Dim loL As ListObject: Set loL = GetLO("tblLeave")
    Dim lv As Variant: lv = loL.DataBodyRange.Value2
    Dim cLt As Long, cLs As Long, cLd As Long
    cLt = ColIx(loL, "LeaveType"): cLs = ColIx(loL, "StartDate")
    cLd = ColIx(loL, "Days")
    Dim sick12 As Double
    For i = 1 To UBound(lv, 1)
        If lv(i, cLt) = "Sick" And lv(i, cLs) > cutoff And lv(i, cLs) <= asD Then
            sick12 = sick12 + lv(i, cLd)
        End If
    Next i
    Dim absent As Double
    If hc > 0 Then absent = sick12 / (hc * 22# * 12#)

    ' ---- write cards ---------------------------------------------------------------------
    Dim tgtSaudi As Double: tgtSaudi = CfgVal("SaudizationTarget")
    Dim tgtAttr As Double: tgtAttr = CfgVal("AttritionTargetMax")
    Dim tgtTtf As Double: tgtTtf = CfgVal("TimeToFillTargetDays")
    Dim tgtTrn As Double: tgtTrn = CfgVal("TrainingCompletionTarget")
    Dim tgtAbs As Double: tgtAbs = CfgVal("AbsenteeismMax")
    Dim saudiPct As Double: If hc > 0 Then saudiPct = saudi / hc
    Dim ttf As Double: If ttfN > 0 Then ttf = ttfSum / ttfN
    Dim complPct As Double: If trnAll > 0 Then complPct = trnDone / trnAll

    SetCard ws, CARD_R1, 0, hc, "#,##0", "across 8 facilities", ClrTeal()
    SetCard ws, CARD_R1, 1, saudiPct, "0.0%", _
        "Nitaqat target " & Format(tgtSaudi, "0%") & " " & UpDown(saudiPct >= tgtSaudi), _
        IIf(saudiPct >= tgtSaudi, ClrGreen(), ClrRed())
    SetCard ws, CARD_R1, 2, attr12, "0.0%", _
        "ceiling " & Format(tgtAttr, "0%") & " " & UpDown(attr12 <= tgtAttr), _
        IIf(attr12 <= tgtAttr, ClrGreen(), ClrRed())
    SetCard ws, CARD_R1, 3, tenureSum / IIf(hc = 0, 1, hc), "0.0 ""yrs""", "average service", ClrNavy()
    SetCard ws, CARD_R1, 4, salSum, "#,##0", "monthly base payroll", ClrNavy()
    SetCard ws, CARD_R1, 5, openPos, "#,##0", "requisitions in market", ClrAccent()

    SetCard ws, CARD_R2, 0, fem / IIf(hc = 0, 1, hc), "0.0%", "of active workforce", ClrNavy()
    SetCard ws, CARD_R2, 1, salSum / IIf(hc = 0, 1, hc), "#,##0", "mean monthly base", ClrNavy()
    SetCard ws, CARD_R2, 2, ttf, "0 ""days""", _
        "target " & tgtTtf & "d " & UpDown(ttf <= tgtTtf), _
        IIf(ttf <= tgtTtf, ClrGreen(), ClrRed())
    SetCard ws, CARD_R2, 3, trnHrs, "#,##0", "delivered last 12 months", ClrTeal()
    SetCard ws, CARD_R2, 4, complPct, "0.0%", _
        "target " & Format(tgtTrn, "0%") & " " & UpDown(complPct >= tgtTrn), _
        IIf(complPct >= tgtTrn, ClrGreen(), ClrRed())
    SetCard ws, CARD_R2, 5, absent, "0.00%", _
        "sick days vs capacity " & UpDown(absent <= tgtAbs), _
        IIf(absent <= tgtAbs, ClrGreen(), ClrRed())

    ' ---- refresh stamp ----------------------------------------------------------------------
    ws.Range("B3").Value = "Data through " & Format(AsOfDate, "dd-mmm-yyyy") & _
        "    |    Last refresh " & Format(Now, "dd-mmm-yyyy hh:nn") & _
        "    |    " & Format(hc, "#,##0") & " active employees, " & _
        Format(UBound(a, 1), "#,##0") & " lifetime records"

    BuildScorecard ws, wsC
End Sub

Private Function UpDown(ByVal good As Boolean) As String
    UpDown = IIf(good, ChrW(10003), ChrW(9888))   ' check mark / warning sign
End Function

Private Sub SetCard(ws As Worksheet, ByVal r As Long, ByVal i As Long, _
                    ByVal v As Variant, ByVal fmt As String, _
                    ByVal note As String, ByVal clr As Long)
    Dim c As Long: c = 2 + i * 2
    With ws.Cells(r + 1, c)
        .Value = v
        .NumberFormat = fmt
        .Font.Color = clr
    End With
    ws.Cells(r + 2, c).Value = note
End Sub

Private Sub BuildScorecard(ws As Worksheet, wsC As Worksheet)
    Dim hdr As Variant
    hdr = Array("Department", "Headcount", "Saudization", "Attrition (12M)", _
                "Avg Salary (SAR)", "Avg Rating", "Training Hrs (12M)", "Flag")
    Dim r0 As Long: r0 = SCORE_R + 1
    Dim i As Long, j As Long
    For j = 0 To 7
        With ws.Cells(r0, 2 + j)
            .Value = hdr(j)
            .Font.Bold = True: .Font.Size = 9: .Font.Color = vbWhite
            .Interior.Color = ClrNavy()
        End With
    Next j
    Dim nDep As Long
    nDep = wsC.Cells(wsC.Rows.Count, 10).End(xlUp).Row - 1
    Dim tgtSaudi As Double: tgtSaudi = CfgVal("SaudizationTarget")
    Dim tgtAttr As Double: tgtAttr = CfgVal("AttritionTargetMax")
    Dim sa As Double, at As Double
    For i = 1 To nDep
        ws.Cells(r0 + i, 2).Value = wsC.Cells(1 + i, 10).Value
        ws.Cells(r0 + i, 3).Value = wsC.Cells(1 + i, 11).Value
        ws.Cells(r0 + i, 4).Value = wsC.Cells(1 + i, 12).Value
        ws.Cells(r0 + i, 5).Value = wsC.Cells(1 + i, 14).Value
        ws.Cells(r0 + i, 6).Value = wsC.Cells(1 + i, 15).Value
        ws.Cells(r0 + i, 7).Value = wsC.Cells(1 + i, 16).Value
        ws.Cells(r0 + i, 8).Value = wsC.Cells(1 + i, 17).Value
        sa = ws.Cells(r0 + i, 4).Value
        at = ws.Cells(r0 + i, 5).Value
        If sa < tgtSaudi * 0.5 Or at > tgtAttr Then
            ws.Cells(r0 + i, 9).Value = "WATCH"
            ws.Cells(r0 + i, 9).Font.Color = ClrRed()
            ws.Cells(r0 + i, 9).Font.Bold = True
        Else
            ws.Cells(r0 + i, 9).Value = "OK"
            ws.Cells(r0 + i, 9).Font.Color = ClrGreen()
        End If
    Next i
    Dim body As Range
    Set body = ws.Range(ws.Cells(r0 + 1, 2), ws.Cells(r0 + nDep, 9))
    body.Interior.Color = vbWhite
    body.Font.Size = 9
    ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + nDep, 3)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0 + 1, 4), ws.Cells(r0 + nDep, 4)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(r0 + 1, 5), ws.Cells(r0 + nDep, 5)).NumberFormat = "0.0%"
    ws.Range(ws.Cells(r0 + 1, 6), ws.Cells(r0 + nDep, 6)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0 + 1, 7), ws.Cells(r0 + nDep, 7)).NumberFormat = "0.00"
    ws.Range(ws.Cells(r0 + 1, 8), ws.Cells(r0 + nDep, 8)).NumberFormat = "#,##0"
    ws.Range(ws.Cells(r0, 2), ws.Cells(r0 + nDep, 9)).BorderAround _
        Color:=ClrBorder(), Weight:=xlThin

    ' conditional formats: data bar on headcount, colour scales on saudization & rating
    Dim rng As Range, db As Databar, cs As ColorScale
    Set rng = ws.Range(ws.Cells(r0 + 1, 3), ws.Cells(r0 + nDep, 3))
    rng.FormatConditions.Delete
    Set db = rng.FormatConditions.AddDatabar
    db.BarColor.Color = ClrTeal()
    db.BarFillType = xlDataBarFillSolid

    Set rng = ws.Range(ws.Cells(r0 + 1, 4), ws.Cells(r0 + nDep, 4))
    rng.FormatConditions.Delete
    Set cs = rng.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).FormatColor.Color = ClrRed()
    cs.ColorScaleCriteria(2).FormatColor.Color = ClrAccent()
    cs.ColorScaleCriteria(3).FormatColor.Color = ClrGreen()

    Set rng = ws.Range(ws.Cells(r0 + 1, 7), ws.Cells(r0 + nDep, 7))
    rng.FormatConditions.Delete
    Set cs = rng.FormatConditions.AddColorScale(3)
    cs.ColorScaleCriteria(1).FormatColor.Color = ClrRed()
    cs.ColorScaleCriteria(2).FormatColor.Color = ClrAccent()
    cs.ColorScaleCriteria(3).FormatColor.Color = ClrGreen()
End Sub
