Attribute VB_Name = "modSnapshots"
Option Explicit
'==============================================================================
' THE DATA ENGINE
' Reads the raw HR tables into memory (Variant arrays) and crunches them with
' Scripting.Dictionary aggregation - no worksheet formulas, no loops over cells.
' Produces:
'   * Snapshots sheet : month-end x facility x department fact table
'   * _Calc (hidden)  : monthly trend block, department / facility / leave
'                       aggregates that feed the dashboard cards and charts
' ~53,000 employee-month evaluations run in well under a second.
'==============================================================================

Private mAsOf As Date

' Anchor "today" to the data so the file stays meaningful whenever it is opened
Public Function AsOfDate() As Date
    Dim loE As ListObject, a As Variant, i As Long, mx As Double
    Dim cH As Long, cT As Long
    If mAsOf = 0 Then
        Set loE = GetLO("tblEmployees")
        a = loE.DataBodyRange.Value2
        cH = ColIx(loE, "HireDate"): cT = ColIx(loE, "TerminationDate")
        For i = 1 To UBound(a, 1)
            If a(i, cH) > mx Then mx = a(i, cH)
            If Not IsEmpty(a(i, cT)) Then
                If a(i, cT) > mx Then mx = a(i, cT)
            End If
        Next i
        mAsOf = CDate(mx)
    End If
    AsOfDate = mAsOf
End Function

Public Sub BuildSnapshotsAndCalc()
    Dim loE As ListObject: Set loE = GetLO("tblEmployees")
    Dim a As Variant: a = loE.DataBodyRange.Value2
    Dim n As Long: n = UBound(a, 1)
    Dim cH As Long, cT As Long, cF As Long, cD As Long, cN As Long
    Dim cS As Long, cR As Long, cId As Long
    cH = ColIx(loE, "HireDate"): cT = ColIx(loE, "TerminationDate")
    cF = ColIx(loE, "Facility"): cD = ColIx(loE, "Department")
    cN = ColIx(loE, "NationalityGroup"): cS = ColIx(loE, "MonthlySalarySAR")
    cR = ColIx(loE, "PerformanceRating"): cId = ColIx(loE, "EmployeeID")

    ' ---- combo (month|facility|dept) and monthly dictionaries --------------
    Dim dHC As Object, dSa As Object, dHi As Object, dTe As Object, dPa As Object
    Dim mHC As Object, mSa As Object, mHi As Object, mTe As Object, mPa As Object
    Set dHC = NewDict(): Set dSa = NewDict(): Set dHi = NewDict()
    Set dTe = NewDict(): Set dPa = NewDict()
    Set mHC = NewDict(): Set mSa = NewDict(): Set mHi = NewDict()
    Set mTe = NewDict(): Set mPa = NewDict()

    Dim mFirst As Date: mFirst = DateSerial(2020, 1, 1)
    Dim lastMonth As Date
    lastMonth = DateSerial(Year(AsOfDate), Month(AsOfDate), 1)

    Dim mCur As Date, i As Long
    Dim somD As Double, eomD As Double, ym As String, key As String
    Dim h As Double, t As Double, isActive As Boolean

    mCur = mFirst
    Do While mCur <= lastMonth
        somD = CDbl(mCur)
        eomD = CDbl(DateSerial(Year(mCur), Month(mCur) + 1, 0))
        ym = Format(mCur, "yyyy-mm")
        For i = 1 To n
            h = a(i, cH)
            t = 0
            If Not IsEmpty(a(i, cT)) Then
                If IsNumeric(a(i, cT)) Then t = a(i, cT)
            End If
            isActive = (h <= eomD) And (t = 0 Or t > eomD)
            key = ym & "|" & a(i, cF) & "|" & a(i, cD)
            If isActive Then
                dHC(key) = dHC(key) + 1
                mHC(ym) = mHC(ym) + 1
                dPa(key) = dPa(key) + a(i, cS)
                mPa(ym) = mPa(ym) + a(i, cS)
                If a(i, cN) = "Saudi" Then
                    dSa(key) = dSa(key) + 1
                    mSa(ym) = mSa(ym) + 1
                End If
            End If
            If h >= somD And h <= eomD Then
                dHi(key) = dHi(key) + 1
                mHi(ym) = mHi(ym) + 1
            End If
            If t >= somD And t <= eomD And t > 0 Then
                dTe(key) = dTe(key) + 1
                mTe(ym) = mTe(ym) + 1
            End If
        Next i
        mCur = DateAdd("m", 1, mCur)
    Loop

    ' ---- _Calc monthly trend block A:G --------------------------------------
    Dim wsC As Worksheet: Set wsC = GetWS("_Calc")
    wsC.Visible = xlSheetVisible
    wsC.Cells.Clear
    wsC.Range("A1:G1").Value = Array("MonthEnd", "Headcount", "Hires", _
        "Terminations", "AttritionRate", "Saudization", "Payroll")
    Dim nMonths As Long
    nMonths = DateDiff("m", mFirst, lastMonth) + 1
    Dim trend() As Variant: ReDim trend(1 To nMonths, 1 To 7)
    Dim r As Long: r = 0
    Dim hc As Double
    mCur = mFirst
    Do While mCur <= lastMonth
        ym = Format(mCur, "yyyy-mm")
        hc = mHC(ym) + 0
        r = r + 1
        trend(r, 1) = DateSerial(Year(mCur), Month(mCur) + 1, 0)
        trend(r, 2) = hc
        trend(r, 3) = mHi(ym) + 0
        trend(r, 4) = mTe(ym) + 0
        If hc > 0 Then trend(r, 5) = (mTe(ym) + 0) / hc Else trend(r, 5) = 0
        If hc > 0 Then trend(r, 6) = (mSa(ym) + 0) / hc Else trend(r, 6) = 0
        trend(r, 7) = mPa(ym) + 0
        mCur = DateAdd("m", 1, mCur)
    Loop
    wsC.Range("A2").Resize(nMonths, 7).Value = trend
    wsC.Range("A2:A" & nMonths + 1).NumberFormat = "dd-mmm-yyyy"

    ' ---- Snapshots fact sheet ------------------------------------------------
    Dim k As Variant
    For Each k In dHi.Keys
        If Not dHC.Exists(k) Then dHC(k) = 0
    Next k
    For Each k In dTe.Keys
        If Not dHC.Exists(k) Then dHC(k) = 0
    Next k
    Dim wsS As Worksheet: Set wsS = GetWS("Snapshots")
    Dim loOld As ListObject
    For Each loOld In wsS.ListObjects
        loOld.Unlist
    Next loOld
    wsS.Cells.Clear
    wsS.Range("A1:H1").Value = Array("MonthEnd", "Facility", "Department", _
        "Headcount", "SaudiHeadcount", "Hires", "Terminations", "PayrollSAR")
    Dim out() As Variant: ReDim out(1 To dHC.Count, 1 To 8)
    Dim p() As String, rr As Long
    For Each k In dHC.Keys
        p = Split(CStr(k), "|")
        rr = rr + 1
        out(rr, 1) = DateSerial(CLng(Left(p(0), 4)), CLng(Mid(p(0), 6, 2)) + 1, 0)
        out(rr, 2) = p(1)
        out(rr, 3) = p(2)
        out(rr, 4) = dHC(k) + 0
        out(rr, 5) = dSa(k) + 0
        out(rr, 6) = dHi(k) + 0
        out(rr, 7) = dTe(k) + 0
        out(rr, 8) = dPa(k) + 0
    Next k
    wsS.Range("A2").Resize(rr, 8).Value = out
    wsS.Range("A2:A" & rr + 1).NumberFormat = "dd-mmm-yyyy"
    wsS.Range("H2:H" & rr + 1).NumberFormat = "#,##0"
    Dim loSnap As ListObject
    Set loSnap = wsS.ListObjects.Add(xlSrcRange, wsS.Range("A1").Resize(rr + 1, 8), , xlYes)
    loSnap.Name = "tblSnapshots"
    loSnap.TableStyle = "TableStyleMedium2"
    wsS.Columns("A:H").AutoFit
    wsS.Tab.Color = ClrTeal()

    ' ---- as-of aggregates: department / facility ----------------------------
    Dim asD As Double: asD = CDbl(AsOfDate)
    Dim cutoff As Double: cutoff = asD - 365
    Dim pHC As Object, pSa As Object, pSal As Object, pRat As Object, pTe As Object
    Dim fHC As Object, fSa As Object, eDept As Object
    Set pHC = NewDict(): Set pSa = NewDict(): Set pSal = NewDict()
    Set pRat = NewDict(): Set pTe = NewDict()
    Set fHC = NewDict(): Set fSa = NewDict(): Set eDept = NewDict()
    Dim dep As String, fac As String
    For i = 1 To n
        h = a(i, cH)
        t = 0
        If Not IsEmpty(a(i, cT)) Then
            If IsNumeric(a(i, cT)) Then t = a(i, cT)
        End If
        dep = a(i, cD): fac = a(i, cF)
        eDept(CStr(a(i, cId))) = dep
        If h <= asD And (t = 0 Or t > asD) Then
            pHC(dep) = pHC(dep) + 1
            pSal(dep) = pSal(dep) + a(i, cS)
            pRat(dep) = pRat(dep) + a(i, cR)
            fHC(fac) = fHC(fac) + 1
            If a(i, cN) = "Saudi" Then
                pSa(dep) = pSa(dep) + 1
                fSa(fac) = fSa(fac) + 1
            End If
        End If
        If t > cutoff And t <= asD And t > 0 Then pTe(dep) = pTe(dep) + 1
    Next i

    ' training hours (last 12m) joined to department through EmployeeID
    Dim loT As ListObject: Set loT = GetLO("tblTraining")
    Dim tr As Variant: tr = loT.DataBodyRange.Value2
    Dim cTid As Long, cTdt As Long, cThr As Long
    cTid = ColIx(loT, "EmployeeID"): cTdt = ColIx(loT, "TrainingDate")
    cThr = ColIx(loT, "Hours")
    Dim pTrn As Object: Set pTrn = NewDict()
    For i = 1 To UBound(tr, 1)
        If tr(i, cTdt) > cutoff And tr(i, cTdt) <= asD Then
            If eDept.Exists(CStr(tr(i, cTid))) Then
                dep = eDept(CStr(tr(i, cTid)))
                pTrn(dep) = pTrn(dep) + tr(i, cThr)
            End If
        End If
    Next i

    ' department block J:Q, sorted by headcount desc
    Dim loD As ListObject: Set loD = GetLO("tblDepartments")
    Dim dl As Variant: dl = loD.DataBodyRange.Value2
    Dim nDep As Long: nDep = UBound(dl, 1)
    Dim names() As String, hcs() As Double
    ReDim names(1 To nDep): ReDim hcs(1 To nDep)
    For i = 1 To nDep
        names(i) = dl(i, 1)
        hcs(i) = pHC(names(i)) + 0
    Next i
    Dim j As Long, tmpS As String, tmpD As Double
    For i = 1 To nDep - 1
        For j = i + 1 To nDep
            If hcs(j) > hcs(i) Then
                tmpD = hcs(i): hcs(i) = hcs(j): hcs(j) = tmpD
                tmpS = names(i): names(i) = names(j): names(j) = tmpS
            End If
        Next j
    Next i
    wsC.Range("J1:Q1").Value = Array("Department", "Headcount", "SaudizationPct", _
        "Terms12M", "AttritionRate12M", "AvgSalary", "AvgRating", "TrainingHrs12M")
    For i = 1 To nDep
        dep = names(i)
        hc = pHC(dep) + 0
        wsC.Cells(1 + i, 10).Value = dep
        wsC.Cells(1 + i, 11).Value = hc
        If hc > 0 Then
            wsC.Cells(1 + i, 12).Value = (pSa(dep) + 0) / hc
            wsC.Cells(1 + i, 14).Value = (pTe(dep) + 0) / hc
            wsC.Cells(1 + i, 15).Value = (pSal(dep) + 0) / hc
            wsC.Cells(1 + i, 16).Value = (pRat(dep) + 0) / hc
        End If
        wsC.Cells(1 + i, 13).Value = pTe(dep) + 0
        wsC.Cells(1 + i, 17).Value = pTrn(dep) + 0
    Next i

    ' facility block S:U
    Dim loF As ListObject: Set loF = GetLO("tblFacilities")
    Dim fl As Variant: fl = loF.DataBodyRange.Value2
    wsC.Range("S1:U1").Value = Array("Facility", "Headcount", "SaudizationPct")
    For i = 1 To UBound(fl, 1)
        fac = fl(i, 1)
        hc = fHC(fac) + 0
        wsC.Cells(1 + i, 19).Value = fac
        wsC.Cells(1 + i, 20).Value = hc
        If hc > 0 Then wsC.Cells(1 + i, 21).Value = (fSa(fac) + 0) / hc
    Next i

    ' leave mix (last 12m) W:X
    Dim loL As ListObject: Set loL = GetLO("tblLeave")
    Dim lv As Variant: lv = loL.DataBodyRange.Value2
    Dim cLt As Long, cLd As Long, cLs As Long
    cLt = ColIx(loL, "LeaveType"): cLs = ColIx(loL, "StartDate")
    cLd = ColIx(loL, "Days")
    Dim lMix As Object: Set lMix = NewDict()
    For i = 1 To UBound(lv, 1)
        If lv(i, cLs) > cutoff And lv(i, cLs) <= asD Then
            lMix(CStr(lv(i, cLt))) = lMix(CStr(lv(i, cLt))) + lv(i, cLd)
        End If
    Next i
    wsC.Range("W1:X1").Value = Array("LeaveType", "Days12M")
    r = 1
    For Each k In lMix.Keys
        r = r + 1
        wsC.Cells(r, 23).Value = CStr(k)
        wsC.Cells(r, 24).Value = lMix(k)
    Next k

    wsC.Visible = xlSheetHidden
End Sub
