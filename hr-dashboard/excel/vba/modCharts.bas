Attribute VB_Name = "modCharts"
Option Explicit
'==============================================================================
' Creates the four dashboard charts from the _Calc aggregates. Charts are
' deleted and rebuilt on every refresh so they always match the data.
'==============================================================================

Public Sub BuildCharts()
    Dim ws As Worksheet: Set ws = GetWS("Dashboard")
    Dim wsC As Worksheet: Set wsC = GetWS("_Calc")
    Dim co As ChartObject, i As Long
    For i = ws.ChartObjects.Count To 1 Step -1
        ws.ChartObjects(i).Delete
    Next i

    Dim lastM As Long: lastM = wsC.Cells(wsC.Rows.Count, 1).End(xlUp).Row
    Dim lastD As Long: lastD = wsC.Cells(wsC.Rows.Count, 10).End(xlUp).Row
    Dim lastF As Long: lastF = wsC.Cells(wsC.Rows.Count, 19).End(xlUp).Row
    Dim lastL As Long: lastL = wsC.Cells(wsC.Rows.Count, 23).End(xlUp).Row
    Dim s As Series

    ' ---- 1. headcount columns + attrition line (secondary axis) ---------------
    Set co = ws.ChartObjects.Add(ws.Range("B13").Left, ws.Range("B13").Top, 405, 205)
    co.Name = "chTrend"
    With co.Chart
        .ChartType = xlColumnClustered
        Set s = .SeriesCollection.NewSeries
        s.Name = "Headcount"
        s.XValues = wsC.Range("A2:A" & lastM)
        s.Values = wsC.Range("B2:B" & lastM)
        s.Format.Fill.ForeColor.RGB = ClrTeal()
        Set s = .SeriesCollection.NewSeries
        s.Name = "Attrition rate"
        s.Values = wsC.Range("E2:E" & lastM)
        s.AxisGroup = xlSecondary
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = ClrRed()
        s.Format.Line.Weight = 1.75
        .HasTitle = True
        .ChartTitle.Text = "Headcount vs Monthly Attrition"
        StyleTitle .ChartTitle
        On Error Resume Next
        .Axes(xlValue, xlSecondary).TickLabels.NumberFormat = "0.0%"
        .Axes(xlCategory).TickLabels.NumberFormat = "mmm-yy"
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).TickLabels.Font.Size = 8
        On Error GoTo 0
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .ChartArea.Format.Line.Visible = msoFalse
    End With

    ' ---- 2. headcount by department (bar) ---------------------------------------
    Set co = ws.ChartObjects.Add(ws.Range("J13").Left, ws.Range("J13").Top, 405, 205)
    co.Name = "chDept"
    With co.Chart
        .ChartType = xlBarClustered
        Set s = .SeriesCollection.NewSeries
        s.Name = "Headcount"
        s.XValues = wsC.Range("J2:J" & lastD)
        s.Values = wsC.Range("K2:K" & lastD)
        s.Format.Fill.ForeColor.RGB = ClrNavy()
        s.HasDataLabels = True
        s.DataLabels.Font.Size = 8
        .HasTitle = True
        .ChartTitle.Text = "Active Headcount by Department"
        StyleTitle .ChartTitle
        On Error Resume Next
        .Axes(xlCategory).ReversePlotOrder = True
        .Axes(xlCategory).TickLabels.Font.Size = 8
        .Axes(xlValue).Delete
        On Error GoTo 0
        .Legend.Delete
        .ChartArea.Format.Line.Visible = msoFalse
    End With

    ' ---- 3. saudization by facility vs target -------------------------------------
    Set co = ws.ChartObjects.Add(ws.Range("B28").Left, ws.Range("B28").Top, 405, 205)
    co.Name = "chFacility"
    Dim tgt As Double: tgt = CfgVal("SaudizationTarget")
    Dim nFac As Long: nFac = lastF - 1
    Dim tgtArr() As Variant: ReDim tgtArr(1 To nFac)
    For i = 1 To nFac
        tgtArr(i) = tgt
    Next i
    With co.Chart
        .ChartType = xlColumnClustered
        Set s = .SeriesCollection.NewSeries
        s.Name = "Saudization"
        s.XValues = wsC.Range("S2:S" & lastF)
        s.Values = wsC.Range("U2:U" & lastF)
        For i = 1 To nFac
            If wsC.Cells(1 + i, 21).Value >= tgt Then
                s.Points(i).Format.Fill.ForeColor.RGB = ClrGreen()
            Else
                s.Points(i).Format.Fill.ForeColor.RGB = ClrAccent()
            End If
        Next i
        s.HasDataLabels = True
        s.DataLabels.NumberFormat = "0%"
        s.DataLabels.Font.Size = 8
        Set s = .SeriesCollection.NewSeries
        s.Name = "Nitaqat target"
        s.Values = tgtArr
        s.ChartType = xlLine
        s.Format.Line.ForeColor.RGB = ClrRed()
        s.Format.Line.DashStyle = msoLineDash
        s.Format.Line.Weight = 1.5
        .HasTitle = True
        .ChartTitle.Text = "Saudization by Facility vs 40% Target"
        StyleTitle .ChartTitle
        On Error Resume Next
        .Axes(xlValue).TickLabels.NumberFormat = "0%"
        .Axes(xlValue).TickLabels.Font.Size = 8
        .Axes(xlCategory).TickLabels.Font.Size = 7
        On Error GoTo 0
        .Legend.Position = xlLegendPositionBottom
        .Legend.Font.Size = 8
        .ChartArea.Format.Line.Visible = msoFalse
    End With

    ' ---- 4. leave mix doughnut (last 12 months) -------------------------------------
    Set co = ws.ChartObjects.Add(ws.Range("J28").Left, ws.Range("J28").Top, 405, 205)
    co.Name = "chLeave"
    With co.Chart
        .ChartType = xlDoughnut
        Set s = .SeriesCollection.NewSeries
        s.XValues = wsC.Range("W2:W" & lastL)
        s.Values = wsC.Range("X2:X" & lastL)
        s.HasDataLabels = True
        With s.DataLabels
            .ShowPercentage = True
            .ShowValue = False
            .Font.Size = 8
        End With
        .HasTitle = True
        .ChartTitle.Text = "Leave Days by Type (last 12 months)"
        StyleTitle .ChartTitle
        .Legend.Position = xlLegendPositionRight
        .Legend.Font.Size = 8
        .ChartArea.Format.Line.Visible = msoFalse
    End With
End Sub

Private Sub StyleTitle(ct As ChartTitle)
    ct.Font.Size = 11
    ct.Font.Bold = True
    ct.Font.Color = ClrNavy()
End Sub
