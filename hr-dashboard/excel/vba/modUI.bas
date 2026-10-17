Attribute VB_Name = "modUI"
Option Explicit
'==============================================================================
' Builds the Dashboard sheet from a blank canvas: banner, action buttons,
' twelve KPI card frames, alerts mini-panel and print setup. Values are filled
' afterwards by modKPI / modAlerts; charts by modCharts.
'==============================================================================

Public Const CARD_R1 As Long = 5    ' first KPI card row block (rows 5-7)
Public Const CARD_R2 As Long = 9    ' second KPI card row block (rows 9-11)
Public Const SCORE_R As Long = 43   ' department scorecard title row

Public Sub BuildDashboardLayout()
    Dim ws As Worksheet: Set ws = GetWS("Dashboard")
    Dim sh As Shape, i As Long

    For i = ws.Shapes.Count To 1 Step -1
        ws.Shapes(i).Delete
    Next i
    ws.Cells.Clear
    ws.Cells.Interior.Color = ClrLight()
    ws.Columns("A").ColumnWidth = 2
    ws.Columns("B:R").ColumnWidth = 11.5
    ws.Rows(1).RowHeight = 6
    ws.Rows(2).RowHeight = 32
    ws.Rows(3).RowHeight = 18
    ws.Rows(4).RowHeight = 26

    ' ---- banner -------------------------------------------------------------
    With ws.Range("B2:R2")
        .Merge
        .Value = "CONTOSO HEALTH KSA (FICTIONAL)  -  HR COMMAND CENTER"
        .Interior.Color = ClrNavy()
        .Font.Color = vbWhite: .Font.Size = 18: .Font.Bold = True
        .Font.Name = "Segoe UI Semibold"
        .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With
    With ws.Range("B3:R3")
        .Merge
        .Value = "Initializing..."
        .Interior.Color = ClrTeal()
        .Font.Color = vbWhite: .Font.Size = 10: .Font.Name = "Segoe UI"
        .VerticalAlignment = xlCenter: .IndentLevel = 1
    End With

    ' ---- action buttons (row 4) ----------------------------------------------
    AddButton ws, ws.Cells(4, 2), "REFRESH DATA", "modMain.RefreshAllUI", ClrTeal()
    AddButton ws, ws.Cells(4, 4), "REBUILD ALL", "modMain.BuildAllUI", ClrNavy()
    AddButton ws, ws.Cells(4, 6), "FIND EMPLOYEE", "modTools.EmployeeSearch_UI", ClrAccent()
    AddButton ws, ws.Cells(4, 8), "EXPORT PDF", "modTools.ExportPDF_UI", ClrGreen()
    AddButton ws, ws.Cells(4, 10), "ALERTS", "modMain.GoToAlerts", ClrRed()
    AddButton ws, ws.Cells(4, 12), "PIVOT EXPLORER", "modMain.GoToPivots", ClrGrey()

    ' ---- KPI card frames ------------------------------------------------------
    Dim labels1 As Variant, labels2 As Variant
    labels1 = Array("ACTIVE HEADCOUNT", "SAUDIZATION", "ATTRITION (12M)", _
                    "AVG TENURE", "MONTHLY PAYROLL (SAR)", "OPEN POSITIONS")
    labels2 = Array("FEMALE RATIO", "AVG SALARY (SAR)", "AVG TIME TO FILL", _
                    "TRAINING HRS (12M)", "TRAINING COMPLETION", "ABSENTEEISM")
    For i = 0 To 5
        CardFrame ws, CARD_R1, i, CStr(labels1(i))
        CardFrame ws, CARD_R2, i, CStr(labels2(i))
    Next i
    ws.Rows(CARD_R1 + 1).RowHeight = 28
    ws.Rows(CARD_R2 + 1).RowHeight = 28

    ' ---- alerts mini-panel N5:R11 --------------------------------------------
    With ws.Range(ws.Cells(CARD_R1, 14), ws.Cells(CARD_R2 + 2, 18))
        .Interior.Color = vbWhite
        .BorderAround Color:=ClrBorder(), Weight:=xlThin
    End With
    With ws.Range(ws.Cells(CARD_R1, 14), ws.Cells(CARD_R1, 18))
        .Merge
        .Value = "COMPLIANCE RADAR"
        .Font.Size = 9: .Font.Bold = True: .Font.Color = ClrNavy()
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Interior.Color = ClrLight()
    End With
    For i = 1 To 6
        ws.Range(ws.Cells(CARD_R1 + i, 14), ws.Cells(CARD_R1 + i, 18)).Merge
        ws.Cells(CARD_R1 + i, 14).Font.Size = 9
        ws.Cells(CARD_R1 + i, 14).IndentLevel = 1
    Next i

    ' ---- scorecard title -------------------------------------------------------
    With ws.Cells(SCORE_R, 2)
        .Value = "DEPARTMENT SCORECARD"
        .Font.Size = 12: .Font.Bold = True: .Font.Color = ClrNavy()
    End With

    ' ---- print setup ------------------------------------------------------------
    On Error Resume Next
    Application.PrintCommunication = False
    With ws.PageSetup
        .PrintArea = "$A$1:$S$" & (SCORE_R + 16)
        .Orientation = xlLandscape
        .Zoom = False
        .FitToPagesWide = 1
        .FitToPagesTall = 1
        .CenterHorizontally = True
    End With
    Application.PrintCommunication = True
    On Error GoTo 0
End Sub

Private Sub CardFrame(ws As Worksheet, ByVal r As Long, ByVal i As Long, ByVal labelText As String)
    Dim c As Long: c = 2 + i * 2
    With ws.Range(ws.Cells(r, c), ws.Cells(r + 2, c + 1))
        .Interior.Color = vbWhite
        .BorderAround Color:=ClrBorder(), Weight:=xlMedium
    End With
    With ws.Range(ws.Cells(r, c), ws.Cells(r, c + 1))
        .Merge
        .Value = labelText
        .Font.Size = 8: .Font.Bold = True: .Font.Color = ClrGrey()
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlBottom
    End With
    With ws.Range(ws.Cells(r + 1, c), ws.Cells(r + 1, c + 1))
        .Merge
        .HorizontalAlignment = xlCenter: .VerticalAlignment = xlCenter
        .Font.Size = 18: .Font.Bold = True: .Font.Color = ClrNavy()
    End With
    With ws.Range(ws.Cells(r + 2, c), ws.Cells(r + 2, c + 1))
        .Merge
        .HorizontalAlignment = xlCenter
        .Font.Size = 8: .Font.Color = ClrGrey()
    End With
End Sub

Private Sub AddButton(ws As Worksheet, anchor As Range, ByVal caption As String, _
                      ByVal macroName As String, ByVal fillClr As Long)
    Dim sh As Shape
    Set sh = ws.Shapes.AddShape(msoShapeRoundedRectangle, _
                                anchor.Left + 2, anchor.Top + 3, 118, 20)
    sh.Fill.ForeColor.RGB = fillClr
    sh.Line.Visible = msoFalse
    With sh.TextFrame2
        .TextRange.Text = caption
        .TextRange.Font.Size = 9
        .TextRange.Font.Bold = msoTrue
        .TextRange.Font.Fill.ForeColor.RGB = RGB(255, 255, 255)
        .TextRange.ParagraphFormat.Alignment = msoAlignCenter
        .VerticalAnchor = msoAnchorMiddle
        .MarginLeft = 0: .MarginRight = 0: .MarginTop = 0: .MarginBottom = 0
    End With
    sh.OnAction = macroName
    sh.Name = "btn_" & Replace(macroName, ".", "_")
End Sub
