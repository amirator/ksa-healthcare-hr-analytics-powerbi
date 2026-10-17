Attribute VB_Name = "modMain"
Option Explicit
'==============================================================================
' Orchestration. BuildAll is the full silent pipeline (also run once at build
' time); the *_UI subs are wired to the dashboard buttons.
'==============================================================================

Public Sub BuildAll()
    SpeedOn "Building HR Command Center..."
    On Error GoTo fail
    BuildSnapshotsAndCalc
    BuildDashboardLayout
    ComputeKPIs
    BuildCharts
    BuildPivots
    BuildAlerts
    GetWS("Dashboard").Activate
    On Error Resume Next
    ActiveWindow.DisplayGridlines = False
    ActiveWindow.Zoom = 90
    On Error GoTo 0
    SpeedOff
    Exit Sub
fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    SpeedOff
    Err.Raise n, "BuildAll", d
End Sub

Public Sub BuildAllUI()
    Dim t0 As Single: t0 = Timer
    BuildAll
    MsgBox "Full rebuild complete in " & Format(Timer - t0, "0.0") & "s." & vbCrLf & _
           "Snapshots, KPIs, charts, pivots and alerts are all up to date.", _
           vbInformation, "HR Command Center"
End Sub

' Refresh data-driven content without rebuilding the layout/buttons
Public Sub RefreshAllUI()
    Dim t0 As Single: t0 = Timer
    SpeedOn "Refreshing data..."
    On Error GoTo fail
    BuildSnapshotsAndCalc
    ComputeKPIs
    BuildCharts
    BuildPivots
    BuildAlerts
    GetWS("Dashboard").Activate
    SpeedOff
    MsgBox "Data refreshed in " & Format(Timer - t0, "0.0") & "s  -  " & _
           "~53,000 employee-month evaluations recomputed.", _
           vbInformation, "HR Command Center"
    Exit Sub
fail:
    Dim n As Long, d As String
    n = Err.Number: d = Err.Description
    SpeedOff
    MsgBox "Refresh failed: " & d, vbExclamation, "HR Command Center"
End Sub

Public Sub GoToAlerts()
    GetWS("Alerts").Activate
End Sub

Public Sub GoToPivots()
    GetWS("PivotExplorer").Activate
End Sub
