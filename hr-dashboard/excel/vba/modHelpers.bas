Attribute VB_Name = "modHelpers"
Option Explicit

' ---- Brand palette -------------------------------------------------
Public Function ClrNavy() As Long: ClrNavy = RGB(11, 79, 108): End Function
Public Function ClrTeal() As Long: ClrTeal = RGB(14, 124, 123): End Function
Public Function ClrLight() As Long: ClrLight = RGB(243, 247, 247): End Function
Public Function ClrAccent() As Long: ClrAccent = RGB(244, 162, 89): End Function
Public Function ClrRed() As Long: ClrRed = RGB(196, 69, 54): End Function
Public Function ClrGreen() As Long: ClrGreen = RGB(30, 158, 106): End Function
Public Function ClrGrey() As Long: ClrGrey = RGB(59, 74, 84): End Function
Public Function ClrBorder() As Long: ClrBorder = RGB(220, 231, 231): End Function

' ---- performance switches -------------------------------------------------
Public Sub SpeedOn(Optional ByVal msg As String = "HR engine working...")
    With Application
        .ScreenUpdating = False
        .EnableEvents = False
        .Calculation = xlCalculationManual
        .StatusBar = msg
    End With
End Sub

Public Sub SpeedOff()
    With Application
        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .StatusBar = False
    End With
End Sub

' ---- object access --------------------------------------------------------
Public Function GetWS(ByVal sheetName As String) As Worksheet
    On Error Resume Next
    Set GetWS = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0
    If GetWS Is Nothing Then
        Set GetWS = ThisWorkbook.Worksheets.Add( _
            After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        GetWS.Name = sheetName
    End If
End Function

Public Function GetLO(ByVal tableName As String) As ListObject
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        On Error Resume Next
        Set GetLO = ws.ListObjects(tableName)
        On Error GoTo 0
        If Not GetLO Is Nothing Then Exit Function
    Next ws
    Err.Raise vbObjectError + 513, "GetLO", "Table not found: " & tableName
End Function

Public Function ColIx(ByVal loRef As ListObject, ByVal colName As String) As Long
    ColIx = loRef.ListColumns(colName).Index
End Function

Public Function NewDict() As Object
    Set NewDict = CreateObject("Scripting.Dictionary")
End Function

' Configuration values from the Config sheet (tblConfig)
Public Function CfgVal(ByVal key As String) As Double
    Static cfg As Object
    Dim a As Variant, i As Long
    If cfg Is Nothing Then
        Set cfg = NewDict()
        a = GetLO("tblConfig").DataBodyRange.Value2
        For i = 1 To UBound(a, 1)
            cfg(CStr(a(i, 1))) = CDbl(a(i, 2))
        Next i
    End If
    CfgVal = cfg(key)
End Function
