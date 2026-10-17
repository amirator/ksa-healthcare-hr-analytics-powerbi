Attribute VB_Name = "modTools"
Option Explicit
'==============================================================================
' Interactive tools: instant employee search across all columns of interest,
' and one-click PDF export of the dashboard.
'==============================================================================

Public Sub EmployeeSearch_UI()
    Dim q As String
    q = Trim(InputBox("Search by ID, name, department, job title or facility:" & _
        vbCrLf & vbCrLf & "Examples:  AMG-2021,  Nursing,  Al-Qahtani,  Dammam", _
        "HR - Find Employee"))
    If Len(q) = 0 Then Exit Sub
    q = LCase(q)

    Dim loE As ListObject: Set loE = GetLO("tblEmployees")
    Dim a As Variant: a = loE.DataBodyRange.Value2
    Dim nCols As Long: nCols = UBound(a, 2)
    Dim searchCols As Variant
    searchCols = Array(ColIx(loE, "EmployeeID"), ColIx(loE, "FullName"), _
                       ColIx(loE, "Department"), ColIx(loE, "JobTitle"), _
                       ColIx(loE, "Facility"), ColIx(loE, "Nationality"))
    Dim hits() As Long, nHits As Long
    ReDim hits(1 To UBound(a, 1))
    Dim i As Long, j As Long
    For i = 1 To UBound(a, 1)
        For j = LBound(searchCols) To UBound(searchCols)
            If InStr(1, LCase(CStr(a(i, searchCols(j)))), q) > 0 Then
                nHits = nHits + 1
                hits(nHits) = i
                Exit For
            End If
        Next j
    Next i

    If nHits = 0 Then
        MsgBox "No employees match """ & q & """.", vbInformation, "HR Dashboard"
        Exit Sub
    End If

    Dim ws As Worksheet: Set ws = GetWS("SearchResults")
    Dim loOld As ListObject
    For Each loOld In ws.ListObjects
        loOld.Unlist
    Next loOld
    ws.Cells.Clear
    With ws.Range("A1")
        .Value = "SEARCH RESULTS for """ & q & """  -  " & nHits & " match(es)"
        .Font.Size = 12: .Font.Bold = True: .Font.Color = ClrNavy()
    End With
    Dim hdrRng As Range
    Set hdrRng = loE.HeaderRowRange
    ws.Range("A3").Resize(1, nCols).Value = hdrRng.Value
    Dim out() As Variant: ReDim out(1 To nHits, 1 To nCols)
    For i = 1 To nHits
        For j = 1 To nCols
            out(i, j) = a(hits(i), j)
        Next j
    Next i
    ws.Range("A4").Resize(nHits, nCols).Value = out
    ' restore date formats (Value2 carries serials)
    Dim dateCols As Variant: dateCols = Array("HireDate", "TerminationDate", "BirthDate")
    For j = LBound(dateCols) To UBound(dateCols)
        ws.Range(ws.Cells(4, ColIx(loE, CStr(dateCols(j)))), _
                 ws.Cells(3 + nHits, ColIx(loE, CStr(dateCols(j))))).NumberFormat = "dd-mmm-yyyy"
    Next j
    Dim loNew As ListObject
    Set loNew = ws.ListObjects.Add(xlSrcRange, ws.Range("A3").Resize(nHits + 1, nCols), , xlYes)
    loNew.Name = "tblSearchResults"
    loNew.TableStyle = "TableStyleMedium2"
    ws.Columns("A:T").AutoFit
    ws.Activate
    ws.Range("A4").Select
End Sub

Public Sub ExportPDF_UI()
    Dim f As String
    f = ThisWorkbook.Path & Application.PathSeparator & _
        "HR_Dashboard_" & Format(Now, "yyyymmdd_hhnn") & ".pdf"
    On Error GoTo fail
    GetWS("Dashboard").ExportAsFixedFormat Type:=xlTypePDF, Filename:=f, _
        Quality:=xlQualityStandard, IncludeDocProperties:=True, OpenAfterPublish:=True
    Exit Sub
fail:
    MsgBox "PDF export failed: " & Err.Description, vbExclamation, "HR Dashboard"
End Sub
