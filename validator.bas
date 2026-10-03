Attribute VB_Name = "Validator"
Option Explicit

Private Const COL_DATE As Long = 1
Private Const COL_DESC As Long = 2
Private Const COL_CAT As Long = 3
Private Const COL_TYPE As Long = 4
Private Const COL_AMT As Long = 5
Private Const COL_STATUS As Long = 6

Private wsRep As Worksheet
Private issueRow As Long
Private rowIssueCount As Long
Private errCount As Long
Private warnCount As Long

Public Sub ValidateLedger()
    Dim wsT As Worksheet, wsB As Worksheet, wsRules As Worksheet, wsS As Worksheet
    Dim lastRow As Long, lastBud As Long, r As Long, checked As Long
    Dim maxAmt As Double, tol As Double
    Dim cats As Object, seen As Object
    Dim dt As Variant, amt As Variant
    Dim desc As String, cat As String, typ As String, key As String
    Dim calcInc As Double, calcExp As Double

    On Error GoTo Fail
    Application.ScreenUpdating = False
    Application.Calculate

    Set wsT = ThisWorkbook.Worksheets("Transactions")
    Set wsB = ThisWorkbook.Worksheets("Budget")
    Set wsRules = ThisWorkbook.Worksheets("Rules")
    Set wsS = ThisWorkbook.Worksheets("Summary")
    maxAmt = wsRules.Range("B2").Value
    tol = wsRules.Range("B3").Value

    ClearValidation
    Set wsRep = GetReportSheet()
    wsRep.Range("A4:D4").Value = Array("Source", "Reference", "Severity", "Issue")
    With wsRep.Range("A4:D4")
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = RGB(31, 56, 100)
    End With
    issueRow = 5
    errCount = 0
    warnCount = 0

    Set cats = CreateObject("Scripting.Dictionary")
    cats.CompareMode = vbTextCompare
    lastBud = wsB.Cells(wsB.Rows.Count, 1).End(xlUp).Row
    For r = 2 To lastBud
        cats(Trim$(CStr(wsB.Cells(r, 1).Value))) = True
    Next r

    Set seen = CreateObject("Scripting.Dictionary")
    lastRow = wsT.UsedRange.Row + wsT.UsedRange.Rows.Count - 1

    For r = 2 To lastRow
        dt = wsT.Cells(r, COL_DATE).Value
        desc = Trim$(CStr(wsT.Cells(r, COL_DESC).Value))
        cat = Trim$(CStr(wsT.Cells(r, COL_CAT).Value))
        typ = Trim$(CStr(wsT.Cells(r, COL_TYPE).Value))
        amt = wsT.Cells(r, COL_AMT).Value
        rowIssueCount = 0
        checked = checked + 1

        ' Date
        If IsError(dt) Then
            AddIssue "Row " & r, desc, "Error", "Date contains an Excel error value"
        ElseIf IsEmpty(dt) Then
            AddIssue "Row " & r, desc, "Error", "Date is missing"
        ElseIf VarType(dt) <> vbDate Then
            AddIssue "Row " & r, desc, "Error", "Date is not a valid date: '" & dt & "'"
        ElseIf dt > Date Then
            AddIssue "Row " & r, desc, "Error", "Date is in the future (" & Format(dt, "yyyy-mm-dd") & ")"
        End If

        ' Type
        If typ <> "Income" And typ <> "Expense" Then
            AddIssue "Row " & r, desc, "Error", "Type must be Income or Expense"
        End If

        ' Category
        If Len(cat) = 0 Then
            AddIssue "Row " & r, desc, "Error", "Category is missing"
        ElseIf typ = "Income" Then
            If StrComp(cat, "Income", vbTextCompare) <> 0 Then
                AddIssue "Row " & r, desc, "Error", "Income rows must use the category 'Income'"
            End If
        ElseIf typ = "Expense" Then
            If Not cats.Exists(cat) Then
                AddIssue "Row " & r, desc, "Error", "Category '" & cat & "' is not on the Budget sheet"
            End If
        End If

        ' Amount
        If IsError(amt) Then
            AddIssue "Row " & r, desc, "Error", "Amount contains an Excel error value"
        ElseIf IsEmpty(amt) Then
            AddIssue "Row " & r, desc, "Error", "Amount is missing"
        ElseIf VarType(amt) = vbString Then
            AddIssue "Row " & r, desc, "Error", "Amount is stored as text: '" & amt & "'"
        Else
            If amt < 0 Then
                AddIssue "Row " & r, desc, "Warning", "Negative amount; record refunds as Income"
            ElseIf amt = 0 Then
                AddIssue "Row " & r, desc, "Error", "Amount is zero"
            End If
            If amt > maxAmt Then
                AddIssue "Row " & r, desc, "Error", "Amount exceeds the maximum of " & Format(maxAmt, "#,##0")
            End If
        End If

        ' Possible duplicates
        If VarType(dt) = vbDate And (VarType(amt) = vbDouble Or VarType(amt) = vbCurrency Or VarType(amt) = vbLong) Then
            key = Format(dt, "yyyy-mm-dd") & "|" & LCase$(desc) & "|" & CStr(amt)
            If seen.Exists(key) Then
                AddIssue "Row " & r, desc, "Warning", "Possible duplicate of row " & seen(key)
            Else
                seen.Add key, r
            End If
        End If

        ' Row status (text and colour)
        If rowIssueCount = 0 Then
            wsT.Cells(r, COL_STATUS).Value = "OK"
            wsT.Cells(r, COL_STATUS).Interior.Color = RGB(198, 239, 206)
        Else
            wsT.Cells(r, COL_STATUS).Value = "Check (" & rowIssueCount & ")"
            wsT.Cells(r, COL_STATUS).Interior.Color = RGB(255, 199, 206)
        End If
    Next r

    ' Independent totals versus Summary sheet
    With Application.WorksheetFunction
        calcInc = .SumIfs(wsT.Range("E2:E" & lastRow), wsT.Range("D2:D" & lastRow), "Income")
        calcExp = .SumIfs(wsT.Range("E2:E" & lastRow), wsT.Range("D2:D" & lastRow), "Expense")
    End With
    If Abs(calcInc - wsS.Range("B1").Value) > tol Then _
        AddIssue "Summary", "B1", "Error", "Total income differs from the recalculated total"
    If Abs(calcExp - wsS.Range("B2").Value) > tol Then _
        AddIssue "Summary", "B2", "Error", "Total expenses differ from the recalculated total"
    If Abs((calcInc - calcExp) - wsS.Range("B5").Value) > tol Then _
        AddIssue "Summary", "B5", "Error", "Reported net (" & Format(wsS.Range("B5").Value, "#,##0.00") & _
            ") does not match calculated net (" & Format(calcInc - calcExp, "#,##0.00") & ")"

    ' Budget limits
    For r = 9 To 15
        If wsS.Cells(r, 4).Value < 0 Then
            AddIssue "Summary", CStr(wsS.Cells(r, 1).Value), "Warning", _
                "Over budget by " & Format(-wsS.Cells(r, 4).Value, "#,##0.00")
        End If
    Next r

    wsRep.Range("A1").Value = "Life Ledger Validation Report"
    wsRep.Range("A1").Font.Bold = True
    wsRep.Range("A1").Font.Size = 14
    wsRep.Range("A2").Value = "Run: " & Format(Now, "yyyy-mm-dd hh:nn") & "  |  Rows checked: " & checked & _
        "  |  Errors: " & errCount & "  |  Warnings: " & warnCount
    wsRep.Columns("A:D").AutoFit

    Application.ScreenUpdating = True
    MsgBox "Validation complete." & vbCrLf & "Rows checked: " & checked & vbCrLf & _
        "Errors: " & errCount & vbCrLf & "Warnings: " & warnCount, vbInformation
    Exit Sub

Fail:
    Application.ScreenUpdating = True
    MsgBox "Validation stopped: " & Err.Description, vbCritical
End Sub

Public Sub ClearValidation()
    Dim wsT As Worksheet, lastRow As Long
    Set wsT = ThisWorkbook.Worksheets("Transactions")
    lastRow = wsT.UsedRange.Row + wsT.UsedRange.Rows.Count - 1
    If lastRow < 2 Then Exit Sub
    wsT.Range(wsT.Cells(2, COL_STATUS), wsT.Cells(lastRow, COL_STATUS)).ClearContents
    wsT.Range(wsT.Cells(2, COL_STATUS), wsT.Cells(lastRow, COL_STATUS)).Interior.ColorIndex = xlNone
End Sub

Private Sub AddIssue(ByVal src As String, ByVal ref As String, ByVal sev As String, ByVal issue As String)
    wsRep.Cells(issueRow, 1).Value = src
    wsRep.Cells(issueRow, 2).Value = ref
    wsRep.Cells(issueRow, 3).Value = sev
    wsRep.Cells(issueRow, 4).Value = issue
    issueRow = issueRow + 1
    rowIssueCount = rowIssueCount + 1
    If sev = "Error" Then errCount = errCount + 1 Else warnCount = warnCount + 1
End Sub

Private Function GetReportSheet() As Worksheet
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets("Report")
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = "Report"
    Else
        ws.Cells.Clear
    End If
    Set GetReportSheet = ws
End Function
