Attribute VB_Name = "modCalendarEvents"
Option Explicit

' ============================================================
' modCalendarEvents - Handles Calendar worksheet changes
'
' When a name is selected/removed from a slot cell on the Calendar,
' this module handles leave deduction, reversal, and slot bumping.
'
' IMPORTANT: Paste this into the Calendar sheet code module:
'
'   Private Sub Worksheet_Change(ByVal Target As Range)
'       modCalendarEvents.HandleCalendarChange Target
'   End Sub
'
' ============================================================

' Module-level variable to track old value (set by SelectionChange)
Private mOldValue As String
Private mOldRow As Long
Private mOldCol As Long

' Call this from Calendar's Worksheet_SelectionChange
Public Sub CacheOldValue(ByVal Target As Range)
    If Target.Cells.Count = 1 Then
        If Target.Column >= CAL_SLOT_FIRST_COL And Target.Column <= CAL_SLOT_LAST_COL Then
            If IsCalendarDataRow(Target.Row) Then
                mOldValue = Trim(CStr(Target.Value))
                mOldRow = Target.Row
                mOldCol = Target.Column
                Exit Sub
            End If
        End If
    End If
    mOldValue = ""
    mOldRow = 0
    mOldCol = 0
End Sub

' Main handler called from Calendar's Worksheet_Change
Public Sub HandleCalendarChange(ByVal Target As Range)
    ' Only process single-cell changes in slot columns
    If Target.Cells.Count <> 1 Then Exit Sub
    If Target.Column < CAL_SLOT_FIRST_COL Or Target.Column > CAL_SLOT_LAST_COL Then Exit Sub
    If Not IsCalendarDataRow(Target.Row) Then Exit Sub

    Dim newValue As String
    newValue = Trim(CStr(Target.Value))

    ' Determine old value
    Dim oldValue As String
    If mOldRow = Target.Row And mOldCol = Target.Column Then
        oldValue = mOldValue
    Else
        oldValue = ""
    End If

    ' Nothing changed
    If oldValue = newValue Then Exit Sub

    Application.EnableEvents = False
    Application.ScreenUpdating = False
    On Error GoTo Cleanup

    Dim calDate As Date
    calDate = CDate(ThisWorkbook.Sheets(WS_CALENDAR).Cells(Target.Row, CAL_COL_DATE).Value)

    ' --- CASE 1: Name was removed (cell cleared) ---
    If newValue = "" And oldValue <> "" Then
        ' Reverse leave deduction
        ReverseDeduction oldValue, calDate
        ' Bump everyone below up
        BumpSlotsUp Target.Row, Target.Column
        ' Refresh balance for that person
        RefreshPersonBalance oldValue
    End If

    ' --- CASE 2: Name was added (cell was empty) ---
    If newValue <> "" And oldValue = "" Then
        ' Check for duplicate on same date
        If IsAlreadyOnDate(newValue, Target.Row, Target.Column) Then
            MsgBox newValue & " is already on leave on " & Format(calDate, "DD-MMM") & ".", _
                   vbExclamation, "Duplicate"
            Target.Value = ""
            GoTo Cleanup
        End If
        ' Deduct leave
        Dim ok As Boolean
        ok = DeductLeaveForBooking(newValue, calDate)
        If Not ok Then
            MsgBox "Insufficient leave balance for " & newValue & "." & vbCrLf & _
                   "Need " & GetShiftCost() & " days.", vbExclamation, "Cannot Book"
            Target.Value = ""
            GoTo Cleanup
        End If
        RefreshPersonBalance newValue
    End If

    ' --- CASE 3: Name was swapped ---
    If newValue <> "" And oldValue <> "" And newValue <> oldValue Then
        ' Reverse old
        ReverseDeduction oldValue, calDate
        RefreshPersonBalance oldValue
        ' Check duplicate
        If IsAlreadyOnDate(newValue, Target.Row, Target.Column) Then
            MsgBox newValue & " is already on leave on " & Format(calDate, "DD-MMM") & ".", _
                   vbExclamation, "Duplicate"
            Target.Value = ""
            GoTo Cleanup
        End If
        ' Deduct new
        Dim ok2 As Boolean
        ok2 = DeductLeaveForBooking(newValue, calDate)
        If Not ok2 Then
            MsgBox "Insufficient leave balance for " & newValue & ".", vbExclamation, "Cannot Book"
            Target.Value = ""
            ' Re-deduct old person since we already reversed them
            DeductLeaveForBooking oldValue, calDate
            Target.Value = oldValue
            GoTo Cleanup
        End If
        RefreshPersonBalance newValue
    End If

Cleanup:
    Application.EnableEvents = True
    Application.ScreenUpdating = True
End Sub

' Check if a name already appears in another slot on the same row
Private Function IsAlreadyOnDate(displayName As String, row As Long, excludeCol As Long) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_CALENDAR)
    Dim c As Long
    For c = CAL_SLOT_FIRST_COL To CAL_SLOT_LAST_COL
        If c <> excludeCol Then
            If Trim(CStr(ws.Cells(row, c).Value)) = displayName Then
                IsAlreadyOnDate = True
                Exit Function
            End If
        End If
    Next c
    IsAlreadyOnDate = False
End Function

' ============================================================
' SLOT BUMPING: Move everyone below the cleared slot up by one
' ============================================================

Private Sub BumpSlotsUp(calRow As Long, clearedCol As Long)
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_CALENDAR)

    Dim c As Long
    For c = clearedCol To CAL_SLOT_LAST_COL - 1
        ws.Cells(calRow, c).Value = ws.Cells(calRow, c + 1).Value
    Next c
    ' Clear the last slot
    ws.Cells(calRow, CAL_SLOT_LAST_COL).Value = ""
End Sub

' ============================================================
' LEAVE DEDUCTION (soonest-expiry-first)
' ============================================================

Private Function DeductLeaveForBooking(displayName As String, calDate As Date) As Boolean
    Dim cost As Double
    cost = GetShiftCost()  ' 2

    ' Get sorted credit pools (soonest expiry first)
    Dim pools As Collection
    Set pools = GetSortedCreditPools(displayName)

    ' Check total available
    Dim total As Double
    total = 0
    Dim p As Variant
    For Each p In pools
        total = total + p(2)  ' remaining
    Next p
    If total < cost Then
        DeductLeaveForBooking = False
        Exit Function
    End If

    ' Plan deductions
    Dim remaining As Double
    remaining = cost
    Dim debitPlan As New Collection

    For Each p In pools
        If remaining <= 0 Then Exit For
        Dim amt As Double
        If CDbl(p(2)) < remaining Then amt = CDbl(p(2)) Else amt = remaining
        debitPlan.Add Array(p(0), p(1), amt, p(3))  ' sourceID, leaveType, amount, expiry
        remaining = remaining - amt
    Next p

    If remaining > 0.001 Then
        DeductLeaveForBooking = False
        Exit Function
    End If

    ' Write DEBIT entries to Leave Credits table
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim debit As Variant
    For Each debit In debitPlan
        Dim newRow As Long
        newRow = FindNextEmptyCreditRow()
        Dim newID As Long
        newID = GetNextCreditID()

        wsP.Cells(newRow, LC_COL_ID).Value = newID
        wsP.Cells(newRow, LC_COL_NAME).Value = displayName
        wsP.Cells(newRow, LC_COL_TXN).Value = TXN_DEBIT
        wsP.Cells(newRow, LC_COL_LEAVE).Value = CStr(debit(1))
        wsP.Cells(newRow, LC_COL_DAYS).Value = CDbl(debit(2))
        wsP.Cells(newRow, LC_COL_DATE).Value = Now()
        wsP.Cells(newRow, LC_COL_DATE).NumberFormat = "DD-MMM-YYYY"
        wsP.Cells(newRow, LC_COL_EXPIRY).Value = CDate(debit(3))
        wsP.Cells(newRow, LC_COL_EXPIRY).NumberFormat = "DD-MMM-YYYY"
        wsP.Cells(newRow, LC_COL_SOURCE).Value = CLng(debit(0))
        wsP.Cells(newRow, LC_COL_CALREF).Value = calDate
        wsP.Cells(newRow, LC_COL_CALREF).NumberFormat = "DD-MMM-YYYY"
        wsP.Cells(newRow, LC_COL_REMARKS).Value = "Shift leave " & Format(calDate, "DD-MMM-YYYY")
    Next debit

    DeductLeaveForBooking = True
End Function

' ============================================================
' REVERSAL: undo deductions for a person on a specific date
' ============================================================

Private Sub ReverseDeduction(displayName As String, calDate As Date)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If CStr(wsP.Cells(r, LC_COL_NAME).Value) = displayName And _
           CStr(wsP.Cells(r, LC_COL_TXN).Value) = TXN_DEBIT Then
            ' Check calendar reference matches
            If Not IsEmpty(wsP.Cells(r, LC_COL_CALREF).Value) Then
                If CDate(wsP.Cells(r, LC_COL_CALREF).Value) = calDate Then
                    ' Create reversal CREDIT
                    Dim newRow As Long
                    newRow = FindNextEmptyCreditRow()
                    Dim newID As Long
                    newID = GetNextCreditID()

                    wsP.Cells(newRow, LC_COL_ID).Value = newID
                    wsP.Cells(newRow, LC_COL_NAME).Value = displayName
                    wsP.Cells(newRow, LC_COL_TXN).Value = TXN_CREDIT
                    wsP.Cells(newRow, LC_COL_LEAVE).Value = wsP.Cells(r, LC_COL_LEAVE).Value
                    wsP.Cells(newRow, LC_COL_DAYS).Value = wsP.Cells(r, LC_COL_DAYS).Value
                    wsP.Cells(newRow, LC_COL_DATE).Value = Now()
                    wsP.Cells(newRow, LC_COL_DATE).NumberFormat = "DD-MMM-YYYY"
                    wsP.Cells(newRow, LC_COL_EXPIRY).Value = wsP.Cells(r, LC_COL_EXPIRY).Value
                    wsP.Cells(newRow, LC_COL_EXPIRY).NumberFormat = "DD-MMM-YYYY"
                    wsP.Cells(newRow, LC_COL_CALREF).Value = calDate
                    wsP.Cells(newRow, LC_COL_CALREF).NumberFormat = "DD-MMM-YYYY"
                    wsP.Cells(newRow, LC_COL_REMARKS).Value = "Reversal: " & Format(calDate, "DD-MMM-YYYY")
                End If
            End If
        End If
    Next r
End Sub

' Find next empty row in the credits table
Private Function FindNextEmptyCreditRow() As Long
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If Trim(CStr(wsP.Cells(r, LC_COL_ID).Value)) = "" Then
            FindNextEmptyCreditRow = r
            Exit Function
        End If
    Next r
    FindNextEmptyCreditRow = LC_DATA_START + 200  ' overflow
End Function
