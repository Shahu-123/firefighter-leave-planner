Attribute VB_Name = "modLeaveEngine"
Option Explicit

' ============================================================
' modLeaveEngine - Credit pools, balance calculation, pro-ration
' Station Leave Manager (Event-Driven Model)
' ============================================================

Private Type CreditPool
    CreditID As Long
    LeaveType As String
    Remaining As Double
    ExpiryDate As Date
End Type

' ============================================================
' Get sorted credit pools (soonest-expiry-first)
' Returns Collection of arrays: (CreditID, LeaveType, Remaining, ExpiryDate)
' ============================================================

Public Function GetSortedCreditPools(displayName As String, Optional asOfDate As Date = 0) As Collection
    If asOfDate = 0 Then asOfDate = Date

    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    ' Gather CREDIT entries and compute remaining for each
    Dim pools() As CreditPool
    ReDim pools(1 To 200)
    Dim poolCount As Long
    poolCount = 0

    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If CStr(wsP.Cells(r, LC_COL_NAME).Value) = displayName And _
           CStr(wsP.Cells(r, LC_COL_TXN).Value) = TXN_CREDIT Then

            If IsEmpty(wsP.Cells(r, LC_COL_EXPIRY).Value) Then GoTo NextRow
            Dim expDate As Date
            expDate = CDate(wsP.Cells(r, LC_COL_EXPIRY).Value)
            If expDate < asOfDate Then GoTo NextRow  ' expired

            Dim creditID As Long
            creditID = CLng(wsP.Cells(r, LC_COL_ID).Value)
            Dim creditDays As Double
            creditDays = CDbl(wsP.Cells(r, LC_COL_DAYS).Value)

            ' Sum debits referencing this credit
            Dim debited As Double
            debited = SumDebitsForCredit(displayName, creditID)

            Dim rem As Double
            rem = creditDays - debited
            If rem > 0.001 Then
                poolCount = poolCount + 1
                pools(poolCount).CreditID = creditID
                pools(poolCount).LeaveType = CStr(wsP.Cells(r, LC_COL_LEAVE).Value)
                pools(poolCount).Remaining = rem
                pools(poolCount).ExpiryDate = expDate
            End If
        End If
NextRow:
    Next r

    ' Sort by expiry ascending, then OIL < PHOL < VL for same date
    Dim i As Long, j As Long
    Dim tmp As CreditPool
    For i = 1 To poolCount - 1
        For j = 1 To poolCount - i
            Dim doSwap As Boolean
            doSwap = False
            If pools(j).ExpiryDate > pools(j + 1).ExpiryDate Then
                doSwap = True
            ElseIf pools(j).ExpiryDate = pools(j + 1).ExpiryDate Then
                If LeavePriority(pools(j).LeaveType) > LeavePriority(pools(j + 1).LeaveType) Then
                    doSwap = True
                End If
            End If
            If doSwap Then
                tmp = pools(j)
                pools(j) = pools(j + 1)
                pools(j + 1) = tmp
            End If
        Next j
    Next i

    ' Build collection
    Dim result As New Collection
    For i = 1 To poolCount
        result.Add Array(pools(i).CreditID, pools(i).LeaveType, pools(i).Remaining, pools(i).ExpiryDate)
    Next i
    Set GetSortedCreditPools = result
End Function

Private Function LeavePriority(lt As String) As Long
    Select Case lt
        Case LEAVE_OIL: LeavePriority = 1
        Case LEAVE_PHOL: LeavePriority = 2
        Case LEAVE_VL: LeavePriority = 3
        Case Else: LeavePriority = 99
    End Select
End Function

Private Function SumDebitsForCredit(displayName As String, creditID As Long) As Double
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim total As Double
    total = 0
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If CStr(wsP.Cells(r, LC_COL_NAME).Value) = displayName And _
           CStr(wsP.Cells(r, LC_COL_TXN).Value) = TXN_DEBIT Then
            If wsP.Cells(r, LC_COL_SOURCE).Value <> "" Then
                If CLng(wsP.Cells(r, LC_COL_SOURCE).Value) = creditID Then
                    total = total + CDbl(wsP.Cells(r, LC_COL_DAYS).Value)
                End If
            End If
        End If
    Next r
    SumDebitsForCredit = total
End Function

' ============================================================
' Balance Calculation
' ============================================================

Public Function GetBalance(displayName As String, leaveType As String) As Double
    Dim pools As Collection
    Set pools = GetSortedCreditPools(displayName)
    Dim total As Double
    total = 0
    Dim p As Variant
    For Each p In pools
        If CStr(p(1)) = leaveType Then
            total = total + CDbl(p(2))
        End If
    Next p
    GetBalance = total
End Function

Public Function GetTotalBalance(displayName As String) As Double
    Dim pools As Collection
    Set pools = GetSortedCreditPools(displayName)
    Dim total As Double
    total = 0
    Dim p As Variant
    For Each p In pools
        total = total + CDbl(p(2))
    Next p
    GetTotalBalance = total
End Function

' Get total credits (lifetime) for a type
Public Function GetTotalCredits(displayName As String, leaveType As String) As Double
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim total As Double
    total = 0
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If CStr(wsP.Cells(r, LC_COL_NAME).Value) = displayName And _
           CStr(wsP.Cells(r, LC_COL_TXN).Value) = TXN_CREDIT And _
           CStr(wsP.Cells(r, LC_COL_LEAVE).Value) = leaveType Then
            total = total + CDbl(wsP.Cells(r, LC_COL_DAYS).Value)
        End If
    Next r
    GetTotalCredits = total
End Function

' Get total debits (lifetime) for a type
Public Function GetTotalDebits(displayName As String, leaveType As String) As Double
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim total As Double
    total = 0
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If CStr(wsP.Cells(r, LC_COL_NAME).Value) = displayName And _
           CStr(wsP.Cells(r, LC_COL_TXN).Value) = TXN_DEBIT And _
           CStr(wsP.Cells(r, LC_COL_LEAVE).Value) = leaveType Then
            total = total + CDbl(wsP.Cells(r, LC_COL_DAYS).Value)
        End If
    Next r
    GetTotalDebits = total
End Function

' ============================================================
' Refresh balance columns on Personnel sheet for one person
' ============================================================

Public Sub RefreshPersonBalance(displayName As String)
    Dim pRow As Long
    pRow = FindPersonRow(displayName)
    If pRow = 0 Then Exit Sub

    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    ' VL
    wsP.Cells(pRow, P_COL_VL_USED).Value = GetTotalDebits(displayName, LEAVE_VL)
    wsP.Cells(pRow, P_COL_VL_BAL).Value = GetBalance(displayName, LEAVE_VL)

    ' PHOL
    wsP.Cells(pRow, P_COL_PHOL_CR).Value = GetTotalCredits(displayName, LEAVE_PHOL)
    wsP.Cells(pRow, P_COL_PHOL_USED).Value = GetTotalDebits(displayName, LEAVE_PHOL)
    wsP.Cells(pRow, P_COL_PHOL_BAL).Value = GetBalance(displayName, LEAVE_PHOL)

    ' OIL
    wsP.Cells(pRow, P_COL_OIL_CR).Value = GetTotalCredits(displayName, LEAVE_OIL)
    wsP.Cells(pRow, P_COL_OIL_USED).Value = GetTotalDebits(displayName, LEAVE_OIL)
    wsP.Cells(pRow, P_COL_OIL_BAL).Value = GetBalance(displayName, LEAVE_OIL)

    ' Total
    wsP.Cells(pRow, P_COL_TOTAL_BAL).Value = GetTotalBalance(displayName)
End Sub

' Refresh ALL personnel balances
Public Sub RefreshAllBalances()
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn <> "" Then
            RefreshPersonBalance dn
        End If
    Next r
End Sub

' ============================================================
' Pro-rated VL
' ============================================================

Public Function CalcProRatedVL(joinDate As Date, Optional ordDate As Date = 0) As Double
    Dim fullYear As Long
    fullYear = CLng(GetConfig("VLDaysFullYear"))
    Dim yr As Long
    yr = GetYear()

    Dim startD As Date
    If joinDate > DateSerial(yr, 1, 1) Then startD = joinDate Else startD = DateSerial(yr, 1, 1)
    Dim endD As Date
    If ordDate > 0 And ordDate < DateSerial(yr, 12, 31) Then endD = ordDate Else endD = DateSerial(yr, 12, 31)

    Dim months As Long
    months = DateDiff("m", startD, endD) + 1
    If months < 0 Then months = 0
    If months > 12 Then months = 12

    CalcProRatedVL = Application.WorksheetFunction.RoundUp(fullYear * months / 12, 0)
End Function

' ============================================================
' Credit leave in the ledger
' ============================================================

Public Sub CreditLeave(displayName As String, leaveType As String, days As Double, _
                        creditDate As Date, expiryDate As Date, remarks As String)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    Dim newRow As Long
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If Trim(CStr(wsP.Cells(r, LC_COL_ID).Value)) = "" Then
            newRow = r
            Exit For
        End If
    Next r
    If newRow = 0 Then Exit Sub

    wsP.Cells(newRow, LC_COL_ID).Value = GetNextCreditID()
    wsP.Cells(newRow, LC_COL_NAME).Value = displayName
    wsP.Cells(newRow, LC_COL_TXN).Value = TXN_CREDIT
    wsP.Cells(newRow, LC_COL_LEAVE).Value = leaveType
    wsP.Cells(newRow, LC_COL_DAYS).Value = days
    wsP.Cells(newRow, LC_COL_DATE).Value = creditDate
    wsP.Cells(newRow, LC_COL_DATE).NumberFormat = "DD-MMM-YYYY"
    wsP.Cells(newRow, LC_COL_EXPIRY).Value = expiryDate
    wsP.Cells(newRow, LC_COL_EXPIRY).NumberFormat = "DD-MMM-YYYY"
    wsP.Cells(newRow, LC_COL_REMARKS).Value = remarks
End Sub
