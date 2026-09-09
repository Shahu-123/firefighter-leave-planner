Attribute VB_Name = "modInit"
Option Explicit

' ============================================================
' modInit - Workbook initialization and utility macros
' Station Leave Manager
'
' IMPORTANT: Add this to ThisWorkbook code module:
'
'   Private Sub Workbook_Open()
'       modInit.OnWorkbookOpen
'   End Sub
'
' ============================================================

' Called on workbook open
Public Sub OnWorkbookOpen()
    ' Refresh the name lists for Calendar dropdowns
    RefreshNameLists

    ' Auto-credit VL for any personnel that don't have it yet
    CheckAndCreditVL

    ' Refresh all balances
    RefreshAllBalances
End Sub

' Manual refresh - can be assigned to a button or run from Alt+F8
Public Sub RefreshAll()
    Application.ScreenUpdating = False
    Application.EnableEvents = False

    RefreshNameLists
    CheckAndCreditVL
    RefreshAllBalances
    RefreshDashboard

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox "All data refreshed.", vbInformation, "Station Leave Manager"
End Sub

' Quick add personnel helper (can be called from Immediate window or button)
Public Sub QuickAddPerson(pName As String, pRank As String, pShift As Long, _
                           joinDate As Date, Optional ordDate As Date = 0, _
                           Optional vlEntitlement As Double = -1)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    ' Find next empty row
    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        If Trim(CStr(wsP.Cells(r, P_COL_NAME).Value)) = "" Then
            Exit For
        End If
    Next r
    If r > P_DATA_START + 49 Then
        MsgBox "Personnel table is full (max 50).", vbExclamation
        Exit Sub
    End If

    ' Calculate VL if not provided
    Dim vlDays As Double
    If vlEntitlement < 0 Then
        vlDays = CalcProRatedVL(joinDate, ordDate)
    Else
        vlDays = vlEntitlement
    End If

    wsP.Cells(r, P_COL_NAME).Value = pName
    wsP.Cells(r, P_COL_RANK).Value = pRank
    wsP.Cells(r, P_COL_SHIFT).Value = pShift
    wsP.Cells(r, P_COL_JOIN).Value = joinDate
    wsP.Cells(r, P_COL_JOIN).NumberFormat = "DD-MMM-YYYY"
    If ordDate > 0 Then
        wsP.Cells(r, P_COL_ORD).Value = ordDate
        wsP.Cells(r, P_COL_ORD).NumberFormat = "DD-MMM-YYYY"
    End If
    wsP.Cells(r, P_COL_STATUS).Value = "Active"
    wsP.Cells(r, P_COL_VL_ENT).Value = vlDays

    ' Build display name and credit VL
    Dim dn As String
    dn = pRank & " " & pName
    CreditLeave dn, LEAVE_VL, vlDays, joinDate, DateSerial(GetYear(), 12, 31), "Annual VL credit"

    ' Refresh
    RefreshNameLists
    RefreshPersonBalance dn

    MsgBox "Added: " & dn & " (Shift " & pShift & ", " & vlDays & " VL)", vbInformation
End Sub
