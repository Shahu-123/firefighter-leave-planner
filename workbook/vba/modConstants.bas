Attribute VB_Name = "modConstants"
Option Explicit

' ============================================================
' modConstants - Global constants and helpers
' Station Leave Manager (Event-Driven Model)
' ============================================================

' Sheet names
Public Const WS_CALENDAR As String = "Calendar"
Public Const WS_PERSONNEL As String = "Personnel"
Public Const WS_DASHBOARD As String = "Dashboard"
Public Const WS_CONFIG As String = "Config"
Public Const WS_NAMELISTS As String = "NameLists"

' Leave types
Public Const LEAVE_VL As String = "VL"
Public Const LEAVE_PHOL As String = "PHOL"
Public Const LEAVE_OIL As String = "OIL"

' Transaction types
Public Const TXN_CREDIT As String = "CREDIT"
Public Const TXN_DEBIT As String = "DEBIT"

' Calendar slot columns (D=4 through I=9)
Public Const CAL_SLOT_FIRST_COL As Long = 4   ' Column D
Public Const CAL_SLOT_LAST_COL As Long = 9    ' Column I
Public Const CAL_COL_DATE As Long = 1
Public Const CAL_COL_DAY As Long = 2
Public Const CAL_COL_SHIFT As Long = 3
Public Const CAL_COL_ONLEAVE As Long = 10
Public Const CAL_COL_AVAIL As Long = 11

' Personnel table layout (starts row 4)
Public Const P_DATA_START As Long = 4
Public Const P_COL_NAME As Long = 1
Public Const P_COL_RANK As Long = 2
Public Const P_COL_SHIFT As Long = 3
Public Const P_COL_JOIN As Long = 4
Public Const P_COL_ORD As Long = 5
Public Const P_COL_STATUS As Long = 6
Public Const P_COL_VL_ENT As Long = 7
Public Const P_COL_VL_USED As Long = 8
Public Const P_COL_VL_BAL As Long = 9
Public Const P_COL_PHOL_CR As Long = 10
Public Const P_COL_PHOL_USED As Long = 11
Public Const P_COL_PHOL_BAL As Long = 12
Public Const P_COL_OIL_CR As Long = 13
Public Const P_COL_OIL_USED As Long = 14
Public Const P_COL_OIL_BAL As Long = 15
Public Const P_COL_TOTAL_BAL As Long = 16

' Leave Credits table layout (on Personnel sheet)
' Starts at row 57 (P_DATA_START + 50 + 3)
Public Const LC_DATA_START As Long = 59
Public Const LC_COL_ID As Long = 1
Public Const LC_COL_NAME As Long = 2
Public Const LC_COL_TXN As Long = 3
Public Const LC_COL_LEAVE As Long = 4
Public Const LC_COL_DAYS As Long = 5
Public Const LC_COL_DATE As Long = 6
Public Const LC_COL_EXPIRY As Long = 7
Public Const LC_COL_SOURCE As Long = 8
Public Const LC_COL_CALREF As Long = 9
Public Const LC_COL_REMARKS As Long = 10

' ============================================================
' Config Getters
' ============================================================

Public Function GetConfig(key As String) As Variant
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_CONFIG)
    Dim r As Long
    For r = 3 To 20
        If CStr(ws.Cells(r, 1).Value) = key Then
            GetConfig = ws.Cells(r, 2).Value
            Exit Function
        End If
    Next r
    GetConfig = ""
End Function

Public Function GetYear() As Long
    GetYear = CLng(GetConfig("Year"))
End Function

Public Function GetShiftCost() As Double
    GetShiftCost = CDbl(GetConfig("ShiftLeaveCostDays"))
End Function

' ============================================================
' Helpers
' ============================================================

Public Function GetLastRow(ws As Worksheet, col As Long) As Long
    GetLastRow = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

' Find a person's row in the Personnel table by display name
Public Function FindPersonRow(displayName As String) As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim cellName As String
        cellName = GetDisplayName(r)
        If cellName = displayName And cellName <> "" Then
            FindPersonRow = r
            Exit Function
        End If
    Next r
    FindPersonRow = 0
End Function

' Build display name "Rank Name" for a Personnel row
Public Function GetDisplayName(personnelRow As Long) As String
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim n As String, rk As String
    n = Trim(CStr(ws.Cells(personnelRow, P_COL_NAME).Value))
    rk = Trim(CStr(ws.Cells(personnelRow, P_COL_RANK).Value))
    If n = "" Then
        GetDisplayName = ""
    ElseIf rk = "" Then
        GetDisplayName = n
    Else
        GetDisplayName = rk & " " & n
    End If
End Function

' Get the next available ID for the Leave Credits table
Public Function GetNextCreditID() As Long
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim maxID As Long
    maxID = 0
    Dim r As Long
    For r = LC_DATA_START To LC_DATA_START + 199
        If ws.Cells(r, LC_COL_ID).Value <> "" Then
            Dim v As Long
            On Error Resume Next
            v = CLng(ws.Cells(r, LC_COL_ID).Value)
            On Error GoTo 0
            If v > maxID Then maxID = v
        End If
    Next r
    GetNextCreditID = maxID + 1
End Function

' Check if a Calendar row is a data row (has a date in column A and shift in C)
Public Function IsCalendarDataRow(row As Long) As Boolean
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(WS_CALENDAR)
    If row < 3 Then
        IsCalendarDataRow = False
        Exit Function
    End If
    ' Check if column C has a numeric shift value (1, 2, or 3)
    Dim v As Variant
    v = ws.Cells(row, CAL_COL_SHIFT).Value
    If IsNumeric(v) Then
        If CLng(v) >= 1 And CLng(v) <= 3 Then
            IsCalendarDataRow = True
            Exit Function
        End If
    End If
    IsCalendarDataRow = False
End Function

' Slot rank (1-6) from column index
Public Function ColToSlotRank(col As Long) As Long
    ColToSlotRank = col - CAL_SLOT_FIRST_COL + 1
End Function

' Slot type from rank
Public Function SlotRankToType(rank As Long) As String
    Select Case rank
        Case 1, 2: SlotRankToType = "L1"
        Case 3, 4: SlotRankToType = "L2"
        Case 5, 6: SlotRankToType = "KIV"
    End Select
End Function
