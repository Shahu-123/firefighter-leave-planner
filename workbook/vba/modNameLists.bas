Attribute VB_Name = "modNameLists"
Option Explicit

' ============================================================
' modNameLists - Maintains shift-filtered name lists for dropdowns
' Station Leave Manager
'
' Called on workbook open and after personnel changes.
' ============================================================

Public Sub RefreshNameLists()
    Dim wsNL As Worksheet
    Set wsNL = ThisWorkbook.Sheets(WS_NAMELISTS)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    ' Clear existing names (keep headers)
    Dim r As Long
    For r = 2 To 51
        wsNL.Cells(r, 1).Value = ""
        wsNL.Cells(r, 2).Value = ""
        wsNL.Cells(r, 3).Value = ""
    Next r

    ' Counters per shift
    Dim cnt(1 To 3) As Long
    cnt(1) = 0: cnt(2) = 0: cnt(3) = 0

    ' Walk personnel table
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextP

        Dim status As String
        status = Trim(CStr(wsP.Cells(r, P_COL_STATUS).Value))
        If status <> "Active" Then GoTo NextP

        Dim shift As Variant
        shift = wsP.Cells(r, P_COL_SHIFT).Value
        If Not IsNumeric(shift) Then GoTo NextP
        Dim s As Long
        s = CLng(shift)
        If s < 1 Or s > 3 Then GoTo NextP

        cnt(s) = cnt(s) + 1
        wsNL.Cells(cnt(s) + 1, s).Value = dn  ' row 2 onwards
NextP:
    Next r

    ' Update named ranges to match actual data extent
    UpdateNameRange "Shift1_Names", "A", cnt(1)
    UpdateNameRange "Shift2_Names", "B", cnt(2)
    UpdateNameRange "Shift3_Names", "C", cnt(3)
End Sub

Private Sub UpdateNameRange(rangeName As String, colLetter As String, count As Long)
    If count = 0 Then count = 1  ' at least 1 row for valid named range

    ' Delete existing named range if it exists
    On Error Resume Next
    ThisWorkbook.Names(rangeName).Delete
    On Error GoTo 0

    ' Add new named range
    ThisWorkbook.Names.Add Name:=rangeName, _
        RefersTo:="=NameLists!$" & colLetter & "$2:$" & colLetter & "$" & (count + 1)
End Sub

' Auto-credit VL when a new person is added to Personnel
Public Sub CheckAndCreditVL()
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextP2

        Dim status As String
        status = Trim(CStr(wsP.Cells(r, P_COL_STATUS).Value))
        If status <> "Active" Then GoTo NextP2

        ' Check if VL already credited for this person
        Dim vlEnt As Double
        vlEnt = 0
        If Not IsEmpty(wsP.Cells(r, P_COL_VL_ENT).Value) And wsP.Cells(r, P_COL_VL_ENT).Value <> "" Then
            vlEnt = CDbl(wsP.Cells(r, P_COL_VL_ENT).Value)
        End If

        ' Check if VL CREDIT exists in ledger
        Dim hasVL As Boolean
        hasVL = False
        Dim cr As Long
        For cr = LC_DATA_START To LC_DATA_START + 199
            If CStr(wsP.Cells(cr, LC_COL_NAME).Value) = dn And _
               CStr(wsP.Cells(cr, LC_COL_TXN).Value) = TXN_CREDIT And _
               CStr(wsP.Cells(cr, LC_COL_LEAVE).Value) = LEAVE_VL Then
                hasVL = True
                Exit For
            End If
        Next cr

        If Not hasVL And vlEnt > 0 Then
            ' Auto-credit VL
            Dim joinDate As Date
            joinDate = DateSerial(GetYear(), 1, 1)
            If Not IsEmpty(wsP.Cells(r, P_COL_JOIN).Value) Then
                joinDate = CDate(wsP.Cells(r, P_COL_JOIN).Value)
            End If
            Dim expiryDate As Date
            expiryDate = DateSerial(GetYear(), 12, 31)
            CreditLeave dn, LEAVE_VL, vlEnt, joinDate, expiryDate, "Annual VL credit"
        End If
NextP2:
    Next r
End Sub
