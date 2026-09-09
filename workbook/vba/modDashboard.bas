Attribute VB_Name = "modDashboard"
Option Explicit

' ============================================================
' modDashboard - Populates the Dashboard analytics
' Station Leave Manager
' ============================================================

Public Sub RefreshDashboard()
    Application.StatusBar = "Refreshing dashboard..."
    Application.ScreenUpdating = False

    RefreshAllBalances

    RefreshKeyMetrics
    RefreshORDTracker
    RefreshMonthlyOverview
    RefreshExpiryWatch
    RefreshUtilization

    Application.ScreenUpdating = True
    Application.StatusBar = False
End Sub

' ============================================================
' Key Metrics (row 6)
' ============================================================

Private Sub RefreshKeyMetrics()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(WS_DASHBOARD)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    Dim activeCount As Long, totalBal As Double, ordIn30 As Long
    Dim expiringCount As Long
    activeCount = 0: totalBal = 0: ordIn30 = 0

    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextM
        If CStr(wsP.Cells(r, P_COL_STATUS).Value) <> "Active" Then GoTo NextM

        activeCount = activeCount + 1
        totalBal = totalBal + CDbl(wsP.Cells(r, P_COL_TOTAL_BAL).Value)

        If Not IsEmpty(wsP.Cells(r, P_COL_ORD).Value) And wsP.Cells(r, P_COL_ORD).Value <> "" Then
            If CDate(wsP.Cells(r, P_COL_ORD).Value) <= Date + 30 And _
               CDate(wsP.Cells(r, P_COL_ORD).Value) > Date Then
                ordIn30 = ordIn30 + 1
            End If
        End If
NextM:
    Next r

    ' Count expiring within 30 days
    For r = P_DATA_START To P_DATA_START + 49
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextM2
        If CStr(wsP.Cells(r, P_COL_STATUS).Value) <> "Active" Then GoTo NextM2
        Dim pools As Collection
        Set pools = GetSortedCreditPools(dn)
        Dim p As Variant
        For Each p In pools
            If CLng(CDate(p(3)) - Date) <= 30 And CDbl(p(2)) > 0 Then
                expiringCount = expiringCount + CDbl(p(2))
            End If
        Next p
NextM2:
    Next r

    ' Count slots filled this month on Calendar
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Sheets(WS_CALENDAR)
    Dim totalSlots As Long, filledSlots As Long
    totalSlots = 0: filledSlots = 0
    Dim lastCalRow As Long
    lastCalRow = GetLastRow(wsCal, CAL_COL_DATE)
    For r = 3 To lastCalRow
        If IsCalendarDataRow(r) Then
            If Not IsEmpty(wsCal.Cells(r, CAL_COL_DATE).Value) Then
                Dim calDate As Date
                On Error Resume Next
                calDate = CDate(wsCal.Cells(r, CAL_COL_DATE).Value)
                On Error GoTo 0
                If Month(calDate) = Month(Date) And Year(calDate) = Year(Date) Then
                    totalSlots = totalSlots + 6
                    Dim c As Long
                    For c = CAL_SLOT_FIRST_COL To CAL_SLOT_LAST_COL
                        If Trim(CStr(wsCal.Cells(r, c).Value)) <> "" Then filledSlots = filledSlots + 1
                    Next c
                End If
            End If
        End If
    Next r

    ' Write metrics (row 6)
    wsDash.Cells(6, 1).Value = activeCount
    If activeCount > 0 Then
        wsDash.Cells(6, 3).Value = Format(totalBal / activeCount, "0.0")
    Else
        wsDash.Cells(6, 3).Value = "-"
    End If
    wsDash.Cells(6, 5).Value = Format(expiringCount, "0.#") & "d"
    If totalSlots > 0 Then
        wsDash.Cells(6, 7).Value = Format(CDbl(filledSlots) / CDbl(totalSlots), "0%")
    Else
        wsDash.Cells(6, 7).Value = "-"
    End If
    wsDash.Cells(6, 9).Value = ordIn30
End Sub

' ============================================================
' ORD Leave Clearing Tracker
' ============================================================

Private Sub RefreshORDTracker()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(WS_DASHBOARD)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Sheets(WS_CALENDAR)

    ' Find start row (row 10 based on layout)
    Dim startRow As Long
    startRow = 10

    ' Clear
    Dim clr As Long
    For clr = startRow To startRow + 29
        Dim cc As Long
        For cc = 1 To 10
            wsDash.Cells(clr, cc).Value = ""
            wsDash.Cells(clr, cc).Interior.Pattern = xlNone
        Next cc
    Next clr

    Dim outRow As Long
    outRow = startRow

    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextO
        If CStr(wsP.Cells(r, P_COL_STATUS).Value) <> "Active" Then GoTo NextO

        ' Only show people with ORD dates
        If IsEmpty(wsP.Cells(r, P_COL_ORD).Value) Or wsP.Cells(r, P_COL_ORD).Value = "" Then GoTo NextO

        Dim ordDate As Date
        ordDate = CDate(wsP.Cells(r, P_COL_ORD).Value)
        If ordDate <= Date Then GoTo NextO ' already ORD'd

        Dim shift As Long
        shift = CLng(wsP.Cells(r, P_COL_SHIFT).Value)
        Dim bal As Double
        bal = CDbl(wsP.Cells(r, P_COL_TOTAL_BAL).Value)
        Dim daysToORD As Long
        daysToORD = CLng(ordDate - Date)

        ' Shifts to clear = ceil(balance / 2)
        Dim shiftsToClear As Long
        If bal > 0 Then
            shiftsToClear = Application.WorksheetFunction.RoundUp(bal / 2, 0)
        Else
            shiftsToClear = 0
        End If

        ' Count shifts available (person's shift days from now to ORD)
        Dim shiftsAvail As Long
        shiftsAvail = CountShiftDaysForShift(shift, Date + 1, ordDate)

        ' Count open slots available (shift days where <6 slots filled)
        Dim openSlots As Long
        openSlots = CountOpenSlotDays(shift, Date + 1, ordDate)

        ' Difficulty ratio
        Dim difficulty As String
        Dim diffRatio As Double
        If openSlots > 0 Then
            diffRatio = CDbl(shiftsToClear) / CDbl(openSlots)
            difficulty = Format(diffRatio, "0%")
        Else
            diffRatio = 1
            difficulty = "N/A"
        End If

        ' Status
        Dim status As String
        If bal <= 0 Then
            status = "Cleared"
        ElseIf diffRatio >= 0.8 Then
            status = "Critical"
        ElseIf diffRatio >= 0.6 Then
            status = "Hard"
        ElseIf diffRatio >= 0.3 Then
            status = "Moderate"
        Else
            status = "Easy"
        End If

        ' Write
        wsDash.Cells(outRow, 1).Value = dn
        wsDash.Cells(outRow, 2).Value = shift
        wsDash.Cells(outRow, 3).Value = ordDate
        wsDash.Cells(outRow, 3).NumberFormat = "DD-MMM-YY"
        wsDash.Cells(outRow, 4).Value = daysToORD
        wsDash.Cells(outRow, 5).Value = bal
        wsDash.Cells(outRow, 6).Value = shiftsToClear
        wsDash.Cells(outRow, 7).Value = shiftsAvail
        wsDash.Cells(outRow, 8).Value = openSlots
        wsDash.Cells(outRow, 9).Value = difficulty
        wsDash.Cells(outRow, 10).Value = status

        ' Color status
        Select Case status
            Case "Cleared", "Easy"
                wsDash.Cells(outRow, 10).Interior.Color = RGB(198, 239, 206)
            Case "Moderate"
                wsDash.Cells(outRow, 10).Interior.Color = RGB(255, 230, 153)
            Case "Hard"
                wsDash.Cells(outRow, 10).Interior.Color = RGB(255, 200, 150)
            Case "Critical"
                wsDash.Cells(outRow, 10).Interior.Color = RGB(244, 204, 204)
        End Select

        outRow = outRow + 1
        If outRow > startRow + 29 Then Exit For
NextO:
    Next r
End Sub

Private Function CountShiftDaysForShift(shift As Long, startDate As Date, endDate As Date) As Long
    ' Uses the anchor-based shift calculation
    Dim anchorDate As Date
    anchorDate = CDate(GetConfig("CycleAnchorDate"))
    Dim anchorShift As Long
    anchorShift = CLng(GetConfig("AnchorShift"))

    Dim count As Long
    count = 0
    Dim d As Date
    For d = startDate To endDate
        Dim daysDiff As Long
        daysDiff = CLng(d - anchorDate)
        Dim cyclePos As Long
        cyclePos = daysDiff Mod 3
        If cyclePos < 0 Then cyclePos = cyclePos + 3
        Dim resultShift As Long
        resultShift = ((anchorShift - 1 + cyclePos) Mod 3) + 1
        If resultShift = shift Then count = count + 1
    Next d
    CountShiftDaysForShift = count
End Function

Private Function CountOpenSlotDays(shift As Long, startDate As Date, endDate As Date) As Long
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Sheets(WS_CALENDAR)
    Dim lastRow As Long
    lastRow = GetLastRow(wsCal, CAL_COL_DATE)

    Dim count As Long
    count = 0
    Dim r As Long
    For r = 3 To lastRow
        If Not IsCalendarDataRow(r) Then GoTo NextR
        Dim calDate As Date
        On Error Resume Next
        calDate = CDate(wsCal.Cells(r, CAL_COL_DATE).Value)
        On Error GoTo 0
        If calDate < startDate Or calDate > endDate Then GoTo NextR
        If CLng(wsCal.Cells(r, CAL_COL_SHIFT).Value) <> shift Then GoTo NextR

        ' Count filled slots
        Dim filled As Long
        filled = 0
        Dim c As Long
        For c = CAL_SLOT_FIRST_COL To CAL_SLOT_LAST_COL
            If Trim(CStr(wsCal.Cells(r, c).Value)) <> "" Then filled = filled + 1
        Next c
        If filled < 6 Then count = count + 1
NextR:
    Next r
    CountOpenSlotDays = count
End Function

' ============================================================
' Monthly Rota Overview
' ============================================================

Private Sub RefreshMonthlyOverview()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(WS_DASHBOARD)
    Dim wsCal As Worksheet
    Set wsCal = ThisWorkbook.Sheets(WS_CALENDAR)

    Dim startRow As Long
    startRow = 10  ' Same as ORD tracker but in col L
    Dim startCol As Long
    startCol = 12

    ' Clear
    Dim clr As Long
    For clr = startRow To startRow + 8
        Dim cc As Long
        For cc = startCol To startCol + 6
            wsDash.Cells(clr, cc).Value = ""
        Next cc
    Next clr

    Dim monthNames As Variant
    monthNames = Array("", "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec")

    Dim lastRow As Long
    lastRow = GetLastRow(wsCal, CAL_COL_DATE)

    Dim startMonth As Long
    startMonth = CLng(GetConfig("Year"))
    ' Actually get start month from config - it's the first month on the calendar
    startMonth = Month(CDate(wsCal.Cells(5, CAL_COL_DATE).Value))  ' first data row date

    Dim outRow As Long
    outRow = startRow
    Dim m As Long
    For m = startMonth To 12
        Dim s1 As Long, s2 As Long, s3 As Long, totalAvail As Long, dayCount As Long
        s1 = 0: s2 = 0: s3 = 0: totalAvail = 0: dayCount = 0

        Dim r As Long
        For r = 3 To lastRow
            If Not IsCalendarDataRow(r) Then GoTo NextMR
            Dim calDate As Date
            On Error Resume Next
            calDate = CDate(wsCal.Cells(r, CAL_COL_DATE).Value)
            On Error GoTo 0
            If Month(calDate) <> m Then GoTo NextMR

            Dim shift As Long
            shift = CLng(wsCal.Cells(r, CAL_COL_SHIFT).Value)
            Dim filled As Long
            filled = 0
            Dim c As Long
            For c = CAL_SLOT_FIRST_COL To CAL_SLOT_LAST_COL
                If Trim(CStr(wsCal.Cells(r, c).Value)) <> "" Then filled = filled + 1
            Next c

            Select Case shift
                Case 1: s1 = s1 + filled
                Case 2: s2 = s2 + filled
                Case 3: s3 = s3 + filled
            End Select
            dayCount = dayCount + 1

            ' Avail from formula
            If IsNumeric(wsCal.Cells(r, CAL_COL_AVAIL).Value) Then
                totalAvail = totalAvail + CLng(wsCal.Cells(r, CAL_COL_AVAIL).Value)
            End If
NextMR:
        Next r

        wsDash.Cells(outRow, startCol).Value = monthNames(m)
        wsDash.Cells(outRow, startCol + 1).Value = s1
        wsDash.Cells(outRow, startCol + 2).Value = s2
        wsDash.Cells(outRow, startCol + 3).Value = s3
        wsDash.Cells(outRow, startCol + 4).Value = s1 + s2 + s3
        If dayCount > 0 Then
            wsDash.Cells(outRow, startCol + 5).Value = Format(CDbl(totalAvail) / CDbl(dayCount), "0.0")
        End If
        Dim totalSlotDays As Long
        totalSlotDays = dayCount * 6
        If totalSlotDays > 0 Then
            wsDash.Cells(outRow, startCol + 6).Value = Format(CDbl(s1 + s2 + s3) / CDbl(totalSlotDays), "0%")
        End If

        outRow = outRow + 1
    Next m
End Sub

' ============================================================
' Leave Expiry Watch
' ============================================================

Private Sub RefreshExpiryWatch()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(WS_DASHBOARD)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    Dim startRow As Long
    startRow = 23  ' Below monthly overview, col L
    Dim startCol As Long
    startCol = 12

    ' Clear
    Dim clr As Long
    For clr = startRow To startRow + 19
        Dim cc As Long
        For cc = startCol To startCol + 5
            wsDash.Cells(clr, cc).Value = ""
            wsDash.Cells(clr, cc).Interior.Pattern = xlNone
        Next cc
    Next clr

    ' Collect warnings
    Dim warnings() As Variant
    ReDim warnings(1 To 100, 1 To 6)
    Dim warnCount As Long
    warnCount = 0

    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextE
        If CStr(wsP.Cells(r, P_COL_STATUS).Value) <> "Active" Then GoTo NextE

        Dim pools As Collection
        Set pools = GetSortedCreditPools(dn)
        Dim p As Variant
        For Each p In pools
            Dim daysLeft As Long
            daysLeft = CLng(CDate(p(3)) - Date)
            If daysLeft <= 60 And CDbl(p(2)) > 0 Then
                warnCount = warnCount + 1
                If warnCount > 100 Then GoTo DoneWarn
                warnings(warnCount, 1) = dn
                warnings(warnCount, 2) = p(1)
                warnings(warnCount, 3) = p(2)
                warnings(warnCount, 4) = CDate(p(3))
                warnings(warnCount, 5) = daysLeft
                If daysLeft <= 14 Then
                    warnings(warnCount, 6) = "URGENT"
                ElseIf daysLeft <= 30 Then
                    warnings(warnCount, 6) = "Warning"
                Else
                    warnings(warnCount, 6) = "Info"
                End If
            End If
        Next p
NextE:
    Next r
DoneWarn:

    ' Sort by days left
    Dim i As Long, j As Long
    For i = 1 To warnCount - 1
        For j = 1 To warnCount - i
            If CLng(warnings(j, 5)) > CLng(warnings(j + 1, 5)) Then
                Dim k As Long
                For k = 1 To 6
                    Dim tmp As Variant
                    tmp = warnings(j, k)
                    warnings(j, k) = warnings(j + 1, k)
                    warnings(j + 1, k) = tmp
                Next k
            End If
        Next j
    Next i

    ' Write top 20
    Dim outRow As Long
    outRow = startRow
    Dim maxW As Long
    If warnCount < 20 Then maxW = warnCount Else maxW = 20
    For i = 1 To maxW
        wsDash.Cells(outRow, startCol).Value = warnings(i, 1)
        wsDash.Cells(outRow, startCol + 1).Value = warnings(i, 2)
        wsDash.Cells(outRow, startCol + 2).Value = warnings(i, 3)
        wsDash.Cells(outRow, startCol + 3).Value = warnings(i, 4)
        wsDash.Cells(outRow, startCol + 3).NumberFormat = "DD-MMM"
        wsDash.Cells(outRow, startCol + 4).Value = warnings(i, 5)
        wsDash.Cells(outRow, startCol + 5).Value = warnings(i, 6)
        Select Case CStr(warnings(i, 6))
            Case "URGENT": wsDash.Cells(outRow, startCol + 5).Interior.Color = RGB(244, 204, 204)
            Case "Warning": wsDash.Cells(outRow, startCol + 5).Interior.Color = RGB(255, 230, 153)
            Case "Info": wsDash.Cells(outRow, startCol + 5).Interior.Color = RGB(198, 239, 206)
        End Select
        outRow = outRow + 1
    Next i
End Sub

' ============================================================
' Utilization Statistics
' ============================================================

Private Sub RefreshUtilization()
    Dim wsDash As Worksheet
    Set wsDash = ThisWorkbook.Sheets(WS_DASHBOARD)
    Dim wsP As Worksheet
    Set wsP = ThisWorkbook.Sheets(WS_PERSONNEL)

    ' Find util start row (after ORD tracker)
    Dim startRow As Long
    startRow = 44  ' approximate, based on layout

    Dim vlEnt(1 To 3) As Double, vlUsed(1 To 3) As Double
    Dim pholCr(1 To 3) As Double, pholUsed(1 To 3) As Double
    Dim oilCr(1 To 3) As Double, oilUsed(1 To 3) As Double
    Dim totalBal(1 To 3) As Double, pCount(1 To 3) As Long
    Dim ordIn30(1 To 3) As Long

    Dim r As Long
    For r = P_DATA_START To P_DATA_START + 49
        Dim dn As String
        dn = GetDisplayName(r)
        If dn = "" Then GoTo NextU
        If CStr(wsP.Cells(r, P_COL_STATUS).Value) <> "Active" Then GoTo NextU

        Dim s As Long
        s = CLng(wsP.Cells(r, P_COL_SHIFT).Value)
        If s < 1 Or s > 3 Then GoTo NextU

        pCount(s) = pCount(s) + 1
        vlEnt(s) = vlEnt(s) + CDbl(wsP.Cells(r, P_COL_VL_ENT).Value)
        vlUsed(s) = vlUsed(s) + CDbl(wsP.Cells(r, P_COL_VL_USED).Value)
        pholCr(s) = pholCr(s) + CDbl(wsP.Cells(r, P_COL_PHOL_CR).Value)
        pholUsed(s) = pholUsed(s) + CDbl(wsP.Cells(r, P_COL_PHOL_USED).Value)
        oilCr(s) = oilCr(s) + CDbl(wsP.Cells(r, P_COL_OIL_CR).Value)
        oilUsed(s) = oilUsed(s) + CDbl(wsP.Cells(r, P_COL_OIL_USED).Value)
        totalBal(s) = totalBal(s) + CDbl(wsP.Cells(r, P_COL_TOTAL_BAL).Value)

        If Not IsEmpty(wsP.Cells(r, P_COL_ORD).Value) And wsP.Cells(r, P_COL_ORD).Value <> "" Then
            If CDate(wsP.Cells(r, P_COL_ORD).Value) <= Date + 30 And CDate(wsP.Cells(r, P_COL_ORD).Value) > Date Then
                ordIn30(s) = ordIn30(s) + 1
            End If
        End If
NextU:
    Next r

    ' Write (col B-E = Shift 1, 2, 3, Overall)
    Dim sh As Long
    For sh = 1 To 3
        If vlEnt(sh) > 0 Then wsDash.Cells(startRow, 1 + sh).Value = Format(vlUsed(sh) / vlEnt(sh), "0%")
        If pholCr(sh) > 0 Then wsDash.Cells(startRow + 1, 1 + sh).Value = Format(pholUsed(sh) / pholCr(sh), "0%")
        If oilCr(sh) > 0 Then wsDash.Cells(startRow + 2, 1 + sh).Value = Format(oilUsed(sh) / oilCr(sh), "0%")
        If pCount(sh) > 0 Then wsDash.Cells(startRow + 3, 1 + sh).Value = Format(totalBal(sh) / pCount(sh), "0.0")
        wsDash.Cells(startRow + 4, 1 + sh).Value = pCount(sh)
        wsDash.Cells(startRow + 5, 1 + sh).Value = ordIn30(sh)
    Next sh

    ' Overall
    Dim tP As Long: tP = pCount(1) + pCount(2) + pCount(3)
    Dim tVE As Double: tVE = vlEnt(1) + vlEnt(2) + vlEnt(3)
    Dim tVU As Double: tVU = vlUsed(1) + vlUsed(2) + vlUsed(3)
    If tVE > 0 Then wsDash.Cells(startRow, 5).Value = Format(tVU / tVE, "0%")
    Dim tPC As Double: tPC = pholCr(1) + pholCr(2) + pholCr(3)
    Dim tPU As Double: tPU = pholUsed(1) + pholUsed(2) + pholUsed(3)
    If tPC > 0 Then wsDash.Cells(startRow + 1, 5).Value = Format(tPU / tPC, "0%")
    Dim tOC As Double: tOC = oilCr(1) + oilCr(2) + oilCr(3)
    Dim tOU As Double: tOU = oilUsed(1) + oilUsed(2) + oilUsed(3)
    If tOC > 0 Then wsDash.Cells(startRow + 2, 5).Value = Format(tOU / tOC, "0%")
    If tP > 0 Then
        wsDash.Cells(startRow + 3, 5).Value = Format((totalBal(1) + totalBal(2) + totalBal(3)) / tP, "0.0")
    End If
    wsDash.Cells(startRow + 4, 5).Value = tP
    wsDash.Cells(startRow + 5, 5).Value = ordIn30(1) + ordIn30(2) + ordIn30(3)
End Sub
