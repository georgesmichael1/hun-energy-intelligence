Attribute VB_Name = "Hungary_Module"
Option Explicit
' ============================================================================
' Hungary Energy Intelligence - data refresh module
' Works on Hungary_Energy_Dashboard.xlsx built by dashboard/build_dashboard.py.
' It downloads this repo's CSV outputs, rewrites the D_* data sheets, resizes
' the dashboard charts to the new row counts and recalculates. No token, no
' password and no PAT is used: the repo is public.
' Missing values stay blank. A failed download keeps the sheet's saved data and
' is listed in the warning; nothing is ever replaced with zero.
' NOT verified in desktop Excel by the author (built without Excel). See README.
' ============================================================================
Private Const FEED_ROOT As String = "https://raw.githubusercontent.com/georgesmichael1/hun-energy-intelligence/main/"
Private Const SYNC_MINUTES As Long = 30
Private gNextSync As Date
Public HungaryLastWarning As String

' column spec tokens: name = copy value; @LOCAL = Budapest local start from Interval_Start_UTC;
' ~name = copy as text; #text = constant text
Private Function Specs() As Variant
    Specs = Array( _
      Array("hupx_day_ahead_prices.csv", "D_Price", "~Delivery_Date,@LOCAL,Price_EUR_MWh,~Attribution"), _
      Array("entsoe_day_ahead_prices.csv", "D_EntsoePrice", "~Delivery_Date,@LOCAL,Price_EUR_MWh,#ENTSO-E A44 (primary owner HUPX)"), _
      Array("mavir_load.csv", "D_Load", "@LOCAL,Gross_Actual_System_Load_MW,Gross_Plan_System_Load_MW,Net_Actual_Load_Settlement_MW"), _
      Array("mavir_generation_15min.csv", "D_Gen", "@LOCAL,Net_Domestic_Generation_Actual_MW,Nuclear_MW,Lignite_MW,Gas_MW,Wind_Onshore_MW,Biomass_MW,Solar_PV_Measured_MW,Hydro_RunOfRiver_MW,Waste_Incineration_MW"), _
      Array("mavir_flows_interconnectors.csv", "D_Flows", "@LOCAL,HU_AT_Actual_MW,HU_SK_Actual_MW,HU_RO_Actual_MW,HU_RS_Actual_MW,HU_HR_Actual_MW,HU_SI_Actual_MW,HU_UA_Actual_MW"), _
      Array("mavir_balancing_offers.csv", "D_Balancing", "@LOCAL,AFRR_Up_Offers_MW,AFRR_Down_Offers_MW,MFRR_Up_Offers_MW,MFRR_Down_Offers_MW"), _
      Array("entsoe_load_forecast.csv", "D_LoadForecast", "@LOCAL,Value_MW"), _
      Array("eurostat_tariffs.csv", "D_Tariffs", "~Period,~Consumer_Group,~Consumption_Band,Price_EUR_Per_kWh"), _
      Array("eurostat_hicp.csv", "D_HICP", "~Period,HICP_Annual_Rate_Pct"), _
      Array("mnb_fx.csv", "D_FX", "~Rate_Date,~Currency,HUF_Per_Unit"), _
      Array("mekh_aggregators.csv", "D_Aggregators", "~Aggregator,~Registration_Decision_No,~City"), _
      Array("hun_substation_capacity.csv", "D_Substations", "~Table_Label,~Substation,Voltage_kV,Releasable_Offtake_MVA,Probable_5y_Offtake_MVA"), _
      Array("mekh_storage_licences.csv", "D_Storage", "~Licensee,~Site,Nominal_Power_MW,Storage_Capacity_MWh,~Revoked_Date"), _
      Array("mavir_installed_capacity.csv", "D_Installed", "~Fuel_Group,~Period,Gross_Installed_MW"), _
      Array("balance_responsible_parties.csv", "D_BRP", "~Balance_Responsible_Party,~EIC_Code,~Valid_From"))
End Function

Public Sub RefreshHungary(Optional ByVal Silent As Boolean = False)
    Dim sp As Variant, i As Long, okN As Long, msg As String
    HungaryLastWarning = ""
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    sp = Specs()
    For i = LBound(sp) To UBound(sp)
        msg = ImportOne(CStr(sp(i)(0)), CStr(sp(i)(1)), CStr(sp(i)(2)))
        If Len(msg) = 0 Then
            okN = okN + 1
        Else
            HungaryLastWarning = HungaryLastWarning & sp(i)(0) & ": " & msg & vbLf
        End If
        If i < UBound(sp) Then Application.Wait Now + TimeSerial(0, 0, 1)
    Next i
    msg = ImportStatus()
    If Len(msg) > 0 Then HungaryLastWarning = HungaryLastWarning & "status.json: " & msg & vbLf
    msg = RefreshRegisterTables()
    If Len(msg) > 0 Then HungaryLastWarning = HungaryLastWarning & "register tables: " & msg & vbLf
    ResizeCharts
    Application.Calculation = xlCalculationAutomatic
    Application.Calculate
    Application.ScreenUpdating = True
    If Not Silent Then
        If Len(HungaryLastWarning) = 0 Then
            MsgBox "Hungary data refreshed (" & okN & " feeds).", vbInformation
        Else
            MsgBox "Refreshed " & okN & " feeds. Kept saved data for:" & vbLf & HungaryLastWarning, vbExclamation
        End If
    End If
End Sub

Public Sub ScheduleHungarySync()
    gNextSync = Now + TimeSerial(0, SYNC_MINUTES, 0)
    Application.OnTime gNextSync, "HungaryAutoSync"
End Sub

Public Sub CancelHungarySync()
    On Error Resume Next
    Application.OnTime gNextSync, "HungaryAutoSync", , False
End Sub

Public Sub HungaryAutoSync()
    RefreshHungary True
    ScheduleHungarySync
End Sub

Private Function Download(ByVal url As String, ByRef body As String) As String
    Dim http As Object
    On Error GoTo Fail
    Set http = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    http.setTimeouts 10000, 10000, 30000, 60000
    http.Open "GET", url & "?nocache=" & Format(Now, "yyyymmddhhnnss"), False
    http.send
    If http.Status <> 200 Then Download = "HTTP " & http.Status: Exit Function
    body = Utf8ToString(http.responseBody)
    Exit Function
Fail:
    Download = "download failed"
End Function

Private Function Utf8ToString(ByVal bytes As Variant) As String
    Dim st As Object
    Set st = CreateObject("ADODB.Stream")
    st.Type = 1: st.Open: st.Write bytes
    st.Position = 0: st.Type = 2: st.Charset = "utf-8"
    Utf8ToString = st.ReadText
    st.Close
End Function

' Quote-aware CSV into a 2-D Variant array (row 1 = header)
Private Function ParseCsv(ByVal s As String) As Variant
    Dim rows As Collection, cur() As String, nCol As Long, field As String
    Dim i As Long, ch As String, inQ As Boolean, r As Long
    Set rows = New Collection
    ReDim cur(0 To 63): nCol = 0
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If inQ Then
            If ch = """" Then
                If Mid$(s, i + 1, 1) = """" Then field = field & """": i = i + 1 Else inQ = False
            Else
                field = field & ch
            End If
        ElseIf ch = """" Then
            inQ = True
        ElseIf ch = "," Then
            cur(nCol) = field: nCol = nCol + 1: field = ""
        ElseIf ch = vbLf Or ch = vbCr Then
            If ch = vbLf Or Mid$(s, i + 1, 1) <> vbLf Then
                If nCol > 0 Or Len(field) > 0 Then
                    cur(nCol) = field: nCol = nCol + 1
                    Dim one() As String: ReDim one(0 To nCol - 1)
                    Dim k As Long
                    For k = 0 To nCol - 1: one(k) = cur(k): Next k
                    rows.Add one
                End If
                nCol = 0: field = ""
            End If
        Else
            field = field & ch
        End If
    Next i
    If nCol > 0 Or Len(field) > 0 Then
        cur(nCol) = field: nCol = nCol + 1
        ReDim one(0 To nCol - 1)
        For k = 0 To nCol - 1: one(k) = cur(k): Next k
        rows.Add one
    End If
    Dim res() As Variant, c As Long, rr As Variant, maxC As Long
    maxC = UBound(rows(1))
    ReDim res(1 To rows.Count, 0 To maxC)
    r = 0
    For Each rr In rows
        r = r + 1
        For c = 0 To UBound(rr)
            If c <= maxC Then res(r, c) = rr(c)
        Next c
    Next rr
    ParseCsv = res
End Function

Private Function ColIndex(ByVal data As Variant, ByVal nm As String) As Long
    Dim c As Long
    ColIndex = -1
    For c = 0 To UBound(data, 2)
        If data(1, c) = nm Then ColIndex = c: Exit Function
    Next c
End Function

Private Function ImportOne(ByVal fileName As String, ByVal sheetName As String, ByVal spec As String) As String
    Dim body As String, dlErr As String, data As Variant, toks() As String
    Dim ws As Worksheet, n As Long, r As Long, c As Long, tok As String
    Dim idx() As Long, outv() As Variant, isText() As Boolean, nm As String, v As String
    dlErr = Download(FEED_ROOT & fileName, body)
    If Len(dlErr) > 0 Then ImportOne = dlErr: Exit Function
    If Len(body) < 20 Then ImportOne = "empty file": Exit Function
    data = ParseCsv(body)
    n = UBound(data, 1) - 1
    If n < 0 Then ImportOne = "no header": Exit Function
    toks = Split(spec, ",")
    ReDim idx(0 To UBound(toks)): ReDim isText(0 To UBound(toks))
    For c = 0 To UBound(toks)
        tok = toks(c)
        isText(c) = (Left$(tok, 1) = "~")
        nm = tok
        If Left$(tok, 1) = "~" Then nm = Mid$(tok, 2)
        If Left$(tok, 1) = "#" Then
            idx(c) = -2
        ElseIf tok = "@LOCAL" Then
            idx(c) = ColIndex(data, "Interval_Start_UTC"): If idx(c) < 0 Then ImportOne = "column Interval_Start_UTC missing": Exit Function
        Else
            idx(c) = ColIndex(data, nm)
            If idx(c) < 0 Then ImportOne = "column " & nm & " missing (schema changed, saved data kept)": Exit Function
        End If
    Next c
    If n = 0 Then ImportOne = "no data rows, saved data kept": Exit Function
    ReDim outv(1 To n, 1 To UBound(toks) + 1)
    For r = 1 To n
        For c = 0 To UBound(toks)
            tok = toks(c)
            If idx(c) = -2 Then
                outv(r, c + 1) = Mid$(tok, 2)
            ElseIf tok = "@LOCAL" Then
                outv(r, c + 1) = BudapestLocal(CStr(data(r + 1, idx(c))))
            Else
                v = CStr(data(r + 1, idx(c)))
                If Len(v) = 0 Then
                    outv(r, c + 1) = Empty                       ' missing stays blank
                ElseIf isText(c) Then
                    outv(r, c + 1) = v
                ElseIf IsNumber(v) Then
                    outv(r, c + 1) = Val(v)                      ' Val uses '.' regardless of locale
                Else
                    outv(r, c + 1) = v
                End If
            End If
        Next c
    Next r
    If sheetName = "D_Price" Then outv = WithGapDays(outv): n = UBound(outv, 1)
    If sheetName = "D_EntsoePrice" Then outv = LatestDayOnly(outv): n = UBound(outv, 1)
    Set ws = ThisWorkbook.Worksheets(sheetName)
    ws.Range(ws.Cells(2, 1), ws.Cells(ws.Rows.Count, UBound(toks) + 1)).ClearContents
    For c = 0 To UBound(toks)
        ws.Range(ws.Cells(2, c + 1), ws.Cells(n + 1, c + 1)).NumberFormat = IIf(isText(c) Or Left$(toks(c), 1) = "#", "@", IIf(toks(c) = "@LOCAL", "yyyy-mm-dd hh:mm", "General"))
    Next c
    ws.Range(ws.Cells(2, 1), ws.Cells(n + 1, UBound(toks) + 1)).Value2 = outv
End Function

' HUPX price rows only exist for days the source published. Insert 96 blank quarter-hour rows for each
' omitted whole day (Delivery_Date text in column 1, local time in column 2, price blank) so the chart
' shows an explicit gap instead of joining the neighbouring days. Blank, never zero, never estimated.
Private Function WithGapDays(ByVal a As Variant) As Variant
    Dim n As Long, r As Long, k As Long, q As Long, d0 As Date, d1 As Date, dd As Date
    Dim missing As Long, extra As Long, outv() As Variant, o As Long, prev As String, cur As String, cols As Long
    n = UBound(a, 1): cols = UBound(a, 2)
    If n < 2 Then WithGapDays = a: Exit Function
    For r = 2 To n
        prev = CStr(a(r - 1, 1)): cur = CStr(a(r, 1))
        If prev <> cur Then
            d0 = DateSerial(CLng(Left$(prev, 4)), CLng(Mid$(prev, 6, 2)), CLng(Mid$(prev, 9, 2)))
            d1 = DateSerial(CLng(Left$(cur, 4)), CLng(Mid$(cur, 6, 2)), CLng(Mid$(cur, 9, 2)))
            If d1 - d0 > 1 Then missing = missing + CLng(d1 - d0 - 1): extra = extra + ExtraQuarters(d0, d1)
        End If
    Next r
    If missing = 0 Then WithGapDays = a: Exit Function
    ReDim outv(1 To n + missing * 96 + extra, 1 To cols)
    o = 0
    For r = 1 To n
        If r > 1 Then
            prev = CStr(a(r - 1, 1)): cur = CStr(a(r, 1))
            If prev <> cur Then
                d0 = DateSerial(CLng(Left$(prev, 4)), CLng(Mid$(prev, 6, 2)), CLng(Mid$(prev, 9, 2)))
                d1 = DateSerial(CLng(Left$(cur, 4)), CLng(Mid$(cur, 6, 2)), CLng(Mid$(cur, 9, 2)))
                For dd = d0 + 1 To d1 - 1
                    For q = 0 To QuartersInDay(dd) - 1
                        o = o + 1
                        outv(o, 1) = Format$(dd, "yyyy-mm-dd")
                        outv(o, 2) = LocalOfUtc(UtcMidnight(dd) + TimeSerial(0, 15 * q, 0))
                    Next q
                Next dd
            End If
        End If
        o = o + 1
        For k = 1 To cols: outv(o, k) = a(r, k): Next k
    Next r
    WithGapDays = outv
End Function

' Budapest delivery day length: 92 quarter-hours on the March clock change, 100 on the October one, else 96.
Private Function QuartersInDay(ByVal d As Date) As Long
    Dim y As Long
    y = Year(d)
    QuartersInDay = 96
    If d = LastSunday(y, 3) Then QuartersInDay = 92
    If d = LastSunday(y, 10) Then QuartersInDay = 100
End Function

' sum of (QuartersInDay - 96) over the days strictly between d0 and d1
Private Function ExtraQuarters(ByVal d0 As Date, ByVal d1 As Date) As Long
    Dim dd As Date
    For dd = d0 + 1 To d1 - 1
        ExtraQuarters = ExtraQuarters + QuartersInDay(dd) - 96
    Next dd
End Function

' UTC instant of local midnight at the start of Budapest day d
Private Function UtcMidnight(ByVal d As Date) As Date
    If QuartersInDay(d) = 100 Then
        UtcMidnight = d - TimeSerial(2, 0, 0)       ' CEST at midnight on the October change day
    ElseIf d >= LastSunday(Year(d), 3) And d < LastSunday(Year(d), 10) Then
        UtcMidnight = d - TimeSerial(IIf(d = LastSunday(Year(d), 3), 1, 2), 0, 0)
    Else
        UtcMidnight = d - TimeSerial(1, 0, 0)
    End If
End Function

' UTC time -> Europe/Budapest wall-clock (EU rule), numeric twin of BudapestLocal
Private Function LocalOfUtc(ByVal d As Date) As Date
    Dim y As Long
    y = Year(d)
    If d >= LastSunday(y, 3) + TimeSerial(1, 0, 0) And d < LastSunday(y, 10) + TimeSerial(1, 0, 0) Then
        LocalOfUtc = d + TimeSerial(2, 0, 0)
    Else
        LocalOfUtc = d + TimeSerial(1, 0, 0)
    End If
End Function

' The ENTSO-E chart shows one delivery day only: keep the rows of the latest Delivery_Date (column 1).
Private Function LatestDayOnly(ByVal a As Variant) As Variant
    Dim n As Long, r As Long, k As Long, o As Long, mx As String, cols As Long, outv() As Variant
    n = UBound(a, 1): cols = UBound(a, 2)
    For r = 1 To n
        If CStr(a(r, 1)) > mx Then mx = CStr(a(r, 1))
    Next r
    For r = 1 To n
        If CStr(a(r, 1)) = mx Then o = o + 1
    Next r
    If o = 0 Then LatestDayOnly = a: Exit Function
    ReDim outv(1 To o, 1 To cols)
    o = 0
    For r = 1 To n
        If CStr(a(r, 1)) = mx Then
            o = o + 1
            For k = 1 To cols: outv(o, k) = a(r, k): Next k
        End If
    Next r
    LatestDayOnly = outv
End Function

Private Function IsNumber(ByVal s As String) As Boolean
    Dim i As Long, ch As String, dots As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = "." Then
            dots = dots + 1
            If dots > 1 Then Exit Function
        ElseIf ch = "-" Or ch = "+" Then
            If i > 1 And LCase$(Mid$(s, i - 1, 1)) <> "e" Then Exit Function
        ElseIf ch = "e" Or ch = "E" Then
            If i = 1 Then Exit Function
        ElseIf ch < "0" Or ch > "9" Then
            Exit Function
        End If
    Next i
    IsNumber = True
End Function

' ISO UTC text -> Europe/Budapest clock time (EU rule: last Sunday of March/October at 01:00 UTC)
Private Function BudapestLocal(ByVal iso As String) As Variant
    Dim d As Date, y As Long
    On Error GoTo Bad
    d = DateSerial(CLng(Mid$(iso, 1, 4)), CLng(Mid$(iso, 6, 2)), CLng(Mid$(iso, 9, 2))) + _
        TimeSerial(CLng(Mid$(iso, 12, 2)), CLng(Mid$(iso, 15, 2)), CLng(Mid$(iso, 18, 2)))
    y = Year(d)
    If d >= LastSunday(y, 3) + TimeSerial(1, 0, 0) And d < LastSunday(y, 10) + TimeSerial(1, 0, 0) Then
        BudapestLocal = d + TimeSerial(2, 0, 0)
    Else
        BudapestLocal = d + TimeSerial(1, 0, 0)
    End If
    Exit Function
Bad:
    BudapestLocal = Empty
End Function

Private Function LastSunday(ByVal y As Long, ByVal m As Long) As Date
    Dim d As Date
    d = DateSerial(y, m + 1, 0)
    LastSunday = d - (Weekday(d, vbSunday) - 1)
End Function

' dataset states from status.json (regex over the dataset_details block)
Private Function ImportStatus() As String
    Dim body As String, dlErr As String, p As Long, rx As Object, ms As Object, m As Object, ws As Worksheet, r As Long
    dlErr = Download(FEED_ROOT & "status.json", body)
    If Len(dlErr) > 0 Then ImportStatus = dlErr: Exit Function
    p = InStr(body, """dataset_details""")
    If p = 0 Then ImportStatus = "dataset_details missing": Exit Function
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = """([a-z0-9_]+)"": \{\s*""state"": ""([A-Z_]+)""[\s\S]*?""detail"": ""((?:[^""\\]|\\.)*)""[\s\S]*?(?:""last_success_at"": (?:null|""([^""]*)""))"
    Set ms = rx.Execute(Mid$(body, p))
    If ms.Count = 0 Then ImportStatus = "no dataset states parsed, saved data kept": Exit Function
    Set ws = ThisWorkbook.Worksheets("D_Status")
    ws.Range("A2:D200").ClearContents
    r = 2
    For Each m In ms
        ws.Cells(r, 1).Value2 = m.SubMatches(0)
        ws.Cells(r, 2).Value2 = m.SubMatches(1)
        ws.Cells(r, 3).Value2 = Left$(Replace(m.SubMatches(2), "\""", """"), 240)
        ws.Cells(r, 4).Value2 = m.SubMatches(3)
        r = r + 1
    Next m
End Function

' extend every chart series that reads a D_* sheet to that sheet's last filled row
Private Sub ResizeCharts()
    Dim ws As Worksheet, co As ChartObject, s As Series, rx As Object, f As String, m As Object, sh As String, lastR As Long, out As String
    Set rx = CreateObject("VBScript.RegExp")
    rx.Global = True
    rx.Pattern = "(D_[A-Za-z]+)!\$([A-Z]+)\$(\d+):\$([A-Z]+)\$\d+"
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 2) <> "D_" Then
            For Each co In ws.ChartObjects
                co.Chart.DisplayBlanksAs = xlNotPlotted      ' blanks are gaps, never zero, on every chart
                For Each s In co.Chart.SeriesCollection
                    f = s.Formula
                    out = f
                    For Each m In rx.Execute(f)
                        sh = m.SubMatches(0)
                        If sh = "D_Tariffs" Then GoTo NextMatch   ' tariff charts use fixed per-group row blocks
                        lastR = ThisWorkbook.Worksheets(sh).Cells(ThisWorkbook.Worksheets(sh).Rows.Count, 1).End(xlUp).Row
                        If lastR < 2 Then lastR = 2
                        out = Replace(out, m.Value, sh & "!$" & m.SubMatches(1) & "$" & m.SubMatches(2) & ":$" & m.SubMatches(3) & "$" & lastR)
NextMatch:
                    Next m
                    If out <> f Then On Error Resume Next: s.Formula = out: On Error GoTo 0
                Next s
            Next co
        End If
    Next ws
End Sub

' rewrites the value tables on the Registers page from the refreshed D_* sheets (top storage, E.ON substations, installed capacity).
' All new values are computed first (staging), the old cells are backed up, and any error restores the backup.
Private Function RefreshRegisterTables() As String
    Dim ws As Worksheet, a As Range, aS As Range, aE As Range, aI As Range, d As Worksheet
    Dim r As Long, i As Long, k As Long, best As Long, lastR As Long, nI As Long
    Dim used() As Boolean, bp As Double, lp As String, g As String, tot As Double
    Dim stoOut(1 To 10, 1 To 4) As Variant, subOut(1 To 12, 1 To 4) As Variant
    Dim bkS As Variant, bkE As Variant, bkT As Variant, bkI As Variant, bkH As Variant
    Dim insC() As Variant, totTxt As String, wrote As Boolean
    On Error GoTo Fail
    For Each ws In ThisWorkbook.Worksheets
        If Left$(ws.Name, 2) <> "D_" Then
            Set aS = ws.Columns(2).Find("Largest unrevoked storage", LookIn:=xlValues, LookAt:=xlPart)
            If Not aS Is Nothing Then Exit For
        End If
    Next ws
    If aS Is Nothing Then RefreshRegisterTables = "Registers page not found": Exit Function
    Set aE = ws.Columns(2).Find("E.ON offtake", LookIn:=xlValues, LookAt:=xlPart)
    Set aI = ws.Columns(2).Find("Installed capacity by primary source", LookIn:=xlValues, LookAt:=xlPart)
    If aE Is Nothing Or aI Is Nothing Then RefreshRegisterTables = "register table headings not found": Exit Function
    ' count installed-capacity rows on the page
    i = aI.Row + 2
    Do While Len(CStr(ws.Cells(i, 2).Value2)) > 0 And Left$(CStr(ws.Cells(i, 2).Value2), 6) <> "Latest"
        i = i + 1
    Loop
    nI = i - (aI.Row + 2)
    ' ---- staging: top 10 unrevoked storage licences by MW
    Set d = ThisWorkbook.Worksheets("D_Storage")
    lastR = d.Cells(d.Rows.Count, 1).End(xlUp).Row
    ReDim used(1 To IIf(lastR < 2, 2, lastR))
    For k = 1 To 10
        best = 0: bp = -1
        For r = 2 To lastR
            If Len(CStr(d.Cells(r, 5).Value2)) = 0 And Not IsEmpty(d.Cells(r, 3).Value2) And Not used(r) Then
                If IsNumeric(d.Cells(r, 3).Value2) Then
                    If CDbl(d.Cells(r, 3).Value2) > bp Then bp = CDbl(d.Cells(r, 3).Value2): best = r
                End If
            End If
        Next r
        If best > 0 Then
            used(best) = True
            stoOut(k, 1) = Left$(CStr(d.Cells(best, 1).Value2), 34)
            stoOut(k, 2) = d.Cells(best, 3).Value2
            stoOut(k, 3) = d.Cells(best, 4).Value2
            stoOut(k, 4) = Left$(CStr(d.Cells(best, 2).Value2), 60)
        End If                                      ' unused slots stay Empty and clear the old row
    Next k
    ' ---- staging: top 12 E.ON substations by releasable MVA
    Set d = ThisWorkbook.Worksheets("D_Substations")
    lastR = d.Cells(d.Rows.Count, 1).End(xlUp).Row
    ReDim used(1 To IIf(lastR < 2, 2, lastR))
    tot = 0
    For r = 2 To lastR
        If Not IsEmpty(d.Cells(r, 4).Value2) Then If IsNumeric(d.Cells(r, 4).Value2) Then tot = tot + CDbl(d.Cells(r, 4).Value2)
    Next r
    For k = 1 To 12
        best = 0: bp = -1
        For r = 2 To lastR
            If Not IsEmpty(d.Cells(r, 4).Value2) And Not used(r) Then
                If IsNumeric(d.Cells(r, 4).Value2) Then
                    If CDbl(d.Cells(r, 4).Value2) > bp Then bp = CDbl(d.Cells(r, 4).Value2): best = r
                End If
            End If
        Next r
        If best > 0 Then
            used(best) = True
            subOut(k, 1) = d.Cells(best, 2).Value2: subOut(k, 2) = d.Cells(best, 3).Value2
            subOut(k, 3) = d.Cells(best, 4).Value2: subOut(k, 4) = d.Cells(best, 5).Value2
        End If
    Next k
    totTxt = "Total of " & Application.WorksheetFunction.Max(lastR - 1, 0) & " published substation rows: " & Format$(tot, "0") & " MVA releasable now. MVM Halozat, MVM Emasz and MAVIR substation data are not parsed."
    ' ---- staging: installed capacity (2025 and latest period per fuel group)
    Set d = ThisWorkbook.Worksheets("D_Installed")
    lastR = d.Cells(d.Rows.Count, 1).End(xlUp).Row
    lp = ""
    For r = 2 To lastR
        If CStr(d.Cells(r, 2).Value2) > lp Then lp = CStr(d.Cells(r, 2).Value2)
    Next r
    ReDim insC(1 To IIf(nI < 1, 1, nI), 1 To 2)
    For i = 1 To nI
        g = CStr(ws.Cells(aI.Row + 1 + i, 2).Value2)
        For r = 2 To lastR
            If Replace(CStr(d.Cells(r, 1).Value2), "_", " ") = g Then
                If CStr(d.Cells(r, 2).Value2) = "2025" Then insC(i, 1) = d.Cells(r, 3).Value2
                If CStr(d.Cells(r, 2).Value2) = lp Then insC(i, 2) = d.Cells(r, 3).Value2
            End If
        Next r
    Next i
    ' ---- backup of every output block
    bkS = ws.Range(ws.Cells(aS.Row + 2, 2), ws.Cells(aS.Row + 11, 5)).Value2
    bkE = ws.Range(ws.Cells(aE.Row + 2, 2), ws.Cells(aE.Row + 13, 5)).Value2
    bkT = ws.Cells(aE.Row + 15, 2).Value2
    bkH = ws.Cells(aI.Row + 1, 4).Value2
    If nI > 0 Then bkI = ws.Range(ws.Cells(aI.Row + 2, 3), ws.Cells(aI.Row + 1 + nI, 4)).Value2
    ' ---- write
    wrote = True
    ws.Range(ws.Cells(aS.Row + 2, 2), ws.Cells(aS.Row + 11, 5)).Value2 = stoOut
    ws.Range(ws.Cells(aE.Row + 2, 2), ws.Cells(aE.Row + 13, 5)).Value2 = subOut
    ws.Cells(aE.Row + 15, 2).Value2 = totTxt
    ws.Cells(aI.Row + 1, 4).Value2 = "Latest (MW) " & lp
    If nI > 0 Then ws.Range(ws.Cells(aI.Row + 2, 3), ws.Cells(aI.Row + 1 + nI, 4)).Value2 = insC
    Exit Function
Fail:
    RefreshRegisterTables = "error " & Err.Number & ": " & Err.Description
    If wrote Then
        On Error Resume Next
        ws.Range(ws.Cells(aS.Row + 2, 2), ws.Cells(aS.Row + 11, 5)).Value2 = bkS
        ws.Range(ws.Cells(aE.Row + 2, 2), ws.Cells(aE.Row + 13, 5)).Value2 = bkE
        ws.Cells(aE.Row + 15, 2).Value2 = bkT
        ws.Cells(aI.Row + 1, 4).Value2 = bkH
        If nI > 0 Then ws.Range(ws.Cells(aI.Row + 2, 3), ws.Cells(aI.Row + 1 + nI, 4)).Value2 = bkI
        RefreshRegisterTables = RefreshRegisterTables & " (previous tables restored)"
    Else
        RefreshRegisterTables = RefreshRegisterTables & " (nothing written, saved tables kept)"
    End If
End Function
