Attribute VB_Name = "modTables"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology (VBA) - tables drawn in the DGN and report exports
' Summary table, legend and data-sources table (lines + text on level USGS_StreamStats_Table),
' plus CSV, Excel and KML exports saved beside the active DGN.
'==============================================================================================
Private Const LEVEL_TABLE As String = "USGS_StreamStats_Table"
Private mTableIds As Collection       ' element IDs of the tables drawn by the last batch

Public Sub ResetTableIds()
    ' The tables were deleted (Delete All Drawings); nothing to replace next time.
    Set mTableIds = Nothing
End Sub


'----------------------------------------------------------------------------------------------
' Table builder
'----------------------------------------------------------------------------------------------
Private Function EstWidth(ByVal s As String, ByVal th As Double) As Double
    Dim i As Long, ch As String, w As Double
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If ch = " " Then
            w = w + 0.38
        ElseIf InStr(1, "il.,:;'|!()[]/-I1", ch) > 0 Then
            w = w + 0.42
        ElseIf (ch >= "A" And ch <= "Z") Or InStr(1, "%&@mw", ch) > 0 Then
            w = w + 0.92
        Else
            w = w + 0.74
        End If
    Next i
    EstWidth = w * th
End Function

Private Function CellAlign(ByVal s As String, ByVal r As Long, ByVal c As Long) As String
    ' C for headers and placeholders, R for numbers (so decimals line up), L for text
    Dim t As String
    t = Trim$(s)
    If r = 0 Or t = "" Or t = "-" Or t = "N/A" Then
        CellAlign = "C"
    ElseIf c > 0 And IsNumeric(Replace(t, ",", "")) Then
        CellAlign = "R"
    Else
        CellAlign = "L"
    End If
End Function

Private Sub PlaceCellText(ByVal s As String, ByVal xa As Double, ByVal yc As Double, ByVal th As Double, _
                          ByVal align As String, ByVal lvl As Level, ByVal colorIndex As Long)
    ' Places text, then shifts it using its real extents so it sits leftX / centre / right of xa and is
    ' centred on yc, whatever the active font and justification are.
    Dim t As TextElement, r As Range3d, w As Double, h As Double, est As Double, wasOn As Boolean
    Dim dx As Double, dy As Double, leftX As Double
    If Len(s) = 0 Then Exit Sub
    wasOn = ActiveSettings.AnnotationScaleEnabled
    On Error GoTo Restore
    ActiveSettings.AnnotationScaleEnabled = False
    Set t = CreateTextElement1(Nothing, s, Point3dFromXY(xa, yc), Matrix3dIdentity)
    t.TextStyle.Height = th
    t.TextStyle.Width = th
    AddStyled t, lvl, colorIndex, 0
    r = t.Range
    w = r.High.x - r.Low.x
    h = r.High.y - r.Low.y
    est = EstWidth(s, th)
    If w > 0.3 * est And w < 3# * est And h > 0.3 * th And h < 3# * th Then
        If align = "R" Then
            leftX = xa - w
        ElseIf align = "C" Then
            leftX = xa - w / 2#
        Else
            leftX = xa
        End If
        dx = leftX - r.Low.x
        dy = (yc - h / 2#) - r.Low.y
        t.Move Point3dFromXY(dx, dy)
        t.Rewrite
    End If
Restore:
    ActiveSettings.AnnotationScaleEnabled = wasOn
End Sub

Private Function CellWidth(ByVal s As String, ByVal th As Double) As Double
    CellWidth = EstWidth(s, th) * 1.08
End Function

Public Function BuildTable(ByRef rows As Collection, ByVal tx As Double, ByVal ty As Double, ByVal title As String, _
                           ByVal th As Double, Optional ByVal symbols As Variant, Optional ByVal rowColors As Variant, _
                           Optional ByVal minFirstCol As Double = 0#) As Double
    ' Bordered table of lines + text. rows = Collection of string arrays (row 1 = header).
    ' symbols(r) names a legend symbol drawn in the first cell; rowColors(r) = text color (-1 = default).
    ' Returns the Y of the bottom border.
    Dim lvl As Level, ncols As Long, nrows As Long, r As Long, c As Long, v As Variant
    Dim widths() As Double, w As Double, pad As Double, rowH As Double, titleTh As Double, titleH As Double
    Dim totalW As Double, bodyTop As Double, bottom As Double, x As Double, x0 As Double, x1 As Double, yc As Double
    Dim s As String, al As String, col As Long, hasSym As Boolean, hasCol As Boolean
    Set lvl = EnsureLevel(LEVEL_TABLE, 0, 1)
    titleTh = 1.33 * th
    pad = 0.9 * th
    rowH = 2.2 * th
    titleH = 2.4 * titleTh
    nrows = rows.Count
    For Each v In rows
        If UBound(v) + 1 > ncols Then ncols = UBound(v) + 1
    Next v
    ReDim widths(0 To ncols - 1)
    For c = 0 To ncols - 1
        widths(c) = 4# * th
    Next c
    For Each v In rows
        For c = 0 To UBound(v)
            w = CellWidth(CStr(v(c)), th) + 2# * pad
            If w > widths(c) Then widths(c) = w
        Next c
    Next v
    If minFirstCol > widths(0) Then widths(0) = minFirstCol
    For c = 0 To ncols - 1
        totalW = totalW + widths(c)
    Next c
    w = CellWidth(title, titleTh) + 2# * pad
    If w > totalW Then
        widths(ncols - 1) = widths(ncols - 1) + (w - totalW)
        totalW = w
    End If
    bodyTop = ty - titleH
    bottom = bodyTop - rowH * nrows
    ' outer border (heavy), header separator (medium), body grid (thin)
    AddLine tx, ty, tx + totalW, ty, lvl, 0, 3
    AddLine tx, bottom, tx + totalW, bottom, lvl, 0, 3
    AddLine tx, ty, tx, bottom, lvl, 0, 3
    AddLine tx + totalW, ty, tx + totalW, bottom, lvl, 0, 3
    AddLine tx, bodyTop, tx + totalW, bodyTop, lvl, 0, 2
    For r = 1 To nrows - 1
        AddLine tx, bodyTop - r * rowH, tx + totalW, bodyTop - r * rowH, lvl, 0, IIf(r = 1, 2, 0)
    Next r
    x = tx
    For c = 0 To ncols - 2
        x = x + widths(c)
        AddLine x, bodyTop, x, bottom, lvl, 0, 0
    Next c
    PlaceCellText title, tx + pad, ty - titleH / 2#, titleTh, "L", lvl, 0
    hasSym = Not IsMissing(symbols)
    hasCol = Not IsMissing(rowColors)
    r = 0
    For Each v In rows
        yc = bodyTop - (r + 0.5) * rowH
        x0 = tx
        col = 0
        If hasCol And r > 0 Then
            If rowColors(r) >= 0 Then col = rowColors(r)
        End If
        For c = 0 To ncols - 1
            x1 = x0 + widths(c)
            s = ""
            If c <= UBound(v) Then s = CStr(v(c))
            If Len(s) > 0 Then
                al = CellAlign(s, r, c)
                If al = "L" Then
                    PlaceCellText s, x0 + pad, yc, th, al, lvl, col
                ElseIf al = "R" Then
                    PlaceCellText s, x1 - pad, yc, th, al, lvl, col
                Else
                    PlaceCellText s, (x0 + x1) / 2#, yc, th, al, lvl, col
                End If
            End If
            If c = 0 And hasSym Then
                If Len(CStr(symbols(r))) > 0 Then DrawSymbol CStr(symbols(r)), x0 + pad, x1 - pad, yc, rowH, lvl
            End If
            x0 = x1
        Next c
        r = r + 1
    Next v
    BuildTable = bottom
End Function

Private Sub DrawSymbol(ByVal kind As String, ByVal x0 As Double, ByVal x1 As Double, ByVal yc As Double, _
                       ByVal h As Double, ByVal lvl As Level)
    ' Legend samples drawn the same way as in the drawing.
    Dim xm As Double, a As Double, el As Element, ls As LineStyle
    xm = (x0 + x1) / 2#
    Select Case kind
        Case "inlet"
            a = 0.3 * h
            AddLine xm - a, yc - a, xm + a, yc - a, lvl, 6, 2
            AddLine xm + a, yc - a, xm + a, yc + a, lvl, 6, 2
            AddLine xm + a, yc + a, xm - a, yc + a, lvl, 6, 2
            AddLine xm - a, yc + a, xm - a, yc - a, lvl, 6, 2
            AddLine xm - a, yc - a, xm + a, yc + a, lvl, 6, 1
            AddLine xm + a, yc - a, xm - a, yc + a, lvl, 6, 1
        Case "pour"
            a = 0.25 * h
            Set el = CreateEllipseElement2(Nothing, Point3dFromXY(xm, yc), a, a, Matrix3dIdentity, msdFillModeNotFilled)
            AddStyled el, lvl, 2, 3
            AddLine xm - 1.5 * a, yc, xm + 1.5 * a, yc, lvl, 2, 1
            AddLine xm, yc - 1.5 * a, xm, yc + 1.5 * a, lvl, 2, 1
        Case "basin"
            Set el = CreateLineElement2(Nothing, Point3dFromXY(x0, yc), Point3dFromXY(x1, yc))
            On Error Resume Next
            Set ls = ActiveDesignFile.LineStyles.Find("2", Nothing)
            If Not ls Is Nothing Then Set el.LineStyle = ls
            On Error GoTo 0
            AddStyled el, lvl, 3, 2
        Case "stream"
            AddLine x0, yc, x1, yc, lvl, 1, 2
        Case "arrow", "rd_arrow"
            a = 0.35 * h
            AddLine x0, yc, x1, yc, lvl, IIf(kind = "arrow", 1, 7), 1
            AddFilledTriangle xm + a, yc, xm - a, yc + 0.6 * a, xm - a, yc - 0.6 * a, lvl, IIf(kind = "arrow", 1, 7)
        Case "lfp"
            AddLine x0, yc, x1, yc, lvl, 5, 3
        Case "rd_overland"
            Set el = CreateLineElement2(Nothing, Point3dFromXY(x0, yc), Point3dFromXY(x1, yc))
            On Error Resume Next
            Set ls = ActiveDesignFile.LineStyles.Find("2", Nothing)
            If Not ls Is Nothing Then Set el.LineStyle = ls
            On Error GoTo 0
            AddStyled el, lvl, 4, 2
        Case "rd_channel"
            AddLine x0, yc, x1, yc, lvl, 7, 3
        Case "rd_point"
            AddCircleX xm, yc, 0.32 * h, lvl, 7
    End Select
End Sub

'----------------------------------------------------------------------------------------------
' Batch summary + legend + data sources
'----------------------------------------------------------------------------------------------
Public Sub DefaultTableOrigin(ByRef queue As Collection, ByRef tx As Double, ByRef ty As Double)
    ' Top-left of the table: just right of the combined drainage-area extents, level with their top.
    Dim it As Variant, res As Object, ring As Variant, i As Long, x As Double, y As Double, found As Boolean
    Dim maxx As Double, maxy As Double
    maxx = -1E+30
    maxy = -1E+30
    For Each it In queue
        If Not it("res") Is Nothing Then
            Set res = it("res")
            Set ring = res("rings")(1)
            For i = 1 To ring.Count
                LonLatToDgn CDbl(ring(i)(1)), CDbl(ring(i)(2)), x, y
                If x > maxx Then maxx = x
                If y > maxy Then maxy = y
                found = True
            Next i
        End If
    Next it
    If found Then
        ' Clear of the marker labels, which run to the right of a culvert on the basin edge.
        tx = maxx + 34# * TextHeightFt()
        ty = maxy
    End If
End Sub

Public Sub PlaceBatchTables(ByRef queue As Collection, ByVal tx As Double, ByVal ty As Double, ByVal replacePrevious As Boolean)
    ' Summary table, legend and data sources. With replacePrevious the tables from the last batch are removed first.
    Dim rows As Collection, it As Variant, res As Object, qD As Variant, th As Double, bottom As Double
    Dim n As Long, colors() As Long, syms() As String, i As Long, id As Variant
    Dim anyBc As Boolean, anyNlcd As Boolean, anyNss As Boolean, anyUrban As Boolean, anyAtlas As Boolean, anySoil As Boolean
    Dim qB As Variant, qA As Variant, anyAlt As Boolean, anyAnd As Boolean, anyLfpNhd As Boolean, anyLfpSs As Boolean
    Dim anySoilUc As Boolean, anyCat As Boolean
    On Error GoTo Bad
    th = TextHeightFt()
    If replacePrevious And Not mTableIds Is Nothing Then
        On Error Resume Next
        For Each id In mTableIds
            ActiveModelReference.RemoveElement ActiveModelReference.GetElementByID(DLongFromString(CStr(id)))
        Next id
        On Error GoTo Bad
    End If
    Set gRecIds = New Collection

    Set rows = New Collection
    rows.Add Array("CULVERT ID", "DRAINAGE (AC)", "SLOPE (FT/MI)", "Q10 (CFS)", "Q25 (CFS)", "Q50 (CFS)", "Q100 (CFS)", _
                   "GOVERNING METHOD", "BASIN DATA", "RAINFALL")
    For Each it In queue
        n = n + 1
        If it("res") Is Nothing Then
            rows.Add Array(CStr(it("id")), "FAILED", "-", "-", "-", "-", "-", "Error", "-", "-")
        Else
            Set res = it("res")
            qD = res("q_des")
            qB = res("q_urban")
            rows.Add Array(CStr(it("id")), Format$(res("acres"), "#,##0.0"), Format$(res("slope"), "#,##0.0"), _
                           Format$(qD(2), "#,##0.0"), Format$(qD(3), "#,##0.0"), Format$(qD(4), "#,##0.0"), Format$(qD(5), "#,##0.0"), _
                           CStr(res("gov")), IIf(res("have_bc"), "STREAMSTATS", "ALTERNATE"), IIf(res("have_atlas"), "ATLAS 14", "BUILT-IN"))
            qA = res("q_and")
            If res("have_bc") Then anyBc = True Else anyAlt = True
            If res("have_nlcd") Then anyNlcd = True
            If res("cover_src") = "streamcat" Then anyCat = True
            If res("have_nss") Then anyNss = True
            If res("have_atlas") Then anyAtlas = True
            If Not res("hsg") Is Nothing Then
                If InStr(CStr(res("hsg_src")), "UC Davis") > 0 Then anySoilUc = True Else anySoil = True
            End If
            If InStr(CStr(res("lfp_method")), "StreamStats") > 0 Then
                anyLfpSs = True
            ElseIf Len(CStr(res("lfp_method"))) > 0 Then
                anyLfpNhd = True
            End If
            For i = 0 To 7
                If qB(i) > 0# Then anyUrban = True
                If qA(i) > 0# Then anyAnd = True
            Next i
        End If
    Next it
    bottom = BuildTable(rows, tx, ty, "BATCH CULVERT HYDROLOGY SUMMARY  -  VDOT design peak flows  (" & n & " culverts, " & _
                        Format$(Date, "yyyy-mm-dd") & ")", th)

    ' Legend
    Set rows = New Collection
    rows.Add Array("SYMBOL", "FEATURE", "DESCRIPTION")
    rows.Add Array("", "CULVERT INLET", "Culvert location (picked or imported)")
    rows.Add Array("", "POUR POINT", "StreamStats snapped outlet")
    rows.Add Array("", "DRAINAGE AREA", "StreamStats watershed boundary")
    rows.Add Array("", "NHD STREAM", "NHDPlus flowline within the drainage area")
    rows.Add Array("", "FLOW DIRECTION", "Arrowheads point downstream")
    rows.Add Array("", "LONGEST FLOWPATH", "Hydraulic length used for time of concentration")
    rows.Add Array("DRAINAGE (AC)", "COLUMN", "Delineated drainage area, acres")
    rows.Add Array("SLOPE (FT/MI)", "COLUMN", "Basin mean channel slope, feet per mile")
    rows.Add Array("Q10 - Q100", "COLUMN", "VDOT design peak discharge (cfs) for the 10- to 100-yr storm")
    rows.Add Array("RATIONAL", "GOVERNING METHOD", "VDOT Rational Q = Cf x C x i x A, drainage area <= 200 ac")
    rows.Add Array("REGRESSION", "GOVERNING METHOD", "USGS regional regression (NSS), drainage area > 200 ac")
    rows.Add Array("STREAMSTATS", "BASIN DATA", "Slope and land cover from USGS StreamStats")
    rows.Add Array("ALTERNATE", "BASIN DATA", "StreamStats down: NLCD 2021 land cover, 3DEP slope, polygon area")
    rows.Add Array("ATLAS 14", "RAINFALL", "NOAA Atlas 14 intensity at the pour point")
    rows.Add Array("BUILT-IN", "RAINFALL", "Statewide IDF coefficients (placeholder) - verify before design")
    ReDim colors(0 To rows.Count - 1)
    ReDim syms(0 To rows.Count - 1)
    For i = 0 To rows.Count - 1
        colors(i) = -1
    Next i
    syms(1) = "inlet"
    colors(1) = 6
    syms(2) = "pour"
    colors(2) = 2
    syms(3) = "basin"
    colors(3) = 3
    syms(4) = "stream"
    syms(5) = "arrow"
    syms(6) = "lfp"
    colors(6) = 5
    colors(13) = 4
    colors(15) = 4
    bottom = BuildTable(rows, tx, bottom - 2# * th, "LEGEND", th, syms, colors, 6# * th)

    ' Data sources actually used by this batch
    Set rows = New Collection
    rows.Add Array("DATA", "SOURCE", "WEBSITE")
    rows.Add Array("Watershed boundary, pour point", "USGS StreamStats", "https://streamstats.usgs.gov")
    If anyBc Then rows.Add Array("Basin characteristics", "USGS StreamStats", "https://streamstats.usgs.gov")
    If anyNlcd Then rows.Add Array("Land cover, impervious, % urban", "MRLC National Land Cover Database 2021", "https://www.mrlc.gov")
    If anyCat Then rows.Add Array("Land cover (StreamStats and MRLC down)", "EPA StreamCat (NLCD 2019)", "https://www.epa.gov/national-aquatic-resource-surveys/streamcat-dataset")
    If anyAlt Then rows.Add Array("Channel slope (StreamStats down)", "USGS 3DEP elevation", "https://elevation.nationalmap.gov")
    rows.Add Array("Streams, flow direction", "USGS National Hydrography Dataset (NHDPlus HR)", "https://hydro.nationalmap.gov")
    If anyLfpNhd Then rows.Add Array("Longest flowpath", "Traced from USGS NHD flowlines", "https://hydro.nationalmap.gov")
    If anyLfpSs Then rows.Add Array("Longest flowpath", "USGS StreamStats", "https://streamstats.usgs.gov")
    If anyNss Then rows.Add Array("Rural regression peak flows", "USGS National Streamflow Statistics (NSS)", "https://streamstats.usgs.gov/nssservices")
    If anyUrban Then rows.Add Array("Urban regression peak flows", "USGS SIR 2014-5090 equations", "https://pubs.usgs.gov/publication/sir20145090")
    If anyAnd Then rows.Add Array("Anderson method", "USGS Water-Supply Paper 2001-C", "https://pubs.usgs.gov/publication/wsp01C")
    If anyAtlas Then rows.Add Array("Rainfall intensity and depth", "NOAA Atlas 14 Precipitation Frequency Data Server", "https://hdsc.nws.noaa.gov/pfds")
    If anySoil Then rows.Add Array("Hydrologic soil groups", "USDA NRCS SSURGO (Soil Data Access)", "https://sdmdataaccess.sc.egov.usda.gov")
    If anySoilUc Then rows.Add Array("Hydrologic soil groups", "UC Davis SoilWeb ISSR-800 grid (from SSURGO)", "https://casoilresource.lawr.ucdavis.edu")
    rows.Add Array("Methods, size ranges, Cf factors", "VDOT Drainage Manual, Chapter 6 (Rev. 7/26)", "https://www.vdot.virginia.gov")
    bottom = BuildTable(rows, tx, bottom - 2# * th, "DATA SOURCES", th)

    Set mTableIds = gRecIds
    Set gRecIds = Nothing
    LogLine "Placed the summary table, legend and data sources in the DGN (" & mTableIds.Count & " elements on level '" & LEVEL_TABLE & "')."
    Exit Sub
Bad:
    Set mTableIds = gRecIds
    Set gRecIds = Nothing
    LogLine "[Batch Table Error] " & Err.Description
End Sub

'----------------------------------------------------------------------------------------------
' Exports (saved beside the active DGN)
'----------------------------------------------------------------------------------------------
Public Function DgnStemPath() As String
    Dim nm As String, p As Long
    nm = ActiveDesignFile.Name
    p = InStrRev(nm, ".")
    If p > 1 Then nm = Left$(nm, p - 1)
    DgnStemPath = ActiveDesignFile.path
    If Right$(DgnStemPath, 1) <> "\" Then DgnStemPath = DgnStemPath & "\"
    DgnStemPath = DgnStemPath & nm
End Function

Private Function CsvQ(ByVal s As String) As String
    If InStr(s, ",") > 0 Or InStr(s, """") > 0 Then
        CsvQ = """" & Replace(s, """", """""") & """"
    Else
        CsvQ = s
    End If
End Function

Private Function QTxt(ByVal v As Variant) As String
    If CDbl(v) > 0# Then QTxt = NumFix(CDbl(v), 1)
End Function

Public Function ExportBatchCsv(ByRef queue As Collection) As String
    ' Batch report CSV. Returns the file path ("" on failure).
    Dim f As Integer, path As String, it As Variant, res As Object, qU As Variant, qD As Variant, qR As Variant, qN As Variant
    Dim qB As Variant
    On Error GoTo Bad
    path = DgnStemPath() & "_Batch_Culvert_Hydrology_VBA.csv"
    f = FreeFile
    Open path For Output As #f
    Print #f, "BATCH CULVERT HYDROLOGY & DRAINAGE SUMMARY (VBA edition)"
    Print #f, "Date," & Format$(Now, "yyyy-mm-dd hh:nn:ss") & ",Region,VA"
    Print #f, ""
    Print #f, "Culvert ID,Status,Drainage (sq.mi),Drainage (acres),Slope (ft/mi),VDOT C,Tc (min)," & _
              "Q2 USGS (cfs),Q10 USGS (cfs),Q25 USGS (cfs),Q50 USGS (cfs),Q100 USGS (cfs)," & _
              "Q10 Rational (cfs),Q10 NRCS (cfs),Q10 Urban (cfs)," & _
              "Q10 VDOT (cfs),Q25 VDOT (cfs),Q50 VDOT (cfs),Q100 VDOT (cfs),Governing Method,Exit Lon,Exit Lat," & _
              "Basin Data Source,Rainfall Source,Curve Number,Curve Number Source,Percent Urban," & _
              "Q10 Anderson (cfs),Tc Length (ft),Tc Length Source,Longest Flowpath (ft),Longest Flowpath Method,Soil Source"
    For Each it In queue
        If it("res") Is Nothing Then
            Print #f, CsvQ(CStr(it("id"))) & ",FAILED"
        Else
            Set res = it("res")
            qU = res("q_usgs")
            qD = res("q_des")
            qR = res("q_rat")
            qN = res("q_nrcs")
            qB = res("q_urban")
            Print #f, CsvQ(CStr(it("id"))) & ",Success," & NumFix(res("sqmi"), 4) & "," & NumFix(res("acres"), 1) & "," & _
                      NumFix(res("slope"), 1) & "," & NumFix(res("c"), 2) & "," & NumFix(res("tc"), 1) & "," & _
                      QTxt(qU(0)) & "," & QTxt(qU(2)) & "," & QTxt(qU(3)) & "," & QTxt(qU(4)) & "," & QTxt(qU(5)) & "," & _
                      QTxt(qR(2)) & "," & QTxt(qN(2)) & "," & QTxt(qB(2)) & "," & _
                      QTxt(qD(2)) & "," & QTxt(qD(3)) & "," & QTxt(qD(4)) & "," & QTxt(qD(5)) & "," & CsvQ(CStr(res("gov"))) & "," & _
                      NumFix(res("s_lon"), 6) & "," & NumFix(res("s_lat"), 6) & "," & _
                      CsvQ(CStr(res("bc_src"))) & "," & _
                      IIf(res("have_atlas"), "NOAA Atlas 14", "Built-in IDF (placeholder)") & "," & _
                      IIf(res("cn") > 0#, NumFix(res("cn"), 1), "") & "," & CsvQ(CStr(res("cn_src"))) & "," & _
                      IIf(res("urban_pct") >= 0#, NumFix(res("urban_pct"), 1), "") & "," & _
                      QTxt(res("q_and")(2)) & "," & IIf(res("tc_len") > 0#, NumFix(res("tc_len"), 1), "") & "," & CStr(res("tc_len_src")) & "," & _
                      IIf(res("lfp_ft") > 0#, NumFix(res("lfp_ft"), 1), "") & "," & CsvQ(CStr(res("lfp_method"))) & "," & CsvQ(CStr(res("hsg_src")))
        End If
    Next it
    Close #f
    ExportBatchCsv = path
    LogLine "Saved batch CSV report: " & path
    Exit Function
Bad:
    On Error Resume Next
    Close #f
    LogLine "[Export Error] CSV report: " & Err.Description
End Function

Public Function ExportExcel(ByRef queue As Collection) As String
    ' Results & Comparison workbook (Summary sheet + one sheet per culvert). Needs Microsoft Excel.
    Dim xl As Object, wb As Object, ws As Object, cs As Object, it As Variant, res As Object, path As String
    Dim r As Long, i As Long, yrs As Variant, qU As Variant, qD As Variant, qR As Variant, qN As Variant, qB As Variant
    Dim picks As Variant, rain As Variant, inten As Variant, nm As String, n As Long
    On Error GoTo Bad
    path = DgnStemPath() & "_Results_Comparison_VBA.xlsx"
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    Set xl = CreateObject("Excel.Application")
    xl.Visible = False
    xl.DisplayAlerts = False
    Set wb = xl.Workbooks.Add
    Set ws = wb.Worksheets(1)
    ws.Name = "Summary"
    ws.cells(1, 1).Value = "USGS StreamStats vs VDOT - Results & Comparison (VBA edition)"
    ws.cells(1, 1).Font.Bold = True
    ws.cells(2, 1).Value = "Generated"
    ws.cells(2, 2).Value = Format$(Now, "yyyy-mm-dd hh:nn")
    r = 4
    ws.Range("A4:M4").Value = Array("Culvert ID", "Drainage (sq mi)", "Drainage (ac)", "Slope (ft/mi)", "Rational C", "Tc (min)", _
                                    "Governing Method", "Q10 USGS", "Q100 USGS", "Q10 VDOT", "Q25 VDOT", "Q50 VDOT", "Q100 VDOT")
    ws.Range("A4:M4").Font.Bold = True
    For Each it In queue
        If Not it("res") Is Nothing Then
            Set res = it("res")
            qU = res("q_usgs")
            qD = res("q_des")
            qR = res("q_rat")
            qN = res("q_nrcs")
            qB = res("q_urban")
            picks = res("picks")
            rain = res("rain24")
            inten = res("inten")
            r = r + 1
            ws.cells(r, 1).Value = CStr(it("id"))
            ws.cells(r, 2).Value = CDbl(res("sqmi"))
            ws.cells(r, 3).Value = CDbl(res("acres"))
            ws.cells(r, 4).Value = CDbl(res("slope"))
            ws.cells(r, 5).Value = CDbl(res("c"))
            ws.cells(r, 6).Value = CDbl(res("tc"))
            ws.cells(r, 7).Value = CStr(res("gov"))
            If qU(2) > 0# Then ws.cells(r, 8).Value = CDbl(qU(2))
            If qU(5) > 0# Then ws.cells(r, 9).Value = CDbl(qU(5))
            ws.cells(r, 10).Value = CDbl(qD(2))
            ws.cells(r, 11).Value = CDbl(qD(3))
            ws.cells(r, 12).Value = CDbl(qD(4))
            ws.cells(r, 13).Value = CDbl(qD(5))
            n = n + 1
            nm = Left$(Replace(Replace(Replace(CStr(it("id")), "/", "_"), "\", "_"), ":", "_"), 28)
            Set cs = wb.Worksheets.Add(, wb.Worksheets(wb.Worksheets.Count))
            On Error Resume Next
            cs.Name = nm
            On Error GoTo Bad
            cs.cells(1, 1).Value = "Culvert " & CStr(it("id")) & " - Results & Comparison"
            cs.cells(1, 1).Font.Bold = True
            cs.Range("A3:C3").Value = Array("Basin Metric", "Value", "Units")
            cs.Range("A3:C3").Font.Bold = True
            cs.Range("A4:C4").Value = Array("Drainage Area", CDbl(res("sqmi")), "sq mi")
            cs.Range("A5:C5").Value = Array("Drainage Area", CDbl(res("acres")), "ac")
            cs.Range("A6:C6").Value = Array("Basin Slope", CDbl(res("slope")), "ft/mi")
            cs.Range("A7:C7").Value = Array("Impervious", CDbl(res("imp")), "%")
            cs.Range("A8:C8").Value = Array("Forest", CDbl(res("forest")), "%")
            cs.Range("A9:C9").Value = Array("Rational C", CDbl(res("c")), "")
            cs.Range("A10:C10").Value = Array("Time of Concentration", CDbl(res("tc")), "min")
            cs.Range("A11:C11").Value = Array("Governing Method", CStr(res("gov")), "")
            cs.Range("A12:C12").Value = Array("Basin Data Source", CStr(res("bc_src")), "")
            cs.Range("A13:C13").Value = Array("Rainfall Source", IIf(res("have_atlas"), "NOAA Atlas 14 (at the pour point)", "Built-in IDF (placeholder)"), "")
            If res("urban_pct") >= 0# Then cs.Range("A14:C14").Value = Array("Percent Urban (developed)", CDbl(res("urban_pct")), "%")
            If res("cn") > 0# Then
                cs.Range("A15:C15").Value = Array("NRCS Curve Number", CDbl(res("cn")), "")
                cs.Range("A16:B16").Value = Array("Curve Number Source", CStr(res("cn_src")))
            End If
            cs.Range("A17:C17").Value = Array("Tc Flow Length", IIf(res("tc_len") > 0#, CDbl(res("tc_len")), CStr(res("tc_len_src"))), _
                                              IIf(res("tc_len") > 0#, "ft  (" & CStr(res("lfp_method")) & ")", ""))
            cs.Range("A18:J18").Value = Array("Event", "USGS Rural Regression (cfs)", "Rational (cfs)", "NRCS TR-55 (cfs)", _
                                              "USGS Urban Regression (cfs)", "VDOT Design (cfs)", "Design Method", _
                                              "Rainfall Intensity (in/hr)", "24-hr Rainfall (in)", "Anderson (cfs)")
            cs.Range("A18:J18").Font.Bold = True
            For i = 0 To 7
                cs.cells(19 + i, 1).Value = "Q" & yrs(i) & " (" & yrs(i) & "-yr)"
                If qU(i) > 0# Then cs.cells(19 + i, 2).Value = CDbl(qU(i))
                If qR(i) > 0# Then cs.cells(19 + i, 3).Value = CDbl(qR(i))
                If qN(i) > 0# Then cs.cells(19 + i, 4).Value = CDbl(qN(i))
                If qB(i) > 0# Then cs.cells(19 + i, 5).Value = CDbl(qB(i))
                cs.cells(19 + i, 6).Value = CDbl(qD(i))
                cs.cells(19 + i, 7).Value = MethodLabel(CStr(picks(i)), CDbl(res("acres")))
                cs.cells(19 + i, 8).Value = CDbl(inten(i))
                If rain(i) > 0# Then cs.cells(19 + i, 9).Value = CDbl(rain(i))
                If res("q_and")(i) > 0# Then cs.cells(19 + i, 10).Value = CDbl(res("q_and")(i))
            Next i
            cs.cells(28, 1).Value = "Verify against the VDOT Drainage Manual before design use."
            cs.Columns("A:J").AutoFit
        End If
    Next it
    ws.Columns("A:M").AutoFit
    ws.Activate
    If n = 0 Then
        wb.Close False
        xl.Quit
        LogLine "[Export] No results to export yet."
        Exit Function
    End If
    wb.SaveAs path, 51
    wb.Close False
    xl.Quit
    ExportExcel = path
    LogLine "Exported Results & Comparison: " & path
    Exit Function
Bad:
    LogLine "[Excel Export Error] " & Err.Description
    On Error Resume Next
    wb.Close False
    xl.Quit
End Function

Public Function ExportKml(ByRef queue As Collection) As String
    ' Drainage areas and pour points as a KML file (opens in Google Earth / ArcGIS).
    Dim f As Integer, path As String, it As Variant, res As Object, ring As Variant, i As Long, n As Long, s As String
    Dim qD As Variant
    On Error GoTo Bad
    path = DgnStemPath() & "_StreamStats_VBA.kml"
    f = FreeFile
    Open path For Output As #f
    Print #f, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Print #f, "<kml xmlns=""http://www.opengis.net/kml/2.2""><Document><name>USGS StreamStats drainage areas</name>"
    Print #f, "<Style id=""basin""><LineStyle><color>ff0000ff</color><width>2</width></LineStyle><PolyStyle><color>330000ff</color></PolyStyle></Style>"
    For Each it In queue
        If Not it("res") Is Nothing Then
            Set res = it("res")
            qD = res("q_des")
            Set ring = res("rings")(1)
            s = ""
            For i = 1 To ring.Count
                s = s & NumFix(CDbl(ring(i)(1)), 6) & "," & NumFix(CDbl(ring(i)(2)), 6) & ",0 "
            Next i
            Print #f, "<Placemark><name>" & XmlEsc(CStr(it("id"))) & " drainage area</name><styleUrl>#basin</styleUrl>"
            Print #f, "<description>" & XmlEsc(Format$(res("acres"), "#,##0.0") & " ac; Q10 " & Format$(qD(2), "#,##0.0") & " cfs; Q100 " & _
                      Format$(qD(5), "#,##0.0") & " cfs; " & CStr(res("gov"))) & "</description>"
            Print #f, "<Polygon><outerBoundaryIs><LinearRing><coordinates>" & s & "</coordinates></LinearRing></outerBoundaryIs></Polygon></Placemark>"
            Print #f, "<Placemark><name>" & XmlEsc(CStr(it("id"))) & " PP</name><Point><coordinates>" & NumFix(res("s_lon"), 6) & "," & _
                      NumFix(res("s_lat"), 6) & ",0</coordinates></Point></Placemark>"
            n = n + 1
        End If
    Next it
    Print #f, "</Document></kml>"
    Close #f
    If n = 0 Then
        Kill path
        LogLine "[Export] No results to export yet."
        Exit Function
    End If
    ExportKml = path
    LogLine "Exported KML (" & n & " drainage area(s)): " & path
    Exit Function
Bad:
    On Error Resume Next
    Close #f
    LogLine "[KML Export Error] " & Err.Description
End Function

Private Function XmlEsc(ByVal s As String) As String
    s = Replace(s, "&", "&amp;")
    s = Replace(s, "<", "&lt;")
    s = Replace(s, ">", "&gt;")
    XmlEsc = s
End Function
