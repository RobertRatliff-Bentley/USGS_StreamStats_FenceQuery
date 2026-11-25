Attribute VB_Name = "modMethods"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology (VBA) - peak-discharge methods, land cover and soils
' VDOT Drainage Manual, Chapter 6 (Rev. 7/26), Table 6-2. Same equations as the Python tool.
'==============================================================================================

' Method settings (Settings tab). MethodDefaults sets them the way the VDOT manual recommends.
Public gMethodsInit As Boolean
Public gUseRational As Boolean
Public gUseNrcs As Boolean
Public gUseRural As Boolean
Public gUseUrban As Boolean
Public gUseRanges As Boolean       ' only calculate each method inside its Table 6-2 size range
Public gDesign As String           ' auto, rational, nrcs, rural or urban
Public gCN As Double               ' NRCS curve number for every culvert; 0 = estimate per basin
Public gHsg As String              ' auto (USDA SSURGO) or A, B, C, D

Private Const URL_MRLC As String = "https://www.mrlc.gov/geoserver/ows"
Private Const URL_SSURGO As String = "https://sdmdataaccess.sc.egov.usda.gov/Tabular/post.rest"
Private Const NLCD_MAX_CELLS As Long = 500
Private mSoilDownUntil As Date
Private mNlcdDownUntil As Date

Public Sub MethodDefaults()
    gUseRational = True
    gUseNrcs = True
    gUseRural = True
    gUseUrban = True
    gUseRanges = True
    gDesign = "auto"
    gCN = 0#
    gHsg = "auto"
    gUseAnderson = False
    gAndersonType = "natural"
    gMethodsInit = True
End Sub

Public Sub SetMethodsServiceDown(ByVal key As String, ByVal untilTime As Date)
    ' Used by the service check: untilTime 0 = try the service again at once.
    If key = "ssurgo" Then mSoilDownUntil = untilTime
    If key = "nlcd" Then mNlcdDownUntil = untilTime
End Sub

Public Function MethodLabel(ByVal key As String, ByVal acres As Double) As String
    Select Case key
        Case "rational"
            MethodLabel = "VDOT Rational (<=200 ac)"
        Case "nrcs"
            MethodLabel = "NRCS Graphical Peak (TR-55)"
        Case "rural"
            If acres > 200# Then
                MethodLabel = "USGS/VDOT Regression (>200 ac)"
            Else
                MethodLabel = "USGS Rural Regression"
            End If
        Case "urban"
            MethodLabel = "USGS Urban Regression"
        Case "anderson"
            MethodLabel = "Anderson Method (USGS)"
        Case Else
            MethodLabel = "Built-in regression check (placeholder)"
    End Select
End Function

'----------------------------------------------------------------------------------------------
' NRCS graphical peak discharge (TR-55, Type II storm)
'----------------------------------------------------------------------------------------------
Public Function NrcsQ(ByVal rain24 As Double, ByVal cn As Double, ByVal sqmi As Double, ByVal tcMin As Double) As Double
    ' qp = qu x A x Q (no pond/swamp adjustment). Tc is limited to TR-55's 0.1-10 hr range. -1 = not available.
    Dim s As Double, ia As Double, q As Double, ratio As Double, lt As Double, f As Double
    Dim r As Variant, c0 As Variant, c1 As Variant, c2 As Variant, i As Long, lo As Double, hi As Double
    NrcsQ = -1#
    If rain24 <= 0# Or cn <= 0# Or sqmi <= 0# Then Exit Function
    If cn < 30# Then cn = 30#
    If cn > 98# Then cn = 98#
    s = 1000# / cn - 10#
    ia = 0.2 * s
    If rain24 <= ia Then
        NrcsQ = 0#
        Exit Function
    End If
    q = (rain24 - ia) ^ 2 / (rain24 - ia + s)
    ratio = ia / rain24
    If ratio < 0.1 Then ratio = 0.1
    If ratio > 0.5 Then ratio = 0.5
    lt = tcMin / 60#
    If lt < 0.1 Then lt = 0.1
    If lt > 10# Then lt = 10#
    lt = Log(lt) / Log(10#)
    r = Array(0.1, 0.3, 0.35, 0.4, 0.45, 0.5)
    c0 = Array(2.55323, 2.46532, 2.41896, 2.36409, 2.29238, 2.20282)
    c1 = Array(-0.61512, -0.62257, -0.61594, -0.59857, -0.57005, -0.51599)
    c2 = Array(-0.16403, -0.11657, -0.0882, -0.05621, -0.02281, -0.01259)
    For i = 0 To 4
        If ratio >= r(i) And ratio <= r(i + 1) Then
            f = (ratio - r(i)) / (r(i + 1) - r(i))
            lo = c0(i) + c1(i) * lt + c2(i) * lt * lt
            hi = c0(i + 1) + c1(i + 1) * lt + c2(i + 1) * lt * lt
            NrcsQ = (10# ^ (lo + f * (hi - lo))) * sqmi * q
            Exit Function
        End If
    Next i
End Function

'----------------------------------------------------------------------------------------------
' USGS urban regression for Virginia (SIR 2014-5090; VDOT section 6.4.3.4.4)
'----------------------------------------------------------------------------------------------
Public Function UrbanQ(ByVal yr As Long, ByVal sqmi As Double, ByVal urbanPct As Double) As Double
    Dim b As Variant, la As Double, lq As Double
    UrbanQ = -1#
    If sqmi <= 0# Then Exit Function
    Select Case yr
        Case 2: b = Array(2.027, 40.29, 1.216, -0.00414, 0.00468, -0.366)
        Case 5: b = Array(2.229, 39.37, 1.139, -0.00346, 0.00487, -0.338)
        Case 10: b = Array(2.373, 38.706, 1.103, -0.00313, 0.0047, -0.334)
        Case 25: b = Array(2.557, 39.168, 1.083, -0.00224, 0.00434, -0.332)
        Case 50: b = Array(2.697, 39.168, 1.083, -0.00219, 0.0039, -0.343)
        Case 100: b = Array(2.776, 38.765, 1.07, -0.00242, 0.00434, -0.342)
        Case 200: b = Array(2.863, 39.063, 1.057, -0.00223, 0.00465, -0.329)
        Case 500: b = Array(2.961, 39.287, 0.904, -0.00049, 0.00636, -0.317)
        Case Else: Exit Function
    End Select
    la = Log(sqmi) / Log(10#)
    lq = b(0) + (urbanPct - b(1)) * ((la - b(2)) * b(3)) + urbanPct * b(4) + la * b(5)
    UrbanQ = (10# ^ lq) * sqmi
End Function

'----------------------------------------------------------------------------------------------
' Curve number from NLCD land cover and hydrologic soil groups
'----------------------------------------------------------------------------------------------
Private Function CnFor(ByVal nlcdClass As Long, ByVal soilIndex As Long) As Double
    ' Runoff curve numbers by NLCD class for soil groups A, B, C, D (TR-55 Table 2-2 cover types). 0 = no value.
    Dim v As Variant
    Select Case nlcdClass
        Case 11: v = Array(98, 98, 98, 98)
        Case 21: v = Array(49, 69, 79, 84)
        Case 22: v = Array(57, 72, 81, 86)
        Case 23: v = Array(77, 85, 90, 92)
        Case 24: v = Array(89, 92, 94, 95)
        Case 31: v = Array(77, 86, 91, 94)
        Case 41, 42, 43: v = Array(30, 55, 70, 77)
        Case 52: v = Array(35, 56, 70, 77)
        Case 71: v = Array(30, 58, 71, 78)
        Case 81: v = Array(39, 61, 74, 80)
        Case 82: v = Array(67, 78, 85, 89)
        Case 90: v = Array(77, 77, 77, 77)
        Case 95: v = Array(78, 78, 78, 78)
        Case Else: Exit Function
    End Select
    CnFor = v(soilIndex)
End Function

Public Function EstimateCN(ByRef classes As Object, ByRef hsg As Object, ByVal hsgSource As String, ByRef hsgText As String) As Double
    ' Area-weighted CN from NLCD class fractions and soil-group fractions. The soil group comes from
    ' Settings (A-D), else the basin's soil groups, else group B.
    Dim w(0 To 3) As Double, letters As Variant, i As Long, k As Variant, tot As Double, totC As Double
    Dim cn As Double, rowCn As Double, frac As Double
    letters = Array("A", "B", "C", "D")
    If gHsg = "A" Or gHsg = "B" Or gHsg = "C" Or gHsg = "D" Then
        For i = 0 To 3
            If letters(i) = gHsg Then w(i) = 1#
        Next i
        hsgText = "soil group " & gHsg & " (Settings)"
    ElseIf Not hsg Is Nothing Then
        hsgText = hsgSource & " soils"
        For i = 0 To 3
            If hsg.Exists(letters(i)) Then
                w(i) = CDbl(hsg(letters(i)))
                If w(i) > 0# Then hsgText = hsgText & IIf(Right$(hsgText, 5) = "soils", " ", ", ") & letters(i) & " " & Format$(w(i) * 100#, "0") & "%"
            End If
        Next i
    Else
        w(1) = 1#
        hsgText = "soil group B assumed (SSURGO unavailable)"
    End If
    For i = 0 To 3
        tot = tot + w(i)
    Next i
    If tot <= 0# Then Exit Function
    For Each k In classes.keys
        If CnFor(CLng(k), 0) > 0# Then totC = totC + CDbl(classes(k))
    Next k
    If totC <= 0# Then Exit Function
    For Each k In classes.keys
        If CnFor(CLng(k), 0) > 0# Then
            rowCn = 0#
            For i = 0 To 3
                rowCn = rowCn + (w(i) / tot) * CnFor(CLng(k), i)
            Next i
            frac = CDbl(classes(k)) / totC
            cn = cn + frac * rowCn
        End If
    Next k
    EstimateCN = RoundTo(cn, 1)
End Function

'----------------------------------------------------------------------------------------------
' NLCD 2021 (MRLC) clipped to the basin: impervious %, forest %, developed %, class fractions
'----------------------------------------------------------------------------------------------
Public Function NlcdBasinCover(ByRef rings As Collection, ByRef info As Object) As Boolean
    Dim ring As Variant, i As Long, n As Long, r As Long, c As Long, c0 As Long, c1 As Long, k As Long, j As Long
    Dim minx As Double, miny As Double, maxx As Double, maxy As Double, side As Double, sc As Double
    Dim lc() As Long, im() As Long, nrows As Long, ncols As Long, nr2 As Long, nc2 As Long
    Dim b(0 To 3) As Double, b2(0 To 3) As Double, dx As Double, dy As Double, yc As Double
    Dim xs() As Double, nx As Long, x1 As Double, y1 As Double, x2 As Double, y2 As Double, t As Double
    Dim cells As Long, forest As Long, dev As Long, impSum As Double, classes As Object, cls As Long, key As Variant
    If Now < mNlcdDownUntil Then
        LogLine "  MRLC NLCD land cover service is down (retrying in " & Format$((mNlcdDownUntil - Now) * 1440#, "0") & " min)."
        Exit Function
    End If
    On Error GoTo NoNlcd
    minx = 1E+30
    miny = 1E+30
    maxx = -1E+30
    maxy = -1E+30
    For Each ring In rings
        For i = 1 To ring.Count
            x1 = CDbl(ring(i)(1))
            y1 = CDbl(ring(i)(2))
            If x1 < minx Then minx = x1
            If x1 > maxx Then maxx = x1
            If y1 < miny Then miny = y1
            If y1 > maxy Then maxy = y1
        Next i
    Next ring
    minx = minx - 0.0006
    miny = miny - 0.0006
    maxx = maxx + 0.0006
    maxy = maxy + 0.0006
    side = MaxD((maxx - minx) * Cos((miny + maxy) / 2# * 3.14159265358979 / 180#), maxy - miny) * 111320# / 30#
    sc = MinD(1#, NLCD_MAX_CELLS / MaxD(side, 1#))
    If Not NlcdGrid("mrlc_download__NLCD_2021_Land_Cover_L48", minx, miny, maxx, maxy, sc, lc, nrows, ncols, b) Then Exit Function
    If Not NlcdGrid("mrlc_download__NLCD_2021_Impervious_L48", minx, miny, maxx, maxy, sc, im, nr2, nc2, b2) Then Exit Function
    If nr2 <> nrows Or nc2 <> ncols Then Err.Raise vbObjectError + 30, , "NLCD land-cover and impervious grids do not match"
    dx = (b(2) - b(0)) / ncols
    dy = (b(3) - b(1)) / nrows
    Set classes = CreateObject("Scripting.Dictionary")
    ReDim xs(0 To 63)
    For r = 0 To nrows - 1
        ' x-crossings of every ring edge with this row of cell centres (even-odd, so holes are excluded)
        yc = b(3) - (r + 0.5) * dy
        nx = 0
        For Each ring In rings
            n = ring.Count
            For i = 1 To n - 1
                y1 = CDbl(ring(i)(2))
                y2 = CDbl(ring(i + 1)(2))
                If (y1 > yc) <> (y2 > yc) Then
                    x1 = CDbl(ring(i)(1))
                    x2 = CDbl(ring(i + 1)(1))
                    If nx > UBound(xs) Then ReDim Preserve xs(0 To UBound(xs) * 2 + 1)
                    xs(nx) = x1 + (yc - y1) * (x2 - x1) / (y2 - y1)
                    nx = nx + 1
                End If
            Next i
        Next ring
        For i = 1 To nx - 1          ' insertion sort (few crossings per row)
            t = xs(i)
            j = i - 1
            Do While j >= 0
                If xs(j) <= t Then Exit Do
                xs(j + 1) = xs(j)
                j = j - 1
            Loop
            xs(j + 1) = t
        Next i
        For k = 0 To nx - 2 Step 2
            c0 = -Int(-((xs(k) - b(0)) / dx - 0.5))
            c1 = -Int(-((xs(k + 1) - b(0)) / dx - 0.5)) - 1
            If c0 < 0 Then c0 = 0
            If c1 > ncols - 1 Then c1 = ncols - 1
            For c = c0 To c1
                cls = lc(r, c)
                If cls <> 0 And im(r, c) >= 0 And im(r, c) <= 100 Then
                    cells = cells + 1
                    impSum = impSum + im(r, c)
                    classes(cls) = classes(cls) + 1
                    If cls = 41 Or cls = 42 Or cls = 43 Then forest = forest + 1
                    If cls >= 21 And cls <= 24 Then dev = dev + 1
                End If
            Next c
        Next k
    Next r
    If cells = 0 Then Err.Raise vbObjectError + 30, , "no NLCD cells inside the basin"
    For Each key In classes.keys
        classes(key) = classes(key) / cells
    Next key
    Set info = CreateObject("Scripting.Dictionary")
    info("imperv_pct") = impSum / cells
    info("forest_pct") = 100# * forest / cells
    info("dev_pct") = 100# * dev / cells
    info("cells") = cells
    Set info("classes") = classes
    NlcdBasinCover = True
    LogLine "  NLCD 2021 (MRLC): impervious " & Format$(impSum / cells, "0.0") & "%, forest " & Format$(100# * forest / cells, "0.0") & _
            "%, developed " & Format$(100# * dev / cells, "0.0") & "% from " & Format$(cells, "#,##0") & " cells in the basin."
    Exit Function
NoNlcd:
    If Not gCancel Then mNlcdDownUntil = Now + 10# / 1440#
    LogLine "  [WARNING] MRLC NLCD land cover unavailable (" & Err.Description & ")."
End Function

Private Function NlcdGrid(ByVal coverage As String, ByVal minx As Double, ByVal miny As Double, ByVal maxx As Double, _
                          ByVal maxy As Double, ByVal sc As Double, ByRef g() As Long, ByRef nrows As Long, _
                          ByRef ncols As Long, ByRef b() As Double) As Boolean
    ' NLCD cells for a WGS84 box (north row first). b = min lon, min lat, max lon, max lat of the grid returned.
    Dim url As String, txt As String, p As Long, q As Long, head As String, body As String, tok As String, ch As String
    Dim lines() As String, parts() As String, i As Long, r As Long, c As Long, n As Long
    Dim v As Double, lons(0 To 1) As Double, lats(0 To 1) As Double, nl As Long, nt As Long, rowsText() As String
    url = URL_MRLC & "?service=WCS&version=2.0.1&request=GetCoverage&coverageId=" & coverage & _
          "&format=text/plain&subsettingCrs=http://www.opengis.net/def/crs/EPSG/0/4326" & _
          "&subset=Long(" & NumFix(minx, 6) & "," & NumFix(maxx, 6) & ")&subset=Lat(" & NumFix(miny, 6) & "," & NumFix(maxy, 6) & ")"
    If sc < 1# Then url = url & "&scalefactor=" & NumFix(sc, 4)
    txt = HttpText("GET", url, "", 90)
    p = InStr(1, txt, "Grid bounds:")
    q = InStr(1, txt, "Band 0:")
    If p = 0 Or q = 0 Then Err.Raise vbObjectError + 30, , "unexpected response from the MRLC coverage service"
    head = Mid$(txt, p, InStr(p, txt, vbLf) - p)
    For i = 1 To Len(head) + 1
        ch = Mid$(head, i, 1)
        If Len(ch) > 0 And InStr(1, "-0123456789.eE", ch) > 0 Then
            tok = tok & ch
        Else
            If InStr(1, tok, ".") > 0 And IsNumeric(tok) Then
                v = Val(tok)
                If v < 0# Then
                    If nl < 2 Then lons(nl) = v
                    nl = nl + 1
                Else
                    If nt < 2 Then lats(nt) = v
                    nt = nt + 1
                End If
            End If
            tok = ""
        End If
    Next i
    If nl <> 2 Or nt <> 2 Then Err.Raise vbObjectError + 30, , "unexpected grid bounds from the MRLC coverage service"
    b(0) = MinD(lons(0), lons(1))
    b(2) = MaxD(lons(0), lons(1))
    b(1) = MinD(lats(0), lats(1))
    b(3) = MaxD(lats(0), lats(1))
    body = Replace(Mid$(txt, q + 7), vbCr, "")
    lines = Split(body, vbLf)
    ReDim rowsText(0 To UBound(lines))
    For i = 0 To UBound(lines)
        If Len(Trim$(lines(i))) > 0 Then
            rowsText(n) = Trim$(lines(i))
            n = n + 1
        End If
    Next i
    If n = 0 Then Err.Raise vbObjectError + 30, , "MRLC coverage service returned no cells"
    nrows = n
    parts = Split(rowsText(0), " ")
    ncols = UBound(parts) + 1
    ReDim g(0 To nrows - 1, 0 To ncols - 1)
    For r = 0 To nrows - 1
        parts = Split(rowsText(r), " ")
        If UBound(parts) + 1 < ncols Then Err.Raise vbObjectError + 30, , "ragged NLCD grid"
        For c = 0 To ncols - 1
            g(r, c) = CLng(Val(parts(c)))
        Next c
    Next r
    NlcdGrid = True
End Function

'----------------------------------------------------------------------------------------------
' Hydrologic soil groups inside the basin (USDA SSURGO, Soil Data Access)
'----------------------------------------------------------------------------------------------
Public Function SoilGroups(ByRef rings As Collection, ByRef hsg As Object) As Boolean
    ' Area fractions for groups A-D. Dual groups (A/D, B/D, C/D) count as D, the undrained condition.
    Dim ring As Variant, n As Long, stepN As Long, i As Long, wkt As String, q As String, body As String, txt As String
    Dim d As Variant, row As Variant, grp As String, tot As Double, k As Variant, s As String, first As String
    If Now < mSoilDownUntil Then
        LogLine "  USDA SSURGO soils service is down (retrying in " & Format$((mSoilDownUntil - Now) * 1440#, "0") & " min)."
        Exit Function
    End If
    On Error GoTo NoSoil
    Set ring = rings(1)
    n = ring.Count
    stepN = 1
    If n > 400 Then stepN = n \ 400 + 1
    For i = 1 To n Step stepN
        s = NumFix(CDbl(ring(i)(1)), 6) & " " & NumFix(CDbl(ring(i)(2)), 6)
        If Len(first) = 0 Then first = s
        wkt = wkt & s & ","
    Next i
    wkt = "polygon((" & wkt & first & "))"
    q = "~DeclareGeometry(@aoi)~" & vbLf & _
        "select @aoi = geometry::STPolyFromText('" & wkt & "', 4326)" & vbLf & _
        "~DeclareIdGeomTable(@clip)~" & vbLf & _
        "~GetClippedMapunits(@aoi,polygon,geo,@clip)~" & vbLf & _
        "select c.hydgrp, sum(g.geom.STArea() * c.comppct_r) as w from @clip g join component c on c.mukey = g.id " & _
        "where c.hydgrp is not null group by c.hydgrp"
    body = "{""query"":" & JsonQuote(q) & ",""format"":""JSON""}"
    txt = HttpText("POST", URL_SSURGO, body, 90, 1)
    JsonParse txt, d
    If Not HasKey(d, "Table") Then Err.Raise vbObjectError + 31, , "no soil groups returned for this basin"
    Set hsg = CreateObject("Scripting.Dictionary")
    For Each row In d("Table")
        grp = UCase$(Trim$(CStr(row(1))))
        If InStr(grp, "/") > 0 Then grp = Mid$(grp, InStrRev(grp, "/") + 1)
        If grp = "A" Or grp = "B" Or grp = "C" Or grp = "D" Then
            hsg(grp) = hsg(grp) + Val(CStr(row(2)))
            tot = tot + Val(CStr(row(2)))
        End If
    Next row
    If tot <= 0# Then Err.Raise vbObjectError + 31, , "no soil groups returned for this basin"
    s = ""
    For Each k In hsg.keys
        hsg(k) = hsg(k) / tot
    Next k
    For Each k In Array("A", "B", "C", "D")
        If hsg.Exists(k) Then s = s & IIf(Len(s) > 0, ", ", "") & k & " " & Format$(hsg(k) * 100#, "0") & "%"
    Next k
    SoilGroups = True
    LogLine "  Soil groups (USDA SSURGO): " & s & "."
    Exit Function
NoSoil:
    Set hsg = Nothing
    If Not gCancel Then mSoilDownUntil = Now + 10# / 1440#
    LogLine "  [WARNING] USDA SSURGO soils unavailable (" & Err.Description & "); trying the UC Davis SoilWeb soil grid."
End Function

'----------------------------------------------------------------------------------------------
' Self-test of the equations (compare with the Python tool: 344.4, 1688, 65.5)
'----------------------------------------------------------------------------------------------
Public Sub StreamStatsSelfTestMethods()
    Dim classes As Object, hsg As Object, t As String
    Set classes = CreateObject("Scripting.Dictionary")
    Set hsg = CreateObject("Scripting.Dictionary")
    classes(41) = 0.5
    classes(81) = 0.3
    classes(22) = 0.2
    hsg("B") = 0.6
    hsg("C") = 0.4
    MethodDefaults
    LogLine "Methods self-test: NRCS " & NumFix(NrcsQ(6#, 75#, 250# / 640#, 1.53 * 60#), 1) & _
            " | Urban Q10 " & NumFix(UrbanQ(10, 10#, 40#), 0) & " | CN " & NumFix(EstimateCN(classes, hsg, "test", t), 1) & " (" & t & ")"
End Sub
