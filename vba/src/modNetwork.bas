Attribute VB_Name = "modNetwork"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology (VBA) - stream network tools
' NHD streams clipped to the fence, longest flowpath, raindrop trace, backup data sources
' (EPA StreamCat, USGS 3DEP slope, UC Davis soil grid), the Anderson method and the service check.
' Points are Collections (1 = longitude, 2 = latitude); lines are Collections of points.
'==============================================================================================
Private Const URL_NHD_HR As String = "https://hydro.nationalmap.gov/arcgis/rest/services/NHDPlus_HR/MapServer/3/query"
Private Const URL_NHD_STD As String = "https://hydro.nationalmap.gov/arcgis/rest/services/nhd/MapServer/6/query"
Private Const URL_LFP As String = "https://streamstats.usgs.gov/ss-hydro/v1/additional-features/calculate"
Private Const URL_NLDI As String = "https://api.water.usgs.gov/nldi/linked-data"
Private Const URL_FLOWTRACE As String = "https://api.water.usgs.gov/nldi/pygeoapi/processes/nldi-flowtrace/execution"
Private Const URL_DEM As String = "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/ImageServer/getSamples"
Private Const URL_STREAMCAT As String = "https://api.epa.gov/StreamCat/streams/metrics"
Private Const URL_ISSR800 As String = "https://casoilresource.lawr.ucdavis.edu/cgi-bin/mapserv"
Private Const USER_AGENT As String = "MicroStation-USGS-VDOT-Automation/4.0"

Public gComputeLfp As Boolean        ' Settings: compute the longest flowpath
Public gUseAnderson As Boolean       ' Settings: Anderson method (off by default, as in the VDOT manual)
Public gAndersonType As String       ' natural, developed or sewered
Public gAutoGis As Boolean           ' Settings: save KML and Shapefiles beside the DGN after each run
Public gStreams As Collection        ' NHD streams clipped to the fence (Nothing until loaded)
Public gFence As Collection          ' fence outline, longitude / latitude, closed

Private mLfpDownUntil As Date
Private mLfpCache As Collection       ' last flowpath traced from NHD, reused for another culvert at the same outlet
Private mLfpCacheLon As Double, mLfpCacheLat As Double, mLfpCacheN As Long

Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal ms As Long)
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef dest As Any, ByRef src As Any, ByVal n As LongPtr)

'----------------------------------------------------------------------------------------------
' Geometry helpers
'----------------------------------------------------------------------------------------------
Public Sub PauseMs(ByVal ms As Long)
    Sleep ms
End Sub

Public Function Pt(ByVal lon As Double, ByVal lat As Double) As Collection
    Dim c As New Collection
    c.Add lon
    c.Add lat
    Set Pt = c
End Function

Public Function PolyXY(ByRef poly As Variant, ByRef x() As Double, ByRef y() As Double) As Long
    ' Copies a line (Collection of points) into arrays; returns the number of points.
    Dim n As Long, i As Long, p As Variant
    n = poly.Count
    If n = 0 Then Exit Function
    ReDim x(0 To n - 1)
    ReDim y(0 To n - 1)
    For Each p In poly
        x(i) = CDbl(p(1))
        y(i) = CDbl(p(2))
        i = i + 1
    Next p
    PolyXY = n
End Function

Public Function FtDist(ByVal lon1 As Double, ByVal lat1 As Double, ByVal lon2 As Double, ByVal lat2 As Double) As Double
    Dim dx As Double, dy As Double
    dx = (lon2 - lon1) * 364000# * Cos(0.5 * (lat1 + lat2) * 1.74532925199433E-02)
    dy = (lat2 - lat1) * 364000#
    FtDist = Sqr(dx * dx + dy * dy)
End Function

Public Function LineLenFt(ByRef poly As Variant) As Double
    Dim x() As Double, y() As Double, n As Long, i As Long
    n = PolyXY(poly, x, y)
    For i = 0 To n - 2
        LineLenFt = LineLenFt + FtDist(x(i), y(i), x(i + 1), y(i + 1))
    Next i
End Function

Private Function ASin(ByVal v As Double) As Double
    If v >= 1# Then
        ASin = 1.5707963267949
    ElseIf v <= -1# Then
        ASin = -1.5707963267949
    Else
        ASin = Atn(v / Sqr(1# - v * v))
    End If
End Function

Public Function GeoLengthFt(ByRef poly As Variant) As Double
    ' Great-circle length (ft), the same formula the Python tool uses for raindrop paths.
    Dim x() As Double, y() As Double, n As Long, i As Long, p1 As Double, p2 As Double, dp As Double, dl As Double, a As Double
    Const D2R As Double = 1.74532925199433E-02
    n = PolyXY(poly, x, y)
    For i = 0 To n - 2
        p1 = y(i) * D2R
        p2 = y(i + 1) * D2R
        dp = p2 - p1
        dl = (x(i + 1) - x(i)) * D2R
        a = Sin(dp / 2#) ^ 2 + Cos(p1) * Cos(p2) * Sin(dl / 2#) ^ 2
        GeoLengthFt = GeoLengthFt + 2# * 20902231# * ASin(Sqr(a))
    Next i
End Function

Public Function InRing(ByVal x As Double, ByVal y As Double, ByRef px() As Double, ByRef py() As Double, ByVal pn As Long) As Boolean
    ' Even-odd test; px / py is a closed ring (first point repeated at the end).
    Dim j As Long, hit As Boolean
    For j = 0 To pn - 2
        If (py(j) > y) <> (py(j + 1) > y) Then
            If x < px(j) + (y - py(j)) * (px(j + 1) - px(j)) / (py(j + 1) - py(j)) Then hit = Not hit
        End If
    Next j
    InRing = hit
End Function

Public Function ClipPolyline(ByRef lx() As Double, ByRef ly() As Double, ByVal ln As Long, _
                             ByRef px() As Double, ByRef py() As Double, ByVal pn As Long) As Collection
    ' Parts of a line that lie inside a closed ring. Returns a Collection of lines.
    Dim out As New Collection, chain As Collection
    Dim i As Long, j As Long, k As Long, m As Long, nt As Long, t() As Double, tv As Double
    Dim x1 As Double, y1 As Double, x2 As Double, y2 As Double, rx As Double, ry As Double
    Dim sx As Double, sy As Double, rxs As Double, qx As Double, qy As Double, tt As Double, uu As Double
    Dim bx0 As Double, by0 As Double, bx1 As Double, by1 As Double
    Dim ta As Double, tb As Double, tm As Double, ax As Double, ay As Double, bx As Double, by As Double
    Dim lastX As Double, lastY As Double, haveChain As Boolean
    Set ClipPolyline = out
    If ln < 2 Or pn < 4 Then Exit Function
    bx0 = px(0)
    bx1 = px(0)
    by0 = py(0)
    by1 = py(0)
    For j = 1 To pn - 1
        If px(j) < bx0 Then bx0 = px(j)
        If px(j) > bx1 Then bx1 = px(j)
        If py(j) < by0 Then by0 = py(j)
        If py(j) > by1 Then by1 = py(j)
    Next j
    ReDim t(0 To 63)
    For i = 0 To ln - 2
        x1 = lx(i)
        y1 = ly(i)
        x2 = lx(i + 1)
        y2 = ly(i + 1)
        If (x1 < bx0 And x2 < bx0) Or (x1 > bx1 And x2 > bx1) Or (y1 < by0 And y2 < by0) Or (y1 > by1 And y2 > by1) Then GoTo NextSeg
        rx = x2 - x1
        ry = y2 - y1
        t(0) = 0#
        t(1) = 1#
        nt = 2
        For j = 0 To pn - 2
            sx = px(j + 1) - px(j)
            sy = py(j + 1) - py(j)
            rxs = rx * sy - ry * sx
            If Abs(rxs) >= 0.00000000000001 Then
                qx = px(j) - x1
                qy = py(j) - y1
                tt = (qx * sy - qy * sx) / rxs
                If tt >= 0# And tt <= 1# Then
                    uu = (qx * ry - qy * rx) / rxs
                    If uu >= 0# And uu <= 1# Then
                        If nt > UBound(t) Then ReDim Preserve t(0 To 2 * nt)
                        t(nt) = Round(tt, 9)
                        nt = nt + 1
                    End If
                End If
            End If
        Next j
        For k = 1 To nt - 1
            tv = t(k)
            m = k - 1
            Do While m >= 0
                If t(m) <= tv Then Exit Do
                t(m + 1) = t(m)
                m = m - 1
            Loop
            t(m + 1) = tv
        Next k
        For k = 0 To nt - 2
            ta = t(k)
            tb = t(k + 1)
            If tb - ta >= 0.00000001 Then
                tm = 0.5 * (ta + tb)
                If InRing(x1 + tm * rx, y1 + tm * ry, px, py, pn) Then
                    ax = x1 + ta * rx
                    ay = y1 + ta * ry
                    bx = x1 + tb * rx
                    by = y1 + tb * ry
                    If haveChain And Abs(ax - lastX) < 0.00000001 And Abs(ay - lastY) < 0.00000001 Then
                        chain.Add Pt(bx, by)
                    Else
                        Set chain = New Collection
                        chain.Add Pt(ax, ay)
                        chain.Add Pt(bx, by)
                        out.Add chain
                        haveChain = True
                    End If
                    lastX = bx
                    lastY = by
                End If
            End If
        Next k
NextSeg:
    Next i
End Function

Public Function SimplifyRing(ByRef x() As Double, ByRef y() As Double, ByVal n As Long, ByVal maxPts As Long, _
                             ByRef ox() As Double, ByRef oy() As Double) As Long
    ' Douglas-Peucker with a growing tolerance until the ring has at most maxPts points. Returns the count.
    Dim keep() As Boolean, stA() As Long, stB() As Long, sp As Long, a As Long, b As Long, i As Long
    Dim dx As Double, dy As Double, norm As Double, d As Double, dmax As Double, idx As Long, tol As Double
    Dim xmin As Double, xmax As Double, ymin As Double, ymax As Double, cnt As Long, k As Long
    If n <= maxPts Then
        ReDim ox(0 To n - 1)
        ReDim oy(0 To n - 1)
        For i = 0 To n - 1
            ox(i) = x(i)
            oy(i) = y(i)
        Next i
        SimplifyRing = n
        Exit Function
    End If
    xmin = x(0)
    xmax = x(0)
    ymin = y(0)
    ymax = y(0)
    For i = 1 To n - 1
        If x(i) < xmin Then xmin = x(i)
        If x(i) > xmax Then xmax = x(i)
        If y(i) < ymin Then ymin = y(i)
        If y(i) > ymax Then ymax = y(i)
    Next i
    tol = (xmax - xmin)
    If (ymax - ymin) > tol Then tol = ymax - ymin
    tol = tol * 0.00001
    ReDim stA(0 To n)
    ReDim stB(0 To n)
    Do
        ReDim keep(0 To n - 1)
        keep(0) = True
        keep(n - 1) = True
        sp = 1
        stA(0) = 0
        stB(0) = n - 1
        Do While sp > 0
            sp = sp - 1
            a = stA(sp)
            b = stB(sp)
            dx = x(b) - x(a)
            dy = y(b) - y(a)
            norm = Sqr(dx * dx + dy * dy)
            idx = -1
            dmax = tol
            For i = a + 1 To b - 1
                If norm > 0.000000000001 Then
                    d = Abs(dy * (x(i) - x(a)) - dx * (y(i) - y(a))) / norm
                Else
                    d = Sqr((x(i) - x(a)) ^ 2 + (y(i) - y(a)) ^ 2)
                End If
                If d > dmax Then
                    idx = i
                    dmax = d
                End If
            Next i
            If idx > 0 Then
                keep(idx) = True
                stA(sp) = a
                stB(sp) = idx
                sp = sp + 1
                stA(sp) = idx
                stB(sp) = b
                sp = sp + 1
            End If
        Loop
        cnt = 0
        For i = 0 To n - 1
            If keep(i) Then cnt = cnt + 1
        Next i
        tol = tol * 2#
    Loop While cnt > maxPts
    ReDim ox(0 To cnt)
    ReDim oy(0 To cnt)
    k = 0
    For i = 0 To n - 1
        If keep(i) Then
            ox(k) = x(i)
            oy(k) = y(i)
            k = k + 1
        End If
    Next i
    If ox(0) <> ox(k - 1) Or oy(0) <> oy(k - 1) Then
        ox(k) = ox(0)
        oy(k) = oy(0)
        k = k + 1
    End If
    SimplifyRing = k
End Function

Public Function UrlEnc(ByVal s As String) As String
    Dim i As Long, ch As String, o As String, c As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = Asc(ch)
        If (c >= 48 And c <= 57) Or (c >= 65 And c <= 90) Or (c >= 97 And c <= 122) Or ch = "-" Or ch = "_" Or ch = "." Or ch = "~" Then
            o = o & ch
        Else
            o = o & "%" & Right$("0" & Hex$(c), 2)
        End If
    Next i
    UrlEnc = o
End Function

Private Function LineParts(ByRef geom As Variant) As Collection
    ' GeoJSON LineString / MultiLineString -> Collection of coordinate lists
    Dim parts As New Collection, part As Variant, gtype As String
    Set LineParts = parts
    If Not HasKey(geom, "type") Then Exit Function
    If Not HasKey(geom, "coordinates") Then Exit Function
    If TypeName(geom("coordinates")) <> "Collection" Then Exit Function
    gtype = JStr(geom, "type")
    If gtype = "LineString" Then
        parts.Add geom("coordinates")
    ElseIf gtype = "MultiLineString" Then
        For Each part In geom("coordinates")
            parts.Add part
        Next part
    End If
End Function

Private Function GeomPoints(ByRef geom As Variant) As Collection
    ' One ordered point list from a LineString / MultiLineString, without repeated points.
    Dim out As New Collection, part As Variant, c As Variant, lx As Double, ly As Double, have As Boolean
    For Each part In LineParts(geom)
        For Each c In part
            If Not (have And CDbl(c(1)) = lx And CDbl(c(2)) = ly) Then
                lx = CDbl(c(1))
                ly = CDbl(c(2))
                out.Add Pt(lx, ly)
                have = True
            End If
        Next c
    Next part
    Set GeomPoints = out
End Function

'----------------------------------------------------------------------------------------------
' NHD streams inside the fence
'----------------------------------------------------------------------------------------------
Private Function NhdFeatures(ByVal url As String, ByVal q As String, ByVal timeoutSec As Long, ByRef feats As Collection, _
                             ByRef limitHit As Boolean, ByRef errText As String) As Boolean
    Dim txt As String, gj As Variant
    On Error GoTo Bad
    txt = HttpText("GET", url & q, "", timeoutSec, 2)
    JsonParse txt, gj
    If HasKey(gj, "error") Then Err.Raise vbObjectError + 40, , "the service returned an error"
    Set feats = New Collection
    If HasKey(gj, "features") Then
        If TypeName(gj("features")) = "Collection" Then Set feats = gj("features")
    End If
    If HasKey(gj, "exceededTransferLimit") Then
        If Not IsNull(gj("exceededTransferLimit")) Then limitHit = CBool(gj("exceededTransferLimit"))
    End If
    If HasKey(gj, "properties") Then
        If HasKey(gj("properties"), "exceededTransferLimit") Then
            If Not IsNull(gj("properties")("exceededTransferLimit")) Then limitHit = CBool(gj("properties")("exceededTransferLimit"))
        End If
    End If
    NhdFeatures = True
    Exit Function
Bad:
    errText = Err.Description
End Function

Private Function EnvelopeQuery(ByVal minx As Double, ByVal miny As Double, ByVal maxx As Double, ByVal maxy As Double, _
                               ByVal fields As String) As String
    EnvelopeQuery = "?geometry=" & NumFix(minx, 7) & "," & NumFix(miny, 7) & "," & NumFix(maxx, 7) & "," & NumFix(maxy, 7) & _
        "&geometryType=esriGeometryEnvelope&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=" & fields & _
        "&returnGeometry=true&outSR=4326&f=geojson"
End Function

Private Function LowVal(ByRef lp As Object, ByVal k As String) As Variant
    ' Property value, or Empty when missing / null
    LowVal = Empty
    If lp.Exists(k) Then
        If Not IsNull(lp(k)) And Not IsObject(lp(k)) Then LowVal = lp(k)
    End If
End Function

Public Function QueryClipStreams(ByRef fence As Collection) As Collection
    ' NHDPlus HR flowlines (medium-resolution NHD as backup) clipped to the fence polygon.
    Dim fx() As Double, fy() As Double, fn As Long, i As Long
    Dim minx As Double, miny As Double, maxx As Double, maxy As Double, q As String
    Dim feats As Collection, limitHit As Boolean, errText As String
    Dim feat As Variant, props As Variant, lp As Object, k As Variant, part As Variant, seg As Variant
    Dim lx() As Double, ly() As Double, ln As Long, parts As Collection, idx As Long
    Dim s As Object, nm As String, v As Variant, total As Double, sx() As Double, sy() As Double, sn As Long
    Dim out As New Collection, bySeq As Object, nhd As Object, lastSeg As Variant, lastPt As Variant
    Dim sorted As New Collection, j As Long, placed As Boolean, t As Object

    fn = PolyXY(fence, fx, fy)
    If fn < 4 Then Err.Raise vbObjectError + 41, , "the fence needs at least three points"
    minx = fx(0)
    maxx = fx(0)
    miny = fy(0)
    maxy = fy(0)
    For i = 1 To fn - 1
        minx = MinD(minx, fx(i))
        maxx = MaxD(maxx, fx(i))
        miny = MinD(miny, fy(i))
        maxy = MaxD(maxy, fy(i))
    Next i
    q = EnvelopeQuery(minx, miny, maxx, maxy, "*")
    LogLine "Querying USGS NHDPlus HR flowlines in the fence bounding box..."
    If Not NhdFeatures(URL_NHD_HR, q, 90, feats, limitHit, errText) Then
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
        LogLine "  NHDPlus HR unavailable (" & errText & "); falling back to medium-resolution NHD."
        If Not NhdFeatures(URL_NHD_STD, q, 90, feats, limitHit, errText) Then
            Err.Raise vbObjectError + 42, , "USGS NHD streams could not be read (" & errText & ")"
        End If
    End If
    If limitHit Then LogLine "  [WARNING] Server record limit reached - some flowlines may be missing. Use a smaller fence."

    Set bySeq = CreateObject("Scripting.Dictionary")
    idx = 0
    For Each feat In feats
        idx = idx + 1
        Set lp = CreateObject("Scripting.Dictionary")
        If HasKey(feat, "properties") Then
            If TypeName(feat("properties")) = "Dictionary" Then
                Set props = feat("properties")
                For Each k In props.keys
                    If IsObject(props(k)) Then
                        lp(LCase$(CStr(k))) = Null
                    Else
                        lp(LCase$(CStr(k))) = props(k)
                    End If
                Next k
            End If
        End If
        Set parts = New Collection
        If HasKey(feat, "geometry") Then
            For Each part In LineParts(feat("geometry"))
                ln = PolyXY(part, lx, ly)
                If ln >= 2 Then
                    For Each seg In ClipPolyline(lx, ly, ln, fx, fy, fn)
                        parts.Add seg
                    Next seg
                End If
            Next part
        End If
        If parts.Count > 0 Then
            Set s = CreateObject("Scripting.Dictionary")
            nm = ""
            v = LowVal(lp, "gnis_name")
            If Not IsEmpty(v) Then nm = Trim$(CStr(v))
            If Len(nm) = 0 Then nm = "Tributary #" & idx
            s("id") = idx
            s("name") = nm
            s("order") = 1
            v = LowVal(lp, "streamorde")
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) >= 1# Then s("order") = CLng(Int(CDbl(v)))
            End If
            s("da_sqkm") = 0#
            v = LowVal(lp, "totdasqkm")
            If IsNumeric(v) And Not IsEmpty(v) Then s("da_sqkm") = CDbl(v)
            total = 0#
            For Each part In parts
                sn = PolyXY(part, sx, sy)
                For i = 0 To sn - 2
                    total = total + FtDist(sx(i), sy(i), sx(i + 1), sy(i + 1))
                Next i
            Next part
            s("length_ft") = RoundTo(total, 1)
            Set s("parts") = parts
            Set lastSeg = parts(parts.Count)
            Set lastPt = lastSeg(lastSeg.Count)
            Set s("pour") = Pt(CDbl(lastPt(1)), CDbl(lastPt(2)))
            Set nhd = CreateObject("Scripting.Dictionary")
            For Each k In Array("nhdplusid", "permanent_identifier", "reachcode", "ftype", "fcode", "qema", "vema", _
                                "slope", "hydroseq", "dnhydroseq", "flowdir")
                nhd(CStr(k)) = LowVal(lp, CStr(k))
            Next k
            Set s("nhd") = nhd
            s("flows_into") = ""
            v = nhd("hydroseq")
            If IsNumeric(v) And Not IsEmpty(v) Then
                If CDbl(v) <> 0# Then bySeq(Format$(CDbl(v), "0")) = nm
            End If
            out.Add s
        End If
    Next feat
    For Each s In out
        v = s("nhd")("dnhydroseq")
        If IsNumeric(v) And Not IsEmpty(v) Then
            If CDbl(v) = 0# Then
                s("flows_into") = "Network outlet"
            ElseIf bySeq.Exists(Format$(CDbl(v), "0")) Then
                s("flows_into") = CStr(bySeq(Format$(CDbl(v), "0")))
            Else
                s("flows_into") = "Downstream reach outside fence"
            End If
        End If
    Next s
    ' largest drainage area first, then stream order, then length
    For Each s In out
        placed = False
        For j = 1 To sorted.Count
            Set t = sorted(j)
            If StreamRank(s, t) > 0 Then
                sorted.Add s, , j
                placed = True
                Exit For
            End If
        Next j
        If Not placed Then sorted.Add s
    Next s
    Set QueryClipStreams = sorted
End Function

Private Function StreamRank(ByRef a As Object, ByRef b As Object) As Long
    If a("da_sqkm") <> b("da_sqkm") Then
        StreamRank = IIf(a("da_sqkm") > b("da_sqkm"), 1, -1)
    ElseIf a("order") <> b("order") Then
        StreamRank = IIf(a("order") > b("order"), 1, -1)
    ElseIf a("length_ft") <> b("length_ft") Then
        StreamRank = IIf(a("length_ft") > b("length_ft"), 1, -1)
    End If
End Function

'----------------------------------------------------------------------------------------------
' Longest flowpath
'----------------------------------------------------------------------------------------------
Public Sub SetLfpDown(ByVal untilTime As Date)
    mLfpDownUntil = untilTime
End Sub

Public Function UsgsLongestFlowpath(ByRef wsText As String) As Collection
    ' Longest flowpath published by USGS StreamStats for this watershed; Nothing when there is none.
    Dim txt As String, d As Variant, lines As New Collection, ln As Variant, best As Collection, bestLen As Double, l As Double
    If Now < mLfpDownUntil Then
        LogLine "  USGS longest-flowpath service is down (retrying in " & Format$((mLfpDownUntil - Now) * 1440#, "0") & _
                " min); using the NHD-traced flowpath."
        Exit Function
    End If
    On Error GoTo Bad
    txt = HttpText("POST", URL_LFP, wsText, 120, 2)
    JsonParse txt, d
    WalkLfp d, False, lines
    For Each ln In lines
        If ln.Count >= 2 Then
            l = LineLenFt(ln)
            If l > bestLen Then
                bestLen = l
                Set best = ln
            End If
        End If
    Next ln
    If best Is Nothing Then
        LogLine "  USGS does not publish a longest flowpath for this region; tracing NHD flowlines instead."
    End If
    Set UsgsLongestFlowpath = best
    Exit Function
Bad:
    If InStr(Err.Description, "HTTP Error 5") > 0 Or InStr(Err.Description, "timed out") > 0 Then
        If Not gCancel Then mLfpDownUntil = Now + 10# / 1440#
        LogLine "  USGS longest-flowpath service is down (" & Err.Description & "); using the NHD-traced flowpath. USGS will be retried in 10 min."
    Else
        LogLine "  USGS longest flowpath unavailable (" & Err.Description & "); using the NHD-traced flowpath."
    End If
End Function

Private Sub WalkLfp(ByRef obj As Variant, ByVal tagged As Boolean, ByRef lines As Collection)
    Dim tag As Boolean, nm As String, k As Variant, v As Variant, part As Variant, c As Variant, ln As Collection
    If Not IsObject(obj) Then Exit Sub
    If obj Is Nothing Then Exit Sub
    If TypeName(obj) = "Dictionary" Then
        If obj.Exists("name") Then
            nm = JStr(obj, "name")
        Else
            nm = JStr(obj, "id")
        End If
        tag = tagged Or (InStr(Replace(LCase$(nm), "_", ""), "longestflowpath") > 0)
        If tag And HasKey(obj, "geometry") Then
            If HasKey(obj("geometry"), "type") Then
                For Each part In LineParts(obj("geometry"))
                    Set ln = New Collection
                    For Each c In part
                        ln.Add Pt(CDbl(c(1)), CDbl(c(2)))
                    Next c
                    lines.Add ln
                Next part
            End If
        End If
        For Each k In obj.keys
            If IsObject(obj(k)) Then WalkLfp obj(k), tag, lines
        Next k
    ElseIf TypeName(obj) = "Collection" Then
        For Each v In obj
            If IsObject(v) Then WalkLfp v, tagged, lines
        Next v
    End If
End Sub

Public Function NhdLongestFlowpath(ByRef ring As Variant, ByVal outLon As Double, ByVal outLat As Double) As Collection
    ' Longest NHD channel path to the outlet, extended overland to the divide vertex that drains to its head.
    Dim rx() As Double, ry() As Double, rn As Long, cx() As Double, cy() As Double, cn As Long
    Dim minx As Double, miny As Double, maxx As Double, maxy As Double, i As Long
    Dim feats As Collection, limitHit As Boolean, errText As String
    Dim feat As Variant, part As Variant, seg As Variant, lx() As Double, ly() As Double, ln As Long
    Dim sx() As Double, sy() As Double, sn As Long
    Dim ids As Object, nx() As Double, ny() As Double, nn As Long
    Dim eA() As Long, eB() As Long, eW() As Double, ne As Long
    Dim head() As Long, nxt() As Long, eTo() As Long, eWt() As Double, na As Long
    Dim a As Long, b As Long, ka As String, kb As String, d As Double
    Dim dist() As Double, prev() As Long, startN As Long, best As Double
    Dim hk() As Double, hn() As Long, hs As Long, u As Long, du As Double, e As Long, nd As Long
    Dim headN As Long, out As New Collection, path() As Long, pc As Long, ov As Long
    Dim far As Double, farI As Long, channel As Double, attempt As Long, t0 As Single, gotNhd As Boolean, w As Long

    rn = PolyXY(ring, rx, ry)
    If rn < 4 Then Exit Function
    If Not mLfpCache Is Nothing Then
        If mLfpCacheN = rn And FtDist(mLfpCacheLon, mLfpCacheLat, outLon, outLat) < 100# Then
            LogLine "  Same outlet and basin as the previous culvert; reusing its NHD-traced flowpath."
            Set NhdLongestFlowpath = mLfpCache
            Exit Function
        End If
    End If
    cn = SimplifyRing(rx, ry, rn, 600, cx, cy)
    minx = rx(0)
    maxx = rx(0)
    miny = ry(0)
    maxy = ry(0)
    far = -1#
    For i = 0 To rn - 1
        minx = MinD(minx, rx(i))
        maxx = MaxD(maxx, rx(i))
        miny = MinD(miny, ry(i))
        maxy = MaxD(maxy, ry(i))
        d = FtDist(rx(i), ry(i), outLon, outLat)
        If d > far Then
            far = d
            farI = i
        End If
    Next i
    LogLine "  Tracing longest NHD flowpath inside the basin..."
    ' The USGS NHD server often answers on a second or third try, so a failed query is repeated before
    ' falling back to a straight line (which is only a rough stand-in for the real flowpath).
    For attempt = 1 To 3
        t0 = Timer
        gotNhd = NhdFeatures(URL_NHD_STD, EnvelopeQuery(minx, miny, maxx, maxy, "OBJECTID"), 120, feats, limitHit, errText)
        If gotNhd Or gCancel Then Exit For
        If attempt = 3 Or Timer - t0 > 60# Or Timer < t0 Then Exit For
        LogLine "  NHD flowline query failed (" & errText & "); trying again (" & (attempt + 1) & " of 3)..."
        For w = 1 To 50
            DoEvents
            Sleep 100
            If gCancel Then Exit For
        Next w
    Next attempt
    If Not gotNhd Then
        LogLine "  [WARNING] NHD flowline query failed (" & errText & "); using a straight divide-to-outlet line instead of the real flowpath."
        Set feats = New Collection
    End If
    If limitHit Then LogLine "  [WARNING] NHD record limit reached; flowpath may be incomplete."

    Set ids = CreateObject("Scripting.Dictionary")
    ReDim nx(0 To 1023)
    ReDim ny(0 To 1023)
    ReDim eA(0 To 1023)
    ReDim eB(0 To 1023)
    ReDim eW(0 To 1023)
    For Each feat In feats
        If HasKey(feat, "geometry") Then
            For Each part In LineParts(feat("geometry"))
                ln = PolyXY(part, lx, ly)
                If ln >= 2 Then
                    For Each seg In ClipPolyline(lx, ly, ln, cx, cy, cn)
                        sn = PolyXY(seg, sx, sy)
                        a = -1
                        For i = 0 To sn - 1
                            kb = Format$(sx(i) * 1000000#, "0") & "," & Format$(sy(i) * 1000000#, "0")
                            If ids.Exists(kb) Then
                                b = ids(kb)
                            Else
                                b = nn
                                ids(kb) = b
                                If nn > UBound(nx) Then
                                    ReDim Preserve nx(0 To 2 * nn + 1)
                                    ReDim Preserve ny(0 To 2 * nn + 1)
                                End If
                                nx(nn) = CDbl(Format$(sx(i) * 1000000#, "0")) / 1000000#
                                ny(nn) = CDbl(Format$(sy(i) * 1000000#, "0")) / 1000000#
                                nn = nn + 1
                            End If
                            If a >= 0 And a <> b Then
                                If ne > UBound(eA) Then
                                    ReDim Preserve eA(0 To 2 * ne + 1)
                                    ReDim Preserve eB(0 To 2 * ne + 1)
                                    ReDim Preserve eW(0 To 2 * ne + 1)
                                End If
                                eA(ne) = a
                                eB(ne) = b
                                eW(ne) = FtDist(nx(a), ny(a), nx(b), ny(b))
                                ne = ne + 1
                            End If
                            a = b
                        Next i
                    Next seg
                End If
            Next part
        End If
    Next feat

    If nn = 0 Or ne = 0 Then
        out.Add Pt(rx(farI), ry(farI))
        out.Add Pt(outLon, outLat)
        Set NhdLongestFlowpath = out
        Exit Function
    End If

    ' adjacency lists (both directions)
    ReDim head(0 To nn - 1)
    For i = 0 To nn - 1
        head(i) = -1
    Next i
    ReDim nxt(0 To 2 * ne - 1)
    ReDim eTo(0 To 2 * ne - 1)
    ReDim eWt(0 To 2 * ne - 1)
    For i = 0 To ne - 1
        eTo(na) = eB(i)
        eWt(na) = eW(i)
        nxt(na) = head(eA(i))
        head(eA(i)) = na
        na = na + 1
        eTo(na) = eA(i)
        eWt(na) = eW(i)
        nxt(na) = head(eB(i))
        head(eB(i)) = na
        na = na + 1
    Next i
    best = 1E+300
    For i = 0 To nn - 1
        d = FtDist(nx(i), ny(i), outLon, outLat)
        If d < best Then
            best = d
            startN = i
        End If
    Next i
    ' Dijkstra from the node nearest the outlet (binary heap, stale entries skipped)
    ReDim dist(0 To nn - 1)
    ReDim prev(0 To nn - 1)
    For i = 0 To nn - 1
        dist(i) = 1E+300
        prev(i) = -1
    Next i
    ReDim hk(0 To 2 * ne + nn + 8)
    ReDim hn(0 To 2 * ne + nn + 8)
    dist(startN) = 0#
    HeapPush hk, hn, hs, 0#, startN
    Do While hs > 0
        HeapPop hk, hn, hs, du, u
        If du <= dist(u) Then
            e = head(u)
            Do While e >= 0
                nd = eTo(e)
                If du + eWt(e) < dist(nd) Then
                    dist(nd) = du + eWt(e)
                    prev(nd) = u
                    HeapPush hk, hn, hs, dist(nd), nd
                End If
                e = nxt(e)
            Loop
        End If
    Loop
    headN = startN
    channel = 0#
    For i = 0 To nn - 1
        If dist(i) < 1E+299 And dist(i) > channel Then
            channel = dist(i)
            headN = i
        End If
    Next i
    ReDim path(0 To nn)
    u = headN
    Do While u >= 0
        path(pc) = u
        pc = pc + 1
        If u = startN Then Exit Do
        u = prev(u)
    Loop
    ov = OverlandStart(cx, cy, cn, headN, nx, ny, nn, eA, eB, ne)
    If ov >= 0 Then out.Add Pt(cx(ov), cy(ov))
    For i = 0 To pc - 1
        out.Add Pt(nx(path(i)), ny(path(i)))
    Next i
    If FtDist(nx(path(pc - 1)), ny(path(pc - 1)), outLon, outLat) > 1# Then out.Add Pt(outLon, outLat)
    If far > LineLenFt(out) Then
        LogLine "  NHD channel (" & Format$(channel, "#,##0") & " ft) shorter than divide distance; using divide-to-outlet line."
    End If
    Set NhdLongestFlowpath = out
    Set mLfpCache = out
    mLfpCacheLon = outLon
    mLfpCacheLat = outLat
    mLfpCacheN = rn
End Function

Private Sub HeapPush(ByRef hk() As Double, ByRef hn() As Long, ByRef hs As Long, ByVal key As Double, ByVal node As Long)
    Dim i As Long, p As Long
    If hs > UBound(hk) Then
        ReDim Preserve hk(0 To 2 * hs + 1)
        ReDim Preserve hn(0 To 2 * hs + 1)
    End If
    i = hs
    hs = hs + 1
    Do While i > 0
        p = (i - 1) \ 2
        If hk(p) <= key Then Exit Do
        hk(i) = hk(p)
        hn(i) = hn(p)
        i = p
    Loop
    hk(i) = key
    hn(i) = node
End Sub

Private Sub HeapPop(ByRef hk() As Double, ByRef hn() As Long, ByRef hs As Long, ByRef key As Double, ByRef node As Long)
    Dim i As Long, c As Long, lk As Double, lnode As Long
    key = hk(0)
    node = hn(0)
    hs = hs - 1
    If hs = 0 Then Exit Sub
    lk = hk(hs)
    lnode = hn(hs)
    i = 0
    Do
        c = 2 * i + 1
        If c >= hs Then Exit Do
        If c + 1 < hs Then
            If hk(c + 1) < hk(c) Then c = c + 1
        End If
        If hk(c) >= lk Then Exit Do
        hk(i) = hk(c)
        hn(i) = hn(c)
        i = c
    Loop
    hk(i) = lk
    hn(i) = lnode
End Sub

Private Function SegCross(ByVal p1x As Double, ByVal p1y As Double, ByVal p2x As Double, ByVal p2y As Double, _
                          ByVal ax As Double, ByVal ay As Double, ByVal bx As Double, ByVal by As Double) As Boolean
    ' True when the two segments properly cross (touching at an end point does not count).
    Dim s1 As Double, s2 As Double, s3 As Double, s4 As Double
    s1 = (p2x - p1x) * (ay - p1y) - (p2y - p1y) * (ax - p1x)
    s2 = (p2x - p1x) * (by - p1y) - (p2y - p1y) * (bx - p1x)
    If s1 * s2 >= 0# Then Exit Function
    s3 = (bx - ax) * (p1y - ay) - (by - ay) * (p1x - ax)
    s4 = (bx - ax) * (p2y - ay) - (by - ay) * (p2x - ax)
    SegCross = (s3 * s4 < 0#)
End Function

Private Function OverlandStart(ByRef cx() As Double, ByRef cy() As Double, ByVal cn As Long, ByVal headN As Long, _
                               ByRef nx() As Double, ByRef ny() As Double, ByVal nn As Long, _
                               ByRef eA() As Long, ByRef eB() As Long, ByVal ne As Long) As Long
    ' Divide vertex where the overland leg starts: the farthest one that drains to the channel head - no other
    ' stream point is closer to it than the head, and the straight leg crosses neither a stream nor the
    ' basin boundary. -1 when there is no such vertex.
    Dim k As Double, hx As Double, hy As Double, i As Long, j As Long, px As Double, py As Double
    Dim D2 As Double, limit As Double, ok As Boolean, dx As Double, dy As Double
    Dim cand() As Long, cd() As Double, nc As Long, t As Long, td As Double, m As Long
    Dim fx() As Double, fy() As Double, gx() As Double, gy() As Double
    OverlandStart = -1
    k = 364000# * Cos(ny(headN) * 1.74532925199433E-02)
    hx = nx(headN) * k
    hy = ny(headN) * 364000#
    ReDim fx(0 To nn - 1)
    ReDim fy(0 To nn - 1)
    For i = 0 To nn - 1
        fx(i) = nx(i) * k
        fy(i) = ny(i) * 364000#
    Next i
    ReDim gx(0 To cn - 1)
    ReDim gy(0 To cn - 1)
    For i = 0 To cn - 1
        gx(i) = cx(i) * k
        gy(i) = cy(i) * 364000#
    Next i
    ReDim cand(0 To cn - 1)
    ReDim cd(0 To cn - 1)
    For i = 0 To cn - 1
        px = gx(i)
        py = gy(i)
        D2 = (px - hx) ^ 2 + (py - hy) ^ 2
        If D2 >= 1# Then
            limit = D2 - 1#
            ok = True
            For j = 0 To nn - 1
                If j <> headN Then
                    dx = px - fx(j)
                    dy = py - fy(j)
                    If dx * dx + dy * dy < limit Then
                        ok = False
                        Exit For
                    End If
                End If
            Next j
            If ok Then
                cand(nc) = i
                cd(nc) = D2
                nc = nc + 1
            End If
        End If
    Next i
    ' farthest first
    For i = 1 To nc - 1
        t = cand(i)
        td = cd(i)
        m = i - 1
        Do While m >= 0
            If cd(m) >= td Then Exit Do
            cand(m + 1) = cand(m)
            cd(m + 1) = cd(m)
            m = m - 1
        Loop
        cand(m + 1) = t
        cd(m + 1) = td
    Next i
    For i = 0 To nc - 1
        px = gx(cand(i))
        py = gy(cand(i))
        ok = True
        For j = 0 To ne - 1
            If SegCross(px, py, hx, hy, fx(eA(j)), fy(eA(j)), fx(eB(j)), fy(eB(j))) Then
                ok = False
                Exit For
            End If
        Next j
        If ok Then
            For j = 0 To cn - 2
                If Not ((gx(j) = px And gy(j) = py) Or (gx(j + 1) = px And gy(j + 1) = py)) Then
                    If SegCross(px, py, hx, hy, gx(j), gy(j), gx(j + 1), gy(j + 1)) Then
                        ok = False
                        Exit For
                    End If
                End If
            Next j
        End If
        If ok Then
            OverlandStart = cand(i)
            Exit Function
        End If
    Next i
End Function

Public Function LargestRing(ByRef rings As Collection) As Variant
    Dim r As Variant, best As Variant, n As Long
    For Each r In rings
        If r.Count > n Then
            n = r.Count
            Set best = r
        End If
    Next r
    Set LargestRing = best
End Function

'----------------------------------------------------------------------------------------------
' Backup basin data: EPA StreamCat land cover and USGS 3DEP channel slope
'----------------------------------------------------------------------------------------------
Public Function StreamCatCover(ByVal lon As Double, ByVal lat As Double, ByRef impPct As Double, ByRef forest As Double, _
                               ByRef comid As String) As Boolean
    Dim txt As String, d As Variant, item As Variant, props As Variant
    On Error GoTo Bad
    txt = HttpText("GET", URL_NLDI & "/comid/position?coords=" & UrlEnc("POINT(" & NumFix(lon, 6) & " " & NumFix(lat, 6) & ")"), "", 45, 2)
    JsonParse txt, d
    Set props = d("features")(1)("properties")
    comid = JStr(props, "comid")
    If Len(comid) = 0 Then comid = JStr(props, "identifier")
    txt = HttpText("GET", URL_STREAMCAT & "?name=pctimp2019,pctdecid2019,pctconif2019,pctmxfst2019&comid=" & comid & _
                   "&areaOfInterest=ws", "", 45, 2)
    JsonParse txt, d
    Set item = d("items")(1)
    forest = CDbl(item("pctdecid2019ws")) + CDbl(item("pctconif2019ws")) + CDbl(item("pctmxfst2019ws"))
    impPct = CDbl(item("pctimp2019ws"))
    StreamCatCover = True
    Exit Function
Bad:
    LogLine "  [WARNING] EPA StreamCat lookup failed (" & Err.Description & "); using default impervious 2% and forest 50%."
End Function

Public Function ChannelSlope1085(ByRef lfp As Collection, ByVal outLon As Double, ByVal outLat As Double, _
                                 ByRef slope As Double, ByRef lengthFt As Double) As Boolean
    ' Channel slope (ft/mi) between 10% and 85% of the flowpath length above the outlet, from USGS 3DEP.
    Dim x() As Double, y() As Double, n As Long, i As Long, t As Double, total As Double
    Dim ax As Double, ay As Double, bx As Double, by As Double, txt As String, d As Variant, s As Variant
    Dim e0 As Double, e1 As Double, have0 As Boolean, have1 As Boolean, g As String
    On Error GoTo Bad
    n = PolyXY(lfp, x, y)
    If n < 2 Then Err.Raise vbObjectError + 43, , "no flowpath"
    If FtDist(x(0), y(0), outLon, outLat) > FtDist(x(n - 1), y(n - 1), outLon, outLat) Then
        For i = 0 To (n \ 2) - 1
            t = x(i)
            x(i) = x(n - 1 - i)
            x(n - 1 - i) = t
            t = y(i)
            y(i) = y(n - 1 - i)
            y(n - 1 - i) = t
        Next i
    End If
    For i = 0 To n - 2
        total = total + FtDist(x(i), y(i), x(i + 1), y(i + 1))
    Next i
    If total <= 0# Then Err.Raise vbObjectError + 43, , "zero-length flowpath"
    PointAlong x, y, n, 0.1 * total, ax, ay
    PointAlong x, y, n, 0.85 * total, bx, by
    g = "{""points"":[[" & NumFix(ax, 7) & "," & NumFix(ay, 7) & "],[" & NumFix(bx, 7) & "," & NumFix(by, 7) & _
        "]],""spatialReference"":{""wkid"":4326}}"
    txt = HttpText("GET", URL_DEM & "?geometry=" & UrlEnc(g) & "&geometryType=esriGeometryMultipoint&returnFirstValueOnly=true&f=json", "", 90, 2)
    JsonParse txt, d
    For Each s In d("samples")
        If CLng(s("locationId")) = 0 Then
            e0 = CDbl(Val(CStr(s("value")))) * 3.28084
            have0 = True
        ElseIf CLng(s("locationId")) = 1 Then
            e1 = CDbl(Val(CStr(s("value")))) * 3.28084
            have1 = True
        End If
    Next s
    If Not (have0 And have1) Then Err.Raise vbObjectError + 43, , "no elevation returned"
    slope = Abs(e1 - e0) / (0.75 * total / 5280#)
    lengthFt = total
    ChannelSlope1085 = True
    Exit Function
Bad:
    LogLine "  [WARNING] Channel slope from USGS 3DEP failed (" & Err.Description & "); using the default 50 ft/mi."
End Function

Private Sub PointAlong(ByRef x() As Double, ByRef y() As Double, ByVal n As Long, ByVal target As Double, _
                       ByRef ox As Double, ByRef oy As Double)
    Dim i As Long, run As Double, seg As Double, t As Double
    For i = 0 To n - 2
        seg = FtDist(x(i), y(i), x(i + 1), y(i + 1))
        If seg > 0# And run + seg >= target Then
            t = (target - run) / seg
            ox = x(i) + t * (x(i + 1) - x(i))
            oy = y(i) + t * (y(i + 1) - y(i))
            Exit Sub
        End If
        run = run + seg
    Next i
    ox = x(n - 1)
    oy = y(n - 1)
End Sub

'----------------------------------------------------------------------------------------------
' Backup soil groups: UC Davis SoilWeb ISSR-800 grid
'----------------------------------------------------------------------------------------------
Public Function HttpBytes(ByVal url As String, ByVal timeoutSec As Long, ByRef data() As Byte) As Boolean
    ' Binary GET, polled so MicroStation keeps responding. Raises an error when the request fails.
    Dim h As Object, t0 As Single, st As Long
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.setTimeouts 15000, 15000, 30000, timeoutSec * 1000
    h.Open "GET", url, True
    h.setRequestHeader "User-Agent", USER_AGENT
    h.send
    t0 = Timer
    Do While Not h.waitForResponse(0.1)
        DoEvents
        If gCancel Then
            h.abort
            Err.Raise vbObjectError + 4, , "stopped by user"
        End If
        If Timer - t0 > timeoutSec Then
            h.abort
            Err.Raise vbObjectError + 10, , "The operation timed out"
        End If
    Loop
    st = h.Status
    If st < 200 Or st >= 300 Then Err.Raise vbObjectError + 10, , "HTTP Error " & st
    data = h.responseBody
    HttpBytes = True
End Function

Private Function U16(ByRef b() As Byte, ByVal off As Long, ByVal little As Boolean) As Long
    If little Then
        U16 = CLng(b(off)) + 256& * b(off + 1)
    Else
        U16 = CLng(b(off + 1)) + 256& * b(off)
    End If
End Function

Private Function U32(ByRef b() As Byte, ByVal off As Long, ByVal little As Boolean) As Double
    If little Then
        U32 = b(off) + 256# * b(off + 1) + 65536# * b(off + 2) + 16777216# * b(off + 3)
    Else
        U32 = b(off + 3) + 256# * b(off + 2) + 65536# * b(off + 1) + 16777216# * b(off)
    End If
End Function

Private Function F64(ByRef b() As Byte, ByVal off As Long, ByVal little As Boolean) As Double
    Dim t(0 To 7) As Byte, i As Long, d As Double
    For i = 0 To 7
        If little Then
            t(i) = b(off + i)
        Else
            t(i) = b(off + 7 - i)
        End If
    Next i
    CopyMemory d, t(0), 8
    F64 = d
End Function

Private Function TiffTag(ByRef b() As Byte, ByVal ifd As Long, ByVal little As Boolean, ByVal wanted As Long, _
                         ByRef typ As Long, ByRef cnt As Long, ByRef dataOff As Long) As Boolean
    ' Finds a TIFF tag; dataOff is where its values start.
    Dim n As Long, i As Long, e As Long, size As Long
    n = U16(b, ifd, little)
    For i = 0 To n - 1
        e = ifd + 2 + 12 * i
        If U16(b, e, little) = wanted Then
            typ = U16(b, e + 2, little)
            cnt = CLng(U32(b, e + 4, little))
            Select Case typ
                Case 1
                    size = 1
                Case 3
                    size = 2
                Case 4
                    size = 4
                Case 12
                    size = 8
                Case Else
                    Exit Function
            End Select
            If size * cnt <= 4 Then
                dataOff = e + 8
            Else
                dataOff = CLng(U32(b, e + 8, little))
            End If
            TiffTag = True
            Exit Function
        End If
    Next i
End Function

Private Function TiffVal(ByRef b() As Byte, ByVal little As Boolean, ByVal typ As Long, ByVal dataOff As Long, ByVal idx As Long) As Double
    Select Case typ
        Case 1
            TiffVal = b(dataOff + idx)
        Case 3
            TiffVal = U16(b, dataOff + 2 * idx, little)
        Case 4
            TiffVal = U32(b, dataOff + 4 * idx, little)
        Case 12
            TiffVal = F64(b, dataOff + 8 * idx, little)
    End Select
End Function

Public Function Issr800Groups(ByRef rings As Collection, ByRef hsg As Object) As Boolean
    ' Hydrologic soil group fractions sampled from the 800 m ISSR-800 grid inside the basin.
    ' Grid codes: 1 A, 2 A/D, 3 B, 4 B/D, 5 C, 6 C/D, 7 D; dual groups count as D.
    Dim rx() As Double, ry() As Double, rn As Long, i As Long
    Dim minx As Double, miny As Double, maxx As Double, maxy As Double, k As Double, stepD As Double
    Dim nxs As Long, nys As Long, crs As String, url As String, b() As Byte
    Dim little As Boolean, ifd As Long, typ As Long, cnt As Long, off As Long, typ2 As Long, cnt2 As Long, off2 As Long
    Dim gridW As Long, gridH As Long, pix() As Byte, np As Long, s As Long, so As Long, sc As Long, j As Long
    Dim tie(0 To 5) As Double, scl(0 To 2) As Double, west As Double, north As Double, dx As Double, dy As Double
    Dim r As Long, c As Long, v As Long, grp As String, counts(0 To 3) As Long, total As Long, y As Double
    Dim codes As Variant, names As Variant
    On Error GoTo Bad
    codes = Array(-1, 0, 3, 1, 3, 2, 3, 3)
    names = Array("A", "B", "C", "D")
    rn = PolyXY(LargestRing(rings), rx, ry)
    If rn < 3 Then Err.Raise vbObjectError + 44, , "no basin outline"
    minx = rx(0)
    maxx = rx(0)
    miny = ry(0)
    maxy = ry(0)
    For i = 1 To rn - 1
        minx = MinD(minx, rx(i))
        maxx = MaxD(maxx, rx(i))
        miny = MinD(miny, ry(i))
        maxy = MaxD(maxy, ry(i))
    Next i
    k = Cos(0.5 * (miny + maxy) * 1.74532925199433E-02)
    stepD = MaxD((maxx - minx) * k, maxy - miny) / 60#
    nxs = CLng((maxx - minx) * k / stepD)
    nys = CLng((maxy - miny) / stepD)
    If nxs < 1 Then nxs = 1
    If nys < 1 Then nys = 1
    crs = UrlEnc("http://www.opengis.net/def/crs/EPSG/0/4326")
    url = URL_ISSR800 & "?map=issr800-wcs&SERVICE=WCS&VERSION=2.0.1&REQUEST=GetCoverage&COVERAGEID=hydgrp&FORMAT=" & UrlEnc("image/tiff") & _
          "&SUBSETTINGCRS=" & crs & "&OUTPUTCRS=" & crs & _
          "&SUBSET=" & UrlEnc("x(" & NumFix(minx, 6) & "," & NumFix(maxx, 6) & ")") & _
          "&SUBSET=" & UrlEnc("y(" & NumFix(miny, 6) & "," & NumFix(maxy, 6) & ")") & _
          "&SCALESIZE=" & UrlEnc("x(" & nxs & "),y(" & nys & ")")
    HttpBytes url, 60, b
    If UBound(b) < 16 Then Err.Raise vbObjectError + 44, , "soil grid service did not return a TIFF image"
    If b(0) = 73 And b(1) = 73 Then
        little = True
    ElseIf b(0) = 77 And b(1) = 77 Then
        little = False
    Else
        Err.Raise vbObjectError + 44, , "soil grid service did not return a TIFF image"
    End If
    ifd = CLng(U32(b, 4, little))
    If Not TiffTag(b, ifd, little, 256, typ, cnt, off) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    gridW = CLng(TiffVal(b, little, typ, off, 0))
    If Not TiffTag(b, ifd, little, 257, typ, cnt, off) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    gridH = CLng(TiffVal(b, little, typ, off, 0))
    If TiffTag(b, ifd, little, 259, typ, cnt, off) Then
        If TiffVal(b, little, typ, off, 0) <> 1# Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    End If
    If TiffTag(b, ifd, little, 258, typ, cnt, off) Then
        If TiffVal(b, little, typ, off, 0) <> 8# Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    End If
    If Not TiffTag(b, ifd, little, 33922, typ, cnt, off) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    For i = 0 To 5
        tie(i) = TiffVal(b, little, typ, off, i)
    Next i
    If Not TiffTag(b, ifd, little, 33550, typ, cnt, off) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    For i = 0 To 1
        scl(i) = TiffVal(b, little, typ, off, i)
    Next i
    If Not TiffTag(b, ifd, little, 273, typ, cnt, off) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    If Not TiffTag(b, ifd, little, 279, typ2, cnt2, off2) Then Err.Raise vbObjectError + 44, , "unsupported soil grid image format"
    ReDim pix(0 To gridW * gridH - 1)
    For s = 0 To cnt - 1
        so = CLng(TiffVal(b, little, typ, off, s))
        sc = CLng(TiffVal(b, little, typ2, off2, s))
        For j = 0 To sc - 1
            If np <= UBound(pix) Then
                pix(np) = b(so + j)
                np = np + 1
            End If
        Next j
    Next s
    If np < gridW * gridH Then Err.Raise vbObjectError + 44, , "soil grid image is truncated"
    west = tie(3) - tie(0) * scl(0)
    north = tie(4) + tie(1) * scl(1)
    dx = scl(0)
    dy = scl(1)
    If rx(0) <> rx(rn - 1) Or ry(0) <> ry(rn - 1) Then
        ReDim Preserve rx(0 To rn)
        ReDim Preserve ry(0 To rn)
        rx(rn) = rx(0)
        ry(rn) = ry(0)
        rn = rn + 1
    End If
    For r = 0 To gridH - 1
        y = north - (r + 0.5) * dy
        For c = 0 To gridW - 1
            v = pix(r * gridW + c)
            If v >= 1 And v <= 7 Then
                If InRing(west + (c + 0.5) * dx, y, rx, ry, rn) Then
                    counts(codes(v)) = counts(codes(v)) + 1
                    total = total + 1
                End If
            End If
        Next c
    Next r
    If total = 0 Then      ' sliver basin: take the cell at its centre
        r = CLng(Int((north - 0.5 * (miny + maxy)) / dy))
        c = CLng(Int((0.5 * (minx + maxx) - west) / dx))
        If r < 0 Then r = 0
        If r > gridH - 1 Then r = gridH - 1
        If c < 0 Then c = 0
        If c > gridW - 1 Then c = gridW - 1
        v = pix(r * gridW + c)
        If v >= 1 And v <= 7 Then
            counts(codes(v)) = 1
            total = 1
        End If
    End If
    If total = 0 Then Err.Raise vbObjectError + 44, , "no soil group data for this basin"
    Set hsg = CreateObject("Scripting.Dictionary")
    For i = 0 To 3
        If counts(i) > 0 Then hsg(CStr(names(i))) = counts(i) / total
    Next i
    Issr800Groups = True
    Exit Function
Bad:
    Set hsg = Nothing
    LogLine "  [WARNING] UC Davis SoilWeb soil grid unavailable (" & Err.Description & "); soil properties are left blank" & _
            " and soil group B is assumed when the NRCS CN is estimated."
End Function

'----------------------------------------------------------------------------------------------
' Anderson method (USGS WSP 2001-C; VDOT eq. 6.6-6.8, Tables 6-5 and 6-6)
'----------------------------------------------------------------------------------------------
Public Function AndersonQ(ByVal yr As Long, ByVal sqmi As Double, ByVal lengthMi As Double, ByVal slope As Double, _
                          ByVal impPct As Double, ByVal basinType As String) As Double
    ' The 2.33-yr ratio is the mean annual flood, so there is no 2-yr or 200-yr value (-1).
    ' K = 1 + 0.015 I is used; the manual prints 0.15, which does not match Anderson's report.
    Dim coef As Double, ex As Double, lag As Double, k As Double, rn As Double, r100 As Double, rf As Double, q100 As Double
    AndersonQ = -1#
    If yr = 500 Then
        q100 = AndersonQ(100, sqmi, lengthMi, slope, impPct, basinType)
        If q100 > 0# Then AndersonQ = 1.7 * q100
        Exit Function
    End If
    Select Case yr
        Case 5
            rn = 1.65
            r100 = 1.24
        Case 10
            rn = 2.2
            r100 = 1.45
        Case 25
            rn = 3.3
            r100 = 1.8
        Case 50
            rn = 4.4
            r100 = 2#
        Case 100
            rn = 5.5
            r100 = 2.2
        Case Else
            Exit Function
    End Select
    If sqmi <= 0# Or lengthMi <= 0# Or slope <= 0# Then Exit Function
    Select Case basinType
        Case "developed"
            coef = 0.9
            ex = 0.5
        Case "sewered"
            coef = 0.56
            ex = 0.52
        Case Else
            coef = 4.64
            ex = 0.42
    End Select
    lag = coef * (lengthMi / Sqr(slope)) ^ ex
    If impPct < 0# Then impPct = 0#
    If impPct > 100# Then impPct = 100#
    k = 1# + 0.015 * impPct
    rf = (rn + 0.01 * impPct * (2.5 * r100 - rn)) / k
    AndersonQ = rf * 230# * k * sqmi ^ 0.82 * lag ^ -0.48
End Function

'----------------------------------------------------------------------------------------------
' Flow (raindrop) path
'----------------------------------------------------------------------------------------------
Private Sub ProjectOntoPath(ByVal lon As Double, ByVal lat As Double, ByRef path As Variant, _
                            ByRef segIdx As Long, ByRef ox As Double, ByRef oy As Double, ByRef D2 As Double)
    ' Closest point on a line; segIdx is the 0-based segment it falls on.
    Dim x() As Double, y() As Double, n As Long, i As Long, kx As Double, dx As Double, dy As Double
    Dim seg2 As Double, t As Double, px As Double, py As Double, d As Double
    n = PolyXY(path, x, y)
    kx = Cos(lat * 1.74532925199433E-02)
    D2 = 1E+300
    segIdx = 0
    ox = x(0)
    oy = y(0)
    For i = 0 To n - 2
        dx = (x(i + 1) - x(i)) * kx
        dy = y(i + 1) - y(i)
        seg2 = dx * dx + dy * dy
        t = 0#
        If seg2 > 0# Then
            t = (((lon - x(i)) * kx) * dx + (lat - y(i)) * dy) / seg2
            If t < 0# Then t = 0#
            If t > 1# Then t = 1#
        End If
        px = x(i) + (x(i + 1) - x(i)) * t
        py = y(i) + (y(i + 1) - y(i)) * t
        d = ((px - lon) * kx) ^ 2 + (py - lat) ^ 2
        If d < D2 Then
            D2 = d
            segIdx = i
            ox = px
            oy = py
        End If
    Next i
End Sub

Private Function NldiFlowtrace(ByVal lon As Double, ByVal lat As Double) As Collection
    ' USGS NLDI raindrop trace: the overland path from the point to the nearest NHD flowline.
    Dim body As String, txt As String, d As Variant, feat As Variant, fid As String, pts As Collection, pass As Long
    Dim rev As Collection, i As Long, first As Variant, last As Variant
    body = "{""inputs"":[{""id"":""lat"",""value"":""" & NumFix(lat, 7) & """,""type"":""text/plain""}," & _
           "{""id"":""lon"",""value"":""" & NumFix(lon, 7) & """,""type"":""text/plain""}," & _
           "{""id"":""raindroptrace"",""value"":""true"",""type"":""text/plain""}," & _
           "{""id"":""direction"",""value"":""down"",""type"":""text/plain""}]}"
    txt = HttpText("POST", URL_FLOWTRACE, body, 45, 1)
    JsonParse txt, d
    If Not HasKey(d, "features") Then Exit Function
    If TypeName(d("features")) <> "Collection" Then Exit Function
    For pass = 1 To 2       ' the feature named exactly 'raindropPath' first
        For Each feat In d("features")
            fid = LCase$(JStr(feat, "id"))
            If (pass = 1 And fid = "raindroppath") Or (pass = 2 And fid <> "raindroppath" And InStr(fid, "raindrop") > 0 And InStr(fid, "flowline") = 0) Then
                If HasKey(feat, "geometry") Then
                    Set pts = GeomPoints(feat("geometry"))
                    If pts.Count >= 2 Then
                        Set first = pts(1)
                        Set last = pts(pts.Count)
                        If Sqr((last(1) - lon) ^ 2 + (last(2) - lat) ^ 2) < Sqr((first(1) - lon) ^ 2 + (first(2) - lat) ^ 2) Then
                            Set rev = New Collection
                            For i = pts.Count To 1 Step -1
                                rev.Add pts(i)
                            Next i
                            Set pts = rev
                        End If
                        Set first = pts(1)
                        If first(1) <> lon Or first(2) <> lat Then pts.Add Pt(lon, lat), , 1
                        Set NldiFlowtrace = pts
                        Exit Function
                    End If
                End If
            End If
        Next feat
    Next pass
End Function

Private Function NldiDownstream(ByVal hitLon As Double, ByVal hitLat As Double, ByVal maxMiles As Double) As Collection
    ' Mainstem flowlines downstream of the point (USGS NLDI 'DM' navigation), ordered upstream to downstream.
    Dim txt As String, d As Variant, comid As String, feat As Variant, pts As Collection, r As Object
    Dim lines As New Collection, km As Double, ordered As Collection
    txt = HttpText("GET", URL_NLDI & "/comid/position?f=json&coords=" & UrlEnc("POINT(" & NumFix(hitLon, 7) & " " & NumFix(hitLat, 7) & ")"), "", 60, 2)
    JsonParse txt, d
    If Not HasKey(d, "features") Then Err.Raise vbObjectError + 45, , "no NHDPlus catchment found at this point"
    If d("features").Count = 0 Then Err.Raise vbObjectError + 45, , "no NHDPlus catchment found at this point"
    comid = JStr(d("features")(1)("properties"), "comid")
    LogLine "  Start reach NHDPlus COMID " & comid & "; navigating downstream " & NumStr(maxMiles) & " mi..."
    km = MaxD(0.1, maxMiles) * 1.609344
    txt = HttpText("GET", URL_NLDI & "/comid/" & comid & "/navigation/DM/flowlines?f=json&distance=" & NumFix(km, 3), "", 90, 2)
    JsonParse txt, d
    If HasKey(d, "features") Then
        For Each feat In d("features")
            If HasKey(feat, "geometry") Then
                Set pts = GeomPoints(feat("geometry"))
                If pts.Count >= 2 Then
                    Set r = CreateObject("Scripting.Dictionary")
                    r("comid") = ""
                    If HasKey(feat, "properties") Then r("comid") = JStr(feat("properties"), "nhdplus_comid")
                    r("name") = ""
                    Set r("pts") = pts
                    lines.Add r
                End If
            End If
        Next feat
    End If
    Set NldiDownstream = OrderDownstream(lines, comid)
End Function

Private Function EndKey(ByRef p As Variant) As String
    EndKey = Format$(CDbl(p(1)), "0.00000") & "," & Format$(CDbl(p(2)), "0.00000")
End Function

Private Function OrderDownstream(ByRef lines As Collection, ByVal startComid As String) As Collection
    ' Chains flowlines by matching end and start vertices; keeps the service order when they do not chain.
    Dim byStart As Object, r As Variant, cur As Object, ordered As New Collection, seen As Object, pts As Collection, k As String
    Set OrderDownstream = lines
    If lines.Count < 2 Then Exit Function
    Set byStart = CreateObject("Scripting.Dictionary")
    Set seen = CreateObject("Scripting.Dictionary")
    For Each r In lines
        Set pts = r("pts")
        Set byStart(EndKey(pts(1))) = r
        If cur Is Nothing And CStr(r("comid")) = startComid Then Set cur = r
    Next r
    If cur Is Nothing Then Set cur = lines(1)
    Do
        ordered.Add cur
        Set pts = cur("pts")
        seen(EndKey(pts(1))) = True
        k = EndKey(pts(pts.Count))
        If Not byStart.Exists(k) Then Exit Do
        If seen.Exists(k) Then Exit Do
        Set cur = byStart(k)
    Loop
    If ordered.Count = lines.Count Then Set OrderDownstream = ordered
End Function

Private Function TraceClippedStreams(ByVal sLon As Double, ByVal sLat As Double, ByRef streams As Collection) As Collection
    ' Backup: walks the fence-clipped NHD lines downstream by matching end points.
    Dim parts As New Collection, s As Variant, ln As Variant, r As Object, out As New Collection
    Dim best As Object, bestD As Double, seg As Long, ox As Double, oy As Double, D2 As Double
    Dim cur As Object, pts As Collection, e As Variant, f As Variant, found As Boolean, used As Object, i As Long
    Set TraceClippedStreams = out
    If streams Is Nothing Then Exit Function
    For Each s In streams
        For Each ln In s("parts")
            If ln.Count >= 2 Then
                Set r = CreateObject("Scripting.Dictionary")
                r("comid") = CStr(s("id"))
                r("name") = CStr(s("name"))
                Set r("pts") = ln
                parts.Add r
            End If
        Next ln
    Next s
    If parts.Count = 0 Then Exit Function
    bestD = 1E+300
    For Each r In parts
        ProjectOntoPath sLon, sLat, r("pts"), seg, ox, oy, D2
        If D2 < bestD Then
            bestD = D2
            Set best = r
        End If
    Next r
    Set used = CreateObject("Scripting.Dictionary")
    Set cur = best
    Do
        out.Add cur
        Set pts = cur("pts")
        Set e = pts(pts.Count)
        found = False
        i = 0
        For Each r In parts
            i = i + 1
            If Not (r Is cur) And Not used.Exists(i) Then
                Set f = r("pts")(1)
                If Abs(f(1) - e(1)) < 0.00001 And Abs(f(2) - e(2)) < 0.00001 Then
                    used(i) = True
                    Set cur = r
                    found = True
                    Exit For
                End If
            End If
        Next r
        If out.Count > parts.Count Then Exit Do
    Loop While found
End Function

Public Function TraceRaindrop(ByVal lon As Double, ByVal lat As Double, ByVal maxMiles As Double, ByRef res As Object) As Boolean
    ' Traces water from the point overland to the nearest NHDPlus flowline, then down the mainstem.
    ' The result (or res("error")) comes back in res.
    Dim overland As Collection, method As String, reaches As Collection, source As String
    Dim hitLon As Double, hitLat As Double, seg As Long, px As Double, py As Double, D2 As Double
    Dim r As Variant, pts As Collection, np As Collection, i As Long, cum As Double, path As New Collection
    Dim last As Variant, first As Variant, p As Variant, errText As String, skipFirst As Boolean, r0 As Object
    Set res = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    LogLine "Raindrop trace: requesting NLDI overland flow path..."
    Set overland = TryFlowtrace(lon, lat)
    If Not overland Is Nothing Then method = "NLDI raindrop trace (DEM flow direction)"
    If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
    hitLon = lon
    hitLat = lat
    If Not overland Is Nothing Then
        Set last = overland(overland.Count)
        hitLon = CDbl(last(1))
        hitLat = CDbl(last(2))
    End If
    source = "USGS NLDI downstream mainstem"
    Set reaches = TryDownstream(hitLon, hitLat, maxMiles, errText)
    If reaches Is Nothing Then
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
        If gStreams Is Nothing Then Err.Raise vbObjectError + 46, , "Downstream network trace failed: " & errText
        If gStreams.Count = 0 Then Err.Raise vbObjectError + 46, , "Downstream network trace failed: " & errText
        LogLine "  Note: NLDI navigation failed (" & Left$(errText, 90) & "); tracing the fence NHD streams instead."
        Set reaches = TraceClippedStreams(hitLon, hitLat, gStreams)
        source = "Fence NHDPlus HR streams (endpoint connectivity)"
    End If
    If reaches.Count = 0 Then Err.Raise vbObjectError + 47, , "No downstream flowlines were found for this point."

    ' start the first reach at the point where the water reaches it
    Set r0 = reaches(1)
    Set pts = r0("pts")
    ProjectOntoPath hitLon, hitLat, pts, seg, px, py, D2
    Set np = New Collection
    np.Add Pt(px, py)
    For i = seg + 2 To pts.Count
        np.Add pts(i)
    Next i
    Set r0 = reaches(1)
    Set r0("pts") = np
    If overland Is Nothing Then
        Set overland = New Collection
        overland.Add Pt(lon, lat)
        overland.Add Pt(px, py)
        method = "Straight line to nearest flowline (approximation)"
    Else
        Set last = overland(overland.Count)
        If last(1) <> px Or last(2) <> py Then overland.Add Pt(px, py)
    End If
    cum = GeoLengthFt(overland)
    res("overland_ft") = cum
    For Each r In reaches
        r("length_ft") = GeoLengthFt(r("pts"))
        cum = cum + r("length_ft")
        r("cum_ft") = cum
    Next r
    For Each p In overland
        path.Add p
    Next p
    For Each r In reaches
        Set pts = r("pts")
        Set first = pts(1)
        Set last = path(path.Count)
        skipFirst = (first(1) = last(1) And first(2) = last(2))
        i = 0
        For Each p In pts
            i = i + 1
            If Not (skipFirst And i = 1) Then path.Add p
        Next p
    Next r
    res("lon") = lon
    res("lat") = lat
    Set res("overland") = overland
    res("overland_method") = method
    Set res("reaches") = reaches
    Set res("path") = path
    res("total_ft") = cum
    res("source") = source
    res("max_miles") = maxMiles
    TraceRaindrop = True
    Exit Function
Failed:
    res("error") = Err.Description
    LogLine "[Raindrop Error] " & Err.Description
End Function

Private Function TryFlowtrace(ByVal lon As Double, ByVal lat As Double) As Collection
    On Error GoTo Bad
    Set TryFlowtrace = NldiFlowtrace(lon, lat)
    Exit Function
Bad:
    LogLine "  Note: NLDI raindrop service unavailable (" & Left$(Err.Description, 90) & "); using a straight overland leg to the nearest flowline."
End Function

Private Function TryDownstream(ByVal lon As Double, ByVal lat As Double, ByVal maxMiles As Double, ByRef errText As String) As Collection
    On Error GoTo Bad
    Set TryDownstream = NldiDownstream(lon, lat, maxMiles)
    Exit Function
Bad:
    errText = Err.Description
End Function

'----------------------------------------------------------------------------------------------
' Service check
'----------------------------------------------------------------------------------------------
Public Function CheckServices(ByRef labels As Variant, ByRef oks() As Boolean, ByRef details() As String, _
                              ByRef impacts As Variant) As Long
    ' Probes every web service at the same time (15 s limit). Any HTTP status below 500 counts as responding.
    ' Returns the number of services that are down and applies the 10-minute skip for those.
    Dim urls As Variant, n As Long, i As Long, h() As Object, done() As Boolean, t0 As Single, remaining As Long
    Dim st As Long, body As String, down As Long, untilTime As Date
    labels = Array("StreamStats delineation", "StreamStats pour-point snap", "StreamStats basin characteristics", _
                   "USGS NSS regression", "USGS NHD streams", "USGS NLDI (raindrop path)", "MRLC NLCD land cover", _
                   "USGS 3DEP elevation", "NOAA Atlas 14 rainfall", "USDA SSURGO soils", _
                   "UC Davis SoilWeb soil grid (backup)", "USGS NSS region lookup")
    urls = Array("https://streamstats.usgs.gov/ss-delineate/v1/", "https://streamstats.usgs.gov/pourpoint/v2/", _
                 "https://streamstats.usgs.gov/ss-hydro/v1/", "https://streamstats.usgs.gov/nssservices/regions", _
                 "https://hydro.nationalmap.gov/arcgis/rest/services/nhd/MapServer?f=json", _
                 "https://api.water.usgs.gov/nldi/linked-data/", _
                 "https://www.mrlc.gov/geoserver/ows?service=WCS&version=2.0.1&request=DescribeCoverage&coverageId=mrlc_download__NLCD_2021_Impervious_L48", _
                 "https://elevation.nationalmap.gov/arcgis/rest/services/3DEPElevation/ImageServer?f=json", _
                 "https://hdsc.nws.noaa.gov/cgi-bin/hdsc/new/cgi_readH5.py?lat=37.5&lon=-78.5&type=pf&data=intensity&units=english&series=pds", _
                 "https://sdmdataaccess.sc.egov.usda.gov/", "https://casoilresource.lawr.ucdavis.edu/wcs/issr800.json", _
                 "https://streamstats.usgs.gov/nssservices/regressionregions/bylocation")
    impacts = Array("watersheds cannot be delineated", "pour points are not snapped to the stream grid", _
                    "slope and land cover come from MRLC NLCD / USGS 3DEP; the longest flowpath is traced from NHD", _
                    "no USGS regression peak flows", "streams cannot be loaded or traced", "raindrop paths cannot be traced", _
                    "no alternate impervious/forest values", "no alternate channel slope", _
                    "built-in IDF coefficients are used and the NRCS method is not calculated", _
                    "soil groups come from the coarser UC Davis SoilWeb grid", _
                    "no backup soil groups if USDA SSURGO is also down", _
                    "no USGS regression peak flows (each culvert skips them after one 45 s wait)")
    body = "{""type"":""Polygon"",""coordinates"":[[[-78.50,37.50],[-78.49,37.50],[-78.49,37.51],[-78.50,37.51],[-78.50,37.50]]]}"
    n = UBound(labels) + 1
    ReDim h(0 To n - 1)
    ReDim done(0 To n - 1)
    ReDim oks(0 To n - 1)
    ReDim details(0 To n - 1)
    On Error Resume Next
    For i = 0 To n - 1
        Set h(i) = CreateObject("MSXML2.ServerXMLHTTP.6.0")
        h(i).setTimeouts 15000, 15000, 15000, 15000
        If i = n - 1 Then
            h(i).Open "POST", CStr(urls(i)), True
            h(i).setRequestHeader "User-Agent", USER_AGENT
            h(i).setRequestHeader "Content-Type", "application/json"
            h(i).send body
        Else
            h(i).Open "GET", CStr(urls(i)), True
            h(i).setRequestHeader "User-Agent", USER_AGENT
            h(i).send
        End If
        If Err.Number <> 0 Then
            done(i) = True
            details(i) = Left$(Err.Description, 60)
            Err.Clear
        End If
    Next i
    t0 = Timer
    Do
        remaining = 0
        For i = 0 To n - 1
            If Not done(i) Then
                If h(i).readyState = 4 Then
                    done(i) = True
                    st = 0
                    st = h(i).Status
                    If Err.Number <> 0 Then
                        details(i) = "no answer, " & Format$(Timer - t0, "0.0") & " s"
                        Err.Clear
                    Else
                        oks(i) = (st > 0 And st < 500)
                        details(i) = "HTTP " & st & ", " & Format$(Timer - t0, "0.0") & " s"
                    End If
                Else
                    remaining = remaining + 1
                End If
            End If
        Next i
        If remaining = 0 Then Exit Do
        If Timer - t0 > 15# Or Timer < t0 Then Exit Do
        Sleep 50
        DoEvents
    Loop
    For i = 0 To n - 1
        If Not done(i) Then
            h(i).abort
            details(i) = "no answer in 15 s"
        End If
        Set h(i) = Nothing
    Next i
    Err.Clear
    On Error GoTo 0
    ' services found down are skipped for 10 minutes; the others are tried again at once
    untilTime = Now + 10# / 1440#
    SetServiceDown "bc", IIf(oks(2), 0, untilTime)
    SetLfpDown IIf(oks(2), 0, untilTime)
    SetServiceDown "nss", IIf(oks(3) And oks(11), 0, untilTime)
    SetMethodsServiceDown "ssurgo", IIf(oks(9), 0, untilTime)
    SetMethodsServiceDown "nlcd", IIf(oks(6), 0, untilTime)
    For i = 0 To n - 1
        If Not oks(i) Then down = down + 1
    Next i
    CheckServices = down
End Function
