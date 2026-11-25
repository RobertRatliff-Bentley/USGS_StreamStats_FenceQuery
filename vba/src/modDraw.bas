Attribute VB_Name = "modDraw"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology (VBA) - drawing the stream network in the DGN
' Fence outline, clipped NHD streams (labels, flow arrows, Tributary Item Type), longest flowpath
' and the Flow (Raindrop) Path with its table. Levels, colors and sizes match the Python tool.
'==============================================================================================
Private Const RAINDROP_OVERLAND_COLOR As Long = 4      ' yellow, dashed; the channel uses color 7
Private mFenceOutlineId As String

'----------------------------------------------------------------------------------------------
' Delete All Drawings
'----------------------------------------------------------------------------------------------
Public Function ToolLevelNames() As Variant
    ToolLevelNames = Array("USGS_NHD_Streams_Clipped", "USGS_NHD_Streams_Labels", "USGS_NHD_Flow_Arrows", _
        "USGS_StreamStats_Basin", "USGS_StreamStats_Point", "USGS_Culvert_Inlet", "USGS_Longest_Flowpath", _
        "USGS_Raindrop_Path", "USGS_StreamStats_Table", "USGS_Fence_Boundary")
End Function

Public Function DeleteToolDrawings(ByVal countOnly As Boolean) As Long
    ' Counts (countOnly) or deletes every element in the active model on one of this tool's levels.
    ' Deleting also clears the active fence.
    Dim sc As New ElementScanCriteria, names As Variant, i As Long, lv As Level, found As Boolean
    Dim ee As ElementEnumerator, col As New Collection, el As Variant, n As Long
    names = ToolLevelNames()
    sc.ExcludeAllLevels
    For i = LBound(names) To UBound(names)
        Set lv = Nothing
        On Error Resume Next
        Set lv = ActiveDesignFile.Levels(CStr(names(i)))
        On Error GoTo 0
        If Not lv Is Nothing Then
            sc.IncludeLevel lv
            found = True
        End If
    Next i
    If found Then
        sc.ExcludeNonGraphical
        Set ee = ActiveModelReference.Scan(sc)
        Do While ee.MoveNext
            col.Add ee.Current
        Loop
    End If
    If countOnly Then
        DeleteToolDrawings = col.Count
        Exit Function
    End If
    For Each el In col
        On Error Resume Next
        ActiveModelReference.RemoveElement el
        If Err.Number = 0 Then n = n + 1
        Err.Clear
        On Error GoTo 0
    Next el
    mFenceOutlineId = ""
    On Error Resume Next
    ActiveDesignFile.fence.Undefine
    On Error GoTo 0
    DeleteToolDrawings = n
End Function


'----------------------------------------------------------------------------------------------
' Small drawing helpers
'----------------------------------------------------------------------------------------------
Public Function ProjectLine(ByRef poly As Variant, ByRef xs() As Double, ByRef ys() As Double) As Long
    ' Longitude / latitude line -> DGN coordinates. Returns the number of points.
    Dim n As Long, i As Long, p As Variant, x As Double, y As Double
    n = poly.Count
    If n = 0 Then Exit Function
    ReDim xs(0 To n - 1)
    ReDim ys(0 To n - 1)
    For Each p In poly
        LonLatToDgn CDbl(p(1)), CDbl(p(2)), x, y
        xs(i) = x
        ys(i) = y
        i = i + 1
    Next p
    ProjectLine = n
End Function

Public Function AddLineString(ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long, ByVal lvl As Level, _
                              ByVal colorIndex As Long, ByVal weight As Long, Optional ByVal dashed As Boolean = False, _
                              Optional ByVal firstIdx As Long = 0) As Element
    ' Line string through points firstIdx .. firstIdx + n - 1 (at most 4,001 points per element).
    Dim pts() As Point3d, i As Long, el As Element, ls As LineStyle
    If n < 2 Then Exit Function
    ReDim pts(0 To n - 1)
    For i = 0 To n - 1
        pts(i).x = xs(firstIdx + i)
        pts(i).y = ys(firstIdx + i)
        pts(i).z = 0#
    Next i
    Set el = CreateLineElement1(Nothing, pts)
    If dashed Then
        On Error Resume Next
        Set ls = ActiveDesignFile.LineStyles.Find("2", Nothing)
        If Not ls Is Nothing Then Set el.LineStyle = ls
        On Error GoTo 0
    End If
    AddStyled el, lvl, colorIndex, weight
    Set AddLineString = el
End Function

Public Sub AddFilledTriangle(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                             ByVal x3 As Double, ByVal y3 As Double, ByVal lvl As Level, ByVal colorIndex As Long)
    Dim pts(0 To 2) As Point3d, el As ShapeElement
    pts(0) = Point3dFromXY(x1, y1)
    pts(1) = Point3dFromXY(x2, y2)
    pts(2) = Point3dFromXY(x3, y3)
    Set el = CreateShapeElement1(Nothing, pts, msdFillModeFilled)
    On Error Resume Next
    el.FillColor = colorIndex
    On Error GoTo 0
    AddStyled el, lvl, colorIndex, 1
End Sub

Public Sub AddCircleX(ByVal x As Double, ByVal y As Double, ByVal rad As Double, ByVal lvl As Level, ByVal colorIndex As Long)
    ' Circle with an inscribed X (raindrop start / end symbol)
    Dim el As Element, d As Double
    Set el = CreateEllipseElement2(Nothing, Point3dFromXY(x, y), rad, rad, Matrix3dIdentity, msdFillModeNotFilled)
    AddStyled el, lvl, colorIndex, 3
    d = rad * Sqr(0.5)
    AddLine x - d, y - d, x + d, y + d, lvl, colorIndex, 2
    AddLine x - d, y + d, x + d, y - d, lvl, colorIndex, 2
End Sub

Public Sub AddTextRot(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal th As Double, ByVal angleRad As Double, _
                      ByVal lvl As Level, ByVal colorIndex As Long)
    Dim t As TextElement, wasOn As Boolean
    wasOn = ActiveSettings.AnnotationScaleEnabled
    On Error GoTo Restore
    ActiveSettings.AnnotationScaleEnabled = False
    Set t = CreateTextElement1(Nothing, s, Point3dFromXY(x, y), Matrix3dFromAxisAndRotationAngle(2, angleRad))
    t.TextStyle.Height = th
    t.TextStyle.Width = th
    AddStyled t, lvl, colorIndex, 1
Restore:
    ActiveSettings.AnnotationScaleEnabled = wasOn
End Sub

Private Function Atan2(ByVal y As Double, ByVal x As Double) As Double
    Const Pi As Double = 3.14159265358979
    If x > 0# Then
        Atan2 = Atn(y / x)
    ElseIf x < 0# Then
        If y >= 0# Then
            Atan2 = Atn(y / x) + Pi
        Else
            Atan2 = Atn(y / x) - Pi
        End If
    ElseIf y > 0# Then
        Atan2 = Pi / 2#
    ElseIf y < 0# Then
        Atan2 = -Pi / 2#
    End If
End Function

Private Function BearingText(ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long, ByRef bearing As Double) As String
    ' Grid bearing from the first to the last point, as "Flows NE (47 deg grid)". "" when the ends coincide.
    Dim dx As Double, dy As Double, comp As Variant
    bearing = -1#
    If n < 2 Then Exit Function
    dx = xs(n - 1) - xs(0)
    dy = ys(n - 1) - ys(0)
    If dx = 0# And dy = 0# Then Exit Function
    bearing = Atan2(dx, dy) * 180# / 3.14159265358979
    If bearing < 0# Then bearing = bearing + 360#
    bearing = RoundTo(bearing, 1)
    If bearing >= 360# Then bearing = 0#
    comp = Array("N", "NE", "E", "SE", "S", "SW", "W", "NW")
    BearingText = "Flows " & comp(CLng(Int((bearing + 22.5) / 45#)) Mod 8) & " (" & Format$(bearing, "0") & Chr$(176) & " grid)"
End Function

Public Function FlowArrows(ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long, ByVal lvl As Level, ByVal colorIndex As Long) As Long
    ' Filled arrowheads pointing downstream, every 40 text heights and away from the midpoint label.
    Dim seg() As Double, total As Double, alen As Double, th As Double, na As Long, st() As Double, ns As Long
    Dim i As Long, k As Long, acc As Double, f As Double, ux As Double, uy As Double, tx As Double, ty As Double
    Dim bx As Double, by As Double, hw As Double
    If n < 2 Then Exit Function
    ReDim seg(0 To n - 2)
    For i = 0 To n - 2
        seg(i) = Sqr((xs(i + 1) - xs(i)) ^ 2 + (ys(i + 1) - ys(i)) ^ 2)
        total = total + seg(i)
    Next i
    th = TextHeightFt()
    alen = 1.5 * th
    If total < 4# * alen Then Exit Function
    na = CLng(Int(total / (40# * th)))
    If na < 1 Then na = 1
    ReDim st(0 To na)
    For i = 0 To na - 1
        st(i) = total * (i + 0.25) / na
    Next i
    ns = na
    If na = 1 Then
        st(1) = total * 0.75
        ns = 2
    End If
    hw = 0.4 * alen
    For k = 0 To ns - 1
        acc = 0#
        For i = 0 To n - 2
            If acc + seg(i) >= st(k) And seg(i) > 0.000001 Then
                f = (st(k) - acc) / seg(i)
                ux = (xs(i + 1) - xs(i)) / seg(i)
                uy = (ys(i + 1) - ys(i)) / seg(i)
                tx = xs(i) + f * (xs(i + 1) - xs(i))
                ty = ys(i) + f * (ys(i + 1) - ys(i))
                bx = tx - ux * alen
                by = ty - uy * alen
                AddFilledTriangle tx, ty, bx - uy * hw, by + ux * hw, bx + uy * hw, by - ux * hw, lvl, colorIndex
                FlowArrows = FlowArrows + 1
                Exit For
            End If
            acc = acc + seg(i)
        Next i
    Next k
End Function

Private Function StreamLabel(ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long, ByVal label As String, _
                             ByVal lvl As Level, ByVal colorIndex As Long) As Boolean
    ' Label rotated along the line at its midpoint, one text height off the line.
    Dim total As Double, i As Long, l As Double, half As Double, acc As Double, f As Double
    Dim mX As Double, mY As Double, ang As Double, th As Double
    Const Pi As Double = 3.14159265358979
    If n < 2 Or Len(label) = 0 Then Exit Function
    For i = 0 To n - 2
        total = total + Sqr((xs(i + 1) - xs(i)) ^ 2 + (ys(i + 1) - ys(i)) ^ 2)
    Next i
    If total < 25# Then Exit Function
    half = 0.5 * total
    mX = xs(0)
    mY = ys(0)
    For i = 0 To n - 2
        l = Sqr((xs(i + 1) - xs(i)) ^ 2 + (ys(i + 1) - ys(i)) ^ 2)
        If acc + l >= half And l > 0.000001 Then
            f = (half - acc) / l
            mX = xs(i) + f * (xs(i + 1) - xs(i))
            mY = ys(i) + f * (ys(i + 1) - ys(i))
            ang = Atan2(ys(i + 1) - ys(i), xs(i + 1) - xs(i))
            Exit For
        End If
        acc = acc + l
    Next i
    If ang > Pi / 2# Then
        ang = ang - Pi
    ElseIf ang < -Pi / 2# Then
        ang = ang + Pi
    End If
    th = TextHeightFt()
    AddTextRot label, mX - Sin(ang) * th, mY + Cos(ang) * th, th, ang, lvl, colorIndex
    StreamLabel = True
End Function

'----------------------------------------------------------------------------------------------
' Fence
'----------------------------------------------------------------------------------------------
Public Function ReadFence() As Collection
    ' Active fence outline as a closed longitude / latitude line. Raises an error when there is no usable fence.
    Dim f As fence, el As Element, v() As Point3d, i As Long, out As New Collection, lon As Double, lat As Double
    Dim xs() As Double, ys() As Double, n As Long
    Set f = ActiveDesignFile.fence
    If Not f.IsDefined Then
        Err.Raise vbObjectError + 50, , "No active Fence in MicroStation. Click Place Fence in DGN (or place a Fence) first."
    End If
    On Error GoTo NoVerts
    Set el = f.CreateElement
    v = el.AsVertexList.GetVertices
    On Error GoTo 0
    n = UBound(v) - LBound(v) + 1
    If n < 3 Then GoTo NoVerts
    ReDim xs(0 To n)
    ReDim ys(0 To n)
    For i = 0 To n - 1
        xs(i) = v(LBound(v) + i).x
        ys(i) = v(LBound(v) + i).y
    Next i
    If xs(0) <> xs(n - 1) Or ys(0) <> ys(n - 1) Then
        xs(n) = xs(0)
        ys(n) = ys(0)
        n = n + 1
    End If
    For i = 0 To n - 1
        DgnToLonLat xs(i), ys(i), lon, lat
        out.Add Pt(lon, lat)
    Next i
    DrawFenceOutline xs, ys, n
    Set ReadFence = out
    Exit Function
NoVerts:
    Err.Raise vbObjectError + 51, , "Unable to read fence vertices. Re-place the fence as a Block or Shape and try again."
End Function

Private Sub DrawFenceOutline(ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long)
    ' MicroStation has no fence-color setting, so the boundary is drawn in red on its own level.
    ' It replaces the outline drawn earlier in this session.
    Dim lvl As Level, el As Element
    On Error GoTo Bad
    If Len(mFenceOutlineId) > 0 Then
        On Error Resume Next
        ActiveModelReference.RemoveElement ActiveModelReference.GetElementByID(DLongFromString(mFenceOutlineId))
        On Error GoTo Bad
        mFenceOutlineId = ""
    End If
    Set lvl = EnsureLevel("USGS_Fence_Boundary", 3, 2)
    Set el = AddLineString(xs, ys, n, lvl, 3, 2)
    If Not el Is Nothing Then
        mFenceOutlineId = DLongToString(el.id)
        LogLine "Fence boundary drawn in red on level 'USGS_Fence_Boundary'."
    End If
    Exit Sub
Bad:
    LogLine "[DGN Draw Error] Fence outline: " & Err.Description
End Sub

'----------------------------------------------------------------------------------------------
' Clipped NHD streams
'----------------------------------------------------------------------------------------------
Private Function StreamPropNames() As Variant
    StreamPropNames = Split("StreamName,NHDPlusID,ReachCode,FeatureType,StreamOrder,FlowDirection,FlowBearing_Deg,FlowsInto," & _
                            "MeanAnnualFlow_cfs,MeanVelocity_fps,DrainageArea_SqMi,Slope_FtPerFt,ClippedLength_Ft,Source,QueryDate", ",")
End Function

Private Sub PutNum(ByRef vals As Object, ByVal nm As String, ByVal v As Variant, ByVal decimals As Long)
    ' NHDPlus uses -9998 / -9999 for missing values, so negatives are left blank.
    If IsEmpty(v) Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    If CDbl(v) < 0# Then Exit Sub
    vals(nm) = RoundTo(CDbl(v), decimals)
End Sub

Private Function StreamItemValues(ByRef s As Object, ByRef xs() As Double, ByRef ys() As Double, ByVal n As Long) As Object
    Dim vals As Object, nhd As Object, bearing As Double, dirTxt As String, v As Variant, ft As String
    Set vals = CreateObject("Scripting.Dictionary")
    Set nhd = s("nhd")
    dirTxt = BearingText(xs, ys, n, bearing)
    vals("StreamName") = CStr(s("name"))
    v = nhd("nhdplusid")
    If IsEmpty(v) Then v = nhd("permanent_identifier")
    If Not IsEmpty(v) Then
        If IsNumeric(v) Then
            vals("NHDPlusID") = Format$(CDbl(v), "0")
        Else
            vals("NHDPlusID") = CStr(v)
        End If
    End If
    If Not IsEmpty(nhd("reachcode")) Then vals("ReachCode") = CStr(nhd("reachcode"))
    If Not IsEmpty(nhd("ftype")) Then
        ft = CStr(nhd("ftype"))
        Select Case ft
            Case "460"
                ft = "Stream/River"
            Case "558"
                ft = "Artificial Path"
            Case "336"
                ft = "Canal/Ditch"
            Case "334"
                ft = "Connector"
            Case "420"
                ft = "Underground Conduit"
            Case "428"
                ft = "Pipeline"
            Case "566"
                ft = "Coastline"
        End Select
        vals("FeatureType") = ft
    ElseIf Not IsEmpty(nhd("fcode")) Then
        vals("FeatureType") = CStr(nhd("fcode"))
    End If
    vals("StreamOrder") = CDbl(s("order"))
    If Len(dirTxt) > 0 Then
        vals("FlowDirection") = dirTxt
        vals("FlowBearing_Deg") = bearing
    Else
        vals("FlowDirection") = "Downstream (digitized direction)"
    End If
    If Len(CStr(s("flows_into"))) > 0 Then vals("FlowsInto") = CStr(s("flows_into"))
    PutNum vals, "MeanAnnualFlow_cfs", nhd("qema"), 2
    PutNum vals, "MeanVelocity_fps", nhd("vema"), 2
    If CDbl(s("da_sqkm")) > 0# Then vals("DrainageArea_SqMi") = RoundTo(CDbl(s("da_sqkm")) * 0.386102, 4)
    PutNum vals, "Slope_FtPerFt", nhd("slope"), 5
    vals("ClippedLength_Ft") = CDbl(s("length_ft"))
    vals("Source") = "USGS NHDPlus HR"
    vals("QueryDate") = Format$(Date, "yyyy-mm-dd")
    Set StreamItemValues = vals
End Function

Public Sub DrawStreams(ByRef streams As Collection)
    ' Draws the clipped streams with rotated labels, flow arrows and the Tributary Item Type.
    Dim lvlS As Level, lvlL As Level, lvlA As Level, s As Variant, part As Variant, sd As Object
    Dim xs() As Double, ys() As Double, n As Long, el As Element, i As Long
    Dim cnt As Long, lbl As Long, arrows As Long, items As Long, label As String
    If streams Is Nothing Then Exit Sub
    Set lvlS = EnsureLevel("USGS_NHD_Streams_Clipped", 1, 2)
    Set lvlL = EnsureLevel("USGS_NHD_Streams_Labels", 1, 1)
    Set lvlA = EnsureLevel("USGS_NHD_Flow_Arrows", 1, 1)
    For Each s In streams
        Set sd = s
        label = CStr(sd("name")) & " (Order " & CStr(sd("order")) & ")"
        For Each part In sd("parts")
            n = ProjectLine(part, xs, ys)
            If n >= 2 Then
                On Error Resume Next
                For i = 0 To n - 2 Step 4000
                    Set el = Nothing
                    Set el = AddLineString(xs, ys, IIf(n - i > 4001, 4001, n - i), lvlS, 1, 2, False, i)
                    If Not el Is Nothing Then
                        cnt = cnt + 1
                        If AttachItem(el, "Tributary", StreamPropNames(), "ssssdsdsdddddss", StreamItemValues(sd, xs, ys, n), True) Then items = items + 1
                    End If
                Next i
                If StreamLabel(xs, ys, n, label, lvlL, 1) Then lbl = lbl + 1
                arrows = arrows + FlowArrows(xs, ys, n, lvlA, 1)
                On Error GoTo 0
            End If
        Next part
        DoEvents
    Next s
    LogLine "  Added " & cnt & " clipped stream reaches, " & lbl & " rotated stream labels and " & arrows & " flow-direction arrows to DGN."
    If items > 0 Then LogLine "  Attached Item Type 'USGS_VDOT_Hydrology:Tributary' to " & items & " tributary line(s)."
End Sub

'----------------------------------------------------------------------------------------------
' Longest flowpath
'----------------------------------------------------------------------------------------------
Public Sub DrawFlowpath(ByRef res As Object)
    Dim lfp As Collection, xs() As Double, ys() As Double, n As Long, lvl As Level, el As Element
    Dim vals As Object, bearing As Double, dirTxt As String, i As Long, nseg As Long, seg As Long, items As Long
    Dim first As Variant, last As Variant, lengthFt As Double
    If Not res.Exists("lfp") Then Exit Sub
    If res("lfp") Is Nothing Then Exit Sub
    Set lfp = res("lfp")
    n = ProjectLine(lfp, xs, ys)
    If n < 2 Then Exit Sub
    Set lvl = EnsureLevel("USGS_Longest_Flowpath", 5, 3)
    lengthFt = CDbl(res("lfp_ft"))
    dirTxt = BearingText(xs, ys, n, bearing)
    Set first = lfp(1)
    Set last = lfp(lfp.Count)
    nseg = (n - 2) \ 4000 + 1
    For i = 0 To n - 2 Step 4000
        seg = seg + 1
        Set el = AddLineString(xs, ys, IIf(n - i > 4001, 4001, n - i), lvl, 5, 3, False, i)
        Set vals = CreateObject("Scripting.Dictionary")
        vals("CulvertID") = CStr(res("culvert_id"))
        vals("Length_Ft") = RoundTo(lengthFt, 1)
        vals("Length_Mi") = RoundTo(lengthFt / 5280#, 3)
        vals("Method") = CStr(res("lfp_method"))
        If Len(dirTxt) > 0 Then
            vals("FlowDirection") = dirTxt & " to outlet"
            vals("FlowBearing_Deg") = bearing
        End If
        vals("Segment") = seg & " of " & nseg
        vals("BasinSlope_FtPerMi") = RoundTo(CDbl(res("slope")), 2)
        vals("Tc_Min") = CDbl(res("tc"))
        vals("TcMethod") = "Kirpich"
        vals("DrainageArea_Acres") = RoundTo(CDbl(res("acres")), 2)
        vals("Upstream_Lat") = RoundTo(CDbl(first(2)), 6)
        vals("Upstream_Lon") = RoundTo(CDbl(first(1)), 6)
        vals("Outlet_Lat") = RoundTo(CDbl(last(2)), 6)
        vals("Outlet_Lon") = RoundTo(CDbl(last(1)), 6)
        vals("CalcDate") = Format$(Date, "yyyy-mm-dd")
        If AttachItem(el, "LongestFlowpath", Split("CulvertID,Length_Ft,Length_Mi,Method,FlowDirection,FlowBearing_Deg,Segment," & _
                      "BasinSlope_FtPerMi,Tc_Min,TcMethod,DrainageArea_Acres,Upstream_Lat,Upstream_Lon,Outlet_Lat,Outlet_Lon,CalcDate", ","), _
                      "sddssdsddsddddds", vals, True) Then items = items + 1
    Next i
    LogLine "  Drew longest flowpath (" & Format$(lengthFt, "#,##0") & " ft) on level 'USGS_Longest_Flowpath'."
    If items > 0 Then LogLine "  Attached Item Type 'USGS_VDOT_Hydrology:LongestFlowpath' to the longest flowpath (" & CStr(res("culvert_id")) & ")."
End Sub

'----------------------------------------------------------------------------------------------
' Flow (raindrop) path
'----------------------------------------------------------------------------------------------
Public Function RaindropRows(ByRef res As Object) As Collection
    ' One entry per segment: Array(seq, segment, comid, points, length ft, cumulative ft, method)
    Dim rows As New Collection, r As Variant, i As Long
    rows.Add Array(1, "OVERLAND", "", res("overland"), CDbl(res("overland_ft")), CDbl(res("overland_ft")), CStr(res("overland_method")))
    i = 2
    For Each r In res("reaches")
        rows.Add Array(i, "STREAM", CStr(r("comid")), r("pts"), CDbl(r("length_ft")), CDbl(r("cum_ft")), CStr(res("source")))
        i = i + 1
    Next r
    Set RaindropRows = rows
End Function

Public Sub DrawLabeledPoint(ByVal x As Double, ByVal y As Double, ByVal label As String)
    ' Circle-and-X marker with a label, used for the picked raindrop start point.
    Dim lvl As Level, th As Double, rad As Double
    Set lvl = EnsureLevel("USGS_Raindrop_Path", 7, 3)
    th = TextHeightFt()
    rad = 4# * th / 7#
    AddCircleX x, y, rad, lvl, 7
    AddText label, x + rad + 0.6 * th, y - 0.5 * th, th, lvl, 7
End Sub

Private Function RaindropValues(ByRef res As Object, ByVal traceId As String, ByVal kind As String, ByRef xs() As Double, _
                                ByRef ys() As Double, ByVal firstIdx As Long, ByVal n As Long, ByVal seg As Long, ByVal nseg As Long) As Object
    Dim vals As Object, bx() As Double, by() As Double, i As Long, bearing As Double, dirTxt As String, segLen As Double
    Dim reaches As Collection, path As Collection, last As Variant, total As Double
    Set vals = CreateObject("Scripting.Dictionary")
    ReDim bx(0 To n - 1)
    ReDim by(0 To n - 1)
    For i = 0 To n - 1
        bx(i) = xs(firstIdx + i)
        by(i) = ys(firstIdx + i)
        If i > 0 Then segLen = segLen + Sqr((bx(i) - bx(i - 1)) ^ 2 + (by(i) - by(i - 1)) ^ 2)
    Next i
    dirTxt = BearingText(bx, by, n, bearing)
    Set reaches = res("reaches")
    Set path = res("path")
    Set last = path(path.Count)
    total = CDbl(res("total_ft"))
    vals("TraceID") = traceId
    If nseg <= 1 Then
        vals("Segment") = kind
    Else
        vals("Segment") = kind & " " & seg & " of " & nseg
    End If
    If Left$(kind, 8) = "Overland" Then
        vals("Method") = CStr(res("overland_method"))
    Else
        vals("Method") = CStr(res("source"))
    End If
    If Len(dirTxt) > 0 Then
        vals("FlowDirection") = dirTxt
        vals("FlowBearing_Deg") = bearing
    End If
    vals("SegmentLength_Ft") = RoundTo(segLen, 1)
    vals("OverlandLength_Ft") = RoundTo(CDbl(res("overland_ft")), 1)
    vals("TotalLength_Ft") = RoundTo(total, 1)
    vals("TotalLength_Mi") = RoundTo(total / 5280#, 3)
    vals("MaxDownstream_Mi") = RoundTo(CDbl(res("max_miles")), 2)
    vals("ReachCount") = CDbl(reaches.Count)
    If reaches.Count > 0 Then
        If Len(CStr(reaches(1)("comid"))) > 0 Then vals("StartReachCOMID") = CStr(reaches(1)("comid"))
        If Len(CStr(reaches(reaches.Count)("comid"))) > 0 Then vals("EndReachCOMID") = CStr(reaches(reaches.Count)("comid"))
    End If
    vals("Start_Lat") = RoundTo(CDbl(res("lat")), 6)
    vals("Start_Lon") = RoundTo(CDbl(res("lon")), 6)
    vals("End_Lat") = RoundTo(CDbl(last(2)), 6)
    vals("End_Lon") = RoundTo(CDbl(last(1)), 6)
    vals("CalcDate") = Format$(Date, "yyyy-mm-dd")
    Set RaindropValues = vals
End Function

Private Function RaindropPropNames() As Variant
    RaindropPropNames = Split("TraceID,Segment,Method,FlowDirection,FlowBearing_Deg,SegmentLength_Ft,OverlandLength_Ft,TotalLength_Ft," & _
                              "TotalLength_Mi,MaxDownstream_Mi,ReachCount,StartReachCOMID,EndReachCOMID,Start_Lat,Start_Lon,End_Lat,End_Lon,CalcDate", ",")
End Function

Public Sub DrawRaindrop(ByRef res As Object, ByVal traceId As String, ByVal drawStart As Boolean)
    ' Dashed overland leg, solid downstream reaches, flow arrows, start / end markers and the legend table.
    Const PTYPES As String = "ssssdddddddssdddds"
    Dim lvl As Level, th As Double, rad As Double, ox() As Double, oy() As Double, no As Long
    Dim sx() As Double, sy() As Double, ns As Long, r As Variant, px() As Double, py() As Double, np As Long
    Dim i As Long, j As Long, el As Element, items As Long, arrows As Long, nseg As Long, seg As Long, cnt As Long
    Dim x0 As Double, y0 As Double, x1 As Double, y1 As Double, maxx As Double, maxy As Double, path As Collection, last As Variant
    Dim cx() As Double, cy() As Double
    Set path = res("path")
    If path.Count < 2 Then Exit Sub
    Set lvl = EnsureLevel("USGS_Raindrop_Path", 7, 3)
    th = TextHeightFt()
    rad = 4# * th / 7#
    maxx = -1E+300
    maxy = -1E+300
    no = ProjectLine(res("overland"), ox, oy)
    ReDim sx(0 To path.Count)
    ReDim sy(0 To path.Count)
    For Each r In res("reaches")
        np = ProjectLine(r("pts"), px, py)
        For j = 0 To np - 1
            If Not (ns > 0 And j = 0 And px(0) = sx(IIf(ns > 0, ns - 1, 0)) And py(0) = sy(IIf(ns > 0, ns - 1, 0))) Then
                If ns > UBound(sx) Then
                    ReDim Preserve sx(0 To 2 * ns + 1)
                    ReDim Preserve sy(0 To 2 * ns + 1)
                End If
                sx(ns) = px(j)
                sy(ns) = py(j)
                ns = ns + 1
            End If
        Next j
    Next r
    If no >= 2 Then
        Set el = AddLineString(ox, oy, no, lvl, RAINDROP_OVERLAND_COLOR, 2, True)
        If AttachItem(el, "RaindropPath", RaindropPropNames(), PTYPES, RaindropValues(res, traceId, "Overland leg", ox, oy, 0, no, 1, 1), True) Then items = items + 1
        arrows = arrows + FlowArrows(ox, oy, no, lvl, RAINDROP_OVERLAND_COLOR)
    End If
    If ns >= 2 Then
        nseg = (ns - 2) \ 4000 + 1
        For i = 0 To ns - 2 Step 4000
            seg = seg + 1
            cnt = IIf(ns - i > 4001, 4001, ns - i)
            Set el = AddLineString(sx, sy, cnt, lvl, 7, 3, False, i)
            If AttachItem(el, "RaindropPath", RaindropPropNames(), PTYPES, RaindropValues(res, traceId, "Downstream channel", sx, sy, i, cnt, seg, nseg), True) Then items = items + 1
            ReDim cx(0 To cnt - 1)
            ReDim cy(0 To cnt - 1)
            For j = 0 To cnt - 1
                cx(j) = sx(i + j)
                cy(j) = sy(i + j)
            Next j
            arrows = arrows + FlowArrows(cx, cy, cnt, lvl, 7)
        Next i
    End If
    LonLatToDgn CDbl(res("lon")), CDbl(res("lat")), x0, y0
    If drawStart Then
        AddCircleX x0, y0, rad, lvl, 7
        AddText traceId & " RAINDROP START", x0 + rad + 0.6 * th, y0 - 0.5 * th, th, lvl, 7
    End If
    Set last = path(path.Count)
    LonLatToDgn CDbl(last(1)), CDbl(last(2)), x1, y1
    AddCircleX x1, y1, rad, lvl, 7
    AddText traceId & " END " & Format$(CDbl(res("total_ft")) / 5280#, "0.00") & " MI", x1 + rad + 0.6 * th, y1 - 0.5 * th, th, lvl, 7
    LogLine "  Drew raindrop path (" & Format$(CDbl(res("total_ft")), "#,##0") & " ft) on level 'USGS_Raindrop_Path' with " & arrows & " flow-direction arrow(s)."
    If items > 0 Then LogLine "  Attached Item Type 'USGS_VDOT_Hydrology:RaindropPath' to " & items & " raindrop path line(s)."
    For i = 0 To no - 1
        If ox(i) > maxx Then maxx = ox(i)
        If oy(i) > maxy Then maxy = oy(i)
    Next i
    For i = 0 To ns - 1
        If sx(i) > maxx Then maxx = sx(i)
        If sy(i) > maxy Then maxy = sy(i)
    Next i
    If x0 > maxx Then maxx = x0
    If x1 > maxx Then maxx = x1
    If y0 > maxy Then maxy = y0
    If y1 > maxy Then maxy = y1
    PlaceRaindropTable res, traceId, maxx + 4# * th, maxy
End Sub

Public Sub PlaceRaindropTable(ByRef res As Object, ByVal traceId As String, ByVal tx As Double, ByVal ty As Double)
    ' Bordered table for a raindrop trace. The LEGEND column shows each line type and symbol in its own color.
    Dim rows As New Collection, row As Variant, syms() As String, colors() As Long, n As Long, i As Long, th As Double
    Dim src As Collection, title As String, m As String, comid As String
    On Error GoTo Bad
    th = TextHeightFt()
    Set src = RaindropRows(res)
    n = src.Count + 4
    ReDim syms(0 To n - 1)
    ReDim colors(0 To n - 1)
    For i = 0 To n - 1
        colors(i) = -1
    Next i
    rows.Add Array("LEGEND", "SEGMENT", "NHDPLUS COMID", "LENGTH (FT)", "CUMULATIVE (MI)", "METHOD")
    i = 1
    For Each row In src
        m = CStr(row(6))
        If Len(m) = 0 Then m = "-"
        comid = CStr(row(2))
        If Len(comid) = 0 Then comid = "-"
        If CStr(row(1)) = "OVERLAND" Then
            rows.Add Array("", "OVERLAND FLOW", comid, Format$(row(4), "#,##0"), Format$(row(5) / 5280#, "#,##0.000"), m)
            syms(i) = "rd_overland"
            colors(i) = RAINDROP_OVERLAND_COLOR
        Else
            rows.Add Array("", "STREAM REACH " & (CLng(row(0)) - 1), comid, Format$(row(4), "#,##0"), Format$(row(5) / 5280#, "#,##0.000"), m)
            syms(i) = "rd_channel"
            colors(i) = 7
        End If
        i = i + 1
    Next row
    rows.Add Array("", "TOTAL", "-", Format$(CDbl(res("total_ft")), "#,##0"), Format$(CDbl(res("total_ft")) / 5280#, "#,##0.000"), _
                   "Max downstream " & NumStr(CDbl(res("max_miles"))) & " mi")
    i = i + 1
    rows.Add Array("", "FLOW DIRECTION", "", "", "", "Arrowheads point downstream")
    syms(i) = "rd_arrow"
    colors(i) = 7
    i = i + 1
    rows.Add Array("", "START / END POINT", "", "", "", "Raindrop start and trace end")
    syms(i) = "rd_point"
    colors(i) = 7
    title = traceId & " FLOW (RAINDROP) PATH  -  start " & NumFix(CDbl(res("lat")), 6) & ", " & NumFix(CDbl(res("lon")), 6) & _
            "  (" & Format$(Date, "yyyy-mm-dd") & ")"
    BuildTable rows, tx, ty, title, th, syms, colors, 6# * th
    LogLine "  Placed raindrop table for " & traceId & " (" & (rows.Count - 1) & " rows incl. legend)."
    Exit Sub
Bad:
    LogLine "  Raindrop table skipped (" & Err.Description & ")"
End Sub
