Attribute VB_Name = "modStreamStats"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology - MicroStation VBA, Phase 2 (engine)
'
' Macros:
'   StreamStatsShow      Opens the main window (tabs like the Python tool).
'   StreamStatsCulvert   No window: click a culvert location; the macro delineates the watershed with USGS
'                        StreamStats, computes VDOT design peak flows, draws the drainage area and
'                        attaches the DrainageArea and DataSources Item Types.
'   StreamStatsSetZone   Choose the Virginia State Plane zone (South or North) of the DGN.
'   StreamStatsSelfTest  Runs a fixed test point without any prompts (results go to the log file).
'
' Phase 1 covers: pour-point snap, delineation, StreamStats basin characteristics, NOAA Atlas 14
' rainfall, the Rational method (<= 200 ac) and the USGS rural regression (NSS, > 200 ac).
' Not yet in VBA (still in the Python tool): NRCS, Anderson and urban methods, longest flowpath,
' streams, fence, raindrop path, batch queue, tables, GIS/Excel export, alternate data sources.
'
' VBA cannot read the DGN geographic coordinate system, so the DGN must be in NAD83 Virginia
' State Plane (South or North zone) with US survey feet as master units.
'==============================================================================================

Private Const APP_NAME As String = "USGS StreamStats VDOT (VBA)"
Private Const APP_VERSION As String = "1.0.1"
Private Const BUILD_DATE As String = "2026-11-17"
Private Const BUILT_BY As String = "Robert Ratliff"
Private Const REGION As String = "VA"
Private Const USER_AGENT As String = "MicroStation-USGS-VDOT-Automation/4.0"
Private Const URL_SNAP As String = "https://streamstats.usgs.gov/pourpoint/v2/snap/streamgrid/"
Private Const URL_DELINEATE As String = "https://streamstats.usgs.gov/ss-delineate/v1/delineate/sshydro/"
Private Const URL_BASINCHAR As String = "https://streamstats.usgs.gov/ss-hydro/v1/basin-characteristics/calculate"
Private Const URL_NSS As String = "https://streamstats.usgs.gov/nssservices"
Private Const URL_ATLAS14 As String = "https://hdsc.nws.noaa.gov/cgi-bin/hdsc/new/cgi_readH5.py"
Private Const LIB_NAME As String = "USGS_VDOT_Hydrology"
Private Const LEVEL_BASIN As String = "USGS_StreamStats_Basin"
Private Const MAX_SHAPE_VERTICES As Long = 4500
Private Const TEST_LON As Double = -81.818623
Private Const TEST_LAT As Double = 36.625806

' Lambert Conformal Conic constants for the chosen State Plane zone (metres / radians)
Private mLccReady As Boolean
Private mLccN As Double, mLccF As Double, mLccRho0 As Double, mLccLon0 As Double
Private mLccFE As Double, mLccFN As Double
Private Const GRS80_A As Double = 6378137#
Private Const GRS80_INVF As Double = 298.257222101
Private Const FT_PER_M As Double = 3.28083333333333      ' US survey foot (3937 / 1200)

' JSON parser state
Private mJ As String, mP As Long, mLen As Long

' Results of the last run
Private mLogText As String
Private mLastElementId As String
Private mNssDownUntil As Date      ' the regression lookup is skipped until this time after it fails
Private mBcDownUntil As Date       ' same for StreamStats basin characteristics

' Shared with the window (frmStreamStats)
Public gLogSink As Object          ' object with an AppendLog(text) method; Nothing = log file only
Public gSilent As Boolean          ' True = no message boxes (self-tests)
Public gCancel As Boolean          ' set True to stop a running batch
Public gCOverride As Double        ' Rational C override; 0 = estimate from land cover
Public gDrawingScale As Double     ' feet per plotted inch (50, 100 or 200)
Public gPlotTextIn As Double       ' plotted text height, inches
Public gRecIds As Collection        ' when set, AddStyled records the IDs of the elements it adds

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As String
    pvReserved As LongPtr
    dwReserved As Long
    FlagsEx As Long
End Type
Private Declare PtrSafe Function GetOpenFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long

'----------------------------------------------------------------------------------------------
' Public macros
'----------------------------------------------------------------------------------------------
Public Sub StreamStatsShow()
    ' Opens the main window (same tabs as the Python tool).
    frmStreamStats.Show vbModeless
    frmStreamStats.StartupCheck
End Sub

Public Sub StreamStatsCulvert()
    ' One culvert without the window: click a point, get the flows, draw the drainage area.
    Dim id As String, lon As Double, lat As Double
    Dim msg As CadInputMessage, res As Object
    gSilent = False
    gCancel = False
    If Not EnsureZone() Then Exit Sub
    id = InputBox("Culvert ID:", APP_NAME, "CULV-01")
    If Len(Trim$(id)) = 0 Then Exit Sub
    ShowCommand APP_NAME
    ShowPrompt "Click the culvert location (Reset to cancel)"
    Set msg = CadInputQueue.GetInput(msdCadInputTypeDataPoint, msdCadInputTypeReset, _
                                     msdCadInputTypeReset, msdCadInputTypeReset)
    If msg.InputType <> msdCadInputTypeDataPoint Then
        ShowPrompt "Cancelled."
        Exit Sub
    End If
    DgnToLonLat msg.Point.x, msg.Point.y, lon, lat
    If ComputeCulvert(lon, lat, Trim$(id), res) Then
        DrawCulvert res
        MsgBox SummaryText(res), vbInformation, APP_NAME
    Else
        MsgBox CStr(res("error")), vbExclamation, APP_NAME
    End If
End Sub

Public Sub StreamStatsSetZone()
    Dim z As String
    z = InputBox("Virginia State Plane zone of this DGN (NAD83, US survey feet)." & vbCrLf & _
                 "Enter S for South or N for North:", APP_NAME, GetZone())
    z = UCase$(Left$(Trim$(z), 1))
    If z = "S" Or z = "N" Then
        SaveSetting "USGS_StreamStats_VBA", "Settings", "Zone", z
        mLccReady = False
        MsgBox "State Plane zone set to " & IIf(z = "S", "South", "North") & ".", vbInformation, APP_NAME
    End If
End Sub

Public Sub StreamStatsSelfTest()
    Dim res As Object
    gSilent = True
    gCancel = False
    If ComputeCulvert(TEST_LON, TEST_LAT, "VBA-TEST", res) Then DrawCulvert res
End Sub

'----------------------------------------------------------------------------------------------
' Engine (used by the macros above and by the window)
'----------------------------------------------------------------------------------------------
Public Function AppBuildDate() As String
    AppBuildDate = BUILD_DATE
End Function

Public Function AppBuiltBy() As String
    AppBuiltBy = BUILT_BY
End Function

Public Function AppVersion() As String
    AppVersion = APP_VERSION
End Function

Public Function TextHeightFt() As Double
    ' DGN text height (ft) = plotted height (in) x drawing scale (ft per in)
    If gDrawingScale <= 0# Then gDrawingScale = 100#
    If gPlotTextIn <= 0# Then gPlotTextIn = 0.2
    TextHeightFt = gDrawingScale * gPlotTextIn
End Function

Public Sub DgnToLonLat(ByVal x As Double, ByVal y As Double, ByRef lon As Double, ByRef lat As Double)
    LccInit GetZone()
    XYToLonLat x, y, lon, lat
End Sub

Public Sub LonLatToDgn(ByVal lon As Double, ByVal lat As Double, ByRef x As Double, ByRef y As Double)
    LccInit GetZone()
    LonLatToXY lon, lat, x, y
End Sub

Public Function ComputeCulvert(ByVal lon As Double, ByVal lat As Double, ByVal culvertId As String, ByRef res As Object) As Boolean
    ' Delineates the watershed and computes the flows. Results (or res("error")) come back in res.
    Dim sLon As Double, sLat As Double
    Dim stage As String, wsText As String, geomJson As String, wsId As String
    Dim ws As Variant, rings As Collection, outer As Collection, params As Object
    Dim haveBc As Boolean, haveAtlas As Boolean, haveNss As Boolean, snapped As Boolean
    Dim atlas() As Double, qUsgs(0 To 7) As Double, qRat(0 To 7) As Double, qDes(0 To 7) As Double
    Dim inten(0 To 7) As Double
    Dim sqmi As Double, acres As Double, slope As Double, impPct As Double, forest As Double
    Dim cVal As Double, tcMin As Double, lengthFt As Double
    Dim i As Long, yrs As Variant, cf As Variant, gov As String
    Dim t0 As Single
    Dim nlcd As Object, hsg As Object, haveNlcd As Boolean, hsgSrc As String, hsgTxt As String
    Dim urbanPct As Double, urbanSrc As String, cn As Double, cnSrc As String, k As Variant, m As Variant, v As Double
    Dim useRat As Boolean, useNrcs As Boolean, useRural As Boolean, useUrban As Boolean, order As Variant, ord2() As String
    Dim qNrcs(0 To 7) As Double, qUrb(0 To 7) As Double, rain24(0 To 7) As Double, picks(0 To 7) As String, qRatAny As Double
    Dim n As Long
    Dim bigRing As Variant, nhdLfp As Collection, lfp As Collection, lfpFt As Double, lfpMethod As String, tcLen As Double
    Dim coverSrc As String, slopeSrc As String, hsgKind As String, v1 As Double, v2 As Double, comid As String
    Dim nssParams As Object, useAnd As Boolean, qAnd(0 To 7) As Double, bcSrc As String

    Set res = CreateObject("Scripting.Dictionary")
    On Error GoTo Failed
    If Not gMethodsInit Then MethodDefaults
    t0 = Timer
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    cf = Array(1#, 1#, 1#, 1.1, 1.2, 1.25, 1.25, 1.25)
    LogLine "Calling USGS StreamStats for " & culvertId & " (Lon " & NumFix(lon, 5) & ", Lat " & NumFix(lat, 5) & ")..."

    stage = "checking the location"
    If lat < 36.4 Or lat > 39.6 Or lon < -83.8 Or lon > -75.1 Then
        Err.Raise vbObjectError + 1, , "The point is outside Virginia. Check the State Plane zone " & _
            "(Settings tab) and that the DGN master units are US survey feet."
    End If

    stage = "snapping the pour point"
    snapped = SnapPourPoint(lon, lat, sLon, sLat)
    If snapped Then
        LogLine "  Pour point snapped to the stream grid: Lon " & NumFix(sLon, 5) & ", Lat " & NumFix(sLat, 5)
    Else
        LogLine "  [WARNING] The point did not snap to a StreamStats stream cell; results may be invalid."
    End If

    stage = "delineating the watershed (USGS StreamStats)"
    ShowStatus "Delineating watershed for " & culvertId & "..."
    wsText = HttpText("GET", URL_DELINEATE & REGION & "?lat=" & NumFix(sLat, 6) & "&lon=" & NumFix(sLon, 6), "", 240)
    JsonParse wsText, ws
    Set rings = New Collection
    Set outer = New Collection
    ExtractWatershed ws, rings, outer, wsId
    If rings.Count = 0 Then Err.Raise vbObjectError + 2, , "StreamStats returned no watershed polygon for this point."
    geomJson = WatershedGeometryJson(wsText)

    stage = "reading basin characteristics"
    ShowStatus "Computing basin characteristics for " & culvertId & "..."
    Set params = CreateObject("Scripting.Dictionary")
    haveBc = BasinCharacteristics(wsText, params)
    If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"

    stage = "reading land cover and soils"
    ShowStatus "Reading land cover and soils for " & culvertId & "..."
    haveNlcd = NlcdBasinCover(rings, nlcd)
    ' Soil groups feed the NRCS curve number and the drainage-area Item Type, so they are read for every basin.
    If SoilGroups(rings, hsg) Then
        hsgSrc = "USDA SSURGO (Soil Data Access); dual groups counted as D"
        hsgKind = "USDA SSURGO"
    ElseIf Not gCancel Then
        If Issr800Groups(rings, hsg) Then
            hsgSrc = "UC Davis SoilWeb ISSR-800 grid (800 m, from SSURGO); dual groups counted as D"
            hsgKind = "UC Davis SoilWeb ISSR-800 grid"
            hsgTxt = ""
            For Each k In Array("A", "B", "C", "D")
                If hsg.Exists(k) Then hsgTxt = hsgTxt & IIf(Len(hsgTxt) > 0, ", ", "") & k & " " & Format$(hsg(k) * 100#, "0") & "%"
            Next k
            LogLine "  Soil groups (UC Davis SoilWeb ISSR-800 grid): " & hsgTxt & "."
        End If
    End If
    If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
    If Not params.Exists("DRNAREA") Then
        params("DRNAREA") = PolygonAreaSqMi(rings, outer)
        LogLine "  Drainage area computed from the basin polygon: " & Format$(params("DRNAREA") * 640#, "#,##0.0") & " ac."
    End If
    Set bigRing = LargestRing(rings)
    bcSrc = "USGS StreamStats"
    If Not haveBc Then
        ' StreamStats basin characteristics are unavailable: land cover and channel slope from other public services.
        stage = "reading alternate basin data"
        bcSrc = "Alternate (MRLC NLCD 2021 / USGS 3DEP / basin polygon)"
        LogLine "  StreamStats basin characteristics are down; using alternate sources (MRLC NLCD 2021, USGS 3DEP)."
        If haveNlcd Then
            params("LC21IMP") = RoundTo(CDbl(nlcd("imperv_pct")), 2)
            params("LC21FOREST") = RoundTo(CDbl(nlcd("forest_pct")), 2)
            params("LC21DEV") = RoundTo(CDbl(nlcd("dev_pct")), 2)
            coverSrc = "nlcd"
        ElseIf StreamCatCover(sLon, sLat, v1, v2, comid) Then
            params("LC19IMP") = RoundTo(v1, 2)
            params("LC19FOREST") = RoundTo(v2, 2)
            coverSrc = "streamcat"
            LogLine "  EPA StreamCat (NHDPlus watershed " & comid & "): impervious " & Format$(v1, "0.0") & "%, forest " & Format$(v2, "0.0") & "%."
        End If
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
        Set nhdLfp = SafeNhdLfp(bigRing, sLon, sLat)
        If Not nhdLfp Is Nothing Then
            If ChannelSlope1085(nhdLfp, sLon, sLat, v1, v2) Then
                params("CSL10_85") = RoundTo(v1, 2)
                slopeSrc = "3dep"
                LogLine "  Channel slope 10-85 (USGS 3DEP): " & Format$(v1, "#,##0.0") & " ft/mi along a " & Format$(v2, "#,##0") & " ft flowpath."
            End If
        End If
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
    End If

    stage = "computing the longest flowpath"
    If gComputeLfp Then
        ShowStatus "Computing the longest flowpath for " & culvertId & "..."
        LogLine "  Computing longest flowpath..."
        Set lfp = UsgsLongestFlowpath(wsText)
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
        If Not lfp Is Nothing Then
            lfpMethod = "USGS StreamStats"
        Else
            If nhdLfp Is Nothing Then Set nhdLfp = SafeNhdLfp(bigRing, sLon, sLat)
            Set lfp = nhdLfp
            If Not lfp Is Nothing Then
                If lfp.Count > 2 Then
                    lfpMethod = "NHD channel + overland (approx.)"
                Else
                    lfpMethod = "Straight line divide-to-outlet (approx.)"
                End If
            End If
        End If
        If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"
        If lfp Is Nothing Then
            LogLine "  [WARNING] Longest flowpath could not be determined."
        Else
            lfpFt = RoundTo(LineLenFt(lfp), 1)
            tcLen = lfpFt
            LogLine "  Longest flowpath: " & Format$(lfpFt, "#,##0") & " ft (" & Format$(lfpFt / 5280#, "0.00") & " mi) via " & lfpMethod & "."
        End If
    ElseIf Not nhdLfp Is Nothing Then
        tcLen = RoundTo(LineLenFt(nhdLfp), 1)
    End If

    stage = "computing basin metrics"
    BasinMetrics params, sqmi, acres, slope, impPct, forest
    If gCOverride > 0# Then
        cVal = MinD(MaxD(gCOverride, 0.05), 0.98)
    Else
        cVal = RoundTo(0.9 * impPct / 100# + 0.2 * forest / 100# + 0.35 * MaxD(0#, 1# - impPct / 100# - forest / 100#), 2)
        cVal = MinD(MaxD(cVal, 0.15), 0.95)
    End If
    ' Kirpich Tc: the measured longest-flowpath length when there is one, else a length estimated from the area.
    If tcLen > 0# Then
        lengthFt = tcLen
    Else
        lengthFt = 1.4 * (MaxD(sqmi, 0.005) ^ 0.6) * 5280#
    End If
    tcMin = 0.0078 * (lengthFt ^ 0.77) * (MaxD(slope / 5280#, 0.0005) ^ -0.385)
    tcMin = RoundTo(MinD(MaxD(tcMin, 5#), 360#), 1)

    stage = "reading NOAA Atlas 14 rainfall"
    ShowStatus "Reading NOAA Atlas 14 rainfall for " & culvertId & "..."
    haveAtlas = Atlas14Intensities(sLon, sLat, atlas)

    stage = "computing USGS regression peak flows"
    For i = 0 To 7
        qUsgs(i) = -1#
    Next i
    If Len(geomJson) > 0 Then
        ShowStatus "Estimating USGS regression peak flows for " & culvertId & "..."
        If haveBc Then
            haveNss = NssPeakFlows(geomJson, params, qUsgs)
        Else
            ' Alternate slope and land cover are approximations, so only the drainage area feeds the regression.
            Set nssParams = CreateObject("Scripting.Dictionary")
            nssParams("DRNAREA") = params("DRNAREA")
            haveNss = NssPeakFlows(geomJson, nssParams, qUsgs)
        End If
    End If
    If gCancel Then Err.Raise vbObjectError + 4, , "stopped by user"

    urbanPct = -1#
    For Each k In params.keys
        If Len(CStr(k)) = 7 And Left$(CStr(k), 2) = "LC" And Right$(CStr(k), 3) = "DEV" Then urbanPct = CDbl(params(k))
    Next k
    urbanSrc = "USGS StreamStats"
    If urbanPct < 0# And haveNlcd Then
        urbanPct = CDbl(nlcd("dev_pct"))
        urbanSrc = "NLCD 2021 developed classes 21-24"
    End If
    If gCN > 0# Then
        cn = gCN
        cnSrc = "Entered in Settings"
    ElseIf haveNlcd Then
        cn = EstimateCN(nlcd("classes"), hsg, IIf(Len(hsgKind) > 0, hsgKind, "USDA SSURGO"), hsgTxt)
        cnSrc = "Estimated: NLCD 2021 land cover, " & hsgTxt
    End If

    stage = "computing VDOT design flows"
    ' Which methods run (VDOT Table 6-2 size ranges), and the order they supply the design flow:
    ' Rational up to 200 ac; above that the USGS regression, urban when the basin is 30% or more developed.
    useRat = gUseRational And (Not gUseRanges Or (acres > 0# And acres <= 200#))
    useNrcs = gUseNrcs And (Not gUseRanges Or acres <= 2000#) And cn > 0# And haveAtlas
    useRural = gUseRural And (Not gUseRanges Or acres > 200#)
    useUrban = gUseUrban And (Not gUseRanges Or acres > 200#) And urbanPct >= 10# And sqmi >= 1.2 And sqmi <= 2400#
    useAnd = gUseAnderson And (Not gUseRanges Or acres > 200#) And tcLen > 0#
    If acres > 0# And acres <= 200# Then
        order = Array("rational", "nrcs", "rural", "urban", "anderson")
    ElseIf urbanPct >= 30# Then
        order = Array("urban", "rural", "nrcs", "anderson")
    Else
        order = Array("rural", "urban", "nrcs", "anderson")
    End If
    ReDim ord2(0 To UBound(order) + 1)
    n = 0
    If gDesign <> "auto" And Len(gDesign) > 0 Then
        ord2(0) = gDesign
        n = 1
    End If
    For Each m In order
        If CStr(m) <> gDesign Then
            ord2(n) = CStr(m)
            n = n + 1
        End If
    Next m
    For i = 0 To 7
        inten(i) = -1#
        rain24(i) = -1#
        If haveAtlas Then
            inten(i) = AtlasIntensity(atlas, CLng(yrs(i)), tcMin)
            rain24(i) = AtlasIntensity(atlas, CLng(yrs(i)), 1440#) * 24#
        End If
        If inten(i) <= 0# Then inten(i) = BuiltinIntensity(CLng(yrs(i)), tcMin)
        qRatAny = RoundTo(MinD(cf(i) * cVal, 1#) * inten(i) * acres, 1)
        qRat(i) = -1#
        qNrcs(i) = -1#
        qUrb(i) = -1#
        qAnd(i) = -1#
        If useRat Then qRat(i) = qRatAny
        If useNrcs And rain24(i) > 0# Then qNrcs(i) = RoundTo(NrcsQ(rain24(i), cn, sqmi, tcMin), 1)
        If useUrban Then qUrb(i) = RoundTo(UrbanQ(CLng(yrs(i)), sqmi, urbanPct), 1)
        If useAnd Then
            v = AndersonQ(CLng(yrs(i)), sqmi, tcLen / 5280#, slope, impPct, gAndersonType)
            If v > 0# Then qAnd(i) = RoundTo(v, 1)
        End If
        picks(i) = ""
        For n = 0 To UBound(ord2)
            v = -1#
            Select Case ord2(n)
                Case "rational"
                    v = qRat(i)
                Case "nrcs"
                    v = qNrcs(i)
                Case "rural"
                    If useRural Then v = qUsgs(i)
                Case "urban"
                    v = qUrb(i)
                Case "anderson"
                    v = qAnd(i)
            End Select
            If v > 0# Then
                picks(i) = ord2(n)
                qDes(i) = RoundTo(v, 1)
                Exit For
            End If
        Next n
        If Len(picks(i)) = 0 Then      ' nothing enabled gave a value: Rational up to 200 ac, else the check equation
            If acres > 0# And acres <= 200# Then
                picks(i) = "rational"
                qDes(i) = qRatAny
            Else
                picks(i) = "builtin"
                qDes(i) = BuiltinRegression(CLng(yrs(i)), sqmi)
            End If
        End If
    Next i
    gov = MethodLabel(picks(2), acres)

    res("culvert_id") = culvertId
    res("lon") = lon
    res("lat") = lat
    res("s_lon") = sLon
    res("s_lat") = sLat
    res("ws_id") = wsId
    Set res("rings") = rings
    Set res("outer") = outer
    res("sqmi") = sqmi
    res("acres") = acres
    res("slope") = slope
    res("imp") = impPct
    res("forest") = forest
    res("c") = cVal
    res("tc") = tcMin
    res("gov") = gov
    res("have_bc") = haveBc
    res("have_atlas") = haveAtlas
    res("have_nss") = haveNss
    res("q_usgs") = qUsgs
    res("q_rat") = qRat
    res("q_des") = qDes
    res("inten") = inten
    res("q_nrcs") = qNrcs
    res("q_urban") = qUrb
    res("picks") = picks
    res("rain24") = rain24
    res("cn") = cn
    res("cn_src") = cnSrc
    res("urban_pct") = urbanPct
    res("urban_src") = urbanSrc
    res("have_nlcd") = haveNlcd
    Set res("hsg") = hsg
    res("hsg_src") = hsgSrc
    res("hsg_kind") = hsgKind
    res("q_and") = qAnd
    res("bc_src") = bcSrc
    res("cover_src") = coverSrc
    res("slope_src") = slopeSrc
    Set res("lfp") = lfp
    res("lfp_ft") = lfpFt
    res("lfp_method") = lfpMethod
    res("tc_len") = tcLen
    res("tc_len_src") = IIf(tcLen > 0#, "Measured longest flowpath", "Estimated from drainage area")
    res("drawn") = False
    LogLine "  Tc " & Format$(tcMin, "0.0") & " min (" & LCase$(CStr(res("tc_len_src"))) & "); basin data: " & bcSrc & "."
    LogLine "  " & culvertId & ": " & Format$(acres, "#,##0.0") & " ac, slope " & Format$(slope, "0.0") & " ft/mi, C " & _
            Format$(cVal, "0.00") & ", Tc " & Format$(tcMin, "0.0") & " min, Q10 " & Format$(qDes(2), "#,##0.0") & " cfs (" & gov & _
            "); rainfall: " & IIf(haveAtlas, "NOAA Atlas 14", "built-in IDF (placeholder)") & ". " & Format$(Timer - t0, "0") & " s."
    ShowStatus "StreamStats: " & culvertId & " done - Q10 " & Format$(qDes(2), "#,##0.0") & " cfs"
    ComputeCulvert = True
    Exit Function

Failed:
    res("error") = "While " & stage & ": " & Err.Description
    LogLine "  [ERROR] " & culvertId & " - " & CStr(res("error"))
    ShowStatus "StreamStats failed: " & Err.Description
End Function

Private Function SafeNhdLfp(ByRef ring As Variant, ByVal lon As Double, ByVal lat As Double) As Collection
    On Error GoTo Bad
    Set SafeNhdLfp = NhdLongestFlowpath(ring, lon, lat)
    If Not SafeNhdLfp Is Nothing Then
        If SafeNhdLfp.Count < 2 Then Set SafeNhdLfp = Nothing
    End If
    Exit Function
Bad:
    LogLine "  [WARNING] Longest flowpath failed: " & Err.Description
End Function

Public Sub SetServiceDown(ByVal key As String, ByVal untilTime As Date)
    ' Used by the service check: untilTime 0 = try the service again at once.
    If key = "nss" Then mNssDownUntil = untilTime
    If key = "bc" Then mBcDownUntil = untilTime
End Sub

Public Sub DrawCulvert(ByRef res As Object)
    ' Draws the drainage area for a computed culvert and attaches the Item Types.
    Dim vals As Object, src As Object, rings As Collection
    Dim i As Long, yrs As Variant, qU As Variant, qD As Variant, d As String
    Dim hsg As Object, k As Variant, w As Double, best As String, bestV As Double, qB As Variant, anyUrban As Boolean
    Dim qA As Variant, anyAnd As Boolean
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    qU = res("q_usgs")
    qD = res("q_des")
    qB = res("q_urban")
    d = Format$(Date, "yyyy-mm-dd")
    Set vals = CreateObject("Scripting.Dictionary")
    vals("CulvertID") = CStr(res("culvert_id"))
    vals("StateCode") = REGION
    vals("WorkspaceID") = IIf(Len(CStr(res("ws_id"))) > 0, CStr(res("ws_id")), "N/A")
    vals("DrainageArea_SqMi") = RoundTo(CDbl(res("sqmi")), 4)
    vals("DrainageArea_Acres") = RoundTo(CDbl(res("acres")), 2)
    vals("Slope_FtPerMi") = RoundTo(CDbl(res("slope")), 2)
    vals("Impervious_Pct") = RoundTo(CDbl(res("imp")), 2)
    vals("Forest_Pct") = RoundTo(CDbl(res("forest")), 2)
    vals("RunoffCoeff_C") = CDbl(res("c"))
    vals("Tc_Min") = CDbl(res("tc"))
    vals("GoverningMethod") = CStr(res("gov"))
    vals("CalcDate") = d
    If Not res("lfp") Is Nothing Then
        vals("LongestFlowpath_Ft") = CDbl(res("lfp_ft"))
        vals("LongestFlowpath_Method") = CStr(res("lfp_method"))
    End If
    If res("cn") > 0# Then
        vals("CurveNumber_CN") = RoundTo(CDbl(res("cn")), 1)
        vals("CurveNumber_Source") = CStr(res("cn_src"))
    End If
    If Not res("hsg") Is Nothing Then
        Set hsg = res("hsg")
        best = ""
        bestV = -1#
        For Each k In Array("A", "B", "C", "D")
            w = 0#
            If hsg.Exists(k) Then w = CDbl(hsg(k))
            vals("SoilGroup_" & k & "_Pct") = RoundTo(w * 100#, 1)
            If w > bestV Then
                bestV = w
                best = CStr(k)
            End If
        Next k
        vals("DominantSoilGroup") = best
        vals("SoilSource") = CStr(res("hsg_src"))
    End If
    For i = 0 To 7
        If qU(i) > 0# Then vals("Q" & yrs(i) & "_USGS_cfs") = RoundTo(CDbl(qU(i)), 1)
        vals("Q" & yrs(i) & "_VDOT_cfs") = CDbl(qD(i))
    Next i
    Set src = CreateObject("Scripting.Dictionary")
    src("CulvertID") = CStr(res("culvert_id"))
    src("CalcDate") = d
    src("Watershed_Source") = "USGS StreamStats"
    src("Watershed_Website") = "https://streamstats.usgs.gov"
    If res("have_bc") Then
        src("BasinCharacteristics_Source") = "USGS StreamStats"
        src("BasinCharacteristics_Website") = "https://streamstats.usgs.gov"
    End If
    If res("have_nss") Then
        src("RuralRegression_Source") = "USGS National Streamflow Statistics (NSS)"
        src("RuralRegression_Website") = "https://streamstats.usgs.gov/nssservices"
    End If
    If res("have_atlas") Then
        src("Rainfall_Source") = "NOAA Atlas 14 Precipitation Frequency Data Server"
        src("Rainfall_Website") = "https://hdsc.nws.noaa.gov/pfds"
    End If
    If res("have_nlcd") Then
        src("LandCover_Source") = "MRLC National Land Cover Database 2021"
        src("LandCover_Website") = "https://www.mrlc.gov"
    ElseIf res("cover_src") = "streamcat" Then
        src("LandCover_Source") = "EPA StreamCat (NLCD 2019, NHDPlus watershed)"
        src("LandCover_Website") = "https://www.epa.gov/national-aquatic-resource-surveys/streamcat-dataset"
    End If
    If res("slope_src") = "3dep" Then
        src("ChannelSlope_Source") = "USGS 3DEP elevation"
        src("ChannelSlope_Website") = "https://elevation.nationalmap.gov"
    End If
    src("Streams_Source") = "USGS National Hydrography Dataset (NHDPlus HR)"
    src("Streams_Website") = "https://hydro.nationalmap.gov"
    If InStr(CStr(res("lfp_method")), "StreamStats") > 0 Then
        src("LongestFlowpath_Source") = "USGS StreamStats"
        src("LongestFlowpath_Website") = "https://streamstats.usgs.gov"
    ElseIf Len(CStr(res("lfp_method"))) > 0 Then
        src("LongestFlowpath_Source") = "Traced from USGS NHD flowlines"
        src("LongestFlowpath_Website") = "https://hydro.nationalmap.gov"
    End If
    If Not res("hsg") Is Nothing Then
        If InStr(CStr(res("hsg_src")), "UC Davis") > 0 Then
            src("SoilGroups_Source") = "UC Davis SoilWeb ISSR-800 grid (from SSURGO)"
            src("SoilGroups_Website") = "https://casoilresource.lawr.ucdavis.edu"
        Else
            src("SoilGroups_Source") = "USDA NRCS SSURGO (Soil Data Access)"
            src("SoilGroups_Website") = "https://sdmdataaccess.sc.egov.usda.gov"
        End If
    End If
    qA = res("q_and")
    For i = 0 To 7
        If qA(i) > 0# Then anyAnd = True
    Next i
    If anyAnd Then
        src("Anderson_Source") = "USGS Water-Supply Paper 2001-C"
        src("Anderson_Website") = "https://pubs.usgs.gov/publication/wsp01C"
    End If
    For i = 0 To 7
        If qB(i) > 0# Then anyUrban = True
    Next i
    If anyUrban Then
        src("UrbanRegression_Source") = "USGS SIR 2014-5090 equations"
        src("UrbanRegression_Website") = "https://pubs.usgs.gov/publication/sir20145090"
    End If
    src("Methods_Source") = "VDOT Drainage Manual, Chapter 6 (Rev. 7/26)"
    src("Methods_Website") = "https://www.vdot.virginia.gov"
    Set rings = res("rings")
    ShowStatus "Drawing the drainage area..."
    DrawBasin rings, vals, src
    DrawFlowpath res
    res("drawn") = True
    res("element_id") = mLastElementId
End Sub

Public Function SummaryText(ByRef res As Object) As String
    Dim s As String, i As Long, yrs As Variant, qU As Variant, qR As Variant, qD As Variant, qN As Variant
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    qN = res("q_nrcs")
    qU = res("q_usgs")
    qR = res("q_rat")
    qD = res("q_des")
    s = "Culvert " & CStr(res("culvert_id")) & vbCrLf & _
        "Drainage area: " & Format$(res("acres"), "#,##0.0") & " ac  (" & Format$(res("sqmi"), "0.000") & " sq mi)" & vbCrLf & _
        "Slope: " & Format$(res("slope"), "0.0") & " ft/mi     C: " & Format$(res("c"), "0.00") & "     Tc: " & Format$(res("tc"), "0.0") & " min" & vbCrLf & _
        "Design method: " & CStr(res("gov")) & vbCrLf & _
        "Rainfall: " & IIf(res("have_atlas"), "NOAA Atlas 14", "built-in IDF coefficients (placeholder)") & vbCrLf & _
        "Basin data: " & IIf(res("have_bc"), "USGS StreamStats", "polygon area and defaults (StreamStats basin characteristics unavailable)") & vbCrLf & vbCrLf & _
        "Event      USGS rural     Rational     NRCS     VDOT design (cfs)"
    For i = 0 To 7
        s = s & vbCrLf & "Q" & yrs(i) & vbTab & vbTab & IIf(qU(i) > 0#, Format$(qU(i), "#,##0.0"), "N/A") & _
            vbTab & vbTab & IIf(qR(i) > 0#, Format$(qR(i), "#,##0.0"), "-") & vbTab & vbTab & _
            IIf(qN(i) > 0#, Format$(qN(i), "#,##0.0"), "-") & vbTab & vbTab & Format$(qD(i), "#,##0.0")
    Next i
    SummaryText = s
End Function

Public Sub DrawMarkers(ByVal ix As Double, ByVal iy As Double, ByVal px As Double, ByVal py As Double, ByVal culvertId As String)
    ' Culvert inlet (square with an X) at ix, iy and pour point (circle with crosshairs) at px, py.
    ' Sizes follow the Settings drawing scale, like the Python tool.
    Dim th As Double, u As Double, h As Double, rad As Double
    Dim lvlI As Level, lvlP As Level, pts(0 To 4) As Point3d, el As Element
    th = TextHeightFt()
    u = th / 4#
    Set lvlI = EnsureLevel("USGS_Culvert_Inlet", 6, 2)
    Set lvlP = EnsureLevel("USGS_StreamStats_Point", 2, 3)
    h = 1.6 * u
    pts(0) = Point3dFromXY(ix - h, iy - h)
    pts(1) = Point3dFromXY(ix + h, iy - h)
    pts(2) = Point3dFromXY(ix + h, iy + h)
    pts(3) = Point3dFromXY(ix - h, iy + h)
    pts(4) = pts(0)
    Set el = CreateLineElement1(Nothing, pts)
    AddStyled el, lvlI, 6, 2
    AddLine ix - h, iy - h, ix + h, iy + h, lvlI, 6, 1
    AddLine ix + h, iy - h, ix - h, iy + h, lvlI, 6, 1
    AddText culvertId & " INLET  E " & Format$(ix, "#,##0.0") & "  N " & Format$(iy, "#,##0.0"), ix + 4# * u, iy - 1.5 * th, th, lvlI, 6
    rad = 2.5 * u
    Set el = CreateEllipseElement2(Nothing, Point3dFromXY(px, py), rad, rad, Matrix3dIdentity, msdFillModeNotFilled)
    AddStyled el, lvlP, 2, 3
    AddLine px - 1.5 * rad, py, px + 1.5 * rad, py, lvlP, 2, 1
    AddLine px, py - 1.5 * rad, px, py + 1.5 * rad, lvlP, 2, 1
    AddText culvertId & " PP", px + 1.6 * rad, py + 0.3 * th, th, lvlP, 2
End Sub

Public Sub AddStyled(ByVal el As Element, ByVal lvl As Level, ByVal colorIndex As Long, ByVal weight As Long)
    Set el.Level = lvl
    el.Color = colorIndex
    el.LineWeight = weight
    ActiveModelReference.AddElement el
    If Not gRecIds Is Nothing Then gRecIds.Add DLongToString(el.id)
End Sub

Public Sub AddLine(ByVal x1 As Double, ByVal y1 As Double, ByVal x2 As Double, ByVal y2 As Double, _
                    ByVal lvl As Level, ByVal colorIndex As Long, ByVal weight As Long)
    Dim el As Element
    Set el = CreateLineElement2(Nothing, Point3dFromXY(x1, y1), Point3dFromXY(x2, y2))
    AddStyled el, lvl, colorIndex, weight
End Sub

Public Sub AddText(ByVal s As String, ByVal x As Double, ByVal y As Double, ByVal th As Double, _
                    ByVal lvl As Level, ByVal colorIndex As Long)
    ' Annotation scale is switched off while the text is created; otherwise a model with annotation
    ' scale on multiplies the height by its scale factor.
    Dim t As TextElement, wasOn As Boolean
    wasOn = ActiveSettings.AnnotationScaleEnabled
    On Error GoTo Restore
    ActiveSettings.AnnotationScaleEnabled = False
    Set t = CreateTextElement1(Nothing, s, Point3dFromXY(x, y), Matrix3dIdentity)
    t.TextStyle.Height = th
    t.TextStyle.Width = th
    AddStyled t, lvl, colorIndex, 0
Restore:
    ActiveSettings.AnnotationScaleEnabled = wasOn
    If Err.Number <> 0 Then LogLine "  [WARNING] Label '" & s & "' could not be placed: " & Err.Description
End Sub

Public Function PickCsvFile() As String
    ' Windows Open dialog for a CSV file; returns "" when cancelled.
    Dim ofn As OPENFILENAME, buf As String
    buf = String$(1024, vbNullChar)
    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFilter = "CSV files (*.csv)" & vbNullChar & "*.csv" & vbNullChar & "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = buf
    ofn.nMaxFile = Len(buf)
    ofn.lpstrTitle = "Import culvert points (ID, X/Lon, Y/Lat)"
    ofn.flags = &H1000 Or &H800 Or &H4        ' file must exist, path must exist, hide read-only
    If GetOpenFileNameA(ofn) <> 0 Then
        PickCsvFile = Left$(ofn.lpstrFile, InStr(1, ofn.lpstrFile, vbNullChar) - 1)
    End If
End Function

'----------------------------------------------------------------------------------------------
' StreamStats services
'----------------------------------------------------------------------------------------------
Private Function SnapPourPoint(ByVal lon As Double, ByVal lat As Double, ByRef sLon As Double, ByRef sLat As Double) As Boolean
    Dim txt As String, d As Variant, c As Object
    sLon = lon
    sLat = lat
    On Error GoTo NoSnap
    txt = HttpText("GET", URL_SNAP & REGION & "?pt=POINT(" & NumFix(lon, 6) & "%20" & NumFix(lat, 6) & ")", "", 60)
    JsonParse txt, d
    If Not HasKey(d, "output") Then Exit Function
    If IsNull(d("scope")) Then Exit Function
    If Len(CStr(d("scope"))) = 0 Then Exit Function
    Set c = d("output")("coordinates")
    If c.Count >= 2 Then
        sLon = CDbl(c(1))
        sLat = CDbl(c(2))
        SnapPourPoint = True
    End If
    Exit Function
NoSnap:
    LogLine "[WARNING] Pour-point snap failed (" & Err.Description & "); using the original point."
End Function

Private Sub ExtractWatershed(ByRef ws As Variant, ByRef rings As Collection, ByRef outer As Collection, ByRef wsId As String)
    Dim wsresp As Object, fcItem As Variant, fc As Variant, feat As Variant, geom As Object
    Dim items As Collection, polys As Collection, poly As Variant, ring As Variant
    Dim nm As String, gtype As String, k As Long
    If Not HasKey(ws, "bcrequest") Then Exit Sub
    If Not HasKey(ws("bcrequest"), "wsresp") Then Exit Sub
    Set wsresp = ws("bcrequest")("wsresp")
    If wsresp.Exists("workspace_id") Then
        If Not IsNull(wsresp("workspace_id")) Then wsId = CStr(wsresp("workspace_id"))
    End If
    If Not wsresp.Exists("featurecollection") Then Exit Sub
    Set items = New Collection
    For Each fcItem In wsresp("featurecollection")
        If TypeName(fcItem) = "Collection" Then
            For Each fc In fcItem
                items.Add fc
            Next fc
        Else
            items.Add fcItem
        End If
    Next fcItem
    For Each fc In items
        If HasKey(fc, "name") And HasKey(fc, "feature") Then
            nm = LCase$(CStr(fc("name")))
            If InStr(nm, "watershed") > 0 And InStr(nm, "point") = 0 And HasKey(fc("feature"), "features") Then
                For Each feat In fc("feature")("features")
                    If HasKey(feat, "geometry") Then
                        Set geom = feat("geometry")
                        gtype = CStr(geom("type"))
                        Set polys = New Collection
                        If gtype = "Polygon" Then
                            polys.Add geom("coordinates")
                        ElseIf gtype = "MultiPolygon" Then
                            For Each poly In geom("coordinates")
                                polys.Add poly
                            Next poly
                        End If
                        For Each poly In polys
                            k = 0
                            For Each ring In poly
                                rings.Add ring
                                outer.Add (k = 0)
                                k = k + 1
                            Next ring
                        Next poly
                    End If
                Next feat
            End If
        End If
    Next fc
End Sub

Private Function WatershedGeometryJson(ByRef wsText As String) As String
    ' Raw text of the watershed geometry object (sent to NSS), cut straight from the response.
    Dim p As Long, g As Long
    p = InStr(1, wsText, """name"":""globalwatershed""")
    If p = 0 Then Exit Function
    g = InStr(p, wsText, """geometry"":")
    If g = 0 Then Exit Function
    WatershedGeometryJson = ExtractJsonObject(wsText, g + 11)
End Function

Private Function BasinCharacteristics(ByRef wsText As String, ByRef params As Object) As Boolean
    Dim txt As String, d As Variant, p As Variant, code As String
    If Now < mBcDownUntil Then
        LogLine "  StreamStats basin characteristics are down (retrying in " & Format$((mBcDownUntil - Now) * 1440#, "0") & " min)."
        Exit Function
    End If
    On Error GoTo NoBc
    LogLine "  Computing basin characteristics..."
    txt = HttpText("POST", URL_BASINCHAR, wsText, 240)
    JsonParse txt, d
    If TypeName(d) <> "Collection" Then Exit Function
    For Each p In d
        If HasKey(p, "code") And HasKey(p, "value") Then
            If Not IsNull(p("value")) Then
                If IsNumeric(p("value")) Then
                    If CDbl(p("value")) > -999# Then
                        code = UCase$(CStr(p("code")))
                        params(code) = CDbl(p("value"))
                    End If
                End If
            End If
        End If
    Next p
    BasinCharacteristics = (params.Count > 0)
    LogLine "StreamStats returned " & params.Count & " basin characteristic(s)."
    Exit Function
NoBc:
    LogLine "  [WARNING] Basin characteristics failed: " & Err.Description
    If Not gCancel Then
        If InStr(Err.Description, "HTTP Error 5") > 0 Or InStr(LCase$(Err.Description), "timed out") > 0 Then mBcDownUntil = Now + 10# / 1440#
    End If
End Function

Private Sub BasinMetrics(ByRef params As Object, ByRef sqmi As Double, ByRef acres As Double, _
                         ByRef slope As Double, ByRef impPct As Double, ByRef forest As Double)
    Dim k As Variant, code As String, v As Double
    sqmi = 0#
    slope = 50#
    impPct = 2#
    forest = 50#
    For Each k In params.keys
        code = CStr(k)
        v = CDbl(params(k))
        If code = "DRNAREA" Then
            sqmi = v
        ElseIf code = "CSL10_85" Or code = "CSL1085LFP" Or code = "SLOP30_10M" Then
            slope = MaxD(v, 1#)
        ElseIf InStr(code, "IMP") > 0 Then
            impPct = MaxD(0#, MinD(v, 100#))
        ElseIf InStr(code, "FOR") > 0 Or InStr(code, "WOOD") > 0 Then
            forest = MaxD(0#, MinD(v, 100#))
        End If
    Next k
    acres = sqmi * 640#
End Sub

Private Function NssPeakFlows(ByRef geomJson As String, ByRef params As Object, ByRef q() As Double) As Boolean
    Dim txt As String, rr As Variant, r As Variant, weights As Object, codes As String, code As String
    If Now < mNssDownUntil Then
        LogLine "  USGS NSS regression service is down (retrying in " & Format$((mNssDownUntil - Now) * 1440#, "0") & _
                " min); no USGS regression flows for this culvert."
        Exit Function
    End If
    On Error GoTo NoNss
    txt = HttpText("POST", URL_NSS & "/regressionregions/bylocation", geomJson, 45, 1)
    JsonParse txt, rr
    If TypeName(rr) <> "Collection" Then Exit Function
    If rr.Count = 0 Then Exit Function
    Set weights = CreateObject("Scripting.Dictionary")
    For Each r In rr
        If HasKey(r, "code") Then
            If Not IsNull(r("code")) Then
                code = CStr(r("code"))
                If Len(code) > 0 Then
                    If HasKey(r, "percentWeight") Then weights(UCase$(code)) = r("percentWeight")
                    codes = codes & IIf(Len(codes) > 0, ",", "") & code
                End If
            End If
        End If
    Next r
    NssPeakFlows = NssEstimate(codes, weights, params, q)
    Exit Function
NoNss:
    LogLine "  [WARNING] Peak-flow regression (NSS) failed: " & Err.Description & ". It is skipped for 10 min so the batch keeps moving."
    If Not gCancel Then mNssDownUntil = Now + 10# / 1440#
End Function

Private Function NssEstimate(ByVal codes As String, ByRef weights As Object, ByRef params As Object, ByRef q() As Double) As Boolean
    Dim txt As String, scen As Variant, est As Variant
    Dim sc As Variant, g As Variant, p As Variant, grp As Variant, reg As Variant, res As Variant
    Dim keep As Collection, code As String, missing As String
    Dim avg(0 To 7) As Double, isAvg As Boolean, idx As Long, i As Long, n As Long
    On Error GoTo NoNss
    txt = HttpText("GET", URL_NSS & "/scenarios?regions=" & REGION & ",NA&statisticgroups=2&regressionregions=" & codes, "", 45)
    JsonParse txt, scen
    If TypeName(scen) <> "Collection" Then Exit Function
    If scen.Count = 0 Then Exit Function
    For Each sc In scen
        Set keep = New Collection
        If HasKey(sc, "regressionRegions") Then
            If TypeName(sc("regressionRegions")) = "Collection" Then
                For Each g In sc("regressionRegions")
                    If InStr(LCase$(JStr(g, "name")), "area-averaged") = 0 Then
                        If g.Exists("results") Then g.Remove "results"
                        code = UCase$(JStr(g, "code"))
                        If weights.Exists(code) Then
                            g("percentWeight") = weights(code)
                        Else
                            g("percentWeight") = Null
                        End If
                        If HasKey(g, "parameters") Then
                            If TypeName(g("parameters")) = "Collection" Then
                                For Each p In g("parameters")
                                    code = UCase$(JStr(p, "code"))
                                    If params.Exists(code) Then
                                        p("value") = params(code)
                                    Else
                                        p("value") = Null
                                        If InStr(missing, code) = 0 Then missing = missing & IIf(Len(missing) > 0, ", ", "") & code
                                    End If
                                Next p
                            End If
                        End If
                        keep.Add g
                    End If
                Next g
            End If
        End If
        Set sc("regressionRegions") = keep
    Next sc
    If Len(missing) > 0 Then
        LogLine "[WARNING] Skipping USGS peak flows - missing basin characteristic(s): " & missing & "."
        Exit Function
    End If
    txt = HttpText("POST", URL_NSS & "/scenarios/estimate?regions=" & REGION & ",NA", JsonStr(scen), 60)
    JsonParse txt, est
    If TypeName(est) <> "Collection" Then Exit Function
    For i = 0 To 7
        avg(i) = -1#
    Next i
    For Each grp In est
        If HasKey(grp, "regressionRegions") Then
            For Each reg In grp("regressionRegions")
                isAvg = (InStr(LCase$(JStr(reg, "name")), "area-averaged") > 0)
                If HasKey(reg, "results") Then
                    If TypeName(reg("results")) = "Collection" Then
                        For Each res In reg("results")
                            idx = PkIndex(JStr(res, "code"))
                            If idx >= 0 And HasKey(res, "value") Then
                                If Not IsNull(res("value")) Then
                                    If IsNumeric(res("value")) Then
                                        If isAvg Then
                                            avg(idx) = CDbl(res("value"))
                                        ElseIf q(idx) < 0# Then
                                            q(idx) = CDbl(res("value"))
                                        End If
                                    End If
                                End If
                            End If
                        Next res
                    End If
                End If
            Next reg
        End If
    Next grp
    For i = 0 To 7
        If avg(i) >= 0# Then q(i) = avg(i)
        If q(i) > 0# Then n = n + 1
    Next i
    NssEstimate = (n > 0)
    LogLine "USGS regression (NSS) returned " & n & " peak flow(s)."
    Exit Function
NoNss:
    LogLine "[WARNING] Peak-flow regression (NSS) failed: " & Err.Description
End Function

Public Sub StreamStatsSelfTestNss()
    ' Checks the regression estimate without the region lookup: one known Virginia region, 0.14 sq mi.
    Dim weights As Object, params As Object, q(0 To 7) As Double, i As Long, s As String
    gSilent = True
    Set weights = CreateObject("Scripting.Dictionary")
    Set params = CreateObject("Scripting.Dictionary")
    weights("GC2088") = 100#
    params("DRNAREA") = 0.14
    For i = 0 To 7
        q(i) = -1#
    Next i
    LogLine "NSS self-test: " & IIf(NssEstimate("GC2088", weights, params, q), "OK", "no result")
    For i = 0 To 7
        s = s & NumFix(q(i), 1) & " "
    Next i
    LogLine "NSS self-test Q2..Q500: " & s
End Sub

Private Function PkIndex(ByVal code As String) As Long
    ' NSS codes such as PK10AEP (10% chance = 10-yr) or PK0_2AEP (500-yr) -> index into 2,5,...,500
    Dim s As String, aep As Double, yr As Long, i As Long, yrs As Variant
    PkIndex = -1
    s = UCase$(Trim$(code))
    If Left$(s, 2) <> "PK" Or Right$(s, 3) <> "AEP" Or Len(s) < 6 Then Exit Function
    s = Replace(Mid$(s, 3, Len(s) - 5), "_", ".")
    If Not IsNumeric(Replace(s, ".", "")) Then Exit Function
    aep = Val(s)
    If aep <= 0# Then Exit Function
    yr = CLng(100# / aep)
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    For i = 0 To 7
        If yrs(i) = yr Then PkIndex = i
    Next i
End Function

'----------------------------------------------------------------------------------------------
' Rainfall and hydrology
'----------------------------------------------------------------------------------------------
Private Function Atlas14Intensities(ByVal lon As Double, ByVal lat As Double, ByRef atlas() As Double) As Boolean
    ' atlas(duration 0..9, return period 0..9) in in/hr; durations 5 min..24 hr, periods 1..1000 yr
    Dim txt As String, a As Long, b As Long, body As String, rows() As String, parts() As String
    Dim r As Long, c As Long, k As Long
    On Error GoTo NoAtlas
    txt = HttpText("GET", URL_ATLAS14 & "?lat=" & NumFix(lat, 5) & "&lon=" & NumFix(lon, 5) & _
                   "&type=pf&data=intensity&units=english&series=pds", "", 45)
    a = InStr(1, txt, "quantiles")
    If a = 0 Then Err.Raise vbObjectError + 3, , "no Atlas 14 data for this location"
    a = InStr(a, txt, "[[")
    b = InStr(a, txt, "]];")
    If a = 0 Or b = 0 Then Err.Raise vbObjectError + 3, , "unexpected Atlas 14 response"
    body = Mid$(txt, a + 2, b - a - 2)
    rows = Split(body, "],")
    If UBound(rows) < 9 Then Err.Raise vbObjectError + 3, , "unexpected Atlas 14 table size"
    ReDim atlas(0 To 9, 0 To 9)
    For r = 0 To 9
        parts = Split(rows(r), "'")
        c = 0
        For k = 1 To UBound(parts) Step 2
            If c <= 9 Then atlas(r, c) = Val(parts(k))
            c = c + 1
        Next k
        If c < 10 Then Err.Raise vbObjectError + 3, , "unexpected Atlas 14 table size"
    Next r
    Atlas14Intensities = True
    LogLine "NOAA Atlas 14 rainfall read for the pour point."
    Exit Function
NoAtlas:
    LogLine "[WARNING] NOAA Atlas 14 rainfall unavailable (" & Err.Description & "); using built-in IDF coefficients."
End Function

Public Function AtlasIntensity(ByRef atlas() As Double, ByVal yr As Long, ByVal tcMin As Double) As Double
    Dim durs As Variant, aris As Variant, col As Long, i As Long, t As Double, f As Double
    AtlasIntensity = -1#
    durs = Array(5#, 10#, 15#, 30#, 60#, 120#, 180#, 360#, 720#, 1440#)
    aris = Array(1, 2, 5, 10, 25, 50, 100, 200, 500, 1000)
    col = -1
    For i = 0 To 9
        If aris(i) = yr Then col = i
    Next i
    If col < 0 Then Exit Function
    t = MinD(MaxD(tcMin, durs(0)), durs(9))
    For i = 0 To 8
        If t >= durs(i) And t <= durs(i + 1) Then
            If atlas(i, col) <= 0# Or atlas(i + 1, col) <= 0# Then Exit Function
            f = (Log(t) - Log(durs(i))) / (Log(durs(i + 1)) - Log(durs(i)))
            AtlasIntensity = Exp(Log(atlas(i, col)) + f * (Log(atlas(i + 1, col)) - Log(atlas(i, col))))
            Exit Function
        End If
    Next i
End Function

Private Function BuiltinIntensity(ByVal yr As Long, ByVal tcMin As Double) As Double
    ' Statewide placeholder IDF coefficients (same as the Python tool); used only without Atlas 14.
    Dim a As Double, b As Double
    Select Case yr
        Case 2: a = 48.5: b = 8.5
        Case 5: a = 63#: b = 9#
        Case 10: a = 75.2: b = 9.5
        Case 25: a = 92.4: b = 10#
        Case 50: a = 106.8: b = 10.5
        Case 100: a = 122#: b = 11#
        Case 200: a = 138#: b = 11.5
        Case Else: a = 160.5: b = 12#
    End Select
    BuiltinIntensity = a / ((tcMin + b) ^ 0.78)
End Function

Private Function BuiltinRegression(ByVal yr As Long, ByVal sqmi As Double) As Double
    ' Placeholder check equation (same as the Python tool); used only when NSS returns no value.
    Dim k As Double, e As Double
    Select Case yr
        Case 2: k = 108#: e = 0.68
        Case 5: k = 185#: e = 0.67
        Case 10: k = 248#: e = 0.66
        Case 25: k = 342#: e = 0.65
        Case 50: k = 425#: e = 0.64
        Case 100: k = 518#: e = 0.63
        Case 200: k = 620#: e = 0.62
        Case Else: k = 775#: e = 0.61
    End Select
    If sqmi > 0# Then BuiltinRegression = RoundTo(k * (sqmi ^ e), 1)
End Function

'----------------------------------------------------------------------------------------------
' Drawing and Item Types
'----------------------------------------------------------------------------------------------
Private Sub DrawBasin(ByRef rings As Collection, ByRef vals As Object, ByRef src As Object)
    Dim lvl As Level, ring As Variant, pts() As Point3d, el As ShapeElement
    Dim n As Long, stepN As Long, i As Long, k As Long, m As Long, x As Double, y As Double
    Dim idx As Long, drawn As Long
    Set lvl = EnsureLevel(LEVEL_BASIN, 3, 2)
    idx = 0
    For Each ring In rings
        n = ring.Count
        If n >= 3 Then
            stepN = 1
            If n > MAX_SHAPE_VERTICES Then stepN = (n \ MAX_SHAPE_VERTICES) + 1
            m = (n - 1) \ stepN + 1
            ReDim pts(0 To m - 1)
            k = 0
            For i = 1 To n Step stepN
                LonLatToXY CDbl(ring(i)(1)), CDbl(ring(i)(2)), x, y
                pts(k).x = x
                pts(k).y = y
                pts(k).z = 0#
                k = k + 1
            Next i
            If k >= 3 Then
                If k < m Then ReDim Preserve pts(0 To k - 1)
                Set el = CreateShapeElement1(Nothing, pts, msdFillModeNotFilled)
                Set el.Level = lvl
                el.Color = 3
                el.LineWeight = 2
                SetDashedStyle el
                ActiveModelReference.AddElement el
                drawn = drawn + 1
                If idx = 0 Then
                    mLastElementId = DLongToString(el.id)
                    AttachItem el, "DrainageArea", BasinPropNames(), BasinPropTypes(), vals
                    AttachItem el, "DataSources", SourcePropNames(), "", src
                End If
            End If
        End If
        idx = idx + 1
    Next ring
    LogLine "Drew " & drawn & " drainage-area boundary element(s) on level '" & LEVEL_BASIN & "' (element ID " & mLastElementId & ")."
End Sub

Private Sub SetDashedStyle(ByVal el As Element)
    Dim ls As LineStyle
    On Error Resume Next
    Set ls = ActiveDesignFile.LineStyles.Find("2", Nothing)
    If Not ls Is Nothing Then Set el.LineStyle = ls
End Sub

Public Function EnsureLevel(ByVal levelName As String, ByVal colorIndex As Long, ByVal weight As Long) As Level
    Dim lvl As Level
    On Error Resume Next
    Set lvl = ActiveDesignFile.Levels.Find(levelName, Nothing)
    On Error GoTo 0
    If lvl Is Nothing Then
        Set lvl = ActiveDesignFile.AddNewLevel(levelName)
        lvl.ElementColor = colorIndex
        lvl.ElementLineWeight = weight
        ActiveDesignFile.Levels.Rewrite
    End If
    Set EnsureLevel = lvl
End Function

Public Function AttachItem(ByVal el As Element, ByVal itemTypeName As String, ByVal names As Variant, _
                            ByVal types As String, ByVal vals As Object, Optional ByVal quiet As Boolean = False) As Boolean
    ' types: one character per property, "d" = number, anything else (or empty) = text
    Dim libs As ItemTypeLibraries, itLib As ItemTypeLibrary, it As ItemType, prop As ItemTypeProperty
    Dim h As ItemTypePropertyHandler, i As Long, nSet As Long, changed As Boolean, nm As String
    On Error GoTo ItemFailed
    Set libs = ActiveDesignFile.ItemTypeLibraries
    Set itLib = Nothing
    On Error Resume Next
    Set itLib = libs.FindByName(LIB_NAME)
    On Error GoTo ItemFailed
    If itLib Is Nothing Then
        Set itLib = libs.CreateLib(LIB_NAME, False)
        changed = True
    End If
    Set it = Nothing
    On Error Resume Next
    Set it = itLib.GetItemTypeByName(itemTypeName)
    On Error GoTo ItemFailed
    If it Is Nothing Then
        Set it = itLib.AddItemType(itemTypeName, False)
        changed = True
    End If
    For i = LBound(names) To UBound(names)
        nm = CStr(names(i))
        Set prop = Nothing
        On Error Resume Next
        Set prop = it.GetPropertyByName(nm)
        On Error GoTo ItemFailed
        If prop Is Nothing Then
            If Mid$(types, i - LBound(names) + 1, 1) = "d" Then
                it.AddProperty nm, ItemPropertyTypeDouble, False
            Else
                it.AddProperty nm, ItemPropertyTypeString, False
            End If
            changed = True
        End If
    Next i
    If changed Then itLib.Write
    Set h = it.AttachItem(el)
    If h Is Nothing Then
        LogLine "[WARNING] Item Type '" & itemTypeName & "' could not be attached."
        Exit Function
    End If
    For i = LBound(names) To UBound(names)
        nm = CStr(names(i))
        If vals.Exists(nm) Then
            If h.SetPropertyValue(nm, vals(nm)) Then nSet = nSet + 1
        End If
    Next i
    el.Rewrite
    If Not quiet Then LogLine "Attached Item Type '" & LIB_NAME & ":" & itemTypeName & "' (" & nSet & " properties)."
    AttachItem = True
    Exit Function
ItemFailed:
    LogLine "[WARNING] Item Type '" & itemTypeName & "' failed: " & Err.Description
End Function

Private Function BasinPropNames() As Variant
    Dim s As String, yrs As Variant, i As Long
    s = "CulvertID,StateCode,WorkspaceID,DrainageArea_SqMi,DrainageArea_Acres,Slope_FtPerMi,Impervious_Pct," & _
        "Forest_Pct,RunoffCoeff_C,Tc_Min,LongestFlowpath_Ft,LongestFlowpath_Method,GoverningMethod," & _
        "SoilGroup_A_Pct,SoilGroup_B_Pct,SoilGroup_C_Pct,SoilGroup_D_Pct,DominantSoilGroup,SoilSource," & _
        "CurveNumber_CN,CurveNumber_Source"
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    For i = 0 To 7
        s = s & ",Q" & yrs(i) & "_USGS_cfs,Q" & yrs(i) & "_VDOT_cfs"
    Next i
    BasinPropNames = Split(s & ",CalcDate", ",")
End Function

Private Function BasinPropTypes() As String
    ' matches BasinPropNames: s = text, d = number
    BasinPropTypes = "sssddddddddss" & "ddddss" & "ds" & String$(16, "d") & "s"
End Function

Private Function SourcePropNames() As Variant
    Dim keys As Variant, s As String, i As Long
    keys = Array("Watershed", "BasinCharacteristics", "LandCover", "ChannelSlope", "Streams", "LongestFlowpath", _
                 "RuralRegression", "UrbanRegression", "Anderson", "Rainfall", "SoilGroups", "Methods")
    s = "CulvertID"
    For i = 0 To UBound(keys)
        s = s & "," & keys(i) & "_Source," & keys(i) & "_Website"
    Next i
    SourcePropNames = Split(s & ",CalcDate", ",")
End Function

'----------------------------------------------------------------------------------------------
' Virginia State Plane (Lambert Conformal Conic, NAD83 / GRS80, US survey feet)
'----------------------------------------------------------------------------------------------
Public Function GetZone() As String
    GetZone = UCase$(GetSetting("USGS_StreamStats_VBA", "Settings", "Zone", "S"))
    If GetZone <> "N" Then GetZone = "S"
End Function

Private Function EnsureZone() As Boolean
    If Len(GetSetting("USGS_StreamStats_VBA", "Settings", "Zone", "")) = 0 Then StreamStatsSetZone
    EnsureZone = (Len(GetSetting("USGS_StreamStats_VBA", "Settings", "Zone", "")) > 0)
End Function

Private Function Pi() As Double
    Pi = 4# * Atn(1#)
End Function

Private Function LccM(ByVal phi As Double, ByVal e As Double) As Double
    LccM = Cos(phi) / Sqr(1# - e * e * Sin(phi) * Sin(phi))
End Function

Private Function LccT(ByVal phi As Double, ByVal e As Double) As Double
    LccT = Tan(Pi() / 4# - phi / 2#) / (((1# - e * Sin(phi)) / (1# + e * Sin(phi))) ^ (e / 2#))
End Function

Private Sub LccInit(ByVal zone As String)
    Dim f As Double, e As Double, lat0 As Double, sp1 As Double, sp2 As Double, d As Double
    Dim m1 As Double, m2 As Double, t0 As Double, t1 As Double, t2 As Double
    d = Pi() / 180#
    If zone = "N" Then        ' EPSG:2283
        lat0 = (37# + 40# / 60#) * d
        sp1 = (38# + 2# / 60#) * d
        sp2 = (39# + 12# / 60#) * d
        mLccFN = 2000000#
    Else                      ' EPSG:2284
        lat0 = (36# + 20# / 60#) * d
        sp1 = (36# + 46# / 60#) * d
        sp2 = (37# + 58# / 60#) * d
        mLccFN = 1000000#
    End If
    mLccFE = 3500000#
    mLccLon0 = -78.5 * d
    f = 1# / GRS80_INVF
    e = Sqr(2# * f - f * f)
    m1 = LccM(sp1, e)
    m2 = LccM(sp2, e)
    t0 = LccT(lat0, e)
    t1 = LccT(sp1, e)
    t2 = LccT(sp2, e)
    mLccN = (Log(m1) - Log(m2)) / (Log(t1) - Log(t2))
    mLccF = m1 / (mLccN * (t1 ^ mLccN))
    mLccRho0 = GRS80_A * mLccF * (t0 ^ mLccN)
    mLccReady = True
End Sub

Private Sub LonLatToXY(ByVal lon As Double, ByVal lat As Double, ByRef x As Double, ByRef y As Double)
    Dim f As Double, e As Double, d As Double, rho As Double, theta As Double
    If Not mLccReady Then LccInit GetZone()
    d = Pi() / 180#
    f = 1# / GRS80_INVF
    e = Sqr(2# * f - f * f)
    rho = GRS80_A * mLccF * (LccT(lat * d, e) ^ mLccN)
    theta = mLccN * (lon * d - mLccLon0)
    x = (mLccFE + rho * Sin(theta)) * FT_PER_M
    y = (mLccFN + mLccRho0 - rho * Cos(theta)) * FT_PER_M
End Sub

Private Sub XYToLonLat(ByVal x As Double, ByVal y As Double, ByRef lon As Double, ByRef lat As Double)
    Dim f As Double, e As Double, d As Double, dx As Double, dy As Double
    Dim rho As Double, t As Double, theta As Double, phi As Double, i As Long
    If Not mLccReady Then LccInit GetZone()
    d = Pi() / 180#
    f = 1# / GRS80_INVF
    e = Sqr(2# * f - f * f)
    dx = x / FT_PER_M - mLccFE
    dy = mLccRho0 - (y / FT_PER_M - mLccFN)
    rho = Sqr(dx * dx + dy * dy)
    If mLccN < 0# Then rho = -rho
    t = (rho / (GRS80_A * mLccF)) ^ (1# / mLccN)
    theta = Atn(dx / dy)
    phi = Pi() / 2# - 2# * Atn(t)
    For i = 1 To 8
        phi = Pi() / 2# - 2# * Atn(t * (((1# - e * Sin(phi)) / (1# + e * Sin(phi))) ^ (e / 2#)))
    Next i
    lon = (theta / mLccN + mLccLon0) / d
    lat = phi / d
End Sub

Private Function PolygonAreaSqMi(ByRef rings As Collection, ByRef outer As Collection) As Double
    ' Area in sq mi from a local equal-area projection (the same formula as the Python tool); holes are subtracted.
    Dim ring As Variant, p As Variant, n As Long, a As Double, total As Double, idx As Long, lat0 As Double, k As Long
    Dim x() As Double, y() As Double
    Const R_MI As Double = 3958.7613
    Const D2R As Double = 1.74532925199433E-02
    idx = 1
    For Each ring In rings
        n = ring.Count
        a = 0#
        If n >= 3 Then
            ReDim x(0 To n - 1)
            ReDim y(0 To n - 1)
            lat0 = 0#
            k = 0
            For Each p In ring
                x(k) = CDbl(p(1))
                y(k) = CDbl(p(2))
                lat0 = lat0 + y(k)
                k = k + 1
            Next p
            lat0 = lat0 / n * D2R
            For k = 0 To n - 2
                a = a + (x(k) * D2R * Cos(lat0) * R_MI) * (y(k + 1) * D2R * R_MI) - (x(k + 1) * D2R * Cos(lat0) * R_MI) * (y(k) * D2R * R_MI)
            Next k
            a = Abs(a) / 2#
        End If
        If outer(idx) Then total = total + a Else total = total - a
        idx = idx + 1
    Next ring
    PolygonAreaSqMi = RoundTo(MaxD(total, 0#), 4)
End Function

'----------------------------------------------------------------------------------------------
' HTTP
'----------------------------------------------------------------------------------------------
Public Function HttpText(ByVal method As String, ByVal url As String, ByRef body As String, ByVal timeoutSec As Long, _
                          Optional ByVal maxAttempts As Long = 2) As String
    Dim attempt As Long, st As Long, lastErr As String, txt As String
    For attempt = 1 To maxAttempts
        st = 0
        lastErr = ""
        txt = ""
        If HttpOnce(method, url, body, timeoutSec, st, txt, lastErr) Then
            HttpText = txt
            Exit Function
        End If
        If gCancel Then Exit For
        If st < 500 And st <> 0 Then Exit For
    Next attempt
    Err.Raise vbObjectError + 10, , "Web request failed (" & HostOf(url) & "): " & lastErr
End Function

Private Function HttpOnce(ByVal method As String, ByVal url As String, ByRef body As String, ByVal timeoutSec As Long, _
                          ByRef st As Long, ByRef txt As String, ByRef errText As String) As Boolean
    ' One request, sent asynchronously and polled so MicroStation and the window keep responding.
    Dim h As Object, t0 As Single
    On Error GoTo Bad
    Set h = CreateObject("MSXML2.ServerXMLHTTP.6.0")
    h.setTimeouts 15000, 15000, 30000, timeoutSec * 1000
    h.Open method, url, True
    h.setRequestHeader "User-Agent", USER_AGENT
    h.setRequestHeader "Accept", "application/json"
    If method = "POST" Then
        h.setRequestHeader "Content-Type", "application/json"
        h.setRequestHeader "X-Is-StreamStats", "true"
        h.send body
    Else
        h.send
    End If
    t0 = Timer
    Do While Not h.waitForResponse(0.1)
        DoEvents
        If gCancel Then
            h.abort
            errText = "stopped by user"
            Exit Function
        End If
        If Timer - t0 > timeoutSec Then
            h.abort
            errText = "The operation timed out"
            Exit Function
        End If
    Loop
    st = h.Status
    If st >= 200 And st < 300 Then
        txt = h.responseText
        HttpOnce = True
    Else
        errText = "HTTP Error " & st
    End If
    Exit Function
Bad:
    errText = Err.Description
End Function

Private Function HostOf(ByVal url As String) As String
    Dim a As Long, b As Long
    a = InStr(1, url, "//")
    If a = 0 Then
        HostOf = url
        Exit Function
    End If
    b = InStr(a + 2, url, "/")
    If b = 0 Then b = Len(url) + 1
    HostOf = Mid$(url, a + 2, b - a - 2)
End Function

'----------------------------------------------------------------------------------------------
' JSON (objects -> Scripting.Dictionary, arrays -> Collection, null -> Null)
'----------------------------------------------------------------------------------------------
Public Sub JsonParse(ByRef text As String, ByRef result As Variant)
    mJ = text
    mLen = Len(mJ)
    mP = 1
    result = Empty
    If mLen = 0 Then Err.Raise vbObjectError + 20, , "empty response"
    ParseValue result
    mJ = ""
End Sub

Private Sub SkipWs()
    Dim c As String
    Do While mP <= mLen
        c = Mid$(mJ, mP, 1)
        If c = " " Or c = vbLf Or c = vbCr Or c = vbTab Then
            mP = mP + 1
        Else
            Exit Do
        End If
    Loop
End Sub

Private Sub ParseValue(ByRef v As Variant)
    Dim c As String, d As Object, col As Collection, k As String, item As Variant, startP As Long
    SkipWs
    If mP > mLen Then Err.Raise vbObjectError + 20, , "unexpected end of JSON"
    c = Mid$(mJ, mP, 1)
    Select Case c
        Case "{"
            Set d = CreateObject("Scripting.Dictionary")
            mP = mP + 1
            SkipWs
            If Mid$(mJ, mP, 1) = "}" Then
                mP = mP + 1
            Else
                Do
                    SkipWs
                    k = ParseString()
                    SkipWs
                    If Mid$(mJ, mP, 1) <> ":" Then Err.Raise vbObjectError + 20, , "bad JSON object at " & mP
                    mP = mP + 1
                    item = Empty
                    ParseValue item
                    If IsObject(item) Then
                        Set d(k) = item
                    Else
                        d(k) = item
                    End If
                    SkipWs
                    c = Mid$(mJ, mP, 1)
                    mP = mP + 1
                    If c = "}" Then Exit Do
                    If c <> "," Then Err.Raise vbObjectError + 20, , "bad JSON object at " & mP
                Loop
            End If
            Set v = d
        Case "["
            Set col = New Collection
            mP = mP + 1
            SkipWs
            If Mid$(mJ, mP, 1) = "]" Then
                mP = mP + 1
            Else
                Do
                    item = Empty
                    ParseValue item
                    col.Add item
                    SkipWs
                    c = Mid$(mJ, mP, 1)
                    mP = mP + 1
                    If c = "]" Then Exit Do
                    If c <> "," Then Err.Raise vbObjectError + 20, , "bad JSON array at " & mP
                Loop
            End If
            Set v = col
        Case """"
            v = ParseString()
        Case "t"
            v = True
            mP = mP + 4
        Case "f"
            v = False
            mP = mP + 5
        Case "n"
            v = Null
            mP = mP + 4
        Case Else
            startP = mP
            Do While mP <= mLen
                If InStr(1, "+-0123456789.eE", Mid$(mJ, mP, 1)) = 0 Then Exit Do
                mP = mP + 1
            Loop
            If mP = startP Then Err.Raise vbObjectError + 20, , "bad JSON value at " & mP
            v = Val(Mid$(mJ, startP, mP - startP))
    End Select
End Sub

Private Function ParseString() As String
    Dim q As Long, c As String, s As String
    If Mid$(mJ, mP, 1) <> """" Then Err.Raise vbObjectError + 20, , "bad JSON string at " & mP
    q = InStr(mP + 1, mJ, """")
    If q = 0 Then Err.Raise vbObjectError + 20, , "unterminated JSON string"
    If InStr(1, Mid$(mJ, mP + 1, q - mP - 1), "\") = 0 Then
        ParseString = Mid$(mJ, mP + 1, q - mP - 1)
        mP = q + 1
        Exit Function
    End If
    mP = mP + 1
    Do While mP <= mLen
        c = Mid$(mJ, mP, 1)
        If c = """" Then
            mP = mP + 1
            Exit Do
        ElseIf c = "\" Then
            mP = mP + 1
            c = Mid$(mJ, mP, 1)
            Select Case c
                Case "n": s = s & vbLf
                Case "r": s = s & vbCr
                Case "t": s = s & vbTab
                Case "b", "f": s = s & " "
                Case "u"
                    s = s & ChrW$(CLng("&H" & Mid$(mJ, mP + 1, 4)))
                    mP = mP + 4
                Case Else: s = s & c
            End Select
            mP = mP + 1
        Else
            s = s & c
            mP = mP + 1
        End If
    Loop
    ParseString = s
End Function

Private Function ExtractJsonObject(ByRef text As String, ByVal fromPos As Long) As String
    ' Text of the first {...} object at or after fromPos (braces matched, strings skipped).
    Dim p As Long, depth As Long, c As String, startP As Long, inQuote As Boolean, n As Long
    n = Len(text)
    p = InStr(fromPos, text, "{")
    If p = 0 Then Exit Function
    startP = p
    Do While p <= n
        c = Mid$(text, p, 1)
        If inQuote Then
            If c = "\" Then
                p = p + 1
            ElseIf c = """" Then
                inQuote = False
            End If
        ElseIf c = """" Then
            inQuote = True
        ElseIf c = "{" Then
            depth = depth + 1
        ElseIf c = "}" Then
            depth = depth - 1
            If depth = 0 Then
                ExtractJsonObject = Mid$(text, startP, p - startP + 1)
                Exit Function
            End If
        End If
        p = p + 1
    Loop
End Function

Public Function JsonStr(ByVal v As Variant) As String
    Dim k As Variant, item As Variant, parts() As String, n As Long
    If IsObject(v) Then
        If v Is Nothing Then
            JsonStr = "null"
        ElseIf TypeName(v) = "Dictionary" Then
            If v.Count = 0 Then
                JsonStr = "{}"
                Exit Function
            End If
            ReDim parts(0 To v.Count - 1)
            For Each k In v.keys
                parts(n) = JsonQuote(CStr(k)) & ":" & JsonStr(v(k))
                n = n + 1
            Next k
            JsonStr = "{" & Join(parts, ",") & "}"
        Else
            If v.Count = 0 Then
                JsonStr = "[]"
                Exit Function
            End If
            ReDim parts(0 To v.Count - 1)
            For Each item In v
                parts(n) = JsonStr(item)
                n = n + 1
            Next item
            JsonStr = "[" & Join(parts, ",") & "]"
        End If
    ElseIf IsNull(v) Or IsEmpty(v) Then
        JsonStr = "null"
    ElseIf VarType(v) = vbBoolean Then
        JsonStr = IIf(v, "true", "false")
    ElseIf VarType(v) = vbString Then
        JsonStr = JsonQuote(CStr(v))
    Else
        JsonStr = NumStr(CDbl(v))
    End If
End Function

Public Function JsonQuote(ByVal s As String) As String
    s = Replace(s, "\", "\\")
    s = Replace(s, """", "\""")
    s = Replace(s, vbCr, "\r")
    s = Replace(s, vbLf, "\n")
    s = Replace(s, vbTab, "\t")
    JsonQuote = """" & s & """"
End Function

Public Function HasKey(ByRef o As Variant, ByVal k As String) As Boolean
    If IsObject(o) Then
        If Not o Is Nothing Then
            If TypeName(o) = "Dictionary" Then HasKey = o.Exists(k)
        End If
    End If
End Function

Public Function JStr(ByRef o As Variant, ByVal k As String) As String
    If HasKey(o, k) Then
        If Not IsNull(o(k)) Then
            If Not IsObject(o(k)) Then JStr = CStr(o(k))
        End If
    End If
End Function

'----------------------------------------------------------------------------------------------
' Small helpers
'----------------------------------------------------------------------------------------------
Public Function NumStr(ByVal d As Double) As String
    ' Locale-independent number text, valid as JSON
    Dim s As String
    s = Trim$(Str$(d))
    If Left$(s, 1) = "." Then s = "0" & s
    If Left$(s, 2) = "-." Then s = "-0" & Mid$(s, 2)
    NumStr = s
End Function

Public Function NumFix(ByVal d As Double, ByVal decimals As Long) As String
    ' Fixed decimals with a "." regardless of the Windows regional settings
    NumFix = Replace(Format$(d, "0." & String$(decimals, "0")), ",", ".")
End Function

Public Function RoundTo(ByVal d As Double, ByVal decimals As Long) As Double
    Dim f As Double
    f = 10# ^ decimals
    RoundTo = Int(d * f + 0.5) / f
End Function

Public Function MaxD(ByVal a As Double, ByVal b As Double) As Double
    If a > b Then MaxD = a Else MaxD = b
End Function

Public Function MinD(ByVal a As Double, ByVal b As Double) As Double
    If a < b Then MinD = a Else MinD = b
End Function

Public Sub LogLine(ByVal s As String)
    Dim f As Integer
    mLogText = mLogText & s & vbCrLf
    Debug.Print s
    On Error Resume Next
    If Not gLogSink Is Nothing Then gLogSink.AppendLog s
    f = FreeFile
    Open Environ$("TEMP") & "\USGS_StreamStats_VBA.log" For Append As #f
    Print #f, Format$(Now, "yyyy-mm-dd hh:nn:ss") & "  " & s
    Close #f
End Sub
