Attribute VB_Name = "modGis"
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology (VBA) - GIS exports
' Styled Google Earth KML / KMZ and ESRI Shapefiles (WGS84) for drainage areas, pour points,
' clipped streams, longest flowpaths and raindrop paths. The KML style file is the same one the
' Python tool reads and writes, so both editions share the colors.
'==============================================================================================
Private Const REG_APP As String = "USGS_StreamStats_VBA"

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
Private Type CHOOSECOLORINFO
    lStructSize As Long
    hwndOwner As LongPtr
    hInstance As LongPtr
    rgbResult As Long
    lpCustColors As LongPtr
    flags As Long
    lCustData As LongPtr
    lpfnHook As LongPtr
    lpTemplateName As LongPtr
End Type
Private Declare PtrSafe Function GetSaveFileNameA Lib "comdlg32.dll" (pOpenfilename As OPENFILENAME) As Long
Private Declare PtrSafe Function ChooseColorA Lib "comdlg32.dll" (pChoosecolor As CHOOSECOLORINFO) As Long

Private mCustColors(0 To 15) As Long
Private mCrc(0 To 255) As Long
Private mCrcReady As Boolean

'----------------------------------------------------------------------------------------------
' Windows dialogs
'----------------------------------------------------------------------------------------------
Public Function PickSaveFile(ByVal title As String, ByVal folder As String, ByVal fileName As String, _
                             ByVal filterName As String, ByVal ext As String, Optional ByVal confirmOverwrite As Boolean = True) As String
    ' Windows Save dialog; ext is "kml", "kmz", "shp" or "json". Returns "" when cancelled.
    Dim ofn As OPENFILENAME, buf As String
    buf = fileName & String$(1024 - Len(fileName), vbNullChar)
    ofn.lStructSize = LenB(ofn)
    ofn.lpstrFilter = filterName & " (*." & ext & ")" & vbNullChar & "*." & ext & vbNullChar & "All files (*.*)" & vbNullChar & "*.*" & vbNullChar & vbNullChar
    ofn.nFilterIndex = 1
    ofn.lpstrFile = buf
    ofn.nMaxFile = Len(buf)
    ofn.lpstrInitialDir = folder
    ofn.lpstrTitle = title
    ofn.lpstrDefExt = ext
    ofn.flags = &H800 Or &H4                 ' path must exist, hide read-only
    If confirmOverwrite Then ofn.flags = ofn.flags Or &H2
    If GetSaveFileNameA(ofn) <> 0 Then
        PickSaveFile = Left$(ofn.lpstrFile, InStr(1, ofn.lpstrFile, vbNullChar) - 1)
    End If
End Function

Public Function PickColor(ByVal rgbStart As Long, ByRef rgbOut As Long) As Boolean
    ' Windows color dialog. Colors are VBA RGB values.
    Dim cc As CHOOSECOLORINFO
    cc.lStructSize = LenB(cc)
    cc.rgbResult = rgbStart
    cc.lpCustColors = VarPtr(mCustColors(0))
    cc.flags = &H1 Or &H2                    ' start with rgbResult, open the full dialog
    If ChooseColorA(cc) <> 0 Then
        rgbOut = cc.rgbResult
        PickColor = True
    End If
End Function

Public Function HexToRgb(ByVal hexColor As String) As Long
    ' "#RRGGBB" -> VBA RGB value
    Dim h As String
    h = Replace(hexColor, "#", "")
    If Len(h) <> 6 Then h = "FFFFFF"
    HexToRgb = RGB(CLng("&H" & Mid$(h, 1, 2)), CLng("&H" & Mid$(h, 3, 2)), CLng("&H" & Mid$(h, 5, 2)))
End Function

Public Function RgbToHex(ByVal c As Long) As String
    RgbToHex = "#" & Right$("0" & Hex$(c And &HFF), 2) & Right$("0" & Hex$((c \ &H100) And &HFF), 2) & Right$("0" & Hex$((c \ &H10000) And &HFF), 2)
End Function

Public Function DgnFolder() As String
    Dim p As String
    p = ActiveDesignFile.path
    If Len(p) = 0 Then p = Environ$("USERPROFILE")
    If Right$(p, 1) = "\" Then p = Left$(p, Len(p) - 1)
    DgnFolder = p
End Function

Public Function DgnStem() As String
    Dim n As String, p As Long
    n = ActiveDesignFile.Name
    p = InStrRev(n, ".")
    If p > 1 Then n = Left$(n, p - 1)
    If Len(n) = 0 Then n = "ActiveDGN"
    DgnStem = n
End Function

Public Function SafeName(ByVal s As String, ByVal fallback As String) As String
    Dim i As Long, ch As String, o As String, lastUnd As Boolean
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        If InStr(1, "<>:""/\|?* " & vbTab, ch) > 0 Then
            If Not lastUnd And Len(o) > 0 Then o = o & "_"
            lastUnd = True
        Else
            o = o & ch
            lastUnd = False
        End If
    Next i
    Do While Right$(o, 1) = "_"
        o = Left$(o, Len(o) - 1)
    Loop
    If Len(o) = 0 Then o = fallback
    SafeName = o
End Function

'----------------------------------------------------------------------------------------------
' KML style (shared with the Python tool)
'----------------------------------------------------------------------------------------------
Public Function KmlStyleDefaultPath() As String
    KmlStyleDefaultPath = Environ$("USERPROFILE") & "\.usgs_streamstats_kml_style.json"
End Function

Public Function KmlStylePath() As String
    KmlStylePath = GetSetting(REG_APP, "Settings", "KmlStyleFile", "")
    If Len(KmlStylePath) = 0 Then KmlStylePath = KmlStyleDefaultPath()
End Function

Public Sub SetKmlStylePath(ByVal p As String)
    If StrComp(p, KmlStyleDefaultPath(), vbTextCompare) = 0 Then p = ""
    SaveSetting REG_APP, "Settings", "KmlStyleFile", p
End Sub

Public Function KmlStyleDefaults() As Object
    Dim st As Object
    Set st = CreateObject("Scripting.Dictionary")
    st("basin_line_color") = "#FF0000"
    st("basin_line_width") = 2.5
    st("basin_fill") = True
    st("basin_fill_color") = "#FF0000"
    st("basin_fill_opacity") = 33#
    st("stream_color") = "#0088FF"
    st("stream_width") = 3#
    st("flowpath_color") = "#FF00FF"
    st("flowpath_width") = 3.5
    st("raindrop_color") = "#00E5FF"
    st("raindrop_width") = 4#
    st("point_color") = "#00FF00"
    st("point_scale") = 1.1
    st("point_icon") = "http://maps.google.com/mapfiles/kml/shapes/placemark_circle.png"
    st("label_scale") = 0.8
    Set KmlStyleDefaults = st
End Function

Public Function LoadKmlStyle(Optional ByVal path As String = "") As Object
    ' Defaults, overridden by whatever the style file holds. A missing or bad file just gives the defaults.
    Dim st As Object, f As Integer, txt As String, d As Variant, k As Variant
    Set st = KmlStyleDefaults()
    Set LoadKmlStyle = st
    If Len(path) = 0 Then path = KmlStylePath()
    On Error GoTo done
    If Len(Dir$(path)) = 0 Then Exit Function
    f = FreeFile
    Open path For Binary Access Read As #f
    txt = Space$(LOF(f))
    Get #f, , txt
    Close #f
    JsonParse txt, d
    If TypeName(d) <> "Dictionary" Then Exit Function
    For Each k In st.keys
        If d.Exists(k) Then
            If Not IsNull(d(k)) And Not IsObject(d(k)) Then
                If VarType(st(k)) = vbBoolean Then
                    st(k) = CBool(d(k))
                ElseIf VarType(st(k)) = vbDouble Then
                    If IsNumeric(d(k)) Then st(k) = CDbl(d(k))
                Else
                    st(k) = CStr(d(k))
                End If
            End If
        End If
    Next k
done:
    On Error Resume Next
    Close #f
End Function

Public Sub SaveKmlStyle(ByRef st As Object, ByVal path As String)
    Dim f As Integer, k As Variant, s As String, first As Boolean
    s = "{" & vbLf
    first = True
    For Each k In st.keys
        If Not first Then s = s & "," & vbLf
        first = False
        If VarType(st(k)) = vbBoolean Then
            s = s & "  " & JsonQuote(CStr(k)) & ": " & IIf(st(k), "true", "false")
        ElseIf VarType(st(k)) = vbDouble Then
            s = s & "  " & JsonQuote(CStr(k)) & ": " & NumStr(CDbl(st(k)))
        Else
            s = s & "  " & JsonQuote(CStr(k)) & ": " & JsonQuote(CStr(st(k)))
        End If
    Next k
    s = s & vbLf & "}"
    f = FreeFile
    Open path For Output As #f
    Print #f, s
    Close #f
End Sub

Private Function KmlColor(ByVal hexColor As String, Optional ByVal opacityPct As Double = 100#) As String
    ' "#RRGGBB" + opacity % -> KML "aabbggrr"
    Dim h As String, a As Long
    h = Replace(hexColor, "#", "")
    If Len(h) <> 6 Then h = "FFFFFF"
    If opacityPct < 0# Then opacityPct = 0#
    If opacityPct > 100# Then opacityPct = 100#
    a = CLng(opacityPct * 2.55)
    KmlColor = LCase$(Right$("0" & Hex$(a), 2) & Mid$(h, 5, 2) & Mid$(h, 3, 2) & Mid$(h, 1, 2))
End Function

Private Function XmlEsc(ByVal s As String) As String
    ' XML escape; characters outside ASCII become numeric entities so the file stays valid UTF-8.
    Dim i As Long, ch As String, o As String, c As Long
    For i = 1 To Len(s)
        ch = Mid$(s, i, 1)
        c = AscW(ch)
        If c < 0 Then c = c + 65536
        Select Case ch
            Case "&"
                o = o & "&amp;"
            Case "<"
                o = o & "&lt;"
            Case ">"
                o = o & "&gt;"
            Case """"
                o = o & "&quot;"
            Case Else
                If c > 126 Or c < 32 Then
                    o = o & "&#" & c & ";"
                Else
                    o = o & ch
                End If
        End Select
    Next i
    XmlEsc = o
End Function

Private Function CoordText(ByRef poly As Variant, ByVal closeRing As Boolean) As String
    Dim p As Variant, parts() As String, n As Long, first As Variant
    If poly.Count = 0 Then Exit Function
    ReDim parts(0 To poly.Count)
    For Each p In poly
        parts(n) = NumFix(CDbl(p(1)), 6) & "," & NumFix(CDbl(p(2)), 6) & ",0"
        n = n + 1
    Next p
    If closeRing Then
        If parts(0) <> parts(n - 1) Then
            parts(n) = parts(0)
            n = n + 1
        End If
    End If
    ReDim Preserve parts(0 To n - 1)
    CoordText = Join(parts, " ")
End Function

Private Function QAt(ByRef res As Object, ByVal key As String, ByVal idx As Long) As Double
    ' Flow from one of the result arrays; 0 when it was not calculated.
    Dim a As Variant
    a = res(key)
    If a(idx) > 0# Then QAt = CDbl(a(idx))
End Function

Private Function F1(ByVal v As Double) As String
    F1 = NumFix(v, 1)
End Function

Private Sub WriteTextFile(ByVal path As String, ByRef text As String, ByVal asKmz As Boolean)
    Dim f As Integer, b() As Byte
    If asKmz Then
        b = StrConv(text, vbFromUnicode)
        WriteStoredZip path, "doc.kml", b
    Else
        f = FreeFile
        Open path For Output As #f
        Print #f, text;
        Close #f
    End If
End Sub

Public Function ExportKmlFile(ByVal path As String, ByRef queue As Collection, ByRef streams As Collection) As String
    ' Styled KML (or KMZ when the name ends in .kmz): drainage areas, pour points, clipped streams, longest flowpaths.
    Dim st As Object, o() As String, n As Long, it As Variant, res As Object, cid As String, ring As Variant
    Dim s As Variant, part As Variant, anyLfp As Boolean, lenFt As Double
    Set st = LoadKmlStyle()
    ReDim o(0 To 255)
    Push o, n, "<?xml version=""1.0"" encoding=""UTF-8""?>"
    Push o, n, "<kml xmlns=""http://www.opengis.net/kml/2.2"">"
    Push o, n, "  <Document>"
    Push o, n, "    <name>USGS StreamStats &amp; VDOT Hydrology Deliverable</name>"
    Push o, n, "    <Style id=""basinPolyStyle"">"
    Push o, n, "      <LineStyle><color>" & KmlColor(st("basin_line_color")) & "</color><width>" & NumStr(st("basin_line_width")) & "</width></LineStyle>"
    Push o, n, "      <PolyStyle><color>" & KmlColor(st("basin_fill_color"), st("basin_fill_opacity")) & "</color><fill>" & IIf(st("basin_fill"), "1", "0") & "</fill><outline>1</outline></PolyStyle>"
    Push o, n, "    </Style>"
    Push o, n, "    <Style id=""pourPointStyle"">"
    Push o, n, "      <IconStyle><color>" & KmlColor(st("point_color")) & "</color><scale>" & NumStr(st("point_scale")) & "</scale><Icon><href>" & XmlEsc(st("point_icon")) & "</href></Icon></IconStyle>"
    Push o, n, "      <LabelStyle><scale>" & NumStr(st("label_scale")) & "</scale></LabelStyle>"
    Push o, n, "    </Style>"
    Push o, n, "    <Style id=""nhdStreamStyle"">"
    Push o, n, "      <LineStyle><color>" & KmlColor(st("stream_color")) & "</color><width>" & NumStr(st("stream_width")) & "</width></LineStyle>"
    Push o, n, "    </Style>"
    Push o, n, "    <Style id=""flowpathStyle"">"
    Push o, n, "      <LineStyle><color>" & KmlColor(st("flowpath_color")) & "</color><width>" & NumStr(st("flowpath_width")) & "</width></LineStyle>"
    Push o, n, "    </Style>"

    Push o, n, "    <Folder>"
    Push o, n, "      <name>1_Watershed_Basins</name>"
    For Each it In queue
        If Not it("res") Is Nothing Then
            Set res = it("res")
            cid = XmlEsc(CStr(it("id")))
            Push o, n, "      <Placemark>"
            Push o, n, "        <name>" & cid & " Basin (" & F1(res("acres")) & " ac)</name>"
            Push o, n, "        <description><![CDATA[<h3>" & cid & " - Watershed Summary</h3><table border='1' cellpadding='4' cellspacing='0'>" & _
                "<tr><td><b>State / Workspace</b></td><td>VA / " & CStr(res("ws_id")) & "</td></tr>" & _
                "<tr><td><b>Drainage Area</b></td><td>" & NumFix(res("sqmi"), 3) & " sq.mi (" & F1(res("acres")) & " ac)</td></tr>" & _
                "<tr><td><b>Channel Slope</b></td><td>" & F1(res("slope")) & " ft/mi</td></tr>" & _
                "<tr><td><b>VDOT C / Tc</b></td><td>C = " & NumFix(res("c"), 2) & ", Tc = " & F1(res("tc")) & " min</td></tr>" & _
                "<tr><td><b>Q10 (USGS / VDOT)</b></td><td>" & F1(QAt(res, "q_usgs", 2)) & " cfs / " & F1(QAt(res, "q_des", 2)) & " cfs</td></tr>" & _
                "<tr><td><b>Q25 (USGS / VDOT)</b></td><td>" & F1(QAt(res, "q_usgs", 3)) & " cfs / " & F1(QAt(res, "q_des", 3)) & " cfs</td></tr>" & _
                "<tr><td><b>Q100 (USGS / VDOT)</b></td><td>" & F1(QAt(res, "q_usgs", 5)) & " cfs / " & F1(QAt(res, "q_des", 5)) & " cfs</td></tr>" & _
                "<tr><td><b>Governing Method</b></td><td>" & XmlEsc(CStr(res("gov"))) & "</td></tr></table>]]></description>"
            Push o, n, "        <styleUrl>#basinPolyStyle</styleUrl>"
            Push o, n, "        <MultiGeometry>"
            For Each ring In res("rings")
                If ring.Count >= 3 Then
                    Push o, n, "          <Polygon>"
                    Push o, n, "            <outerBoundaryIs><LinearRing>"
                    Push o, n, "              <coordinates>" & CoordText(ring, True) & "</coordinates>"
                    Push o, n, "            </LinearRing></outerBoundaryIs>"
                    Push o, n, "          </Polygon>"
                End If
            Next ring
            Push o, n, "        </MultiGeometry>"
            Push o, n, "      </Placemark>"
            If res.Exists("lfp") Then
                If Not res("lfp") Is Nothing Then anyLfp = True
            End If
        End If
    Next it
    Push o, n, "    </Folder>"

    Push o, n, "    <Folder>"
    Push o, n, "      <name>2_Pour_Points_Culverts</name>"
    For Each it In queue
        If Not it("res") Is Nothing Then
            Set res = it("res")
            cid = XmlEsc(CStr(it("id")))
            Push o, n, "      <Placemark>"
            Push o, n, "        <name>" & cid & " Pour Point</name>"
            Push o, n, "        <description><![CDATA[<b>" & cid & "</b><br/>Lon: " & NumFix(res("s_lon"), 6) & ", Lat: " & NumFix(res("s_lat"), 6) & _
                "<br/>Q25 VDOT: " & F1(QAt(res, "q_des", 3)) & " cfs | Q100 VDOT: " & F1(QAt(res, "q_des", 5)) & " cfs]]></description>"
            Push o, n, "        <styleUrl>#pourPointStyle</styleUrl>"
            Push o, n, "        <Point><coordinates>" & NumFix(res("s_lon"), 6) & "," & NumFix(res("s_lat"), 6) & ",0</coordinates></Point>"
            Push o, n, "      </Placemark>"
        End If
    Next it
    Push o, n, "    </Folder>"

    If Not streams Is Nothing Then
        If streams.Count > 0 Then
            Push o, n, "    <Folder>"
            Push o, n, "      <name>3_Clipped_NHDPlus_Streams</name>"
            For Each s In streams
                Push o, n, "      <Placemark>"
                Push o, n, "        <name>" & XmlEsc(CStr(s("name"))) & " (Order " & CStr(s("order")) & ")</name>"
                Push o, n, "        <description><![CDATA[Clipped Length: " & F1(s("length_ft")) & " ft<br/>Drainage Area: " & NumFix(s("da_sqkm"), 2) & " sq.km]]></description>"
                Push o, n, "        <styleUrl>#nhdStreamStyle</styleUrl>"
                Push o, n, "        <MultiGeometry>"
                For Each part In s("parts")
                    If part.Count >= 2 Then Push o, n, "          <LineString><coordinates>" & CoordText(part, False) & "</coordinates></LineString>"
                Next part
                Push o, n, "        </MultiGeometry>"
                Push o, n, "      </Placemark>"
            Next s
            Push o, n, "    </Folder>"
        End If
    End If

    If anyLfp Then
        Push o, n, "    <Folder>"
        Push o, n, "      <name>4_Longest_Flowpaths</name>"
        For Each it In queue
            If Not it("res") Is Nothing Then
                Set res = it("res")
                If res.Exists("lfp") Then
                    If Not res("lfp") Is Nothing Then
                        lenFt = CDbl(res("lfp_ft"))
                        Push o, n, "      <Placemark>"
                        Push o, n, "        <name>" & XmlEsc(CStr(it("id"))) & " Longest Flowpath (" & Format$(lenFt, "#,##0") & " ft)</name>"
                        Push o, n, "        <description><![CDATA[Length: " & Format$(lenFt, "#,##0.0") & " ft (" & NumFix(lenFt / 5280#, 3) & " mi)<br/>Method: " & _
                            XmlEsc(CStr(res("lfp_method"))) & "]]></description>"
                        Push o, n, "        <styleUrl>#flowpathStyle</styleUrl>"
                        Push o, n, "        <LineString><tessellate>1</tessellate><coordinates>" & CoordText(res("lfp"), False) & "</coordinates></LineString>"
                        Push o, n, "      </Placemark>"
                    End If
                End If
            End If
        Next it
        Push o, n, "    </Folder>"
    End If
    Push o, n, "  </Document>"
    Push o, n, "</kml>"
    ReDim Preserve o(0 To n - 1)
    WriteTextFile path, Join(o, vbLf), (LCase$(Right$(path, 4)) = ".kmz")
    ExportKmlFile = path
End Function

Private Sub Push(ByRef o() As String, ByRef n As Long, ByVal s As String)
    If n > UBound(o) Then ReDim Preserve o(0 To 2 * n + 1)
    o(n) = s
    n = n + 1
End Sub

'----------------------------------------------------------------------------------------------
' Shapefiles
'----------------------------------------------------------------------------------------------
Private Function Swap32(ByVal v As Long) As Long
    ' Big-endian byte order for the shapefile headers
    Dim b0 As Long, b1 As Long, b2 As Long, b3 As Long
    b0 = v And &HFF&
    b1 = (v \ &H100&) And &HFF&
    b2 = (v \ &H10000) And &HFF&
    b3 = (v \ &H1000000) And &HFF&
    If b0 >= 128 Then
        Swap32 = ((b0 - 256) * &H1000000) + b1 * &H10000 + b2 * &H100& + b3
    Else
        Swap32 = b0 * &H1000000 + b1 * &H10000 + b2 * &H100& + b3
    End If
End Function

Private Sub PutL(ByVal f As Integer, ByVal v As Long)
    Put #f, , v
End Sub

Private Sub PutI(ByVal f As Integer, ByVal v As Integer)
    Put #f, , v
End Sub

Private Sub PutMainHeader(ByVal f As Integer, ByVal lenWords As Long, ByVal shapeType As Long, _
                          ByVal xmin As Double, ByVal ymin As Double, ByVal xmax As Double, ByVal ymax As Double)
    Dim i As Long, z As Double
    PutL f, Swap32(9994)
    For i = 1 To 5
        PutL f, CLng(0)
    Next i
    PutL f, Swap32(lenWords)
    PutL f, CLng(1000)
    Put #f, , shapeType
    Put #f, , xmin
    Put #f, , ymin
    Put #f, , xmax
    Put #f, , ymax
    For i = 1 To 4
        Put #f, , z
    Next i
End Sub

Public Function WriteShapefile(ByVal basePath As String, ByVal shapeType As Long, ByRef fields As Variant, ByRef recs As Collection) As String
    ' Writes <base>.shp / .shx / .dbf / .prj / .cpg in WGS84. shapeType: 1 point, 3 polyline, 5 polygon.
    ' fields = Array(Array(name, "C" or "N", length, decimals), ...)
    ' recs   = Collection of Array(parts, attributes); parts is a Collection of lines (a point is a one-point line).
    Dim rec As Variant, part As Variant, x() As Double, y() As Double, n As Long, i As Long, j As Long, t As Double
    Dim good As New Collection, gparts As Collection, np As Long, npts As Long, area2 As Double
    Dim gx0 As Double, gy0 As Double, gx1 As Double, gy1 As Double, rx0 As Double, ry0 As Double, rx1 As Double, ry1 As Double
    Dim lens() As Long, offs() As Long, total As Long, k As Long, fs As Integer, fx As Integer, idx As Long, g As Variant
    Dim fd As Integer, hdrLen As Long, recLen As Long, fld As Variant, s As String, b() As Byte, attrs As Object, v As Variant
    If recs Is Nothing Then Exit Function
    gx0 = 1E+30
    gy0 = 1E+30
    gx1 = -1E+30
    gy1 = -1E+30
    For Each rec In recs
        Set gparts = New Collection
        npts = 0
        For Each part In rec(0)
            n = PolyXY(part, x, y)
            If (shapeType = 1 And n >= 1) Or n >= 2 Then
                If shapeType = 5 Then
                    If x(0) <> x(n - 1) Or y(0) <> y(n - 1) Then
                        ReDim Preserve x(0 To n)
                        ReDim Preserve y(0 To n)
                        x(n) = x(0)
                        y(n) = y(0)
                        n = n + 1
                    End If
                    area2 = 0#
                    For i = 0 To n - 2
                        area2 = area2 + (x(i + 1) - x(i)) * (y(i + 1) + y(i))
                    Next i
                    If area2 < 0# Then            ' outer rings must be clockwise
                        For i = 0 To (n \ 2) - 1
                            t = x(i)
                            x(i) = x(n - 1 - i)
                            x(n - 1 - i) = t
                            t = y(i)
                            y(i) = y(n - 1 - i)
                            y(n - 1 - i) = t
                        Next i
                    End If
                End If
                gparts.Add Array(x, y, n)
                npts = npts + n
            End If
        Next part
        If gparts.Count > 0 Then good.Add Array(gparts, rec(1), npts)
    Next rec
    If good.Count = 0 Then Exit Function

    ReDim lens(1 To good.Count)
    ReDim offs(1 To good.Count)
    total = 50
    k = 0
    For Each g In good
        k = k + 1
        If shapeType = 1 Then
            lens(k) = 10
        Else
            lens(k) = (44 + 4 * g(0).Count + 16 * g(2)) \ 2
        End If
        offs(k) = total
        total = total + 4 + lens(k)
        For Each part In g(0)
            For i = 0 To part(2) - 1
                If part(0)(i) < gx0 Then gx0 = part(0)(i)
                If part(0)(i) > gx1 Then gx1 = part(0)(i)
                If part(1)(i) < gy0 Then gy0 = part(1)(i)
                If part(1)(i) > gy1 Then gy1 = part(1)(i)
            Next i
        Next part
    Next g

    On Error Resume Next
    Kill basePath & ".shp"
    Kill basePath & ".shx"
    Kill basePath & ".dbf"
    On Error GoTo 0
    fs = FreeFile
    Open basePath & ".shp" For Binary Access Write As #fs
    PutMainHeader fs, total, shapeType, gx0, gy0, gx1, gy1
    k = 0
    For Each g In good
        k = k + 1
        PutL fs, Swap32(k)
        PutL fs, Swap32(lens(k))
        Put #fs, , shapeType
        If shapeType = 1 Then
            Set gparts = g(0)
            x = gparts(1)(0)
            y = gparts(1)(1)
            Put #fs, , x(0)
            Put #fs, , y(0)
        Else
            rx0 = 1E+30
            ry0 = 1E+30
            rx1 = -1E+30
            ry1 = -1E+30
            For Each part In g(0)
                x = part(0)
                y = part(1)
                For i = 0 To part(2) - 1
                    If x(i) < rx0 Then rx0 = x(i)
                    If x(i) > rx1 Then rx1 = x(i)
                    If y(i) < ry0 Then ry0 = y(i)
                    If y(i) > ry1 Then ry1 = y(i)
                Next i
            Next part
            Put #fs, , rx0
            Put #fs, , ry0
            Put #fs, , rx1
            Put #fs, , ry1
            np = g(0).Count
            Put #fs, , np
            npts = g(2)
            Put #fs, , npts
            idx = 0
            For Each part In g(0)
                Put #fs, , idx
                idx = idx + part(2)
            Next part
            For Each part In g(0)
                x = part(0)
                y = part(1)
                For i = 0 To part(2) - 1
                    Put #fs, , x(i)
                    Put #fs, , y(i)
                Next i
            Next part
        End If
    Next g
    Close #fs

    fx = FreeFile
    Open basePath & ".shx" For Binary Access Write As #fx
    PutMainHeader fx, 50 + 4 * good.Count, shapeType, gx0, gy0, gx1, gy1
    For k = 1 To good.Count
        PutL fx, Swap32(offs(k))
        PutL fx, Swap32(lens(k))
    Next k
    Close #fx

    ' dBASE III attribute table
    hdrLen = 32 + 32 * (UBound(fields) + 1) + 1
    recLen = 1
    For Each fld In fields
        recLen = recLen + CLng(fld(2))
    Next fld
    s = Chr$(3) & Chr$(Year(Date) - 1900) & Chr$(Month(Date)) & Chr$(Day(Date))
    b = StrConv(s, vbFromUnicode)
    fd = FreeFile
    Open basePath & ".dbf" For Binary Access Write As #fd
    Put #fd, , b
    PutL fd, CLng(good.Count)
    PutI fd, CInt(hdrLen)
    PutI fd, CInt(recLen)
    b = StrConv(String$(20, vbNullChar), vbFromUnicode)
    Put #fd, , b
    For Each fld In fields
        s = Left$(CStr(fld(0)), 10)
        s = s & String$(11 - Len(s), vbNullChar) & CStr(fld(1)) & String$(4, vbNullChar) & Chr$(CLng(fld(2))) & Chr$(CLng(fld(3))) & String$(14, vbNullChar)
        b = StrConv(s, vbFromUnicode)
        Put #fd, , b
    Next fld
    b = StrConv(Chr$(13), vbFromUnicode)
    Put #fd, , b
    For Each g In good
        Set attrs = g(1)
        s = " "
        For Each fld In fields
            v = Empty
            If attrs.Exists(CStr(fld(0))) Then v = attrs(CStr(fld(0)))
            s = s & DbfCell(v, CStr(fld(1)), CLng(fld(2)), CLng(fld(3)))
        Next fld
        b = StrConv(s, vbFromUnicode)
        Put #fd, , b
    Next g
    b = StrConv(Chr$(26), vbFromUnicode)
    Put #fd, , b
    Close #fd

    fd = FreeFile
    Open basePath & ".prj" For Output As #fd
    Print #fd, "GEOGCS[""GCS_WGS_1984"",DATUM[""D_WGS_1984"",SPHEROID[""WGS_1984"",6378137.0,298.257223563]],PRIMEM[""Greenwich"",0.0],UNIT[""Degree"",0.0174532925199433]]";
    Close #fd
    fd = FreeFile
    Open basePath & ".cpg" For Output As #fd
    Print #fd, "UTF-8";
    Close #fd
    WriteShapefile = basePath & ".shp"
End Function

Private Function DbfCell(ByVal v As Variant, ByVal ftype As String, ByVal flen As Long, ByVal fdec As Long) As String
    Dim s As String, i As Long, ch As String, o As String
    If ftype = "N" Then
        If IsNumeric(v) And Not IsEmpty(v) Then
            If fdec > 0 Then
                s = NumFix(CDbl(v), fdec)
            Else
                s = Format$(CDbl(v), "0")
            End If
        End If
        DbfCell = Right$(Space$(flen) & Left$(s, flen), flen)
    Else
        If Not IsEmpty(v) And Not IsNull(v) Then s = CStr(v)
        For i = 1 To Len(s)                       ' plain ASCII only, so every character is one byte
            ch = Mid$(s, i, 1)
            If AscW(ch) >= 32 And AscW(ch) <= 126 Then o = o & ch Else o = o & "?"
        Next i
        DbfCell = Left$(o & Space$(flen), flen)
    End If
End Function

Private Function BasinAttrs(ByRef res As Object, ByVal cid As String) As Object
    Dim a As Object
    Set a = CreateObject("Scripting.Dictionary")
    a("CULVERT_ID") = cid
    a("STATE") = "VA"
    a("WORKSPACE") = CStr(res("ws_id"))
    a("DA_SQMI") = CDbl(res("sqmi"))
    a("DA_ACRES") = CDbl(res("acres"))
    a("SLOPE_FTMI") = CDbl(res("slope"))
    a("VDOT_C") = CDbl(res("c"))
    a("TC_MIN") = CDbl(res("tc"))
    a("Q2_USGS") = QAt(res, "q_usgs", 0)
    a("Q10_USGS") = QAt(res, "q_usgs", 2)
    a("Q25_USGS") = QAt(res, "q_usgs", 3)
    a("Q50_USGS") = QAt(res, "q_usgs", 4)
    a("Q100_USGS") = QAt(res, "q_usgs", 5)
    a("Q10_VDOT") = QAt(res, "q_des", 2)
    a("Q25_VDOT") = QAt(res, "q_des", 3)
    a("Q100_VDOT") = QAt(res, "q_des", 5)
    a("GOV_METHOD") = CStr(res("gov"))
    Set BasinAttrs = a
End Function

Private Function OneLine(ByRef ln As Variant) As Collection
    Dim c As New Collection
    c.Add ln
    Set OneLine = c
End Function

Public Function ExportShapefiles(ByVal basePath As String, ByRef queue As Collection, ByRef streams As Collection) As Collection
    ' <base>_Basins, _PourPoints, _ClippedStreams and _LongestFlowpath shapefiles. Returns the .shp paths written.
    Dim made As New Collection, basins As New Collection, pts As New Collection, strm As New Collection, lfps As New Collection
    Dim it As Variant, res As Object, attrs As Object, s As Variant, a As Object, p As String, ptLine As Collection
    Dim bf As Variant
    bf = Array(Array("CULVERT_ID", "C", 18, 0), Array("STATE", "C", 4, 0), Array("WORKSPACE", "C", 24, 0), Array("DA_SQMI", "N", 12, 4), _
               Array("DA_ACRES", "N", 12, 2), Array("SLOPE_FTMI", "N", 10, 2), Array("VDOT_C", "N", 8, 2), Array("TC_MIN", "N", 8, 1), _
               Array("Q2_USGS", "N", 10, 1), Array("Q10_USGS", "N", 10, 1), Array("Q25_USGS", "N", 10, 1), Array("Q50_USGS", "N", 10, 1), _
               Array("Q100_USGS", "N", 10, 1), Array("Q10_VDOT", "N", 10, 1), Array("Q25_VDOT", "N", 10, 1), Array("Q100_VDOT", "N", 10, 1), _
               Array("GOV_METHOD", "C", 32, 0))
    Set ExportShapefiles = made
    If Not queue Is Nothing Then
        For Each it In queue
            If Not it("res") Is Nothing Then
                Set res = it("res")
                Set attrs = BasinAttrs(res, CStr(it("id")))
                basins.Add Array(res("rings"), attrs)
                Set ptLine = New Collection
                ptLine.Add Pt(CDbl(res("s_lon")), CDbl(res("s_lat")))
                pts.Add Array(OneLine(ptLine), attrs)
                If res.Exists("lfp") Then
                    If Not res("lfp") Is Nothing Then
                        Set a = CreateObject("Scripting.Dictionary")
                        a("CULVERT_ID") = CStr(it("id"))
                        a("LEN_FT") = CDbl(res("lfp_ft"))
                        a("LEN_MI") = CDbl(res("lfp_ft")) / 5280#
                        a("METHOD") = CStr(res("lfp_method"))
                        lfps.Add Array(OneLine(res("lfp")), a)
                    End If
                End If
            End If
        Next it
    End If
    If Not streams Is Nothing Then
        For Each s In streams
            Set a = CreateObject("Scripting.Dictionary")
            a("REACH_ID") = CLng(s("id"))
            a("GNIS_NAME") = CStr(s("name"))
            a("STRM_ORDER") = CLng(s("order"))
            a("DA_SQKM") = CDbl(s("da_sqkm"))
            a("LEN_FT") = CDbl(s("length_ft"))
            strm.Add Array(s("parts"), a)
        Next s
    End If
    p = WriteShapefile(basePath & "_Basins", 5, bf, basins)
    If Len(p) > 0 Then made.Add p
    p = WriteShapefile(basePath & "_PourPoints", 1, bf, pts)
    If Len(p) > 0 Then made.Add p
    p = WriteShapefile(basePath & "_ClippedStreams", 3, Array(Array("REACH_ID", "N", 6, 0), Array("GNIS_NAME", "C", 40, 0), _
                       Array("STRM_ORDER", "N", 6, 0), Array("DA_SQKM", "N", 12, 3), Array("LEN_FT", "N", 12, 1)), strm)
    If Len(p) > 0 Then made.Add p
    p = WriteShapefile(basePath & "_LongestFlowpath", 3, Array(Array("CULVERT_ID", "C", 18, 0), Array("LEN_FT", "N", 12, 1), _
                       Array("LEN_MI", "N", 10, 3), Array("METHOD", "C", 40, 0)), lfps)
    If Len(p) > 0 Then made.Add p
End Function

Public Function ExportGis(ByVal basePath As String, ByVal doKml As Boolean, ByVal kmz As Boolean, ByVal doShp As Boolean, _
                          ByRef queue As Collection, ByRef streams As Collection) As Collection
    ' basePath has no extension. Returns the files written.
    Dim made As New Collection, p As Variant
    If doKml Then made.Add ExportKmlFile(basePath & IIf(kmz, ".kmz", ".kml"), queue, streams)
    If doShp Then
        For Each p In ExportShapefiles(basePath, queue, streams)
            made.Add p
        Next p
    End If
    Set ExportGis = made
End Function

Public Sub AutoSaveGis(ByVal tag As String, ByRef queue As Collection, ByRef streams As Collection)
    ' KML and Shapefiles beside the active DGN (Settings option).
    Dim p As Variant
    On Error GoTo Bad
    For Each p In ExportGis(DgnFolder() & "\" & DgnStem() & "_" & SafeName(tag, "Export"), True, False, True, queue, streams)
        LogLine "[Auto-GIS] Saved " & IIf(LCase$(Right$(CStr(p), 4)) = ".kml", "KML: ", "Shapefile: ") & CStr(p)
    Next p
    Exit Sub
Bad:
    LogLine "[Auto-GIS] Could not write GIS files beside the DGN (" & Err.Description & "). Use Export GIS... to choose another folder."
End Sub

Public Function ExportRaindropGis(ByVal basePath As String, ByVal traceId As String, ByRef res As Object, ByVal doKml As Boolean, _
                                  ByVal kmz As Boolean, ByVal doShp As Boolean) As Collection
    ' <base>.kml|.kmz and <base>_RaindropPath.shp for a Flow (Raindrop) Path trace.
    Dim made As New Collection, rows As Collection, row As Variant, recs As New Collection, a As Object, p As String
    Dim st As Object, o() As String, n As Long, col As String, wid As Double, nm As String, tid As String, path As Collection, last As Variant
    Set rows = RaindropRows(res)
    Set ExportRaindropGis = made
    If doKml Then
        Set st = LoadKmlStyle()
        tid = XmlEsc(traceId)
        col = KmlColor(st("raindrop_color"))
        wid = CDbl(st("raindrop_width"))
        ReDim o(0 To 63)
        Push o, n, "<?xml version=""1.0"" encoding=""UTF-8""?>"
        Push o, n, "<kml xmlns=""http://www.opengis.net/kml/2.2"">"
        Push o, n, "  <Document>"
        Push o, n, "    <name>" & tid & " Flow (Raindrop) Path</name>"
        Push o, n, "    <Style id=""rdStream""><LineStyle><color>" & col & "</color><width>" & NumStr(wid) & "</width></LineStyle></Style>"
        Push o, n, "    <Style id=""rdOverland""><LineStyle><color>" & KmlColor(st("raindrop_color"), 70) & "</color><width>" & NumStr(MaxD(1#, wid - 1#)) & "</width></LineStyle></Style>"
        Push o, n, "    <Style id=""rdPoint""><IconStyle><color>" & col & "</color><scale>" & NumStr(st("point_scale")) & "</scale><Icon><href>" & _
            XmlEsc(st("point_icon")) & "</href></Icon></IconStyle><LabelStyle><scale>" & NumStr(st("label_scale")) & "</scale></LabelStyle></Style>"
        For Each row In rows
            If row(3).Count >= 2 Then
                If CStr(row(1)) = "OVERLAND" Then
                    nm = "Overland flow"
                Else
                    nm = "Reach " & (CLng(row(0)) - 1) & " (COMID " & XmlEsc(CStr(row(2))) & ")"
                End If
                Push o, n, "    <Placemark>"
                Push o, n, "      <name>" & nm & "</name>"
                Push o, n, "      <description><![CDATA[Length: " & Format$(row(4), "#,##0") & " ft<br/>Cumulative: " & NumFix(row(5) / 5280#, 3) & " mi<br/>Method: " & _
                    XmlEsc(CStr(row(6))) & "]]></description>"
                Push o, n, "      <styleUrl>#" & IIf(CStr(row(1)) = "OVERLAND", "rdOverland", "rdStream") & "</styleUrl>"
                Push o, n, "      <LineString><tessellate>1</tessellate><coordinates>" & CoordText(row(3), False) & "</coordinates></LineString>"
                Push o, n, "    </Placemark>"
            End If
        Next row
        Set path = res("path")
        Set last = path(path.Count)
        Push o, n, "    <Placemark>"
        Push o, n, "      <name>" & tid & " Start</name>"
        Push o, n, "      <styleUrl>#rdPoint</styleUrl>"
        Push o, n, "      <Point><coordinates>" & NumFix(res("lon"), 6) & "," & NumFix(res("lat"), 6) & ",0</coordinates></Point>"
        Push o, n, "    </Placemark>"
        Push o, n, "    <Placemark>"
        Push o, n, "      <name>" & tid & " End (" & NumFix(CDbl(res("total_ft")) / 5280#, 2) & " mi)</name>"
        Push o, n, "      <styleUrl>#rdPoint</styleUrl>"
        Push o, n, "      <Point><coordinates>" & NumFix(last(1), 6) & "," & NumFix(last(2), 6) & ",0</coordinates></Point>"
        Push o, n, "    </Placemark>"
        Push o, n, "  </Document>"
        Push o, n, "</kml>"
        ReDim Preserve o(0 To n - 1)
        WriteTextFile basePath & IIf(kmz, ".kmz", ".kml"), Join(o, vbLf), kmz
        made.Add basePath & IIf(kmz, ".kmz", ".kml")
    End If
    If doShp Then
        For Each row In rows
            If row(3).Count >= 2 Then
                Set a = CreateObject("Scripting.Dictionary")
                a("TRACE_ID") = traceId
                a("SEQ") = CLng(row(0))
                a("SEGMENT") = CStr(row(1))
                a("COMID") = CStr(row(2))
                a("LEN_FT") = CDbl(row(4))
                a("CUM_MI") = CDbl(row(5)) / 5280#
                a("METHOD") = Left$(CStr(row(6)), 50)
                recs.Add Array(OneLine(row(3)), a)
            End If
        Next row
        p = WriteShapefile(basePath & "_RaindropPath", 3, Array(Array("TRACE_ID", "C", 18, 0), Array("SEQ", "N", 4, 0), Array("SEGMENT", "C", 10, 0), _
                           Array("COMID", "C", 14, 0), Array("LEN_FT", "N", 12, 1), Array("CUM_MI", "N", 10, 3), Array("METHOD", "C", 50, 0)), recs)
        If Len(p) > 0 Then made.Add p
    End If
End Function

'----------------------------------------------------------------------------------------------
' KMZ = a zip file holding doc.kml (stored without compression)
'----------------------------------------------------------------------------------------------
Private Function Crc32(ByRef b() As Byte) As Long
    Dim i As Long, n As Long, c As Long, k As Long
    If Not mCrcReady Then
        For n = 0 To 255
            c = n
            For k = 1 To 8
                If (c And 1) <> 0 Then
                    c = (((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF) Xor &HEDB88320
                Else
                    c = ((c And &HFFFFFFFE) \ 2) And &H7FFFFFFF
                End If
            Next k
            mCrc(n) = c
        Next n
        mCrcReady = True
    End If
    c = &HFFFFFFFF
    For i = LBound(b) To UBound(b)
        c = (((c And &HFFFFFF00) \ &H100) And &HFFFFFF) Xor mCrc((c And &HFF) Xor b(i))
    Next i
    Crc32 = Not c
End Function

Private Sub WriteStoredZip(ByVal path As String, ByVal entryName As String, ByRef data() As Byte)
    Dim f As Integer, crc As Long, size As Long, nm() As Byte, dosTime As Integer, dosDate As Integer, t As Long, d As Long
    Dim cdOffset As Long, cdSize As Long
    crc = Crc32(data)
    size = UBound(data) - LBound(data) + 1
    nm = StrConv(entryName, vbFromUnicode)
    t = Hour(Now) * 2048& + Minute(Now) * 32& + Second(Now) \ 2
    If t > 32767 Then t = t - 65536
    dosTime = CInt(t)
    d = (Year(Date) - 1980) * 512& + Month(Date) * 32& + Day(Date)
    If d > 32767 Then d = d - 65536
    dosDate = CInt(d)
    On Error Resume Next
    Kill path
    On Error GoTo 0
    f = FreeFile
    Open path For Binary Access Write As #f
    PutL f, CLng(&H4034B50)          ' local file header
    PutI f, CInt(20)
    PutI f, CInt(0)
    PutI f, CInt(0)
    Put #f, , dosTime
    Put #f, , dosDate
    Put #f, , crc
    Put #f, , size
    Put #f, , size
    PutI f, CInt(UBound(nm) + 1)
    PutI f, CInt(0)
    Put #f, , nm
    Put #f, , data
    cdOffset = 30 + UBound(nm) + 1 + size
    PutL f, CLng(&H2014B50)          ' central directory
    PutI f, CInt(20)
    PutI f, CInt(20)
    PutI f, CInt(0)
    PutI f, CInt(0)
    Put #f, , dosTime
    Put #f, , dosDate
    Put #f, , crc
    Put #f, , size
    Put #f, , size
    PutI f, CInt(UBound(nm) + 1)
    PutI f, CInt(0)
    PutI f, CInt(0)
    PutI f, CInt(0)
    PutI f, CInt(0)
    PutL f, CLng(0)
    PutL f, CLng(0)
    Put #f, , nm
    cdSize = 46 + UBound(nm) + 1
    PutL f, CLng(&H6054B50)          ' end of central directory
    PutI f, CInt(0)
    PutI f, CInt(0)
    PutI f, CInt(1)
    PutI f, CInt(1)
    Put #f, , cdSize
    Put #f, , cdOffset
    PutI f, CInt(0)
    Close #f
End Sub
