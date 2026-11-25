VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmStreamStats 
   Caption         =   "USGS StreamStats & VDOT Hydrology - Fence Query (VBA)"
   ClientHeight    =   11120
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   15220
   OleObjectBlob   =   "frmStreamStats.frx":0000
   StartUpPosition =   2  'CenterScreen
End
Attribute VB_Name = "frmStreamStats"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================
' USGS StreamStats / VDOT Hydrology - main window (VBA edition)
' Same tabs, buttons and workflow as the Python tool.
'==============================================================================================
Private Const REG_APP As String = "USGS_StreamStats_VBA"

Private mQueue As Collection          ' Scripting.Dictionary per culvert: id, ctype, x, y, status, res, markers
Private mBusy As Boolean
Private mDark As Boolean
Private mLoading As Boolean
Private mOriginSet As Boolean, mOriginX As Double, mOriginY As Double
Private mStreamsDrawn As Boolean                 ' the clipped streams are drawn once, at the next batch
Private mRd As Object, mRdId As String           ' last raindrop trace
Private mRdStartDrawn As Boolean, mRdStartLon As Double, mRdStartLat As Double
Private mSvcChecked As Boolean, mSvcBusy As Boolean

'----------------------------------------------------------------------------------------------
' Delete All Drawings (header button)
'----------------------------------------------------------------------------------------------
Private Sub btnDeleteAll_Click()
    Dim n As Long, it As Object, r As Object, msg As String
    If mBusy Then Exit Sub
    n = DeleteToolDrawings(True)
    If n = 0 Then
        Notify "Nothing to delete - this tool has no drawings in the active model."
        Exit Sub
    End If
    msg = "Delete all " & Format$(n, "#,##0") & " element(s) this tool drew in the active model?" & vbCrLf & vbCrLf & _
          "Everything on these levels is removed, and the active fence is cleared:" & vbCrLf & _
          "    " & Join(ToolLevelNames(), vbCrLf & "    ") & vbCrLf & vbCrLf & _
          "Use MicroStation Undo to bring them back."
    If Not gSilent Then
        If MsgBox(msg, vbYesNo Or vbExclamation Or vbDefaultButton2, "Delete All Drawings") <> vbYes Then Exit Sub
    End If
    n = DeleteToolDrawings(False)
    ResetTableIds
    For Each it In mQueue
        it("markers") = False
        Set r = it("res")
        If Not r Is Nothing Then r("drawn") = False
    Next it
    mStreamsDrawn = False
    mRdStartDrawn = False
    lblFence.Caption = "Fence: cleared - place a new fence"
    AppendLog "Deleted " & Format$(n, "#,##0") & " element(s) drawn by this tool and cleared the fence. Draw Areas & Markers redraws completed culverts."
End Sub


'----------------------------------------------------------------------------------------------
' Start-up
'----------------------------------------------------------------------------------------------
Private Sub UserForm_Initialize()
    mLoading = True
    Set mQueue = New Collection
    Set gLogSink = Me
    Me.Caption = "USGS StreamStats & VDOT Hydrology - Fence Query (VBA)   v" & AppVersion() & "   -   Created by " & AppBuiltBy() & "   -   Build " & AppBuildDate()
    lblVersion.Caption = "v" & AppVersion() & "  -  Build " & AppBuildDate() & "  -  Created by " & AppBuiltBy() & "  -  VBA edition"
    btnDeleteAll.Font.Name = btnTheme.Font.Name
    btnDeleteAll.Font.size = btnTheme.Font.size
    btnDeleteAll.Font.Bold = btnTheme.Font.Bold
    cboFenceMode.AddItem "Shape"
    cboFenceMode.AddItem "Block"
    cboFenceMode.ListIndex = 0
    cboScale.AddItem "1"" = 50'"
    cboScale.AddItem "1"" = 100'"
    cboScale.AddItem "1"" = 200'"
    Select Case CLng(Val(GetSetting(REG_APP, "Settings", "DrawingScale", "100")))
        Case 50
            cboScale.ListIndex = 0
        Case 200
            cboScale.ListIndex = 2
        Case Else
            cboScale.ListIndex = 1
    End Select
    txtPlotText.text = GetSetting(REG_APP, "Settings", "PlotTextIn", "0.20")
    cboZone.AddItem "South (EPSG 2284)"
    cboZone.AddItem "North (EPSG 2283)"
    If UCase$(GetSetting(REG_APP, "Settings", "Zone", "S")) = "N" Then
        cboZone.ListIndex = 1
    Else
        cboZone.ListIndex = 0
    End If
    If Len(GetSetting(REG_APP, "Settings", "Zone", "")) = 0 Then SaveSetting REG_APP, "Settings", "Zone", "S"
    txtRdId.text = "RD-01"
    txtRdDist.text = "1"
    chkRdAuto.Value = True
    chkAutoMarkers.Value = True
    chkAutoGis.Value = (GetSetting(REG_APP, "Settings", "AutoGis", "0") = "1")
    chkLfp.Value = (GetSetting(REG_APP, "Settings", "Lfp", "1") = "1")
    lblFence.Caption = "Fence: not loaded"
    LoadMethodSettings
    mLoading = False
    ApplyMethodSettings
    ApplyOptionSettings
    ApplyTextSettings
    mDark = (GetSetting(REG_APP, "Settings", "Theme", "light") = "dark")
    ApplyTheme
    RefreshLibraries
    ClearResultView
    SetProgress 0, 0, "Batch progress: idle"
    mpTabs.Value = 1
    AppendLog "USGS StreamStats Fence Query (VBA edition) v" & AppVersion()
    AppendLog "Ready. Place a fence to load streams, or place culvert markers, then Run Batch."
End Sub

Private Sub UserForm_Activate()
    StartupCheck
End Sub

Public Sub StartupCheck()
    ' The web services are checked once, when the window first opens (click the status to check again).
    If mSvcChecked Then Exit Sub
    mSvcChecked = True
    If Not gSilent Then RunServiceCheck
End Sub

Private Sub UserForm_Terminate()
    Set gLogSink = Nothing
End Sub

'----------------------------------------------------------------------------------------------
' Log, status, messages
'----------------------------------------------------------------------------------------------
Public Sub AppendLog(ByVal s As String)
    Dim t As String
    On Error Resume Next
    t = txtLog.text
    If Len(t) > 30000 Then t = Mid$(t, Len(t) - 20000)
    txtLog.text = t & Format$(Now, "hh:nn:ss") & "  " & s & vbCrLf
    txtLog.SelStart = Len(txtLog.text)
End Sub

Private Sub SetStatusBar(ByVal s As String)
    lblStatus.Caption = "  " & s
End Sub

Private Sub Notify(ByVal s As String, Optional ByVal style As VbMsgBoxStyle = vbInformation)
    AppendLog s
    If Not gSilent Then MsgBox s, style, "USGS StreamStats VDOT (VBA)"
End Sub


Private Sub btnClearLog_Click()
    txtLog.text = ""
End Sub

'----------------------------------------------------------------------------------------------
' Theme
'----------------------------------------------------------------------------------------------
Private Sub btnTheme_Click()
    mDark = Not mDark
    SaveSetting REG_APP, "Settings", "Theme", IIf(mDark, "dark", "light")
    ApplyTheme
End Sub

Private Sub ApplyTheme()
    Dim c As Object, bg As Long, card As Long, fg As Long, muted As Long, field As Long, btn As Long
    Dim navy As Long, accent As Long, danger As Long, i As Long
    navy = RGB(15, 31, 61)
    accent = RGB(13, 148, 136)
    danger = RGB(220, 38, 38)
    If mDark Then
        bg = RGB(15, 23, 42)
        card = RGB(30, 41, 59)
        fg = RGB(226, 232, 240)
        muted = RGB(148, 163, 184)
        field = RGB(51, 65, 85)
        btn = RGB(71, 85, 105)
    Else
        bg = RGB(238, 242, 247)
        card = RGB(255, 255, 255)
        fg = RGB(15, 23, 42)
        muted = RGB(100, 116, 139)
        field = RGB(255, 255, 255)
        btn = RGB(226, 232, 240)
    End If
    On Error Resume Next
    Me.BackColor = bg
    mpTabs.BackColor = bg
    For Each c In Me.Controls
        Select Case TypeName(c)
            Case "Label"
                If c.tag = "hdr" Then
                    c.BackColor = navy
                    c.ForeColor = RGB(255, 255, 255)
                ElseIf c.tag = "hdrsub" Then
                    c.BackColor = navy
                    c.ForeColor = RGB(148, 163, 184)
                ElseIf c.tag = "colhdr" Then
                    c.BackColor = navy
                    c.ForeColor = RGB(255, 255, 255)
                ElseIf c.tag = "svc" Then
                    c.BackColor = navy
                ElseIf c.tag = "pagebg" Then
                    c.BackColor = card
                ElseIf c.tag = "progback" Then
                    c.BackColor = btn
                ElseIf c.tag = "progfill" Then
                    c.BackColor = accent
                ElseIf c.tag = "muted" Then
                    c.BackColor = card
                    c.ForeColor = muted
                ElseIf c.tag = "form" Then
                    c.BackColor = bg
                    c.ForeColor = fg
                Else
                    c.BackColor = card
                    c.ForeColor = fg
                End If
            Case "CommandButton"
                If c.tag = "accent" Then
                    c.BackColor = accent
                    c.ForeColor = RGB(255, 255, 255)
                ElseIf c.tag = "danger" Then
                    c.BackColor = danger
                    c.ForeColor = RGB(255, 255, 255)
                ElseIf c.tag = "hdrbtn" Then
                    c.BackColor = RGB(19, 49, 92)
                    c.ForeColor = RGB(226, 232, 240)
                Else
                    c.BackColor = btn
                    c.ForeColor = fg
                End If
            Case "TextBox", "ComboBox", "ListBox"
                If c.tag = "log" Then
                    c.BackColor = RGB(11, 18, 32)
                    c.ForeColor = RGB(203, 213, 225)
                Else
                    c.BackColor = field
                    c.ForeColor = fg
                End If
            Case "CheckBox"
                c.BackColor = card
                c.ForeColor = fg
        End Select
    Next c
    btnTheme.Caption = IIf(mDark, "Light Mode", "Dark Mode")
    Me.Repaint
End Sub

'----------------------------------------------------------------------------------------------
' Settings tab
'----------------------------------------------------------------------------------------------
Private Sub cboScale_Change()
    If Not mLoading Then ApplyTextSettings
End Sub

Private Sub txtPlotText_Change()
    If Not mLoading Then ApplyTextSettings
End Sub

Private Sub ApplyTextSettings()
    ' Drawing scale x plotted text height -> DGN text height for marker labels.
    Dim scaleFt As Double, plotIn As Double
    Select Case cboScale.ListIndex
        Case 0
            scaleFt = 50#
        Case 2
            scaleFt = 200#
        Case Else
            scaleFt = 100#
    End Select
    plotIn = Val(Replace(txtPlotText.text, ",", "."))
    If plotIn < 0.01 Or plotIn > 2# Then
        lblTextFt.Caption = "Enter 0.01 - 2.00 in"
        Exit Sub
    End If
    gDrawingScale = scaleFt
    gPlotTextIn = plotIn
    lblTextFt.Caption = "= " & Format$(TextHeightFt(), "0.0") & " ft in DGN"
    If Not mLoading Then
        SaveSetting REG_APP, "Settings", "DrawingScale", CStr(CLng(scaleFt))
        SaveSetting REG_APP, "Settings", "PlotTextIn", Replace(Format$(plotIn, "0.00"), ",", ".")
    End If
End Sub

Private Sub cboZone_Change()
    If mLoading Then Exit Sub
    SaveSetting REG_APP, "Settings", "Zone", IIf(cboZone.ListIndex = 1, "N", "S")
    AppendLog "State Plane zone set to " & IIf(cboZone.ListIndex = 1, "North", "South") & "."
End Sub

Private Function ReadCOverride(ByRef ok As Boolean) As Double
    Dim t As String, v As Double
    ok = True
    t = Trim$(Replace(txtCOverride.text, ",", "."))
    If Len(t) = 0 Then Exit Function
    v = Val(t)
    If Not IsNumeric(t) Or v <= 0# Or v > 1# Then
        ok = False
        Notify "Rational C override must be a number between 0 and 1 (or blank).", vbExclamation
        Exit Function
    End If
    ReadCOverride = v
End Function


'----------------------------------------------------------------------------------------------
' Libraries tab
'----------------------------------------------------------------------------------------------
Private Sub RefreshLibraries()
    lstLibs.Clear
    AddLibRow "MSXML2.ServerXMLHTTP.6.0", "USGS / NOAA web requests"
    AddLibRow "Scripting.Dictionary", "Reading web service results (JSON)"
End Sub

Private Sub AddLibRow(ByVal progId As String, ByVal purpose As String)
    Dim o As Object, r As Long
    On Error Resume Next
    Set o = CreateObject(progId)
    On Error GoTo 0
    lstLibs.AddItem progId
    r = lstLibs.ListCount - 1
    lstLibs.List(r, 1) = IIf(o Is Nothing, "Not available", "Loaded")
    lstLibs.List(r, 2) = purpose
End Sub

Private Sub btnLibRefresh_Click()
    RefreshLibraries
End Sub

Private Sub btnLibCopy_Click()
    Dim i As Long, s As String, d As Object
    For i = 0 To lstLibs.ListCount - 1
        s = s & lstLibs.List(i, 0) & vbTab & lstLibs.List(i, 1) & vbTab & lstLibs.List(i, 2) & vbCrLf
    Next i
    On Error Resume Next
    Set d = New MSForms.DataObject
    d.SetText s
    d.PutInClipboard
    AppendLog "Library status copied to the clipboard."
End Sub

'----------------------------------------------------------------------------------------------
' Culvert Analysis tab - queue
'----------------------------------------------------------------------------------------------
Private Function QueueHasId(ByVal id As String) As Boolean
    Dim it As Variant
    For Each it In mQueue
        If StrComp(CStr(it("id")), id, vbTextCompare) = 0 Then QueueHasId = True
    Next it
End Function

Private Function NextCulvertId() As String
    Dim n As Long
    n = mQueue.Count + 1
    Do While QueueHasId("CULV-" & Format$(n, "00"))
        n = n + 1
    Loop
    NextCulvertId = "CULV-" & Format$(n, "00")
End Function

Public Sub AddQueueItem(ByVal id As String, ByVal ctype As String, ByVal x As Double, ByVal y As Double, _
                        Optional ByVal markersPlaced As Boolean = False)
    Dim it As Object, r As Long, baseId As String, n As Long
    baseId = id
    n = 2
    Do While QueueHasId(id)
        id = baseId & "_" & n
        n = n + 1
    Loop
    Set it = CreateObject("Scripting.Dictionary")
    it("id") = id
    it("ctype") = ctype
    it("x") = x
    it("y") = y
    it("status") = "Queued"
    it("markers") = markersPlaced
    Set it("res") = Nothing
    mQueue.Add it
    lstQueue.AddItem id
    r = lstQueue.ListCount - 1
    lstQueue.List(r, 1) = ctype
    If ctype = "WGS84_LonLat" Then
        lstQueue.List(r, 2) = Format$(x, "0.000000")
        lstQueue.List(r, 3) = Format$(y, "0.000000")
    Else
        lstQueue.List(r, 2) = Format$(x, "#,##0.00")
        lstQueue.List(r, 3) = Format$(y, "#,##0.00")
    End If
    lstQueue.List(r, 4) = "Queued"
End Sub

Private Sub SetItemStatus(ByVal idx As Long, ByVal s As String)
    mQueue(idx)("status") = s
    lstQueue.List(idx - 1, 4) = s
End Sub

Private Function HasResult(ByVal idx As Long) As Boolean
    HasResult = Not (mQueue(idx)("res") Is Nothing)
End Function

Private Sub btnPlaceMarkers_Click()
    PlaceMarkers
End Sub

Public Sub PlaceMarkers()
    ' Click culvert locations in the DGN; each click adds a culvert (and its markers) to the queue.
    Dim msg As CadInputMessage, id As String, n As Long
    If mBusy Then Exit Sub
    On Error GoTo done
    Me.Hide
    ShowCommand "USGS StreamStats"
    Do
        ShowPrompt "Click a culvert location (Reset / right-click to finish)"
        Set msg = CadInputQueue.GetInput(msdCadInputTypeDataPoint, msdCadInputTypeReset, _
                                         msdCadInputTypeReset, msdCadInputTypeReset)
        If msg.InputType <> msdCadInputTypeDataPoint Then Exit Do
        id = NextCulvertId()
        If chkAutoMarkers.Value Then DrawMarkers msg.Point.x, msg.Point.y, msg.Point.x, msg.Point.y, id
        AddQueueItem id, "DGN_Master", msg.Point.x, msg.Point.y, CBool(chkAutoMarkers.Value)
        n = n + 1
    Loop
done:
    If Err.Number <> 0 Then AppendLog "[Pick Error] " & Err.Description
    ShowPrompt ""
    ShowCommand ""
    Me.Show vbModeless
    AppendLog "Marker placement finished: " & n & " culvert(s) added to the queue."
End Sub

Private Sub btnImportCsv_Click()
    Dim f As String
    If mBusy Then Exit Sub
    f = PickCsvFile()
    If Len(f) > 0 Then AddCsvFile f
End Sub

Public Sub AddCsvFile(ByVal path As String)
    ' CSV rows: ID, X/Lon, Y/Lat. A header row is skipped; lat,lon order is detected for US data;
    ' values outside the lat/lon range are treated as DGN master units.
    Dim fnum As Integer, ln As String, parts() As String, a As Double, b As Double, t As Double, n As Long
    Dim sa As String, sb As String
    On Error GoTo Bad
    fnum = FreeFile
    Open path For Input As #fnum
    Do While Not EOF(fnum)
        Line Input #fnum, ln
        parts = Split(ln, ",")
        If UBound(parts) >= 2 Then
            sa = Trim$(parts(1))
            sb = Trim$(parts(2))
            If IsNumeric(sa) And IsNumeric(sb) And Len(Trim$(parts(0))) > 0 Then
                a = Val(sa)
                b = Val(sb)
                If Abs(a) <= 90# And Abs(b) <= 180# And a > 0# And b < 0# Then
                    t = a
                    a = b
                    b = t
                End If
                If Abs(a) <= 180# And Abs(b) <= 90# Then
                    AddQueueItem Trim$(parts(0)), "WGS84_LonLat", a, b
                Else
                    AddQueueItem Trim$(parts(0)), "DGN_Master", a, b
                End If
                n = n + 1
            End If
        End If
    Loop
    Close #fnum
    AppendLog "Imported " & n & " culvert(s) from " & path
    Exit Sub
Bad:
    On Error Resume Next
    Close #fnum
    Notify "Could not read the CSV file: " & Err.Description, vbExclamation
End Sub

Private Sub btnRemoveSel_Click()
    Dim i As Long
    If mBusy Then Exit Sub
    For i = lstQueue.ListCount - 1 To 0 Step -1
        If lstQueue.Selected(i) Then
            lstQueue.RemoveItem i
            mQueue.Remove i + 1
        End If
    Next i
    RefreshResultList
End Sub

Private Sub btnClearQueue_Click()
    If mBusy Then Exit Sub
    lstQueue.Clear
    Set mQueue = New Collection
    RefreshResultList
End Sub

'----------------------------------------------------------------------------------------------
' Culvert Analysis tab - batch
'----------------------------------------------------------------------------------------------
Private Sub btnRunBatch_Click()
    RunBatch False
End Sub

Private Sub btnRetry_Click()
    RunBatch True
End Sub

Private Sub btnStop_Click()
    If mBusy Then
        gCancel = True
        AppendLog "Stop requested - the batch ends after the current web request."
    End If
End Sub

Private Sub SetBusy(ByVal busy As Boolean)
    mBusy = busy
    On Error Resume Next
    btnRunBatch.Enabled = Not busy
    btnRetry.Enabled = Not busy
    btnPlaceMarkers.Enabled = Not busy
    btnImportCsv.Enabled = Not busy
    btnRemoveSel.Enabled = Not busy
    btnClearQueue.Enabled = Not busy
    btnDrawAreas.Enabled = Not busy
    btnPlaceFence.Enabled = Not busy
    btnReloadFence.Enabled = Not busy
    btnQueueStreams.Enabled = Not busy
    btnRdTrace.Enabled = Not busy
    btnRdPick.Enabled = Not busy
    btnStop.Enabled = busy
    SetStatusBar IIf(busy, "Working...", "Ready")
End Sub

Private Sub SetProgress(ByVal done As Long, ByVal total As Long, ByVal s As String)
    Dim frac As Double
    If total > 0 Then frac = done / total
    lblProg.Caption = s
    lblProgPct.Caption = Format$(frac * 100#, "0") & "%"
    lblProgFill.Width = lblProgBack.Width * frac
    Me.Repaint
End Sub

Public Sub RunBatch(ByVal failedOnly As Boolean)
    Dim i As Long, total As Long, done As Long, okCount As Long, failCount As Long
    Dim it As Object, res As Object, lon As Double, lat As Double, cOk As Boolean, t0 As Single
    If mBusy Then Exit Sub
    If mQueue.Count = 0 Then
        Notify "Add culvert points before running the batch.", vbExclamation
        Exit Sub
    End If
    For i = 1 To mQueue.Count
        If Not (failedOnly And HasResult(i)) Then total = total + 1
    Next i
    If total = 0 Then
        Notify "Every culvert in the queue already has a result.", vbInformation
        Exit Sub
    End If
    gCOverride = ReadCOverride(cOk)
    If Not cOk Then Exit Sub

    gCancel = False
    SetBusy True
    t0 = Timer
    EnsureStreamsDrawn
    AppendLog IIf(failedOnly, "Retrying ", "Starting batch delineation for ") & total & IIf(failedOnly, " failed", "") & " culvert(s)..."
    For i = 1 To mQueue.Count
        If Not (failedOnly And HasResult(i)) Then
            Set it = mQueue(i)
            If gCancel Then
                SetItemStatus i, "Skipped (stopped)"
            Else
                Set it("res") = Nothing
                SetItemStatus i, "Delineating..."
                SetProgress done, total, "Culvert " & CStr(it("id")) & " (" & (done + 1) & "/" & total & ")"
                If CStr(it("ctype")) = "WGS84_LonLat" Then
                    lon = CDbl(it("x"))
                    lat = CDbl(it("y"))
                Else
                    DgnToLonLat CDbl(it("x")), CDbl(it("y")), lon, lat
                End If
                Set res = Nothing
                If ComputeCulvert(lon, lat, CStr(it("id")), res) Then
                    Set it("res") = res
                    If DrawOne(i) Then
                        SetItemStatus i, "Completed"
                    Else
                        SetItemStatus i, "Delineated; DGN draw failed"
                    End If
                    okCount = okCount + 1
                Else
                    If gCancel Then
                        SetItemStatus i, "Skipped (stopped)"
                    Else
                        SetItemStatus i, "Failed: " & CStr(res("error"))
                        failCount = failCount + 1
                    End If
                End If
                done = done + 1
            End If
        End If
    Next i
    SetProgress total, total, IIf(gCancel, "Batch stopped: ", "Batch complete: ") & okCount & " delineated, " & failCount & " failed  (" & Format$(Timer - t0, "0") & " s)"
    SetBusy False
    RefreshResultList
    AppendLog IIf(gCancel, "Batch stopped: ", "Batch complete: ") & okCount & " delineated, " & failCount & " failed."
    If okCount > 0 Then
        PlaceTables failedOnly
        ExportBatchCsv mQueue
        If gAutoGis Then AutoSaveGis "Batch_Watersheds", mQueue, gStreams
        mpTabs.Value = 2
    End If
    If Not gSilent Then MsgBox IIf(gCancel, "Batch stopped.", "Batch complete.") & vbCrLf & vbCrLf & okCount & " delineated, " & failCount & " failed.", vbInformation, "USGS StreamStats VDOT (VBA)"
End Sub

Private Function DrawOne(ByVal idx As Long) As Boolean
    ' Draws the drainage area (once) and, when wanted, the markers for one culvert.
    Dim it As Object, res As Object, ix As Double, iy As Double, px As Double, py As Double
    On Error GoTo Bad
    Set it = mQueue(idx)
    Set res = it("res")
    If Not res("drawn") Then DrawCulvert res
    If chkAutoMarkers.Value And Not it("markers") Then
        LonLatToDgn CDbl(res("lon")), CDbl(res("lat")), ix, iy
        LonLatToDgn CDbl(res("s_lon")), CDbl(res("s_lat")), px, py
        DrawMarkers ix, iy, px, py, CStr(it("id"))
        it("markers") = True
    End If
    DrawOne = True
    Exit Function
Bad:
    AppendLog "[DGN Draw Error] " & CStr(mQueue(idx)("id")) & ": " & Err.Description
End Function

Private Sub btnDrawAreas_Click()
    Dim i As Long, n As Long
    If mBusy Then Exit Sub
    For i = 1 To mQueue.Count
        If HasResult(i) Then
            If DrawOne(i) Then n = n + 1
        End If
    Next i
    If n = 0 Then
        Notify "Run a delineation first, then draw its drainage area and markers.", vbExclamation
    Else
        AppendLog "Drainage areas and markers are in the DGN for " & n & " culvert(s)."
    End If
End Sub






'----------------------------------------------------------------------------------------------
' Results & Comparison tab
'----------------------------------------------------------------------------------------------
Private Sub RefreshResultList()
    Dim i As Long, keep As String
    keep = cboResult.text
    cboResult.Clear
    For i = 1 To mQueue.Count
        If HasResult(i) Then cboResult.AddItem CStr(mQueue(i)("id"))
    Next i
    If cboResult.ListCount = 0 Then
        ClearResultView
        Exit Sub
    End If
    For i = 0 To cboResult.ListCount - 1
        If cboResult.List(i) = keep Then cboResult.ListIndex = i
    Next i
    If cboResult.ListIndex < 0 Then cboResult.ListIndex = 0
End Sub

Private Sub cboResult_Change()
    If cboResult.ListIndex >= 0 Then ShowResult cboResult.text
End Sub

Private Sub ClearResultView()
    lblAreaV.Caption = "-"
    lblSlopeV.Caption = "-"
    lblCV.Caption = "-"
    lblTcV.Caption = "-"
    lblGovV.Caption = "-"
    lstResults.Clear
    lblSources.Caption = "Data sources for the selected culvert are listed here after a delineation."
End Sub

Private Function Q1(ByVal v As Variant) As String
    If IsNumeric(v) Then
        If CDbl(v) > 0# Then Q1 = Format$(CDbl(v), "#,##0.0") Else Q1 = "-"
    Else
        Q1 = "-"
    End If
End Function

Private Sub ShowResult(ByVal id As String)
    Dim i As Long, res As Object, yrs As Variant, qU As Variant, qR As Variant, qD As Variant, r As Long
    Dim qN As Variant, qB As Variant, src As String, qA As Variant, skipped As String
    Dim diff As String
    Set res = Nothing
    For i = 1 To mQueue.Count
        If CStr(mQueue(i)("id")) = id Then
            If HasResult(i) Then Set res = mQueue(i)("res")
        End If
    Next i
    If res Is Nothing Then
        ClearResultView
        Exit Sub
    End If
    lblAreaV.Caption = Format$(res("sqmi"), "0.000") & " sq mi  /  " & Format$(res("acres"), "#,##0.0") & " ac"
    lblSlopeV.Caption = Format$(res("slope"), "0.0") & " ft/mi"
    lblCV.Caption = Format$(res("c"), "0.00")
    lblTcV.Caption = Format$(res("tc"), "0.0") & " min"
    lblGovV.Caption = CStr(res("gov"))
    yrs = Array(2, 5, 10, 25, 50, 100, 200, 500)
    qU = res("q_usgs")
    qR = res("q_rat")
    qD = res("q_des")
    qN = res("q_nrcs")
    qB = res("q_urban")
    qA = res("q_and")
    lstResults.Clear
    For i = 0 To 7
        lstResults.AddItem "Q" & yrs(i) & "  (" & yrs(i) & "-yr)"
        r = lstResults.ListCount - 1
        lstResults.List(r, 1) = Q1(qR(i))
        lstResults.List(r, 2) = Q1(qN(i))
        lstResults.List(r, 3) = Q1(qA(i))
        lstResults.List(r, 4) = Q1(qU(i))
        lstResults.List(r, 5) = Q1(qB(i))
        lstResults.List(r, 6) = Q1(qD(i))
        diff = "N/A"
        If qU(i) > 0# And qD(i) > 0# Then diff = Format$((qD(i) - qU(i)) / qU(i) * 100#, "+0.0;-0.0") & "%"
        lstResults.List(r, 7) = diff
    Next i
    src = "Basin data: " & CStr(res("bc_src")) & _
        "   |   Rainfall: " & IIf(res("have_atlas"), "NOAA Atlas 14 (at the pour point)", "Built-in statewide IDF coefficients (placeholder)") & _
        "   |   Tc length: " & CStr(res("tc_len_src")) & IIf(res("tc_len") > 0#, " (" & Format$(res("tc_len"), "#,##0") & " ft)", "") & _
        "   |   Regression: " & IIf(res("have_nss"), "USGS NSS regression equations", "not available")
    If res("cn") > 0# Then src = src & "   |   NRCS CN " & Format$(res("cn"), "0") & ": " & CStr(res("cn_src"))
    If res("urban_pct") >= 0# Then src = src & "   |   Urban area: " & Format$(res("urban_pct"), "0.0") & "% (" & CStr(res("urban_src")) & ")"
    If qR(2) <= 0# Then skipped = skipped & "; Rational Method"
    If qN(2) <= 0# Then skipped = skipped & "; NRCS TR-55"
    If qU(2) <= 0# Then skipped = skipped & "; USGS rural regression"
    If qB(2) <= 0# Then skipped = skipped & "; USGS urban regression"
    If gUseAnderson And qA(2) <= 0# Then skipped = skipped & "; Anderson"
    If Len(skipped) > 0 Then src = src & "   |   Not calculated (off, outside its VDOT size range, or missing data): " & Mid$(skipped, 3)
    lblSources.Caption = src & "   |   Verify against the VDOT Drainage Manual before design use."
End Sub


'----------------------------------------------------------------------------------------------
' Flow (Raindrop) Path tab
'----------------------------------------------------------------------------------------------
Private Sub btnRdPick_Click()
    Dim msg As CadInputMessage, lon As Double, lat As Double, tid As String, picked As Boolean
    If mBusy Then Exit Sub
    On Error GoTo done
    Me.Hide
    ShowCommand "USGS StreamStats"
    ShowPrompt "Click the raindrop start point (Reset to cancel)"
    Set msg = CadInputQueue.GetInput(msdCadInputTypeDataPoint, msdCadInputTypeReset, _
                                     msdCadInputTypeReset, msdCadInputTypeReset)
    If msg.InputType = msdCadInputTypeDataPoint Then
        DgnToLonLat msg.Point.x, msg.Point.y, lon, lat
        txtRdLon.text = Replace(Format$(lon, "0.000000"), ",", ".")
        txtRdLat.text = Replace(Format$(lat, "0.000000"), ",", ".")
        AppendLog "Raindrop start captured at " & txtRdLon.text & ", " & txtRdLat.text & "."
        tid = Trim$(txtRdId.text)
        If Len(tid) = 0 Then tid = "RD-01"
        DrawLabeledPoint msg.Point.x, msg.Point.y, tid & " START  " & txtRdLat.text & ", " & txtRdLon.text
        mRdStartDrawn = True
        mRdStartLon = lon
        mRdStartLat = lat
        picked = True
    End If
done:
    If Err.Number <> 0 Then AppendLog "[DGN Draw Error] Start point: " & Err.Description
    ShowPrompt ""
    ShowCommand ""
    Me.Show vbModeless
    If picked And Not gSilent Then
        If MsgBox("Start point captured:" & vbCrLf & txtRdLat.text & ", " & txtRdLon.text & vbCrLf & vbCrLf & "Trace flow path " & tid & " now?", _
                  vbQuestion + vbYesNo, "Flow (Raindrop) Path") = vbYes Then RunRaindropTrace
    End If
End Sub





'----------------------------------------------------------------------------------------------
' Tables in the DGN and exports
'----------------------------------------------------------------------------------------------
Private Sub PlaceTables(ByVal replacePrevious As Boolean)
    Dim tx As Double, ty As Double
    If mOriginSet Then
        tx = mOriginX
        ty = mOriginY
    Else
        DefaultTableOrigin mQueue, tx, ty
    End If
    PlaceBatchTables mQueue, tx, ty, replacePrevious
End Sub

Private Sub btnPickOrigin_Click()
    Dim msg As CadInputMessage
    If mBusy Then Exit Sub
    On Error GoTo done
    Me.Hide
    ShowCommand "USGS StreamStats"
    ShowPrompt "Click the top-left corner for the batch results table (Reset to cancel)"
    Set msg = CadInputQueue.GetInput(msdCadInputTypeDataPoint, msdCadInputTypeReset, _
                                     msdCadInputTypeReset, msdCadInputTypeReset)
    If msg.InputType = msdCadInputTypeDataPoint Then
        SetTableOrigin msg.Point.x, msg.Point.y
    End If
done:
    ShowPrompt ""
    ShowCommand ""
    Me.Show vbModeless
End Sub

Public Sub SetTableOrigin(ByVal x As Double, ByVal y As Double)
    mOriginSet = True
    mOriginX = x
    mOriginY = y
    lblOrigin.Caption = "Table origin: " & Format$(x, "#,##0") & ", " & Format$(y, "#,##0")
    AppendLog "Table origin set to " & Format$(x, "#,##0.0") & ", " & Format$(y, "#,##0.0") & "."
End Sub

Private Sub btnExportExcel_Click()
    DoExportExcel
End Sub

Public Sub DoExportExcel()
    Dim p As String
    If mBusy Then Exit Sub
    p = ExportExcel(mQueue)
    If Len(p) > 0 Then
        Notify "Saved: " & p, vbInformation
    Else
        Notify "Nothing was exported. Run a delineation first (Excel must be installed).", vbExclamation
    End If
End Sub



'----------------------------------------------------------------------------------------------
' Settings tab - peak discharge methods (VDOT Drainage Manual Table 6-2)
'----------------------------------------------------------------------------------------------
Private Sub LoadMethodSettings()
    MethodDefaults
    cboDesign.AddItem "VDOT default (by drainage area)"
    cboDesign.AddItem "Rational Method"
    cboDesign.AddItem "NRCS Graphical Peak Discharge (TR-55)"
    cboDesign.AddItem "USGS Regional Regression - Rural"
    cboDesign.AddItem "USGS Regional Regression - Urban"
    cboHsg.AddItem "Auto (USDA SSURGO)"
    cboHsg.AddItem "A"
    cboHsg.AddItem "B"
    cboHsg.AddItem "C"
    cboHsg.AddItem "D"
    cboAnderson.AddItem "Natural rural watershed"
    cboAnderson.AddItem "Developed, partially channelized"
    cboAnderson.AddItem "Completely developed and sewered"
    cboDesign.AddItem "Anderson Method (USGS)"
    chkMAnderson.Value = (GetSetting(REG_APP, "Methods", "Anderson", "0") = "1")
    Select Case GetSetting(REG_APP, "Methods", "AndersonType", "natural")
        Case "developed"
            cboAnderson.ListIndex = 1
        Case "sewered"
            cboAnderson.ListIndex = 2
        Case Else
            cboAnderson.ListIndex = 0
    End Select
    chkMRational.Value = (GetSetting(REG_APP, "Methods", "Rational", "1") = "1")
    chkMNrcs.Value = (GetSetting(REG_APP, "Methods", "Nrcs", "1") = "1")
    chkMRural.Value = (GetSetting(REG_APP, "Methods", "Rural", "1") = "1")
    chkMUrban.Value = (GetSetting(REG_APP, "Methods", "Urban", "1") = "1")
    chkUseRanges.Value = (GetSetting(REG_APP, "Methods", "UseRanges", "1") = "1")
    Select Case GetSetting(REG_APP, "Methods", "Design", "auto")
        Case "rational"
            cboDesign.ListIndex = 1
        Case "nrcs"
            cboDesign.ListIndex = 2
        Case "rural"
            cboDesign.ListIndex = 3
        Case "urban"
            cboDesign.ListIndex = 4
        Case "anderson"
            cboDesign.ListIndex = 5
        Case Else
            cboDesign.ListIndex = 0
    End Select
    txtCN.text = GetSetting(REG_APP, "Methods", "CN", "")
    Select Case GetSetting(REG_APP, "Methods", "Hsg", "auto")
        Case "A"
            cboHsg.ListIndex = 1
        Case "B"
            cboHsg.ListIndex = 2
        Case "C"
            cboHsg.ListIndex = 3
        Case "D"
            cboHsg.ListIndex = 4
        Case Else
            cboHsg.ListIndex = 0
    End Select
End Sub

Private Sub ApplyMethodSettings()
    ' Pushes the method choices to the calculator (used by the next delineation) and remembers them.
    Dim t As String, v As Double
    If mLoading Then Exit Sub
    gUseRational = CBool(chkMRational.Value)
    gUseNrcs = CBool(chkMNrcs.Value)
    gUseRural = CBool(chkMRural.Value)
    gUseUrban = CBool(chkMUrban.Value)
    gUseAnderson = CBool(chkMAnderson.Value)
    Select Case cboAnderson.ListIndex
        Case 1
            gAndersonType = "developed"
        Case 2
            gAndersonType = "sewered"
        Case Else
            gAndersonType = "natural"
    End Select
    gUseRanges = CBool(chkUseRanges.Value)
    Select Case cboDesign.ListIndex
        Case 1
            gDesign = "rational"
        Case 2
            gDesign = "nrcs"
        Case 3
            gDesign = "rural"
        Case 4
            gDesign = "urban"
        Case 5
            gDesign = "anderson"
        Case Else
            gDesign = "auto"
    End Select
    t = Trim$(Replace(txtCN.text, ",", "."))
    gCN = 0#
    If Len(t) = 0 Then
        lblCnNote.Caption = "CN is estimated for each basin from NLCD 2021 land cover and the soil group."
    Else
        v = Val(t)
        If IsNumeric(t) And v >= 30# And v <= 98# Then
            gCN = v
            lblCnNote.Caption = "CN " & t & " is used for every culvert."
        Else
            lblCnNote.Caption = "Enter a curve number from 30 to 98, or leave blank to estimate it."
        End If
    End If
    Select Case cboHsg.ListIndex
        Case 1
            gHsg = "A"
        Case 2
            gHsg = "B"
        Case 3
            gHsg = "C"
        Case 4
            gHsg = "D"
        Case Else
            gHsg = "auto"
    End Select
    gMethodsInit = True
    SaveSetting REG_APP, "Methods", "Rational", IIf(gUseRational, "1", "0")
    SaveSetting REG_APP, "Methods", "Nrcs", IIf(gUseNrcs, "1", "0")
    SaveSetting REG_APP, "Methods", "Rural", IIf(gUseRural, "1", "0")
    SaveSetting REG_APP, "Methods", "Urban", IIf(gUseUrban, "1", "0")
    SaveSetting REG_APP, "Methods", "Anderson", IIf(gUseAnderson, "1", "0")
    SaveSetting REG_APP, "Methods", "AndersonType", gAndersonType
    SaveSetting REG_APP, "Methods", "UseRanges", IIf(gUseRanges, "1", "0")
    SaveSetting REG_APP, "Methods", "Design", gDesign
    SaveSetting REG_APP, "Methods", "CN", IIf(gCN > 0#, t, "")
    SaveSetting REG_APP, "Methods", "Hsg", gHsg
End Sub

Private Sub chkMRational_Click()
    ApplyMethodSettings
End Sub

Private Sub chkMNrcs_Click()
    ApplyMethodSettings
End Sub

Private Sub chkMRural_Click()
    ApplyMethodSettings
End Sub

Private Sub chkMUrban_Click()
    ApplyMethodSettings
End Sub

Private Sub chkMAnderson_Click()
    ApplyMethodSettings
    If chkMAnderson.Value And Not chkLfp.Value And Not mLoading Then
        AppendLog "The Anderson method needs the longest flowpath - turn on 'Compute longest flowpath' as well."
    End If
End Sub

Private Sub cboAnderson_Change()
    ApplyMethodSettings
End Sub

Private Sub chkUseRanges_Click()
    ApplyMethodSettings
End Sub

Private Sub cboDesign_Change()
    ApplyMethodSettings
End Sub

Private Sub cboHsg_Change()
    ApplyMethodSettings
End Sub

Private Sub txtCN_Change()
    ApplyMethodSettings
End Sub

Private Sub btnMethodReset_Click()
    mLoading = True
    chkMRational.Value = True
    chkMNrcs.Value = True
    chkMRural.Value = True
    chkMUrban.Value = True
    chkMAnderson.Value = False
    cboAnderson.ListIndex = 0
    chkUseRanges.Value = True
    cboDesign.ListIndex = 0
    txtCN.text = ""
    cboHsg.ListIndex = 0
    mLoading = False
    ApplyMethodSettings
    AppendLog "Peak-discharge methods reset to the VDOT Drainage Manual defaults."
End Sub

'----------------------------------------------------------------------------------------------
' Settings tab - options
'----------------------------------------------------------------------------------------------
Private Sub ApplyOptionSettings()
    gAutoGis = CBool(chkAutoGis.Value)
    gComputeLfp = CBool(chkLfp.Value)
    If mLoading Then Exit Sub
    SaveSetting REG_APP, "Settings", "AutoGis", IIf(gAutoGis, "1", "0")
    SaveSetting REG_APP, "Settings", "Lfp", IIf(gComputeLfp, "1", "0")
End Sub

Private Sub chkAutoGis_Click()
    ApplyOptionSettings
End Sub

Private Sub chkLfp_Click()
    ApplyOptionSettings
End Sub

Private Sub btnKmlStyle_Click()
    dlgKmlStyle.Show vbModal
End Sub

'----------------------------------------------------------------------------------------------
' Service check (status bar)
'----------------------------------------------------------------------------------------------
Private Sub lblSvc_Click()
    RunServiceCheck
End Sub

Public Sub RunServiceCheck()
    ' Probes the USGS / MRLC / NOAA / USDA services and reports which are down.
    Dim labels As Variant, impacts As Variant, oks() As Boolean, details() As String, down As Long, i As Long, tip As String
    If mSvcBusy Or mBusy Then Exit Sub
    mSvcBusy = True
    On Error GoTo done
    lblSvc.Caption = "Services: checking..."
    lblSvc.ForeColor = RGB(148, 163, 184)
    Me.Repaint
    down = CheckServices(labels, oks, details, impacts)
    For i = 0 To UBound(labels)
        tip = tip & IIf(oks(i), "OK    ", "DOWN  ") & labels(i) & " (" & details(i) & ")" & vbCrLf
        If Not oks(i) Then AppendLog "[Service down] " & labels(i) & " (" & details(i) & ") - " & impacts(i) & "."
    Next i
    lblSvc.ControlTipText = Left$(tip & "Click to check again.", 1000)
    If down > 0 Then
        lblSvc.Caption = "Services: " & down & " down"
        lblSvc.ForeColor = RGB(253, 164, 175)
    Else
        lblSvc.Caption = "Services: all " & (UBound(labels) + 1) & " up"
        lblSvc.ForeColor = RGB(94, 234, 212)
        AppendLog "All " & (UBound(labels) + 1) & " web services are responding."
    End If
done:
    If Err.Number <> 0 Then
        lblSvc.Caption = "Services: check failed"
        AppendLog "[Service check] " & Err.Description
    End If
    mSvcBusy = False
End Sub

'----------------------------------------------------------------------------------------------
' Fence and NHD streams
'----------------------------------------------------------------------------------------------
Private Sub btnPlaceFence_Click()
    PlaceFence
End Sub

Public Sub PlaceFence()
    ' Places a block or shape fence with a red rubber-band outline, then loads the streams inside it.
    Dim tool As clsFenceTool, shapeMode As Boolean, t0 As Single
    If mBusy Then Exit Sub
    shapeMode = (cboFenceMode.ListIndex = 0)
    On Error Resume Next
    ActiveDesignFile.fence.Undefine
    On Error GoTo Failed
    Set tool = New clsFenceTool
    tool.shapeMode = shapeMode
    lblFence.Caption = "Fence: waiting for " & IIf(shapeMode, "shape", "block") & " fence in DGN..."
    AppendLog "Place a " & IIf(shapeMode, "shape fence in the DGN (Reset to close the shape)", "block fence in the DGN (two corner clicks)") & "; streams reload automatically."
    Me.Hide
    CommandState.StartPrimitive tool
    t0 = Timer
    Do While tool.State = 0
        DoEvents
        PauseMs 20
        If Timer - t0 > 300# Or Timer < t0 Then Exit Do
    Loop
    If tool.State = 0 Then CommandState.StartDefaultCommand
    ShowPrompt ""
    ShowCommand ""
    Me.Show vbModeless
    If tool.State = 1 Then
        AppendLog "Fence captured."
        LoadFenceStreams False
    Else
        lblFence.Caption = "Fence: placement cancelled - no active fence"
        If tool.State = 3 Then
            AppendLog "Fence placement failed: " & tool.ErrorText & "."
        Else
            AppendLog "Fence placement cancelled" & IIf(tool.State = 0, " or timed out.", ".")
        End If
    End If
    Exit Sub
Failed:
    Me.Show vbModeless
    AppendLog "[Fence Error] " & Err.Description
End Sub

Private Sub btnReloadFence_Click()
    LoadFenceStreams False
End Sub

Public Function LoadFenceStreams(ByVal thenQueue As Boolean) As Boolean
    ' Reads the active fence, draws its outline in red and clips the NHD streams to it.
    Dim fence As Collection, n As Long, errText As String
    If mBusy Then Exit Function
    On Error GoTo Failed
    Set fence = ReadFence()
    Set gFence = fence
    lblFence.Caption = "Fence: querying NHDPlus streams..."
    gCancel = False
    SetBusy True
    Set gStreams = QueryClipStreams(fence)
    SetBusy False
    mStreamsDrawn = False
    n = gStreams.Count
    lblFence.Caption = "Fence: " & n & " stream reach(es) clipped inside boundary"
    AppendLog "Clipped " & n & " NHDPlus stream(s) inside the active fence."
    LoadFenceStreams = True
    If thenQueue Then QueueStreams
    Exit Function
Failed:
    errText = Err.Description
    If mBusy Then
        SetBusy False
        lblFence.Caption = "Fence: NHD query failed"
    Else
        lblFence.Caption = "Fence: " & errText
    End If
    AppendLog "[Fence Error] " & errText
End Function

Private Sub EnsureStreamsDrawn()
    If gStreams Is Nothing Or mStreamsDrawn Then Exit Sub
    If gStreams.Count = 0 Then Exit Sub
    On Error Resume Next
    DrawStreams gStreams
    If Err.Number <> 0 Then AppendLog "[DGN Draw Error] Streams: " & Err.Description
    mStreamsDrawn = True
End Sub

Private Sub btnQueueStreams_Click()
    If mBusy Then Exit Sub
    If gStreams Is Nothing Then
        LoadFenceStreams True
    ElseIf gStreams.Count = 0 Then
        LoadFenceStreams True
    Else
        QueueStreams
    End If
End Sub

Private Sub QueueStreams()
    ' One culvert at the outlet of every clipped stream.
    Dim s As Variant, nm As String, safe As String, i As Long, ch As String, p As Variant, lastUnd As Boolean
    For Each s In gStreams
        nm = CStr(s("name"))
        safe = ""
        lastUnd = True
        For i = 1 To Len(nm)
            ch = Mid$(nm, i, 1)
            If ch Like "[A-Za-z0-9]" Then
                safe = safe & ch
                lastUnd = False
            ElseIf Not lastUnd Then
                safe = safe & "_"
                lastUnd = True
            End If
        Next i
        Do While Right$(safe, 1) = "_"
            safe = Left$(safe, Len(safe) - 1)
        Loop
        safe = Left$(safe, 12)
        If Len(safe) = 0 Then safe = "Stream"
        Set p = s("pour")
        AddQueueItem "CULV-" & Format$(CLng(s("id")), "00") & "_" & safe, "WGS84_LonLat", CDbl(p(1)), CDbl(p(2))
    Next s
    AppendLog "Queued " & gStreams.Count & " clipped stream outlet(s)."
End Sub

'----------------------------------------------------------------------------------------------
' Flow (Raindrop) Path tab - trace, draw, export
'----------------------------------------------------------------------------------------------
Private Sub btnRdTrace_Click()
    RunRaindropTrace
End Sub

Public Sub RunRaindropTrace()
    Dim lon As Double, lat As Double, dist As Double, tid As String, res As Object, sLon As String, sLat As String, sDist As String
    If mBusy Then Exit Sub
    sLon = Trim$(Replace(txtRdLon.text, ",", "."))
    sLat = Trim$(Replace(txtRdLat.text, ",", "."))
    sDist = Trim$(Replace(txtRdDist.text, ",", "."))
    If Len(sDist) = 0 Then sDist = "5"
    If Len(sLon) = 0 Or Len(sLat) = 0 Then
        Notify "Enter a longitude and latitude, or use Pick Point in DGN.", vbExclamation
        Exit Sub
    End If
    If Not (IsNumeric(sLon) And IsNumeric(sLat) And IsNumeric(sDist)) Then
        Notify "Longitude, latitude and distance must be numbers.", vbExclamation
        Exit Sub
    End If
    lon = Val(sLon)
    lat = Val(sLat)
    dist = Val(sDist)
    If lon < -180# Or lon > 180# Or lat < -90# Or lat > 90# Then
        Notify "Longitude/latitude are outside the valid WGS84 range.", vbExclamation
        Exit Sub
    End If
    If dist <= 0# Or dist > 500# Then
        Notify "Max downstream distance must be between 0 and 500 miles.", vbExclamation
        Exit Sub
    End If
    tid = Trim$(txtRdId.text)
    If Len(tid) = 0 Then tid = "RD-01"
    AppendLog "Tracing flow (raindrop) path " & tid & " from " & NumFix(lon, 6) & ", " & NumFix(lat, 6) & " (" & NumStr(dist) & " mi max)..."
    gCancel = False
    SetBusy True
    lblRdSummary.Caption = "Tracing..."
    If TraceRaindrop(lon, lat, dist, res) Then
        SetBusy False
        Set mRd = res
        mRdId = tid
        ShowRaindrop
        AppendLog "Raindrop trace " & tid & " completed: " & Format$(res("total_ft") / 5280#, "0.00") & " mi through " & res("reaches").Count & " reach(es)."
        If chkRdAuto.Value Then RaindropDraw
    Else
        SetBusy False
        lblRdSummary.Caption = "Trace failed."
        Notify CStr(res("error")), vbCritical
    End If
End Sub

Private Sub ShowRaindrop()
    Dim row As Variant, r As Long, comid As String
    lstFlow.Clear
    For Each row In RaindropRows(mRd)
        lstFlow.AddItem CStr(row(0))
        r = lstFlow.ListCount - 1
        comid = CStr(row(2))
        If Len(comid) = 0 Then comid = "-"
        lstFlow.List(r, 1) = CStr(row(1))
        lstFlow.List(r, 2) = comid
        lstFlow.List(r, 3) = Format$(row(4), "#,##0")
        lstFlow.List(r, 4) = Format$(row(5) / 5280#, "#,##0.000")
        lstFlow.List(r, 5) = CStr(row(6))
    Next row
    lblRdSummary.Caption = mRdId & ":  " & Format$(mRd("total_ft") / 5280#, "#,##0.00") & " mi  (" & Format$(mRd("total_ft"), "#,##0") & " ft)   -   overland " & _
        Format$(mRd("overland_ft"), "#,##0") & " ft   -   " & mRd("reaches").Count & " downstream reach(es)"
End Sub

Private Sub btnRdDraw_Click()
    RaindropDraw
End Sub

Public Sub RaindropDraw()
    Dim startDrawn As Boolean
    If mRd Is Nothing Then
        Notify "Run a trace first.", vbInformation
        Exit Sub
    End If
    On Error GoTo Failed
    startDrawn = mRdStartDrawn And Abs(mRdStartLon - CDbl(mRd("lon"))) < 0.00001 And Abs(mRdStartLat - CDbl(mRd("lat"))) < 0.00001
    DrawRaindrop mRd, mRdId, Not startDrawn
    Exit Sub
Failed:
    AppendLog "[Raindrop Draw Error] " & Err.Description
    Notify Err.Description, vbCritical
End Sub

Private Sub btnRdExport_Click()
    Dim fmt As String, ext As String, p As String, made As Collection, f As Variant, msg As String, doKml As Boolean, doShp As Boolean, kmz As Boolean
    If mRd Is Nothing Then
        Notify "Run a trace first.", vbInformation
        Exit Sub
    End If
    fmt = AskGisFormat("Raindrop path lines with start and end points.")
    If Len(fmt) = 0 Then Exit Sub
    doKml = (InStr(fmt, "KM") > 0)
    doShp = (InStr(fmt, "Shapefile") > 0)
    kmz = (InStr(fmt, "KMZ") > 0)
    ext = IIf(doKml, IIf(kmz, "kmz", "kml"), "shp")
    p = PickSaveFile("Export Raindrop Path - " & fmt, DgnFolder(), DgnStem() & "_" & SafeName(mRdId, "RD") & "." & ext, _
                     IIf(ext = "shp", "ESRI Shapefile", IIf(kmz, "Google Earth KMZ", "Google Earth KML")), ext)
    If Len(p) = 0 Then Exit Sub
    On Error GoTo Failed
    If InStrRev(p, ".") > InStrRev(p, "\") Then p = Left$(p, InStrRev(p, ".") - 1)
    Set made = ExportRaindropGis(p, mRdId, mRd, doKml, kmz, doShp)
    For Each f In made
        AppendLog "Exported GIS file: " & CStr(f)
        msg = msg & vbCrLf & CStr(f)
    Next f
    Notify "Exported files:" & vbCrLf & msg, vbInformation
    Exit Sub
Failed:
    AppendLog "[GIS Export Error] " & Err.Description
    Notify Err.Description, vbCritical
End Sub

Private Sub btnRdClear_Click()
    Set mRd = Nothing
    lstFlow.Clear
    lblRdSummary.Caption = "No trace yet."
End Sub

'----------------------------------------------------------------------------------------------
' GIS export
'----------------------------------------------------------------------------------------------
Private Function AskGisFormat(ByVal note As String) As String
    dlgGisFormat.lblNote.Caption = note
    dlgGisFormat.result = ""
    dlgGisFormat.Show vbModal
    AskGisFormat = dlgGisFormat.result
    Unload dlgGisFormat
End Function

Private Sub btnExportGis_Click()
    ' Drainage areas, pour points, clipped streams and longest flowpaths to KML / KMZ and / or Shapefiles.
    Dim fmt As String, ext As String, p As String, made As Collection, f As Variant, msg As String, n As Long, i As Long
    Dim doKml As Boolean, doShp As Boolean, kmz As Boolean, haveStreams As Boolean
    If mBusy Then Exit Sub
    For i = 1 To mQueue.Count
        If HasResult(i) Then n = n + 1
    Next i
    If Not gStreams Is Nothing Then haveStreams = (gStreams.Count > 0)
    If n = 0 And Not haveStreams Then
        Notify "Delineate a watershed or clip NHDPlus streams before exporting.", vbExclamation
        Exit Sub
    End If
    fmt = AskGisFormat("Watershed polygons, pour points, clipped streams and longest flowpaths.")
    If Len(fmt) = 0 Then Exit Sub
    doKml = (InStr(fmt, "KM") > 0)
    doShp = (InStr(fmt, "Shapefile") > 0)
    kmz = (InStr(fmt, "KMZ") > 0)
    If doKml And Not doShp Then ext = IIf(kmz, "kmz", "kml") Else ext = "shp"
    p = PickSaveFile("Export Watersheds & Streams - " & fmt, DgnFolder(), DgnStem() & "_USGS_Watersheds." & ext, _
                     IIf(ext = "shp", "ESRI Shapefile", IIf(kmz, "Google Earth KMZ", "Google Earth KML")), ext)
    If Len(p) = 0 Then Exit Sub
    On Error GoTo Failed
    If InStrRev(p, ".") > InStrRev(p, "\") Then p = Left$(p, InStrRev(p, ".") - 1)
    Set made = ExportGis(p, doKml, kmz, doShp, mQueue, gStreams)
    For Each f In made
        AppendLog "Exported GIS file: " & CStr(f)
        msg = msg & vbCrLf & CStr(f)
    Next f
    Notify "Exported files:" & vbCrLf & msg, vbInformation
    Exit Sub
Failed:
    AppendLog "[GIS Export Error] " & Err.Description
    Notify Err.Description, vbCritical
End Sub







