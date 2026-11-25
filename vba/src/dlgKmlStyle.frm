VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} dlgKmlStyle 
   Caption         =   "KML / KMZ Symbology"
   ClientHeight    =   6040
   ClientLeft      =   110
   ClientTop       =   450
   ClientWidth     =   7780
   OleObjectBlob   =   "dlgKmlStyle.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "dlgKmlStyle"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================
' KML / KMZ symbology: colors, line widths and the style-file location used for every KML/KMZ
' export. The style file is shared with the Python tool.
'==============================================================================================
Private Const ICON_BASE As String = "http://maps.google.com/mapfiles/kml/"

Private Sub UserForm_Initialize()
    cboIcon.AddItem "Circle"
    cboIcon.AddItem "Square"
    cboIcon.AddItem "Triangle"
    cboIcon.AddItem "Target"
    cboIcon.AddItem "Pushpin"
    cboIcon.AddItem "Water"
    txtPath.text = KmlStylePath()
    LoadInto LoadKmlStyle()
End Sub

Private Function IconUrl(ByVal nm As String) As String
    Select Case nm
        Case "Circle"
            IconUrl = ICON_BASE & "shapes/placemark_circle.png"
        Case "Square"
            IconUrl = ICON_BASE & "shapes/placemark_square.png"
        Case "Triangle"
            IconUrl = ICON_BASE & "shapes/triangle.png"
        Case "Target"
            IconUrl = ICON_BASE & "shapes/target.png"
        Case "Pushpin"
            IconUrl = ICON_BASE & "pushpin/wht-pushpin.png"
        Case "Water"
            IconUrl = ICON_BASE & "shapes/water.png"
        Case Else
            IconUrl = nm
    End Select
End Function

Private Function IconName(ByVal url As String) As String
    Dim i As Long
    IconName = url
    For i = 0 To cboIcon.ListCount - 1
        If IconUrl(cboIcon.List(i)) = url Then IconName = cboIcon.List(i)
    Next i
End Function

Private Function NumText(ByVal v As Variant) As String
    NumText = NumStr(CDbl(v))
End Function

Private Sub LoadInto(ByVal st As Object)
    swBasinLine.BackColor = HexToRgb(st("basin_line_color"))
    txtBasinWidth.text = NumText(st("basin_line_width"))
    swBasinFill.BackColor = HexToRgb(st("basin_fill_color"))
    txtFillOpacity.text = NumText(st("basin_fill_opacity"))
    chkFill.Value = CBool(st("basin_fill"))
    swStream.BackColor = HexToRgb(st("stream_color"))
    txtStreamWidth.text = NumText(st("stream_width"))
    swFlowpath.BackColor = HexToRgb(st("flowpath_color"))
    txtFlowpathWidth.text = NumText(st("flowpath_width"))
    swRaindrop.BackColor = HexToRgb(st("raindrop_color"))
    txtRaindropWidth.text = NumText(st("raindrop_width"))
    swPoint.BackColor = HexToRgb(st("point_color"))
    txtPointScale.text = NumText(st("point_scale"))
    cboIcon.text = IconName(CStr(st("point_icon")))
    txtLabelScale.text = NumText(st("label_scale"))
End Sub

Private Function NumOr(ByVal t As String, ByVal fallback As Double, ByVal lo As Double, ByVal hi As Double) As Double
    t = Trim$(Replace(t, ",", "."))
    NumOr = fallback
    If IsNumeric(t) And Len(t) > 0 Then
        NumOr = Val(t)
        If NumOr < lo Then NumOr = lo
        If NumOr > hi Then NumOr = hi
    End If
End Function

Private Function Collect() As Object
    Dim st As Object, d As Object
    Set d = KmlStyleDefaults()
    Set st = KmlStyleDefaults()
    st("basin_line_color") = RgbToHex(swBasinLine.BackColor)
    st("basin_line_width") = NumOr(txtBasinWidth.text, d("basin_line_width"), 0.5, 10#)
    st("basin_fill") = CBool(chkFill.Value)
    st("basin_fill_color") = RgbToHex(swBasinFill.BackColor)
    st("basin_fill_opacity") = NumOr(txtFillOpacity.text, d("basin_fill_opacity"), 0#, 100#)
    st("stream_color") = RgbToHex(swStream.BackColor)
    st("stream_width") = NumOr(txtStreamWidth.text, d("stream_width"), 0.5, 10#)
    st("flowpath_color") = RgbToHex(swFlowpath.BackColor)
    st("flowpath_width") = NumOr(txtFlowpathWidth.text, d("flowpath_width"), 0.5, 10#)
    st("raindrop_color") = RgbToHex(swRaindrop.BackColor)
    st("raindrop_width") = NumOr(txtRaindropWidth.text, d("raindrop_width"), 0.5, 10#)
    st("point_color") = RgbToHex(swPoint.BackColor)
    st("point_scale") = NumOr(txtPointScale.text, d("point_scale"), 0.2, 4#)
    st("point_icon") = IconUrl(Trim$(cboIcon.text))
    st("label_scale") = NumOr(txtLabelScale.text, d("label_scale"), 0#, 4#)
    Set Collect = st
End Function

Private Sub PickSwatch(ByVal sw As Object)
    Dim c As Long
    If PickColor(sw.BackColor, c) Then sw.BackColor = c
End Sub

Private Sub swBasinLine_Click()
    PickSwatch swBasinLine
End Sub

Private Sub swBasinFill_Click()
    PickSwatch swBasinFill
End Sub

Private Sub swStream_Click()
    PickSwatch swStream
End Sub

Private Sub swFlowpath_Click()
    PickSwatch swFlowpath
End Sub

Private Sub swRaindrop_Click()
    PickSwatch swRaindrop
End Sub

Private Sub swPoint_Click()
    PickSwatch swPoint
End Sub

Private Sub btnBrowse_Click()
    Dim cur As String, p As String, folder As String, nm As String
    cur = txtPath.text
    folder = Left$(cur, InStrRev(cur, "\") - 1)
    nm = Mid$(cur, InStrRev(cur, "\") + 1)
    p = PickSaveFile("KML / KMZ Style File Location", folder, nm, "Style file", "json", False)
    If Len(p) = 0 Then Exit Sub
    txtPath.text = p
    If Len(Dir$(p)) > 0 Then
        If MsgBox("This style file already exists." & vbCrLf & vbCrLf & "Load its colors and widths into this dialog?" & vbCrLf & _
                  "(No keeps the current settings and overwrites it on Save.)", vbQuestion + vbYesNo, "KML Style") = vbYes Then
            LoadInto LoadKmlStyle(p)
        End If
    End If
End Sub

Private Sub btnDefault_Click()
    txtPath.text = KmlStyleDefaultPath()
End Sub

Private Sub btnReset_Click()
    LoadInto KmlStyleDefaults()
End Sub

Public Sub SaveStyle()
    SaveKmlStyle Collect(), txtPath.text
    SetKmlStylePath txtPath.text
    LogLine "Saved KML/KMZ symbology: " & txtPath.text
End Sub

Private Sub btnSave_Click()
    On Error GoTo Bad
    SaveStyle
    Unload Me
    Exit Sub
Bad:
    MsgBox Err.Description, vbCritical, "KML Style"
End Sub

Private Sub btnCancel_Click()
    Unload Me
End Sub


