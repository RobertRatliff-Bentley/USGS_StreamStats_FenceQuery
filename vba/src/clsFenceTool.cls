VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsFenceTool"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'==============================================================================================
' Places a block or shape fence, drawing the rubber-band outline in red while it is being placed
' (MicroStation's own Place Fence tool cannot change its dynamics color).
' State: 0 = placing, 1 = fence placed, 2 = cancelled, 3 = failed (see ErrorText).
'==============================================================================================
Implements IPrimitiveCommandEvents

Public shapeMode As Boolean
Public State As Long
Public ErrorText As String

Private mX() As Double
Private mY() As Double
Private mN As Long
Private mView As View

Private Sub Class_Initialize()
    ReDim mX(0 To 63)
    ReDim mY(0 To 63)
End Sub

Private Sub Prompt()
    If mN = 0 Then
        ShowPrompt "Enter first point"
    ElseIf Not shapeMode Then
        ShowPrompt "Enter opposite corner"
    ElseIf mN < 3 Then
        ShowPrompt "Enter next vertex"
    Else
        ShowPrompt "Enter next vertex, or Reset (right-click) to close the fence"
    End If
End Sub

Private Function Outline(ByVal cx As Double, ByVal cy As Double, ByVal withCursor As Boolean, ByRef pts() As Point3d) As Long
    ' Closed outline through the placed points (and the cursor position). Returns the number of points.
    Dim n As Long, i As Long
    If Not shapeMode Then
        ReDim pts(0 To 4)
        pts(0) = Point3dFromXY(mX(0), mY(0))
        pts(1) = Point3dFromXY(cx, mY(0))
        pts(2) = Point3dFromXY(cx, cy)
        pts(3) = Point3dFromXY(mX(0), cy)
        pts(4) = pts(0)
        Outline = 5
        Exit Function
    End If
    n = mN
    If withCursor Then n = n + 1
    ReDim pts(0 To n)
    For i = 0 To mN - 1
        pts(i) = Point3dFromXY(mX(i), mY(i))
    Next i
    If withCursor Then pts(mN) = Point3dFromXY(cx, cy)
    If n >= 3 Then
        pts(n) = pts(0)
        n = n + 1
    End If
    If n - 1 < UBound(pts) Then ReDim Preserve pts(0 To n - 1)
    Outline = n
End Function

Private Sub Finish(ByVal ok As Boolean, ByVal cx As Double, ByVal cy As Double)
    Dim pts() As Point3d, n As Long
    If State <> 0 Then Exit Sub
    If ok Then
        On Error GoTo Bad
        n = Outline(cx, cy, False, pts)
        ReDim Preserve pts(0 To n - 2)          ' the fence closes itself
        ActiveDesignFile.fence.DefineFromModelPoints mView, pts
        State = 1
        ShowPrompt "Fence placed."
    Else
        State = 2
        ShowPrompt "Fence placement cancelled."
    End If
    CommandState.StartDefaultCommand
    Exit Sub
Bad:
    ErrorText = Err.Description
    State = 3
    CommandState.StartDefaultCommand
End Sub

Private Sub IPrimitiveCommandEvents_Start()
    ShowCommand "Place Fence"
    Prompt
End Sub

Private Sub IPrimitiveCommandEvents_DataPoint(Point As Point3d, ByVal View As View)
    If State <> 0 Then Exit Sub
    Set mView = View
    If mN = 0 Then
        mX(0) = Point.x
        mY(0) = Point.y
        mN = 1
        CommandState.StartDynamics
        Prompt
        Exit Sub
    End If
    If Not shapeMode Then
        If Point.x = mX(0) Or Point.y = mY(0) Then Exit Sub
        Finish True, Point.x, Point.y
        Exit Sub
    End If
    If Point.x <> mX(mN - 1) Or Point.y <> mY(mN - 1) Then
        If mN > UBound(mX) Then
            ReDim Preserve mX(0 To 2 * mN + 1)
            ReDim Preserve mY(0 To 2 * mN + 1)
        End If
        mX(mN) = Point.x
        mY(mN) = Point.y
        mN = mN + 1
    End If
    Prompt
End Sub

Private Sub IPrimitiveCommandEvents_Dynamics(Point As Point3d, ByVal View As View, ByVal DrawMode As MsdDrawingMode)
    Dim pts() As Point3d, n As Long, el As LineElement
    If State <> 0 Or mN = 0 Then Exit Sub
    On Error Resume Next
    n = Outline(Point.x, Point.y, True, pts)
    If n < 2 Then Exit Sub
    Set el = CreateLineElement1(Nothing, pts)
    el.Color = 3
    el.LineWeight = 2
    el.Redraw DrawMode
End Sub

Private Sub IPrimitiveCommandEvents_Reset()
    If shapeMode And mN >= 3 Then
        Finish True, 0#, 0#
    Else
        Finish False, 0#, 0#
    End If
End Sub

Private Sub IPrimitiveCommandEvents_Cleanup()
    ' Another command replaced this tool before the fence was finished: report it as cancelled.
    If State = 0 Then State = 2
End Sub

Private Sub IPrimitiveCommandEvents_Keyin(ByVal Keyin As String)
End Sub
