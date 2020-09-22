VERSION 5.00
Begin VB.UserControl DVBPanel 
   Alignable       =   -1  'True
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   1860
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2355
   ControlContainer=   -1  'True
   ForwardFocus    =   -1  'True
   ScaleHeight     =   1860
   ScaleWidth      =   2355
   ToolboxBitmap   =   "DVBPanel.ctx":0000
   Begin VB.Timer tmrChangeFocus 
      Interval        =   10
      Left            =   120
      Top             =   600
   End
   Begin VB.PictureBox picScroll 
      BorderStyle     =   0  'None
      Height          =   495
      Left            =   0
      ScaleHeight     =   495
      ScaleWidth      =   1215
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   240
      Width           =   1215
   End
   Begin VB.VScrollBar VScroll 
      Height          =   1215
      Left            =   1440
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   120
      Width           =   255
   End
   Begin VB.HScrollBar HScroll 
      Height          =   255
      Left            =   0
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   1080
      Width           =   1215
   End
End
Attribute VB_Name = "DVBPanel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'API Declares
Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long

'API Constants
Private Const SM_CYVSCROLL = 20

Private Type POINTAPI
        X As Long
        Y As Long
End Type
Private Type RECT
        Left As Long
        Top As Long
        Right As Long
        Bottom As Long
End Type

'Default Property Values:
Const m_def_Enabled = True
Const m_def_BackColor = &H8000000F
'Property Variables:
Private m_Enabled As Boolean
Private m_BackColor As OLE_COLOR                'BackColor

Private strControlName As String
Private intControlIndex As Integer
Private ctlControl As Control

Private Sub tmrChangeFocus_Timer()
    'Determine if in design view
    'tmrChangeFocus.Enabled = False
    On Error Resume Next
    If UserControl.Ambient.UserMode = True Then
        Dim strThisControl As String
        Dim intThisIndex As Integer
        Dim lngCtrlTop, lngCalcDiff, lngCtrlValue, lngLastHwnd As Long
        Dim lngCtrlLeft, lngCtrlLeftValue As Long
        On Error Resume Next
        'Get Current Control Name and Index
        strThisControl = Screen.ActiveControl.Name
        intThisIndex = Screen.ActiveControl.Index
        'Determine if control name or index has changed
        If strThisControl <> strControlName Or (intThisIndex <> intControlIndex And Not IsNull(intThisIndex)) Then
            strControlName = strThisControl
            'Determine if Index changed
            If intThisIndex <> intControlIndex And Not IsNull(intThisIndex) Then
                intControlIndex = intThisIndex
            End If
            lngCtrlValue = 0
            Set ctlControl = Screen.ActiveControl
            On Error Resume Next
            Dim objControlContainer As Object
            Set objControlContainer = ctlControl.Container
            Do Until objControlContainer Is Nothing
                lngLastHwnd = objControlContainer.hwnd
                If lngLastHwnd = Me.hwnd Then Exit Do
                lngCtrlValue = lngCtrlValue + objControlContainer.Top
                lngCtrlLeftValue = lngCtrlLeftValue + objControlContainer.Left
                Err.Clear
                Set objControlContainer = objControlContainer.Container
                If Err.Number <> 0 Then Exit Do
            Loop
            'If the Active Control is in the DVBPanel Control
            If lngLastHwnd = Me.hwnd Then
                Dim lngTempValue As Long
                'Determine if control is out of vertical viewing range
                lngCtrlTop = lngCtrlValue + ctlControl.Top - 50
                lngTempValue = ctlControl.Height
                If lngTempValue > VScroll.Height Then lngTempValue = VScroll.Height - 175
                lngCtrlValue = lngCtrlValue + ctlControl.Top + lngTempValue + 75
                'If the top of the Active Control is outside of the viewing
                '   area then change the Vertical Scroll bar to place it in view
                If lngCtrlTop < 0 Then
                    VScroll.Value = VScroll.Value + lngCtrlTop
                ElseIf lngCtrlValue > VScroll.Height Then
                    lngCalcDiff = lngCtrlValue - (VScroll.Height)
                    VScroll.Value = VScroll.Value + lngCalcDiff
                End If
                'Determine if control is out of horizontal viewing range
                lngCtrlLeft = lngCtrlLeftValue + ctlControl.Left - 50
                lngTempValue = ctlControl.Width
                If lngTempValue > HScroll.Width Then lngTempValue = HScroll.Width - 175
                lngCtrlLeftValue = lngCtrlLeftValue + ctlControl.Left + lngTempValue + 75
                'If the left of the Active Control is outside of the viewing
                '   area then change the Horizontal Scroll bar to place it in view
                If lngCtrlLeft < 0 Then
                    HScroll.Value = HScroll.Value + lngCtrlLeft
                ElseIf lngCtrlLeftValue > HScroll.Width Then
                    lngCalcDiff = lngCtrlLeftValue - (HScroll.Width)
                    HScroll.Value = HScroll.Value + lngCalcDiff
                End If
            Else
                Exit Sub
            End If
        End If
    Else
        tmrChangeFocus.Enabled = False
    End If
    DoEvents
End Sub
'MemberInfo=0,0,0,0
Public Property Get Enabled() As Boolean
    Enabled = m_Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    m_Enabled = New_Enabled
    PropertyChanged "Enabled"
End Property
'MemberInfo=7,0,0,0
Public Property Get BackColor() As OLE_COLOR
    BackColor = m_BackColor
End Property
Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    m_BackColor = New_BackColor
    UserControl.BackColor = New_BackColor
    UserControl_Paint
    PropertyChanged "BackColor"
End Property
Public Property Get hwnd() As Long
    hwnd = UserControl.hwnd
End Property

Private Sub UserControl_Initialize()
    tmrChangeFocus.Enabled = True
End Sub
Private Sub UserControl_InitProperties()
    m_Enabled = m_def_Enabled
End Sub
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    m_Enabled = PropBag.ReadProperty("Enabled", m_def_Enabled)
    m_BackColor = PropBag.ReadProperty("BackColor", m_def_BackColor)
    Me.BackColor = m_BackColor
    picScroll.BackColor = m_BackColor
End Sub
Private Sub UserControl_Resize()
    Dim lngBarXSize As Long 'the width of a scrollbar
    Dim lngBarYSize As Long 'the height of a scrollbar
    Dim blnHScroll As Boolean 'whether a horizontal ScrollBar is needed
    Dim blnVScroll As Boolean 'whether a vetical ScrollBar is needed
    'Determine the size of the ScrollBars.
    lngBarXSize = GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelX
    lngBarYSize = GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelY
    On Error Resume Next
    
    If Ambient.UserMode Then 'We only show the scrollbars at run-time
        'Position the controls
        With UserControl
            VScroll.Move .ScaleWidth - lngBarXSize, 0, lngBarXSize, .ScaleHeight + (lngBarYSize * -1)
            HScroll.Move 0, .ScaleHeight - lngBarYSize, .ScaleWidth + (lngBarXSize * -1), lngBarYSize
            'picScroll is used to fill the area at the bottom right of the control so
            'that constituent controls do not show through.
            picScroll.Move .ScaleWidth - lngBarXSize, .ScaleHeight - lngBarYSize, lngBarXSize, lngBarYSize
        End With
        'Set the Visible state of the controls
        HScroll.ZOrder
        VScroll.ZOrder
        picScroll.ZOrder
        'Determine whether we need to enable the scrollbars
        blnHScroll = (ClientWidth > (UserControl.ScaleWidth - VScroll.Width))
        blnVScroll = (ClientHeight > (UserControl.ScaleHeight - HScroll.Height))
        HScroll.Enabled = blnHScroll
        VScroll.Enabled = blnVScroll
        picScroll.Visible = True
        'reposition the ContainedControls
        RecalculateScrollBarValues
    Else
        'Position the controls
        With UserControl
            VScroll.Move .ScaleWidth - lngBarXSize, 0, lngBarXSize, .ScaleHeight + (lngBarYSize * -1)
            HScroll.Move 0, .ScaleHeight - lngBarYSize, .ScaleWidth + (lngBarXSize * -1), lngBarYSize
            'picScroll is used to fill the area at the bottom right of the control so
            'that constituent controls do not show through.
            picScroll.Move .ScaleWidth - lngBarXSize, .ScaleHeight - lngBarYSize, lngBarXSize, lngBarYSize
        End With
        'Set the Visible state of the controls
        picScroll.ZOrder
        HScroll.ZOrder
        VScroll.ZOrder
        HScroll.Enabled = False
        VScroll.Enabled = False
        picScroll.Visible = True
    End If
End Sub

Private Sub UserControl_Paint()
    If VScroll.Width <> (GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelX) Then
        'UserControl_Resize
    End If
End Sub

Private Sub UserControl_Show()
    'Call the resize event to set the visible state of constituent controls
    UserControl_Resize
    'Hook for use with the MouseWheel
    If Ambient.UserMode Then Hook Me.hwnd
End Sub

Private Sub UserControl_Terminate()
    UnHook
    tmrChangeFocus.Enabled = False
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    Call PropBag.WriteProperty("Enabled", m_Enabled, m_def_Enabled)
    Call PropBag.WriteProperty("BackColor", m_BackColor, m_def_BackColor)
End Sub

Private Sub hscroll_Change()
    PositionContainedControls
End Sub
Private Sub hscroll_Scroll()
    PositionContainedControls
End Sub
Private Sub vscroll_Change()
    PositionContainedControls
End Sub
Private Sub vscroll_Scroll()
    PositionContainedControls
End Sub

Private Function InBox() As Boolean
    Dim mpos As POINTAPI
    Dim oRect As RECT
    GetCursorPos mpos
    GetWindowRect Me.hwnd, oRect
    If mpos.X >= oRect.Left And mpos.X <= oRect.Right And _
        mpos.Y >= oRect.Top And mpos.Y <= oRect.Bottom Then
        InBox = True
    Else
        InBox = False
   End If
End Function
Private Function ClientWidth() As Long
    'This procedure calculates the width of a virtual container that is wide
    'enough to accommodate all contained controls.
    Dim ndx As Long
    Dim lngWidth As Long
    Dim ctl As Object
    For Each ctl In UserControl.Parent
        If Not TypeOf ctl Is Menu Then
            If ctl.Container.hwnd = UserControl.hwnd Then
                If lngWidth < (ctl.Left + ctl.Width + (8 * Screen.TwipsPerPixelX)) Then
                    lngWidth = (ctl.Left + ctl.Width + (8 * Screen.TwipsPerPixelX))
                End If
            End If
        End If
    Next
    'Return the calculated width
    If lngWidth <> 0 Then
        ClientWidth = lngWidth
    Else
        'there are no controls so we use the width of the usercontrol, less the
        'width of the scrollbar
        ClientWidth = UserControl.ScaleWidth - (GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelX)
    End If
End Function

Private Function ClientHeight() As Long
    'This procedure calculates the height of a virtual container that is tall
    'enough to accommodate all contained controls.
    Dim ndx As Long
    Dim lngHeight As Long
    Dim ctl As Object
    For Each ctl In UserControl.Parent
        If Not TypeOf ctl Is Menu Then
            If ctl.Container.hwnd = UserControl.hwnd Then
                If lngHeight < (ctl.Top + ctl.Height + (8 * Screen.TwipsPerPixelY)) Then
                    'Debug.Print ctl.Name
                    lngHeight = (ctl.Top + ctl.Height + (8 * Screen.TwipsPerPixelY))
                    DoEvents
                End If
            End If
        End If
    Next
    'return the calculated height
    If lngHeight <> 0 Then
        ClientHeight = lngHeight
    Else
        'there are no controls so we use the height of the usercontrol, less the
        'height of the scrollbar
        ClientHeight = UserControl.ScaleHeight - (GetSystemMetrics(SM_CYVSCROLL) * Screen.TwipsPerPixelY)
    End If
End Function

Private Sub PositionContainedControls()
    Static sHSValue As Long 'previous value of the horizontal scrollbar
    Static sVSValue As Long 'previous value of the vertical scrollbar
    Dim ndx As Long
    Dim ctl As Control
    'Re-position the controls
    For Each ctl In UserControl.Parent
        If Not TypeOf ctl Is Menu Then
            If ctl.Container.hwnd = UserControl.hwnd Then
                With ctl
                    .Move (.Left + sHSValue) - HScroll.Value, (.Top + sVSValue) - VScroll.Value
                End With
            End If
        End If
    Next
    'store the current ScrollBar positions so we can use them to calculate the
    'original position of controls the next time this procedure is called.
    sHSValue = HScroll.Value
    sVSValue = VScroll.Value
    On Error Resume Next
    ctlControl.SetFocus
End Sub

Private Sub RecalculateScrollBarValues()
    'This procedure recalculates the values of the ScrollBars
    Dim lngNewMax As Long
    If Ambient.UserMode Then
        'Calculate the new maximum values
        If HScroll.Enabled Then
            With HScroll
                lngNewMax = ClientWidth - (UserControl.ScaleWidth - VScroll.Width)
                If lngNewMax > 32767 Then lngNewMax = 32767
                .Max = lngNewMax
                If .Value > .Max Then
                    .Value = .Max
                End If
                If .Max > UserControl.ScaleWidth - VScroll.Width Then
                    .LargeChange = (UserControl.ScaleWidth - VScroll.Width)
                Else
                    .LargeChange = .Max
                End If
                .SmallChange = .LargeChange / 10
                .Enabled = True
            End With
        Else
            With HScroll
                .Max = 0
                .Value = 0
            End With
        End If
        If VScroll.Enabled Then
            With VScroll
                Dim test As Variant
                lngNewMax = ClientHeight - (UserControl.ScaleHeight - HScroll.Height)
                If lngNewMax > 32767 Then lngNewMax = 32767
                .Max = lngNewMax
                If .Max > UserControl.ScaleHeight - .Height Then
                    .LargeChange = (UserControl.ScaleHeight - HScroll.Height)
                Else
                    .LargeChange = .Max
                End If
                .LargeChange = (UserControl.ScaleWidth - .Width)
                .SmallChange = .LargeChange / 10
                .Enabled = True
                If .Value > .Max Then
                    .Value = .Max
                End If
            End With
        Else
            With VScroll
                .Max = 0
                .Value = 0
            End With
        End If
        'Re-position the ContainedControls
        strControlName = ""
        intControlIndex = 0
        PositionContainedControls
    End If
End Sub
'Function used by MouseWheel Module
Public Function ScrollUp() As Boolean
    If InBox = True Then
        If VScroll.Value >= VScroll.SmallChange Then
            VScroll.Value = VScroll.Value - VScroll.SmallChange
        ElseIf VScroll.Value = VScroll.Min Then
            If HScroll.Value >= HScroll.SmallChange Then
                HScroll.Value = HScroll.Value - HScroll.SmallChange
            Else
                HScroll.Value = HScroll.Min
            End If
        Else
            VScroll.Value = VScroll.Min
        End If
        ScrollUp = True
    Else
        ScrollUp = False
    End If
End Function
'Function used by MouseWheel Module
Public Function ScrollDown() As Boolean
    If InBox = True Then
        If VScroll.Value <= VScroll.Max - VScroll.SmallChange Then
            VScroll.Value = VScroll.Value + VScroll.SmallChange
        ElseIf VScroll.Value = VScroll.Max Then
            If HScroll.Value <= HScroll.Max - HScroll.SmallChange Then
                HScroll.Value = HScroll.Value + HScroll.SmallChange
            Else
                HScroll.Value = HScroll.Max
            End If
        Else
            VScroll.Value = VScroll.Max
        End If
        ScrollDown = True
    Else
        ScrollDown = False
    End If
End Function

