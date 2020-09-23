VERSION 5.00
Begin VB.UserControl Animation 
   AutoRedraw      =   -1  'True
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ScaleHeight     =   3600
   ScaleWidth      =   4800
   Begin VB.PictureBox PicHol 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1455
      Left            =   600
      ScaleHeight     =   97
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   97
      TabIndex        =   0
      Top             =   0
      Visible         =   0   'False
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   0
      Top             =   2880
   End
End
Attribute VB_Name = "Animation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Private Declare Function StretchBlt Lib "gdi32" (ByVal hDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Const SRCCOPY = &HCC0020 ' (DWORD) dest = source

'Event Declarations:
Event Change()
Event ChangeToFrame(Index As Integer)
Event Resize() 'MappingInfo=UserControl,UserControl,-1,Resize
Event Click() 'MappingInfo=UserControl,UserControl,-1,Click
Event DblClick() 'MappingInfo=UserControl,UserControl,-1,DblClick
Event KeyDown(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyDown
Event KeyPress(KeyAscii As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyPress
Event KeyUp(KeyCode As Integer, Shift As Integer) 'MappingInfo=UserControl,UserControl,-1,KeyUp
Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseDown
Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseMove
Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single) 'MappingInfo=UserControl,UserControl,-1,MouseUp
'Default Property Values:
Const m_def_MousePointer = 0
Const m_def_AutoSizeDetection = True
Const m_def_FrameWidth = 0
Const m_def_FrameHeight = 0
Const m_def_RunOnDesigneTime = True
Const m_def_AutoSize = True
Const m_def_FitPicture = False
Const m_def_Transparent = True
'Property Variables:
Dim m_MousePointer As MousePoin
Dim m_AutoSizeDetection As Boolean
Dim m_FrameWidth As Single
Dim m_FrameHeight As Single
Dim m_RunOnDesigneTime As Boolean
Dim m_AutoSize As Boolean
Dim m_FitPicture As Boolean
Dim m_Transparent As Boolean

Public Enum MousePoin
                    [Default]
                    [Arrow]
                    [Cross]
                    [I beam]
                    [Icon]
                    [Resize]
                    [Size NE SW]
                    [Size N S]
                    [Size NW SE]
                    [Size W E]
                    [Up arrow]
                    [Hourglass]
                    [No drop]
                    [Arrow and Hourglass]
                    [Arrow and Question Mark]
                    [Size All]
                    [Custom] = 99
End Enum

Dim AutoSIZ As Boolean
Dim SizerW, SizerH
Dim SizerWi, SizerHei
Dim Fram As Integer, TimeFram As Integer
Dim ShowNeFrame As Boolean
Dim ShowPreFrame As Boolean

Sub ShowAboutBox()
    frmAbout.Show vbModal
    Unload frmAbout
    Set frmAbout = Nothing
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,BackColor
Public Property Get BackColor() As OLE_COLOR
    BackColor = UserControl.BackColor
End Property

Public Property Let BackColor(ByVal New_BackColor As OLE_COLOR)
    UserControl.BackColor() = New_BackColor
    UserControl.MaskColor = New_BackColor
    PropertyChanged "BackColor"
    PicHol.BackColor = New_BackColor
    DrawFrame TimeFram
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Enabled
Public Property Get Enabled() As Boolean
    Enabled = UserControl.Enabled
End Property

Public Property Let Enabled(ByVal New_Enabled As Boolean)
    UserControl.Enabled() = New_Enabled
    PropertyChanged "Enabled"
    'Timer1.Enabled = Enabled
    If Enabled = True Then
        DrawFrame TimeFram
    Else
        DrawFrame 1
    End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Refresh
Public Sub Refresh()
    DrawFrame TimeFram
    UserControl.Refresh
End Sub

Private Sub Timer1_Timer()
'If ShowPreFrame = True Then TimeFram = TimeFram - 1
If TimeFram < 0 Then TimeFram = Fram
If TimeFram > Fram Then TimeFram = 0
DrawFrame TimeFram
TimeFram = TimeFram + 1
If ShowNeFrame = True Then: ShowNeFrame = False: Timer1.Enabled = False
If ShowPreFrame = True Then ShowPreFrame = False: Timer1.Enabled = False
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_DblClick()
    RaiseEvent DblClick
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyPress(KeyAscii As Integer)
    RaiseEvent KeyPress(KeyAscii)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
    RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, X, Y)
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, X, Y)
End Sub

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Picture
Public Property Get Picture() As Picture
    Set Picture = PicHol.Picture
End Property

Public Property Set Picture(ByVal New_Picture As Picture)
    Set PicHol.Picture = New_Picture
    PropertyChanged "Picture"
    LookForSizes
    DrawFrame TimeFram
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,hWnd
Public Property Get hWnd() As Long
    hWnd = UserControl.hWnd
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get Transparent() As Boolean
    Transparent = m_Transparent
End Property

Public Property Let Transparent(ByVal New_Transparent As Boolean)
    m_Transparent = New_Transparent
    PropertyChanged "Transparent"
    DrawFrame TimeFram
If m_Transparent = True Then
    UserControl.MaskColor = UserControl.BackColor
    Set UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0   'Set control's backstyle to Transparent
Else
    UserControl.BackStyle = 1   'Set control's backstyle to Transparent'
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,False
Public Property Get FitPicture() As Boolean
    FitPicture = m_FitPicture
End Property

Public Property Let FitPicture(ByVal New_FitPicture As Boolean)
    m_FitPicture = New_FitPicture
    PropertyChanged "FitPicture"
    DrawFrame TimeFram
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSize() As Boolean
    AutoSize = m_AutoSize
End Property

Public Property Let AutoSize(ByVal New_AutoSize As Boolean)
    m_AutoSize = New_AutoSize
    PropertyChanged "AutoSize"
    If m_AutoSize = True Then
        'UserControl.Size PicHol.Width, PicHol.Height
        LookForSizes
    End If
        DrawFrame TimeFram
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=Timer1,Timer1,-1,Interval
Public Property Get Interval() As Long
    Interval = Timer1.Interval
End Property

Public Property Let Interval(ByVal New_Interval As Long)
    Timer1.Interval() = New_Interval
    PropertyChanged "Interval"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,Image
Public Property Get Image() As Picture
    Set Image = UserControl.Image
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get RunOnDesigneTime() As Boolean
    RunOnDesigneTime = m_RunOnDesigneTime
End Property

Public Property Let RunOnDesigneTime(ByVal New_RunOnDesigneTime As Boolean)
If Ambient.UserMode = False Then
    m_RunOnDesigneTime = New_RunOnDesigneTime
    PropertyChanged "RunOnDesigneTime"
    Timer1.Enabled = m_RunOnDesigneTime
End If
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get FrameWidth() As Single
    FrameWidth = m_FrameWidth
End Property

Public Property Let FrameWidth(ByVal New_FrameWidth As Single)
    m_FrameWidth = New_FrameWidth
    PropertyChanged "FrameWidth"
    LookForSizes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=12,0,0,0
Public Property Get FrameHeight() As Single
    FrameHeight = m_FrameHeight
End Property

Public Property Let FrameHeight(ByVal New_FrameHeight As Single)
    m_FrameHeight = New_FrameHeight
    PropertyChanged "FrameHeight"
    LookForSizes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=0,0,0,True
Public Property Get AutoSizeDetection() As Boolean
    AutoSizeDetection = m_AutoSizeDetection
End Property

Public Property Let AutoSizeDetection(ByVal New_AutoSizeDetection As Boolean)
    m_AutoSizeDetection = New_AutoSizeDetection
    PropertyChanged "AutoSizeDetection"
    LookForSizes
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MappingInfo=UserControl,UserControl,-1,MouseIcon
Public Property Get MouseIcon() As Picture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Set MouseIcon(ByVal New_MouseIcon As Picture)
    Set UserControl.MouseIcon = New_MouseIcon
    PropertyChanged "MouseIcon"
End Property

'WARNING! DO NOT REMOVE OR MODIFY THE FOLLOWING COMMENTED LINES!
'MemberInfo=7,0,0,0
Public Property Get MousePointer() As MousePoin
    MousePointer = m_MousePointer
End Property

Public Property Let MousePointer(ByVal New_MousePointer As MousePoin)
    m_MousePointer = New_MousePointer
    PropertyChanged "MousePointer"
End Property

'Initialize Properties for User Control
Private Sub UserControl_InitProperties()
    m_Transparent = m_def_Transparent
    m_FitPicture = m_def_FitPicture
    m_AutoSize = m_def_AutoSize
    m_RunOnDesigneTime = m_def_RunOnDesigneTime
    m_FrameWidth = m_def_FrameWidth
    m_FrameHeight = m_def_FrameHeight
    m_AutoSizeDetection = m_def_AutoSizeDetection
    m_MousePointer = m_def_MousePointer
End Sub

'Load property values from storage
Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
TimeFram = 0

    UserControl.BackColor = PropBag.ReadProperty("BackColor", &H8000000F)
    UserControl.Enabled = PropBag.ReadProperty("Enabled", True)
    Set PicHol.Picture = PropBag.ReadProperty("Picture", Nothing)
    m_Transparent = PropBag.ReadProperty("Transparent", m_def_Transparent)
    m_FitPicture = PropBag.ReadProperty("FitPicture", m_def_FitPicture)
    m_AutoSize = PropBag.ReadProperty("AutoSize", m_def_AutoSize)
    PicHol.BackColor = UserControl.BackColor
    Timer1.Interval = PropBag.ReadProperty("Interval", 1)
    m_RunOnDesigneTime = PropBag.ReadProperty("RunOnDesigneTime", m_def_RunOnDesigneTime)
    m_FrameWidth = PropBag.ReadProperty("FrameWidth", m_def_FrameWidth)
    m_FrameHeight = PropBag.ReadProperty("FrameHeight", m_def_FrameHeight)
    m_AutoSizeDetection = PropBag.ReadProperty("AutoSizeDetection", m_def_AutoSizeDetection)
    Set MouseIcon = PropBag.ReadProperty("MouseIcon", Nothing)
    m_MousePointer = PropBag.ReadProperty("MousePointer", m_def_MousePointer)
    LookForSizes
If Ambient.UserMode = False Then
    Timer1.Enabled = m_RunOnDesigneTime
End If
'MakeAutoSize
DrawFrame TimeFram
If m_Transparent = True Then
    UserControl.MaskColor = UserControl.BackColor
    Set UserControl.MaskPicture = UserControl.Image
    UserControl.BackStyle = 0   'Set control's backstyle to Transparent
Else
    UserControl.BackStyle = 1   'Set control's backstyle to Transparent'
End If
End Sub

Private Sub UserControl_Resize()
    RaiseEvent Resize
If AutoSIZ = True Then Exit Sub
'Set UserControl.Picture = Nothing
'Set UserControl.MaskPicture = Nothing
'    UserControl.BackStyle = 1   'Set control's backstyle to Transparent'
DrawFrame TimeFram
'If m_Transparent = True Then
'    UserControl.BackStyle = 0   'Set control's backstyle to Transparent
'Else
'    UserControl.BackStyle = 1   'Set control's backstyle to Transparent'
'End If
End Sub

Private Sub UserControl_Show()
LookForSizes
DrawFrame TimeFram
End Sub

Private Sub UserControl_Terminate()
Timer1.Enabled = False
End Sub

'Write property values to storage
Private Sub UserControl_WriteProperties(PropBag As PropertyBag)

    Call PropBag.WriteProperty("BackColor", UserControl.BackColor, &H8000000F)
    Call PropBag.WriteProperty("Enabled", UserControl.Enabled, True)
    Call PropBag.WriteProperty("Picture", PicHol.Picture, Nothing)
    Call PropBag.WriteProperty("Transparent", m_Transparent, m_def_Transparent)
    Call PropBag.WriteProperty("FitPicture", m_FitPicture, m_def_FitPicture)
    Call PropBag.WriteProperty("AutoSize", m_AutoSize, m_def_AutoSize)
    Call PropBag.WriteProperty("Interval", Timer1.Interval, 1)
    Call PropBag.WriteProperty("RunOnDesigneTime", m_RunOnDesigneTime, m_def_RunOnDesigneTime)
    Call PropBag.WriteProperty("FrameWidth", m_FrameWidth, m_def_FrameWidth)
    Call PropBag.WriteProperty("FrameHeight", m_FrameHeight, m_def_FrameHeight)
    Call PropBag.WriteProperty("AutoSizeDetection", m_AutoSizeDetection, m_def_AutoSizeDetection)
    Call PropBag.WriteProperty("MouseIcon", MouseIcon, Nothing)
    Call PropBag.WriteProperty("MousePointer", m_MousePointer, m_def_MousePointer)
End Sub

Private Sub MakeAutoSize()
If PicHol.Picture > 0 Then
    If m_AutoSize = True Then
            AutoSIZ = True
            UserControl.Size SizerW, SizerH
            AutoSIZ = False
    End If
End If
End Sub

Private Sub LookForSizes()
On Error Resume Next
If PicHol.Picture > 0 And m_AutoSizeDetection = True Then
    If PicHol.Width <= PicHol.Height Then
        SizerHei = PicHol.Width / Screen.TwipsPerPixelY
        SizerWi = 0
        SizerW = PicHol.Width
        SizerH = SizerW
        Fram = Fix(PicHol.Height / SizerW) - 1
    Else
        SizerWi = PicHol.Height / Screen.TwipsPerPixelX
        SizerHei = 0
        SizerH = PicHol.Height
        SizerW = SizerH
        Fram = Fix(PicHol.Width / SizerH) - 1
    End If
Else
    If PicHol.Picture > 0 Then
        SizerW = m_FrameWidth * Screen.TwipsPerPixelX
        SizerH = m_FrameHeight * Screen.TwipsPerPixelY
        If PicHol.Width <= PicHol.Height Then
            SizerHei = m_FrameHeight
            SizerWi = 0
            Fram = Fix(PicHol.ScaleHeight / m_FrameHeight) - 1
'    Debug.Print UserControl.Extender.Name, Fram, " Fram = Fix(PicHol.Height / SizerW) - 1"
        Else
            SizerHei = 0
            SizerWi = m_FrameWidth
            Fram = Fix(PicHol.ScaleWidth / m_FrameWidth) - 1
'    Debug.Print UserControl.Extender.Name, Fram, " Fram = Fix(PicHol.Width / SizerH) - 1"
        End If
    End If
End If
MakeAutoSize
End Sub

Public Sub DrawFrame(Index As Integer)
'SizerW = 20
Set UserControl.Picture = Nothing
If PicHol.Picture > 0 Then
    If m_FitPicture = True Then
        StretchBlt UserControl.hDC, 0, 0, UserControl.ScaleWidth / Screen.TwipsPerPixelX, UserControl.ScaleHeight / Screen.TwipsPerPixelY, _
                    PicHol.hDC, SizerWi * (Index), SizerHei * (Index), SizerW / Screen.TwipsPerPixelX, SizerH / Screen.TwipsPerPixelY, SRCCOPY

    Else
        StretchBlt UserControl.hDC, 0, 0, SizerW / Screen.TwipsPerPixelX, SizerH / Screen.TwipsPerPixelY, _
                    PicHol.hDC, SizerWi * (Index), SizerHei * (Index), SizerW / Screen.TwipsPerPixelX, SizerH / Screen.TwipsPerPixelY, SRCCOPY
    End If
    Set UserControl.MaskPicture = UserControl.Image

    RaiseEvent Change
    RaiseEvent ChangeToFrame(Index)
End If
'Else
    If Ambient.UserMode = False Then
        Dim Dashh As Integer
        Dashh = UserControl.DrawStyle
        UserControl.DrawStyle = 2
        UserControl.Line (0, 0)-(UserControl.ScaleWidth - 15, 0), vbBlack
        UserControl.Line (UserControl.ScaleWidth - 15, 0)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 15), vbBlack
        UserControl.Line (0, UserControl.ScaleHeight - 15)-(UserControl.ScaleWidth - 15, UserControl.ScaleHeight - 15), vbBlack
        UserControl.Line (0, 0)-(0, UserControl.ScaleHeight - 15), vbBlack
        UserControl.DrawStyle = Dashh
    Set UserControl.MaskPicture = UserControl.Image
    End If
End Sub

Public Sub AutoPlay()
    Timer1.Enabled = True
    TimeFram = 0
    DrawFrame TimeFram
End Sub

Public Sub Pause()
    Timer1.Enabled = False
End Sub

Public Sub Continue()
    Timer1.Enabled = True
End Sub

Public Sub StopAutoPlay()
    Timer1.Enabled = False
    TimeFram = 0
    DrawFrame TimeFram
End Sub

Public Sub ShowFrame(Index As Integer)
    Timer1.Enabled = False
    TimeFram = Index
    DrawFrame TimeFram
End Sub

Public Function FameCount()
    FameCount = Fram + 1
End Function

Public Sub ShowNextFrame()
ShowNeFrame = True
Timer1.Enabled = True
'    TimeFram = TimeFram + 1
'    If TimeFram > Fram Then TimeFram = 0
'    DrawFrame TimeFram
End Sub

Public Sub ShowPreviewFrame()
TimeFram = TimeFram - 2
ShowPreFrame = True
Timer1.Enabled = True
'    TimeFram = TimeFram - 1
'    If TimeFram < 0 Then TimeFram = Fram
'    DrawFrame TimeFram
End Sub



