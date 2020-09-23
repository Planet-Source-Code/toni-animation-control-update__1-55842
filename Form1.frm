VERSION 5.00
Object = "*\AAnimator.vbp"
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   7470
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   7470
   ScaleWidth      =   9480
   StartUpPosition =   3  'Windows Default
   Begin Animator.Animation Animation9 
      Height          =   2010
      Left            =   2160
      TabIndex        =   24
      Top             =   5280
      Width           =   2010
      _ExtentX        =   3545
      _ExtentY        =   3545
      BackColor       =   16777215
      Picture         =   "Form1.frx":0000
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation8 
      Height          =   495
      Left            =   0
      TabIndex        =   23
      Top             =   3360
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   16777215
      Picture         =   "Form1.frx":EDEC2
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
      FrameWidth      =   120
      FrameHeight     =   50
   End
   Begin Animator.Animation Animation7 
      Height          =   495
      Left            =   0
      TabIndex        =   22
      Top             =   2880
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   16777215
      Picture         =   "Form1.frx":F2714
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation6 
      Height          =   495
      Left            =   0
      TabIndex        =   21
      Top             =   2400
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   16777215
      Picture         =   "Form1.frx":1E05D6
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation5 
      Height          =   495
      Left            =   0
      TabIndex        =   20
      Top             =   1920
      Width           =   495
      _ExtentX        =   873
      _ExtentY        =   873
      BackColor       =   16777215
      Picture         =   "Form1.frx":1F9BF0
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation4 
      Height          =   465
      Left            =   0
      TabIndex        =   19
      Top             =   1440
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BackColor       =   16777215
      Picture         =   "Form1.frx":228A42
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation3 
      Height          =   465
      Left            =   0
      TabIndex        =   18
      Top             =   960
      Width           =   465
      _ExtentX        =   820
      _ExtentY        =   820
      BackColor       =   16777215
      Picture         =   "Form1.frx":34C98C
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation2 
      Height          =   480
      Left            =   0
      TabIndex        =   17
      Top             =   480
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   847
      BackColor       =   16777215
      Picture         =   "Form1.frx":384C76
      Interval        =   100
      RunOnDesigneTime=   0   'False
   End
   Begin Animator.Animation Animation1 
      Height          =   495
      Left            =   0
      TabIndex        =   16
      Top             =   0
      Width           =   480
      _ExtentX        =   847
      _ExtentY        =   873
      BackColor       =   16777215
      Picture         =   "Form1.frx":396CC8
      FitPicture      =   -1  'True
      AutoSize        =   0   'False
      Interval        =   100
      RunOnDesigneTime=   0   'False
      FrameWidth      =   120
      FrameHeight     =   50
      AutoSizeDetection=   0   'False
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Start"
      Height          =   495
      Left            =   0
      TabIndex        =   7
      Top             =   4080
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Pause"
      Height          =   495
      Left            =   2280
      TabIndex        =   6
      Top             =   4080
      Width           =   1815
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Continue"
      Height          =   495
      Left            =   4200
      TabIndex        =   5
      Top             =   4080
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Caption         =   "StopPlay"
      Height          =   495
      Left            =   6360
      TabIndex        =   4
      Top             =   4080
      Width           =   1935
   End
   Begin VB.CommandButton Command5 
      Caption         =   "ShowFrame"
      Height          =   495
      Left            =   0
      TabIndex        =   3
      Top             =   4680
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      AutoSize        =   -1  'True
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      ScaleHeight     =   735
      ScaleWidth      =   1815
      TabIndex        =   2
      Top             =   5280
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Show Next Frame"
      Height          =   495
      Left            =   2280
      TabIndex        =   1
      Top             =   4680
      Width           =   1815
   End
   Begin VB.CommandButton Command7 
      Caption         =   "Show Preview Frame"
      Height          =   495
      Left            =   4200
      TabIndex        =   0
      Top             =   4680
      Width           =   2055
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   0
      Left            =   480
      TabIndex        =   15
      Top             =   120
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   1
      Left            =   480
      TabIndex        =   14
      Top             =   600
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   2
      Left            =   480
      TabIndex        =   13
      Top             =   1080
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   3
      Left            =   480
      TabIndex        =   12
      Top             =   1560
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   4
      Left            =   480
      TabIndex        =   11
      Top             =   2040
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   5
      Left            =   480
      TabIndex        =   10
      Top             =   2520
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   6
      Left            =   480
      TabIndex        =   9
      Top             =   3000
      Width           =   90
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "0"
      Height          =   195
      Index           =   7
      Left            =   480
      TabIndex        =   8
      Top             =   3480
      Width           =   90
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   9480
      Y1              =   3855
      Y2              =   3855
   End
   Begin VB.Line Line2 
      X1              =   0
      X2              =   9480
      Y1              =   3360
      Y2              =   3360
   End
   Begin VB.Line Line3 
      X1              =   0
      X2              =   9480
      Y1              =   2880
      Y2              =   2880
   End
   Begin VB.Line Line4 
      X1              =   0
      X2              =   9480
      Y1              =   2400
      Y2              =   2400
   End
   Begin VB.Line Line5 
      X1              =   0
      X2              =   9480
      Y1              =   1920
      Y2              =   1920
   End
   Begin VB.Line Line6 
      X1              =   0
      X2              =   9480
      Y1              =   1440
      Y2              =   1440
   End
   Begin VB.Line Line7 
      X1              =   0
      X2              =   9480
      Y1              =   960
      Y2              =   960
   End
   Begin VB.Line Line8 
      X1              =   0
      X2              =   9480
      Y1              =   480
      Y2              =   480
   End
   Begin VB.Shape Shape1 
      BorderColor     =   &H00004000&
      FillColor       =   &H00C0C0FF&
      FillStyle       =   0  'Solid
      Height          =   3975
      Left            =   -480
      Top             =   0
      Width           =   9975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim an1, an2, an3, an4, an5, an6, an7, an8
Dim Anim, Bonus
Private Sub Animation1_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an1 = an1 + 60 '  Animation1.Width / 4
If an1 > Me.ScaleWidth Then an1 = -Animation1.Width: Bathmologia Label1(0)
Animation1.Left = an1
Label1(0).Left = Animation1.Left - Label1(0).Width
End Sub
Private Sub Animation2_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an2 = an2 + 60 '  Animation1.Width / 4
If an2 > Me.ScaleWidth Then an2 = -Animation2.Width: Bathmologia Label1(1)
Animation2.Left = an2
Label1(1).Left = Animation2.Left - Label1(1).Width
End Sub
Private Sub Animation3_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an3 = an3 + 60 '  Animation1.Width / 4
If an3 > Me.ScaleWidth Then an3 = -Animation3.Width: Bathmologia Label1(2)
Animation3.Left = an3
Label1(2).Left = Animation3.Left - Label1(2).Width
End Sub
Private Sub Animation4_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an4 = an4 + 60 '  Animation1.Width / 4
If an4 > Me.ScaleWidth Then an4 = -Animation4.Width: Bathmologia Label1(3)
Animation4.Left = an4
Label1(3).Left = Animation4.Left - Label1(3).Width
End Sub
Private Sub Animation5_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an5 = an5 + 60 '  Animation1.Width / 4
If an5 > Me.ScaleWidth Then an5 = -Animation5.Width: Bathmologia Label1(4)
Animation5.Left = an5
Label1(4).Left = Animation5.Left - Label1(4).Width
End Sub
Private Sub Animation6_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an6 = an6 + 60 '  Animation1.Width / 4
If an6 > Me.ScaleWidth Then an6 = -Animation6.Width: Bathmologia Label1(5)
Animation6.Left = an6
Label1(5).Left = Animation6.Left - Label1(5).Width
End Sub
Private Sub Animation7_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an7 = an7 + 60 '  Animation1.Width / 4
If an7 > Me.ScaleWidth Then an7 = -Animation7.Width: Bathmologia Label1(6)
Animation7.Left = an7
Label1(6).Left = Animation7.Left - Label1(6).Width
End Sub
Private Sub Animation8_ChangeToFrame(Index As Integer)
'Debug.Print Index
'Set Picture1.Picture = Animation1.Image
an8 = an8 + 60 '  Animation1.Width / 4
If an8 > Me.ScaleWidth Then an8 = -Animation8.Width: Bathmologia Label1(7)
Animation8.Left = an8
Label1(7).Left = Animation8.Left - Label1(7).Width
End Sub

Private Sub Animation1_Click()
Debug.Print Animation1.FameCount
End Sub

Private Sub Animation9_Click()
Animation9.AutoPlay
End Sub

Private Sub Command1_Click()
Animation1.AutoPlay
Animation2.AutoPlay
Animation3.AutoPlay
Animation4.AutoPlay
Animation5.AutoPlay
Animation6.AutoPlay
Animation7.AutoPlay
Animation8.AutoPlay

End Sub

Private Sub Command2_Click()
Animation1.Pause
Animation2.Pause
Animation3.Pause
Animation4.Pause
Animation5.Pause
Animation6.Pause
Animation7.Pause
Animation8.Pause
End Sub

Private Sub Command3_Click()
Animation1.Continue
Animation2.Continue
Animation3.Continue
Animation4.Continue
Animation5.Continue
Animation6.Continue
Animation7.Continue
Animation8.Continue
End Sub

Private Sub Command4_Click()
Animation1.StopAutoPlay
Animation2.StopAutoPlay
Animation3.StopAutoPlay
Animation4.StopAutoPlay
Animation5.StopAutoPlay
Animation6.StopAutoPlay
Animation7.StopAutoPlay
Animation8.StopAutoPlay
Animation1.Left = 0
Animation2.Left = 0
Animation3.Left = 0
Animation4.Left = 0
Animation5.Left = 0
Animation6.Left = 0
Animation7.Left = 0
Animation8.Left = 0
For x = 0 To Label1.Count - 1
Label1(x).Left = Animation1.Left + Animation1.Width
Label1(x).Caption = 0
Next
an1 = 0
an2 = 0
an3 = 0
an4 = 0
an5 = 0
an6 = 0
an7 = 0
an8 = 0
Anim = 0
End Sub

Private Sub Command5_Click()
Animation1.ShowFrame Animation1.FameCount - 1
Set Picture1.Picture = Animation1.Image
End Sub

Private Sub Bathmologia(Obj As Object)
Anim = Anim + 1
If Anim >= 9 Then Anim = 1
Obj.Caption = Anim
End Sub

Private Sub Command6_Click()
On Error Resume Next
Dim Control As Object
For Each Control In Form1.Controls
    'If TypeOf Control Is Animation Then
        Control.ShowNextFrame
    'End If
Next
End Sub

Private Sub Command7_Click()
On Error Resume Next
Dim Control As Object
For Each Control In Form1.Controls
    'If TypeOf Control Is Animation Then
        Control.ShowPreviewFrame
    'End If
Next
End Sub

