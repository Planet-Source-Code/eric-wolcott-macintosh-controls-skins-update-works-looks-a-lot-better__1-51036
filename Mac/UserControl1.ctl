VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.UserControl FrmMAC 
   ClientHeight    =   5025
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6105
   ScaleHeight     =   5025
   ScaleWidth      =   6105
   ToolboxBitmap   =   "UserControl1.ctx":0000
   Begin VB.PictureBox Picture4 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   600
      ScaleHeight     =   345
      ScaleWidth      =   285
      TabIndex        =   5
      Top             =   3960
      Width           =   285
      Begin VB.Image Image23 
         Height          =   345
         Left            =   210
         Picture         =   "UserControl1.ctx":0312
         Top             =   0
         Width           =   75
      End
      Begin VB.Image Image24 
         Height          =   105
         Left            =   0
         Picture         =   "UserControl1.ctx":04C4
         Top             =   240
         Width           =   285
      End
      Begin VB.Image Image25 
         Height          =   165
         Left            =   45
         Picture         =   "UserControl1.ctx":06AA
         Top             =   75
         Width           =   165
      End
      Begin VB.Image Image22 
         Height          =   75
         Left            =   0
         Picture         =   "UserControl1.ctx":0878
         Top             =   0
         Width           =   285
      End
      Begin VB.Image Image3 
         Height          =   345
         Left            =   0
         Picture         =   "UserControl1.ctx":09E6
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   345
      Left            =   3180
      ScaleHeight     =   345
      ScaleWidth      =   300
      TabIndex        =   4
      Top             =   3855
      Width           =   300
      Begin VB.Image Image21 
         Height          =   165
         Left            =   75
         Picture         =   "UserControl1.ctx":0B3C
         Top             =   75
         Width           =   165
      End
      Begin VB.Image Image20 
         Height          =   105
         Left            =   0
         Picture         =   "UserControl1.ctx":0D0A
         Top             =   240
         Width           =   300
      End
      Begin VB.Image Image19 
         Height          =   345
         Left            =   240
         Picture         =   "UserControl1.ctx":0EF0
         Top             =   0
         Width           =   60
      End
      Begin VB.Image Image18 
         Height          =   345
         Left            =   0
         Picture         =   "UserControl1.ctx":1046
         Top             =   0
         Width           =   75
      End
      Begin VB.Image Image17 
         Height          =   75
         Left            =   0
         Picture         =   "UserControl1.ctx":11F8
         Top             =   0
         Width           =   300
      End
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Command1"
      Height          =   195
      Left            =   15
      TabIndex        =   3
      Top             =   30
      Width           =   180
   End
   Begin VB.PictureBox Picture2 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   3765
      ScaleHeight     =   405
      ScaleWidth      =   1920
      TabIndex        =   1
      Top             =   3150
      Width           =   1920
      Begin VB.Image Image15 
         Height          =   345
         Left            =   225
         Picture         =   "UserControl1.ctx":1366
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
      Begin VB.Image Image14 
         Height          =   345
         Left            =   0
         Picture         =   "UserControl1.ctx":1460
         Top             =   0
         Width           =   60
      End
      Begin VB.Image Image13 
         Height          =   345
         Left            =   1785
         Picture         =   "UserControl1.ctx":15B6
         Top             =   0
         Width           =   45
      End
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00E0E0E0&
      BorderStyle     =   0  'None
      Height          =   405
      Left            =   195
      ScaleHeight     =   405
      ScaleWidth      =   1920
      TabIndex        =   0
      Top             =   2985
      Width           =   1920
      Begin VB.Image Image12 
         Height          =   345
         Left            =   1770
         Picture         =   "UserControl1.ctx":170C
         Top             =   0
         Width           =   45
      End
      Begin VB.Image Image11 
         Height          =   345
         Left            =   0
         Picture         =   "UserControl1.ctx":1862
         Top             =   0
         Width           =   60
      End
      Begin VB.Image Image10 
         Height          =   345
         Left            =   225
         Picture         =   "UserControl1.ctx":19B8
         Stretch         =   -1  'True
         Top             =   0
         Width           =   1290
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   2295
      Top             =   2355
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   7
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1AB2
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":1E4C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":21E6
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":2580
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":35D2
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":4624
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "UserControl1.ctx":5676
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Image Image16 
      Height          =   345
      Left            =   150
      Picture         =   "UserControl1.ctx":66C8
      Top             =   3585
      Width           =   30
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Label1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   1485
      TabIndex        =   2
      Top             =   3735
      Width           =   1815
   End
   Begin VB.Image Image6 
      Height          =   105
      Left            =   3780
      Picture         =   "UserControl1.ctx":67C2
      Top             =   2700
      Width           =   105
   End
   Begin VB.Image Image7 
      Height          =   105
      Left            =   960
      Picture         =   "UserControl1.ctx":68AC
      Stretch         =   -1  'True
      Top             =   2640
      Width           =   45
   End
   Begin VB.Image Image5 
      Height          =   105
      Left            =   255
      Picture         =   "UserControl1.ctx":6942
      Top             =   2640
      Width           =   135
   End
   Begin VB.Image Image2 
      Height          =   1530
      Left            =   3840
      Picture         =   "UserControl1.ctx":6A48
      Stretch         =   -1  'True
      Top             =   1035
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   1530
      Left            =   240
      Picture         =   "UserControl1.ctx":741A
      Stretch         =   -1  'True
      Top             =   825
      Width           =   90
   End
   Begin VB.Image Image9 
      Height          =   345
      Left            =   615
      Picture         =   "UserControl1.ctx":7C54
      Stretch         =   -1  'True
      Top             =   330
      Width           =   2670
   End
   Begin VB.Image Image8 
      Height          =   1530
      Left            =   555
      Picture         =   "UserControl1.ctx":7D4E
      Stretch         =   -1  'True
      Top             =   690
      Width           =   2940
   End
End
Attribute VB_Name = "FrmMAC"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Type POINTAPI2
        X As Long
        Y As Long
End Type
Private Cur As POINTAPI2
Public Event LeftMenuClick()
Public Event RightMenuClick()
Public Event MouseMove()
Public FormTopZero
Public FormLeftZero
Public FormWidth
Public FormHeight
Public X As Form
Public Function LoadMe(frm As Form)
Set X = frm
UserControl.Width = frm.Width
UserControl.Height = frm.Height
Image8.Left = 0
Image8.Top = 0
Image8.Width = UserControl.Width
Image8.Height = UserControl.Height
Picture3.Top = 0
Picture3.Left = 0
Picture4.Top = 0
Picture4.Left = Image8.Width - Picture4.Width
Image9.Top = 0
Image9.Left = 0
Image9.Width = UserControl.Width
Image5.Left = 0
Image5.Top = Image8.Height - Image5.Height
Image1.Left = 0
Image1.Top = Picture3.Top + Picture3.Height
Image1.Height = Image8.Height - Image1.Top - Image5.Height
Image6.Top = Image8.Height - Image6.Height
Image6.Left = Image8.Width - Image6.Width
Image2.Left = Image8.Width - Image2.Width
Image2.Top = Image1.Top
Image2.Height = Image8.Height - Image2.Top - Image6.Height
Image7.Left = Image1.Left + Image1.Width
Image7.Top = Image8.Height - Image7.Height
Image7.Width = Image8.Width

Picture1.Width = Image8.Width / 4
Picture1.Top = 0
Picture1.Height = Image10.Height
Picture1.Left = Picture3.Left + Picture3.Width + 10
Image10.Left = 0
Image10.Width = Picture1.Width
Image11.Left = 0
Image12.Left = Picture1.Width - Image12.Width

Picture2.Width = Image8.Width / 4
Picture2.Top = 0
Picture2.Height = Image15.Height
Picture2.Left = Picture4.Left - Picture2.Width - 60
Image15.Left = 0
Image15.Width = Picture2.Width
Image14.Left = 0
Image13.Left = Picture2.Width - Image13.Width

Label1.Width = Picture2.Left - (Picture1.Left + Picture1.Width)
Label1.Top = 0
Label1.Left = Picture1.Left + Picture1.Width
Label1.Caption = frm.Caption

Image16.Left = Picture3.Left + Picture3.Width
Image16.Top = 0
Image16.Height = Picture1.Height
Image16.Width = Picture4.Left - (Picture3.Left + Picture3.Width)
Image16.ZOrder 0

FormTopZero = Image9.Height - 50
FormLeftZero = Image1.Left + Image1.Width - 100
FormWidth = UserControl.Width - Image1.Width - Image2.Width + 100
FormHeight = UserControl.Height - Image9.Height - Image7.Height + 50
End Function

Private Sub Image1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbSizeWE
RaiseEvent MouseMove
End Sub
Private Sub Image3_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image5_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image6_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Screen.MousePointer = vbSizeNWSE
    RaiseEvent MouseMove
End Sub
Private Sub Image7_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbSizeNS
RaiseEvent MouseMove
End Sub
Private Sub Image8_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub
Private Sub Image9_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image10_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image11_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image12_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image13_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image14_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image15_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image16_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub
Private Sub Image17_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub
Private Sub Image18_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub
Private Sub Image21_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
End Sub
Private Sub Image23_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub
Private Sub Image24_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub

Private Sub Image10_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image11_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image12_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image13_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image14_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image15_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub
Private Sub Image16_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
MoveForm
End Sub

Public Function MoveForm()
ReleaseCapture
SendMessage X.hwnd, &HA1, 2, 0&
End Function

Private Sub Image2_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cur.X = GetX
End Sub

Private Sub Image2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResizeFormX False, True
End Sub

Private Sub Image21_Click()
RaiseEvent LeftMenuClick

End Sub


Private Sub Image25_Click()
RaiseEvent RightMenuClick
End Sub

Private Sub Image6_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Cur.X = GetX
    Cur.Y = GetY
End Sub

Public Sub Image6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResizeFormX False, False
ResizeFormY False, True
End Sub

Private Sub Image7_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Cur.Y = GetY
End Sub

Private Sub Image7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
ResizeFormY False, True
End Sub

Private Sub Picture4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Screen.MousePointer = vbDefault
RaiseEvent MouseMove
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent MouseMove
End Sub

Private Sub UserControl_Resize()
LoadMe Parent
End Sub

Public Function Resize()
LoadMe Parent
End Function
Public Function ResizeFormX(Visible1 As Boolean, Visible2 As Boolean)
    X.Visible = Visible1
    Cur.X = GetX - Cur.X
    X.Width = X.Width + Cur.X
    Screen.MousePointer = vbDefault
    X.Visible = Visible2
End Function
Public Function ResizeFormY(Visible1 As Boolean, Visible2 As Boolean)
    X.Visible = Visible1
    Cur.Y = GetY - Cur.Y
    X.Height = X.Height + Cur.Y
    Screen.MousePointer = vbDefault
    X.Visible = Visible2
End Function

Public Function CreateMenuButton(Text As String)
Load frmMenu
frmMenu.CreateMenuButton Text
End Function

