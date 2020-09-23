VERSION 5.00
Begin VB.Form Form3 
   BorderStyle     =   0  'None
   Caption         =   "Form3"
   ClientHeight    =   4620
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   7050
   LinkTopic       =   "Form3"
   ScaleHeight     =   4620
   ScaleWidth      =   7050
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   270
      Left            =   3315
      ScaleHeight     =   240
      ScaleWidth      =   2535
      TabIndex        =   3
      Top             =   315
      Width           =   2565
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Label1"
         Height          =   240
         Left            =   30
         TabIndex        =   4
         Top             =   15
         Width           =   2460
      End
   End
   Begin Project1.ScollBar ScollBar1 
      Height          =   4215
      Left            =   6720
      TabIndex        =   2
      Top             =   300
      Width           =   240
      _ExtentX        =   423
      _ExtentY        =   7435
   End
   Begin Project1.MacMenu MacMenu1 
      Height          =   780
      Left            =   150
      TabIndex        =   1
      Top             =   390
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   5080
      _ExtentY        =   1376
   End
   Begin Project1.FrmMAC FrmMAC1 
      Height          =   4620
      Left            =   0
      TabIndex        =   0
      Top             =   -15
      Width           =   7050
      _ExtentX        =   12435
      _ExtentY        =   8149
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
ScollBar1.Top = FrmMAC1.FormTopZero
ScollBar1.Height = FrmMAC1.FormHeight
ScollBar1.Left = FrmMAC1.FormWidth - ScollBar1.Width
Form1.Show
End Sub

Private Sub Form_Resize()
FrmMAC1.Resize
ScollBar1.Top = FrmMAC1.FormTopZero
ScollBar1.Height = FrmMAC1.FormHeight
ScollBar1.Left = FrmMAC1.FormWidth - ScollBar1.Width
ScollBar1.ResizeMe
End Sub

Private Sub FrmMAC1_LeftMenuClick()
MacMenu1.Visible = False
MacMenu1.Clear
MacMenu1.Add "Open Main Window"
MacMenu1.Add "Close Window"
MacMenu1.Add "<Spacer>"
MacMenu1.Add "Quit"
MacMenu1.Top = GetY - Me.Top
MacMenu1.Left = GetX - Me.Left
MacMenu1.ShowMenu
MacMenu1.Visible = True
End Sub

Private Sub FrmMAC1_MouseMove()
MacMenu1.Visible = False
End Sub

Private Sub FrmMAC1_RightMenuClick()
End
End Sub

Private Sub MacMenu1_OnMenuClick(Index As Integer, Caption As String)
Select Case UCase(Caption)
Case "QUIT"
End
End Select
End Sub

Private Sub ScollBar1_BarChange(Percent As Integer)
Picture1.Top = (FrmMAC1.FormHeight) * Percent / 100
Label1.Caption = Percent & "%  ------------------->"
End Sub
