VERSION 5.00
Begin VB.UserControl ScollBar 
   ClientHeight    =   7365
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   3075
   ForwardFocus    =   -1  'True
   ScaleHeight     =   491
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   205
   Begin VB.Timer Timer3 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1545
      Top             =   285
   End
   Begin VB.Timer Timer2 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1035
      Top             =   285
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   510
      Top             =   285
   End
   Begin VB.PictureBox Arrow 
      AutoSize        =   -1  'True
      BorderStyle     =   0  'None
      Height          =   240
      Index           =   1
      Left            =   0
      Picture         =   "ScollBar.ctx":0000
      ScaleHeight     =   16
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   16
      TabIndex        =   2
      Top             =   6075
      Width           =   240
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00C0FFFF&
      BorderStyle     =   0  'None
      Height          =   6390
      Left            =   0
      ScaleHeight     =   426
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   18
      TabIndex        =   0
      Top             =   0
      Width           =   270
      Begin VB.PictureBox Arrow 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "ScollBar.ctx":0542
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   1
         Top             =   5850
         Width           =   240
      End
      Begin VB.PictureBox Thumb 
         AutoSize        =   -1  'True
         BorderStyle     =   0  'None
         Height          =   240
         Index           =   0
         Left            =   0
         Picture         =   "ScollBar.ctx":0A84
         ScaleHeight     =   16
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   16
         TabIndex        =   3
         Top             =   3060
         Width           =   240
      End
      Begin VB.Image Image3 
         Height          =   60
         Left            =   0
         Picture         =   "ScollBar.ctx":0FC6
         Top             =   6315
         Width           =   240
      End
      Begin VB.Image Image2 
         Height          =   45
         Left            =   0
         Picture         =   "ScollBar.ctx":1448
         Top             =   15
         Width           =   240
      End
      Begin VB.Image Image1 
         Height          =   5790
         Left            =   0
         Picture         =   "ScollBar.ctx":18BA
         Stretch         =   -1  'True
         Top             =   60
         Width           =   240
      End
   End
   Begin VB.Image Image7 
      Height          =   240
      Left            =   900
      Picture         =   "ScollBar.ctx":1D1C
      Top             =   5190
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image6 
      Height          =   240
      Left            =   855
      Picture         =   "ScollBar.ctx":225E
      Top             =   4605
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image5 
      Height          =   240
      Left            =   645
      Picture         =   "ScollBar.ctx":27A0
      Top             =   5190
      Visible         =   0   'False
      Width           =   240
   End
   Begin VB.Image Image4 
      Height          =   240
      Left            =   585
      Picture         =   "ScollBar.ctx":2CE2
      Top             =   4605
      Visible         =   0   'False
      Width           =   240
   End
End
Attribute VB_Name = "ScollBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Dim frm As Form
Dim Large As Long
Dim pMAX, Pmin, Value
Dim UAdown As Boolean
Dim Xx, yy
Dim RSHeight
Dim thumbsize
Public mY
Public Event BarChange(Percent As Integer)

Private Sub Arrow_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Timer2.Enabled = True
Arrow(Index).Picture = Image7.Picture
Case 1
Timer3.Enabled = True
Arrow(Index).Picture = Image6.Picture
End Select
End Sub

Private Sub Arrow_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Select Case Index
Case 0
Timer2.Enabled = False
Arrow(Index).Picture = Image5.Picture
Case 1
Timer3.Enabled = False
Arrow(Index).Picture = Image4.Picture
End Select
End Sub

Private Sub Text2_Click()
Text2.Text = InitScrollBar
End Sub

Private Sub Thumb_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
mY = Y - 15
Thumb(Index).BorderStyle = 1
Timer1.Enabled = True
End Sub

Private Sub Thumb_MouseUp(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Thumb(Index).BorderStyle = 0
Timer1.Enabled = False
End Sub

Private Sub Timer1_Timer()
'Shape1.Width = GetX
If Thumb(0).Top > Image1.Top Or Thumb(0).Top < GetY - (Parent.Top / 15) - 60 - mY Then
If GetY - (Parent.Top / 15) - 60 - mY < (Arrow(0).Top - Arrow(0).Height) Then
Thumb(0).Top = GetY - (Parent.Top / 15) - 60 - mY
End If
Else
Thumb(0).Top = Image1.Top
End If
RaiseEvent BarChange(GetPercent)
'GetY
End Sub

Private Sub Timer2_Timer()
If Thumb(0).Top > Image1.Top + 1 Then
Thumb(0).Top = Thumb(0).Top - 1
RaiseEvent BarChange(GetPercent)
End If
End Sub

Private Sub Timer3_Timer()
If Thumb(0).Top < Arrow(0).Top - Thumb(0).Height Then
Thumb(0).Top = Thumb(0).Top + 1
RaiseEvent BarChange(GetPercent)
End If
End Sub

Private Sub UserControl_Resize()
ResizeMe
End Sub

Function InitScrollBar() As Integer
InitScrollBar = ((Picture1.Height - Arrow(0).Top - Thumb(0).Height) / 15) - 35
End Function

Function GetPercent() As String
Value = Int(((Thumb(0).Top - Image1.Top) / InitScrollBar) * 100)
If Value > 100 Then
Value = 100
End If
If Value < 0 Then
Value = 0
End If
GetPercent = Value
End Function

Function ResizeMe()
UserControl.Width = Image1.Width * 15
Picture1.Top = 0
Picture1.Left = 0
Picture1.Height = UserControl.Height
Image1.Top = 0
Image2.Top = 0
Thumb(0).Top = 1
Image1.Height = Picture1.Height / 15
Image3.Top = UserControl.Height
Image3.ZOrder 0
Arrow(1).Top = Image1.Height - Arrow(1).Height
Arrow(1).ZOrder 0
Arrow(0).Top = Arrow(1).Top - Arrow(0).Height
Arrow(0).ZOrder 0
End Function
