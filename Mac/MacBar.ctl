VERSION 5.00
Begin VB.UserControl MacBar 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5520
   ScaleHeight     =   3600
   ScaleWidth      =   5520
   Begin VB.ListBox List1 
      Height          =   1230
      ItemData        =   "MacBar.ctx":0000
      Left            =   120
      List            =   "MacBar.ctx":0002
      TabIndex        =   0
      Top             =   1755
      Visible         =   0   'False
      Width           =   1140
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "File"
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   330
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   -45
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image4 
      Height          =   300
      Index           =   0
      Left            =   30
      Picture         =   "MacBar.ctx":0004
      Stretch         =   -1  'True
      Top             =   0
      Visible         =   0   'False
      Width           =   585
   End
   Begin VB.Image Image2 
      Height          =   300
      Left            =   1515
      Picture         =   "MacBar.ctx":04E6
      Stretch         =   -1  'True
      Top             =   540
      Width           =   90
   End
   Begin VB.Image Image3 
      Height          =   300
      Left            =   1710
      Picture         =   "MacBar.ctx":09C8
      Top             =   1140
      Width           =   105
   End
   Begin VB.Image Image1 
      Height          =   300
      Left            =   600
      Picture         =   "MacBar.ctx":0EAA
      Top             =   915
      Width           =   555
   End
End
Attribute VB_Name = "MacBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public Y
Public Current
Public Event OnMenuClick(MenuIndex As Integer, Caption As String, Left As Long, Top As Long)
Public Event OnMenuOver(MenuIndex As Integer, Caption As String, Left As Long, Top As Long)

Function LoadMe(frm As Form)
UserControl.Width = frm.Width
UserControl.Height = Image1.Height
Image1.Top = 0
Image2.Top = 0
Image3.Top = 0
Image1.Left = 0
Image2.Left = Image1.Width
Image2.Width = frm.Width - Image1.Width

End Function

Private Sub Label1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
RaiseEvent OnMenuClick(Index, Label1(Index).Caption, Label1(Index).Left, Label1(Index).Top)
End Sub

Private Sub Label1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
If Current <> Index Then
Current = Index
RaiseEvent OnMenuOver(Index, Label1(Index).Caption, Label1(Index).Left, Label1(Index).Top)
End If

For Y = 0 To List1.ListCount
If Label1(Y).ForeColor = vbWhite And Index <> Y Then
Label1(Y).ForeColor = vbBlack
Image4(Y).Picture = Image2.Picture
End If
Next

If Label1(Index).ForeColor <> vbWhite Then
Label1(Index).ForeColor = vbWhite
Image4(Index).Picture = Image3.Picture
End If
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
LoadMe Parent
End Sub

Function loadbuttons()
Dim X
For X = 1 To Image4.UBound
Unload Image4(Image4.UBound)
Unload Label1(Label1.UBound)
DoEvents
Next

For X = 0 To List1.ListCount - 1
Load Image4(Image4.UBound + 1)
        With Image4(Image4.UBound)
        .Left = Image4(Image4.UBound - 1).Left + Image4(Image4.UBound - 1).Width + 10
        .ZOrder 0
        .Top = 0
        .Width = Len(List1.List(X)) * 120
        .Picture = Image2.Picture
        .Visible = True
        End With
Load Label1(Label1.UBound + 1)
        With Label1(Label1.UBound)
        .Left = Label1(Label1.UBound - 1).Left + Label1(Label1.UBound - 1).Width + 10
        .ZOrder 0
        .Top = -30
        .Height = 400
        .Width = Len(List1.List(X)) * 120
        .Caption = List1.List(X)
        .Tag = X
        .Visible = True
        End With
Next
End Function

Public Function Clear()
List1.Clear
End Function

Public Function Add(txt As String)
List1.AddItem txt
End Function
