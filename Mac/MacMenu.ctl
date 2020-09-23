VERSION 5.00
Begin VB.UserControl MacMenu 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFC0C0&
   ClientHeight    =   3585
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2895
   ScaleHeight     =   3585
   ScaleWidth      =   2895
   Begin VB.ListBox List1 
      Height          =   1425
      ItemData        =   "MacMenu.ctx":0000
      Left            =   1005
      List            =   "MacMenu.ctx":0002
      TabIndex        =   0
      Top             =   105
      Visible         =   0   'False
      Width           =   960
   End
   Begin VB.Label Label3 
      BackColor       =   &H00DCA674&
      BackStyle       =   0  'Transparent
      Caption         =   "Macintosh Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Index           =   0
      Left            =   30
      TabIndex        =   3
      Top             =   30
      Visible         =   0   'False
      Width           =   2820
   End
   Begin VB.Label Label2 
      BackColor       =   &H00DCA674&
      BackStyle       =   0  'Transparent
      Caption         =   "________________________"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   2
      Top             =   1635
      Width           =   2805
   End
   Begin VB.Label Label1 
      BackColor       =   &H00DCA674&
      BackStyle       =   0  'Transparent
      Caption         =   "Macintosh Menu"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   30
      TabIndex        =   1
      Top             =   15
      Width           =   2805
   End
   Begin VB.Image Image16 
      Height          =   270
      Left            =   2805
      Picture         =   "MacMenu.ctx":0004
      Top             =   1680
      Width           =   90
   End
   Begin VB.Image Image15 
      Height          =   270
      Left            =   15
      Picture         =   "MacMenu.ctx":04D6
      Top             =   1680
      Width           =   45
   End
   Begin VB.Image Image14 
      Height          =   270
      Left            =   60
      Picture         =   "MacMenu.ctx":0960
      Stretch         =   -1  'True
      Top             =   1680
      Width           =   2745
   End
   Begin VB.Image Image10 
      Height          =   225
      Index           =   0
      Left            =   2745
      Picture         =   "MacMenu.ctx":0DEA
      Top             =   30
      Visible         =   0   'False
      Width           =   135
   End
   Begin VB.Image Image9 
      Height          =   225
      Index           =   0
      Left            =   0
      Picture         =   "MacMenu.ctx":12E0
      Top             =   30
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image3 
      Height          =   225
      Index           =   0
      Left            =   60
      Picture         =   "MacMenu.ctx":175E
      Stretch         =   -1  'True
      Top             =   30
      Visible         =   0   'False
      Width           =   2685
   End
   Begin VB.Image Image8 
      Height          =   90
      Index           =   0
      Left            =   2805
      Picture         =   "MacMenu.ctx":1C18
      Top             =   255
      Visible         =   0   'False
      Width           =   75
   End
   Begin VB.Image Image7 
      Height          =   90
      Index           =   0
      Left            =   15
      Picture         =   "MacMenu.ctx":208A
      Top             =   255
      Visible         =   0   'False
      Width           =   60
   End
   Begin VB.Image Image6 
      Height          =   90
      Index           =   0
      Left            =   30
      Picture         =   "MacMenu.ctx":24E4
      Stretch         =   -1  'True
      Top             =   255
      Visible         =   0   'False
      Width           =   2745
   End
   Begin VB.Image Image5 
      Height          =   255
      Left            =   2820
      Picture         =   "MacMenu.ctx":293E
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image4 
      Height          =   255
      Left            =   0
      Picture         =   "MacMenu.ctx":2DC4
      Top             =   0
      Width           =   60
   End
   Begin VB.Image Image1 
      Height          =   255
      Left            =   60
      Picture         =   "MacMenu.ctx":324A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   2775
   End
End
Attribute VB_Name = "MacMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public LastSpace As Boolean
Public TabbText As Boolean
Public Event OnMenuClick(Index As Integer, Caption As String)
Public p
Function ShowMenu()
Dim r
r = Image7.UBound
For p = 1 To r
Unload Image7(Image7.UBound)
Unload Image6(Image6.UBound)
Unload Image8(Image8.UBound)
Next

r = Image3.UBound
For p = 1 To r
Unload Image3(Image3.UBound)
Unload Image9(Image9.UBound)
Unload Image10(Image10.UBound)
Unload Label3(Label3.UBound)
Next

Dim g As Form
Set g = Parent
Dim f
For f = 0 To List1.ListCount
        If f = 0 Then
        If TabbText = True Then
                Label1.Caption = "     " & List1.List(f)
        Else
                Label1.Caption = List1.List(f)
        End If
        End If
        If f = List1.ListCount Then
        If TabbText = True Then
                 Label2.Caption = "     " & List1.List(f - 1)
        Else
                Label2.Caption = List1.List(f - 1)
        End If
        End If
Next
'9,3,10     7,6,8     15,14,16
For f = 1 To List1.ListCount - 2
If List1.List(f) = "<Spacer>" Then
Load Image7(Image7.UBound + 1)
Load Image6(Image6.UBound + 1)
Load Image8(Image8.UBound + 1)
        With Image7(Image7.UBound)
        If LastSpace = True Then
                .Top = Image7(Image7.UBound).Top + Image7(Image7.UBound).Height
        Else
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
        End If
                .Left = 0
                .Visible = True
        End With
        With Image6(Image6.UBound)
        If LastSpace = True Then
                .Top = Image7(Image7.UBound).Top + Image7(Image7.UBound).Height
        Else
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
        End If
                .Left = Image7(Image7.UBound).Left + Image7(Image7.UBound).Width
                .Visible = True
        End With
        With Image8(Image8.UBound)
        If LastSpace = True Then
                .Top = Image7(Image7.UBound).Top + Image7(Image7.UBound).Height
        Else
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
        End If
                .Left = Image6(Image6.UBound).Left + Image6(Image6.UBound).Width
                .Visible = True
        End With
        LastSpace = True
Else
Load Image3(Image3.UBound + 1)
Load Image9(Image9.UBound + 1)
Load Image10(Image10.UBound + 1)
Load Label3(Label3.UBound + 1)
        With Image9(Image9.UBound)
                'If f <> 1 Then
                '.Top = Image1.Top + Image1.Height
                'Else
                .Top = Image9(Image9.UBound - 1).Top + Image9(Image9.UBound - 1).Height
                'End If
                .Left = 0
                .Visible = True
        End With
        With Image3(Image3.UBound)
                .Top = Image9(Image9.UBound).Top
                .Left = Image9(Image9.UBound).Left + Image9(Image9.UBound).Width
                .Visible = True
        End With
        With Image10(Image10.UBound)
                .Top = Image9(Image9.UBound).Top
                .Left = Image3(Image3.UBound).Left + Image3(Image3.UBound).Width
                .Visible = True
        End With
        With Label3(Label3.UBound)
                .Top = Image9(Image9.UBound).Top + 15
                If TabbText = True Then
                .Caption = "     " & List1.List(f)
                Else
                .Caption = List1.List(f)
                End If
                .ZOrder 0
                .Visible = True
        End With
        LastSpace = False
End If
Next
        With Image15
                If LastSpace = False Then
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
                Else
                .Top = Image6(Image6.UBound).Top + Image6(Image6.UBound).Height
                End If
                .Left = 0
        End With
        With Image14
                If LastSpace = False Then
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
                Else
                .Top = Image6(Image6.UBound).Top + Image6(Image6.UBound).Height
                End If
                .Left = Image15.Left + Image15.Width
        End With
        With Image16
                If LastSpace = False Then
                .Top = Image9(Image9.UBound).Top + Image9(Image9.UBound).Height
                Else
                .Top = Image6(Image6.UBound).Top + Image6(Image6.UBound).Height
                End If
                .Left = Image14.Left + Image14.Width
                Label2.Top = Image15.Top - 15
        End With
        For p = 1 To Label3.UBound
        
        If Label3(p).ForeColor = vbWhite Then
                Label3(p).ForeColor = vbBlack
                Label3(p).BackStyle = 0
        End If
        If Label2.ForeColor = vbWhite Then
                Label2.ForeColor = vbBlack
                Label2.BackStyle = 0
        End If
        If Label1.ForeColor = vbWhite Then
                Label1.ForeColor = vbBlack
                Label1.BackStyle = 0
        End If
        Next
        DoEvents
        ResizeMe g
End Function

Private Sub Label1_Click()
If TabbText = True Then
RaiseEvent OnMenuClick(1, Right(Label1.Caption, Len(Label1.Caption) - Len("     ")))
Else
RaiseEvent OnMenuClick(1, Label1.Caption)
End If
End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label1.ForeColor = vbWhite
Label1.BackStyle = 1
        If Label2.ForeColor = vbWhite Then
                Label2.ForeColor = vbBlack
                Label2.BackStyle = 0
        End If
        For p = 1 To Label3.UBound
        If Label3(p).ForeColor = vbWhite Then
                Label3(p).BackStyle = 0
                Label3(p).ForeColor = vbBlack
        End If
        Next
End Sub

Private Sub Label2_Click()
If TabbText = True Then
RaiseEvent OnMenuClick(2, Right(Label2.Caption, Len(Label2.Caption) - Len("     ")))
Else
RaiseEvent OnMenuClick(2, Label2.Caption)
End If
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Label2.ForeColor = vbWhite
Label2.BackStyle = 1
        If Label1.ForeColor = vbWhite Then
                Label1.ForeColor = vbBlack
                Label1.BackStyle = 0
        End If
        For p = 1 To Label3.UBound
        If Label3(p).ForeColor = vbWhite Then
                Label3(p).ForeColor = vbBlack
                Label3(p).BackStyle = 0
        End If
        Next
End Sub

Private Sub Label3_Click(Index As Integer)
If TabbText = True Then
RaiseEvent OnMenuClick(Index, Right(Label3(Index).Caption, Len(Label3(Index).Caption) - Len("     ")))
Else
RaiseEvent OnMenuClick(Index, Label3(Index).Caption)
End If
End Sub

Private Sub Label3_MouseMove(Index As Integer, Button As Integer, Shift As Integer, X As Single, Y As Single)
Label3(Index).ForeColor = vbWhite
Label3(Index).BackStyle = 1
For p = 1 To Label3.UBound
        If Label3(p).ForeColor = vbWhite And p <> Index Then
                Label3(p).BackStyle = 0
                Label3(p).ForeColor = vbBlack
        End If
        If Label1.ForeColor = vbWhite Then
                Label1.BackStyle = 0
                Label1.ForeColor = vbBlack
        End If
        If Label2.ForeColor = vbWhite Then
                Label2.BackStyle = 0
                Label2.ForeColor = vbBlack
        End If
Next

End Sub



Function ResizeMe(frm As Form)
UserControl.Height = Image14.Top + Image14.Height
UserControl.Width = Image16.Width + Image16.Left
'TransparentForm frm
End Function

Function Clear()
List1.Clear
End Function
Function Add(Text As String)
List1.AddItem Text
End Function

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
TabbText = .ReadProperty("TabbText", True)
End With
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
.WriteProperty "TabbText", TabbText, True
End With
End Sub
