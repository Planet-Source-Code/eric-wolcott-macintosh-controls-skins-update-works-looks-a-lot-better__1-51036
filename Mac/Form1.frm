VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   3285
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   6210
   LinkTopic       =   "Form1"
   ScaleHeight     =   3285
   ScaleWidth      =   6210
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text2 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      TabIndex        =   3
      Text            =   "Text2"
      Top             =   2925
      Width           =   6090
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      Height          =   315
      Left            =   60
      TabIndex        =   2
      Text            =   "Text1"
      Top             =   2610
      Width           =   6090
   End
   Begin Project1.MacMenu MacMenu1 
      Height          =   585
      Left            =   720
      TabIndex        =   1
      Top             =   900
      Visible         =   0   'False
      Width           =   2880
      _ExtentX        =   2937
      _ExtentY        =   2117
   End
   Begin Project1.MacBar MacBar1 
      Height          =   300
      Left            =   0
      TabIndex        =   0
      Top             =   15
      Width           =   6210
      _ExtentX        =   10954
      _ExtentY        =   529
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Form_Load()
Me.Caption = "MAC"
MacBar1.Clear
MacBar1.Add "File"
MacBar1.Add "Edit"
MacBar1.Add "Manage"
MacBar1.Add "Diagnostics"
MacBar1.Add "Help"
MacBar1.loadbuttons
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
MacMenu1.Visible = False
End Sub

Private Sub Form_Resize()
MacBar1.LoadMe Me
End Sub


Private Sub MacBar1_OnMenuClick(MenuIndex As Integer, Caption As String, Left As Long, Top As Long)
Text1.Text = "Bar Item Clicked - " & Caption
MacMenu1.Visible = False
MacMenu1.Clear
Select Case MenuIndex
Case 1
MacMenu1.Add "Open Main Window"
MacMenu1.Add "Close Window"
MacMenu1.Add "<Spacer>"
MacMenu1.Add "Quit"
Case 2
MacMenu1.Add "Dustin"
MacMenu1.Add "Has"
MacMenu1.Add "A"
MacMenu1.Add "Small"
MacMenu1.Add "???????"
Case 2
MacMenu1.Add "Dustin"
MacMenu1.Add "Has"
MacMenu1.Add "A"
MacMenu1.Add "Small"
MacMenu1.Add "???????"
End Select
MacMenu1.ShowMenu
MacMenu1.Top = Me.MacBar1.Height - 10 'UserControl.Height
MacMenu1.Left = Left + 5
MacMenu1.Visible = True
End Sub

Private Sub MacBar1_OnMenuOver(MenuIndex As Integer, Caption As String, Left As Long, Top As Long)
Text1.Text = "Bar Item Mouse Over - " & Caption
MacMenu1.Visible = False
MacMenu1.Clear
Select Case MenuIndex
Case 1
MacMenu1.Add "Open Main Window"
MacMenu1.Add "Close Window"
MacMenu1.Add "<Spacer>"
MacMenu1.Add "Quit"
Case 2
MacMenu1.Add "Dustin"
MacMenu1.Add "Has"
MacMenu1.Add "A"
MacMenu1.Add "Small"
MacMenu1.Add "???????"
Case 2
MacMenu1.Add "Dustin"
MacMenu1.Add "Has"
MacMenu1.Add "A"
MacMenu1.Add "Small"
MacMenu1.Add "???????"
End Select
MacMenu1.ShowMenu
MacMenu1.Top = Me.MacBar1.Height - 10 'UserControl.Height
MacMenu1.Left = Left + 5
MacMenu1.Visible = True
End Sub

Private Sub MacMenu1_OnMenuClick(Index As Integer, Caption As String)
Text2.Text = "Menu Item Selected - " & Caption
Select Case UCase(Caption)
Case "QUIT"
End
End Select
End Sub
