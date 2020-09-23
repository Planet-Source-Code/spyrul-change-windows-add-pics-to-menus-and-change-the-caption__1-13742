VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "Window Changer"
   ClientHeight    =   3195
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   Picture         =   "frmMain.frx":0000
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "change caption"
      Height          =   495
      Left            =   3000
      TabIndex        =   9
      Top             =   1920
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   525
      Left            =   2520
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   8
      Top             =   1200
      Width           =   2055
   End
   Begin VB.TextBox Text1 
      Height          =   525
      Left            =   720
      MultiLine       =   -1  'True
      ScrollBars      =   1  'Horizontal
      TabIndex        =   4
      Top             =   120
      Width           =   3255
   End
   Begin VB.ComboBox Combo2 
      Height          =   315
      Left            =   1200
      TabIndex        =   3
      Top             =   1200
      Width           =   615
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   600
      TabIndex        =   2
      Top             =   1200
      Width           =   615
   End
   Begin VB.CommandButton Command1 
      Caption         =   "change menu"
      Height          =   495
      Left            =   600
      TabIndex        =   1
      Top             =   1560
      Width           =   1215
   End
   Begin VB.PictureBox Picture1 
      Height          =   855
      Left            =   720
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   795
      ScaleWidth      =   795
      TabIndex        =   0
      Top             =   2160
      Visible         =   0   'False
      Width           =   855
   End
   Begin VB.Label Label4 
      Caption         =   "New Caption"
      Height          =   495
      Left            =   2520
      TabIndex        =   10
      Top             =   960
      Width           =   1095
   End
   Begin VB.Line Line2 
      X1              =   2400
      X2              =   2400
      Y1              =   720
      Y2              =   3240
   End
   Begin VB.Label Label3 
      Caption         =   "Sub"
      Height          =   495
      Left            =   1200
      TabIndex        =   7
      Top             =   960
      Width           =   615
   End
   Begin VB.Label Label2 
      Caption         =   "Main"
      Height          =   495
      Left            =   600
      TabIndex        =   6
      Top             =   960
      Width           =   615
   End
   Begin VB.Line Line1 
      X1              =   0
      X2              =   4680
      Y1              =   720
      Y2              =   720
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Window Caption"
      Height          =   495
      Left            =   0
      TabIndex        =   5
      Top             =   120
      Width           =   735
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuNew 
         Caption         =   "&New"
      End
      Begin VB.Menu mnuSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuExit 
         Caption         =   "&Exit"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub Command1_Click()
Dim lngMenu As Long
Dim lngSubMenu As Long
Dim lngMenuItemID As Long
Dim lngRet As Long
Dim handle As Long
handle = FindWindow(vbNullString, Text1)

lngMenu = GetMenu(handle)

lngSubMenu = GetSubMenu(lngMenu, Combo1.Text - 1)

lngMenuItemID = GetMenuItemID(lngSubMenu, Combo1.Text - 1)

lngRet = SetMenuItemBitmaps(lngMenu, lngMenuItemID, 0, Picture1.Picture, Picture1.Picture)
End Sub

Private Sub Command2_Click()
a = FindWindow(vbNullString, Text1)
If a = 0 Then Exit Sub

r$ = GetWindowTitle(a)
z$ = Text2.Text
If z$ = "" Then Exit Sub
If z$ = Text1.Text Then Exit Sub
SetWindowText a, z$
End Sub

Private Sub Form_Load()
Dim b As Integer
Dim j As Integer
For b = 1 To 20
    Combo1.AddItem b
Next b
For j = 1 To 35
    Combo2.AddItem j
Next j
End Sub

Private Sub mnuExit_Click()
End
End Sub

Private Sub mnuNew_Click()
Text1.Text = ""
Text2.Text = ""
End Sub
