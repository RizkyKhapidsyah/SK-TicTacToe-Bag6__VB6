VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.ocx"
Begin VB.Form Form1 
   BackColor       =   &H00C0E0FF&
   Caption         =   "Tic Tac Toe"
   ClientHeight    =   4200
   ClientLeft      =   225
   ClientTop       =   870
   ClientWidth     =   5325
   LinkTopic       =   "Form1"
   ScaleHeight     =   4200
   ScaleWidth      =   5325
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1440
      Top             =   1680
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   32
      ImageHeight     =   32
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   3
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0000
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":0454
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Form1.frx":08A8
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   840
      Left            =   0
      TabIndex        =   9
      Top             =   0
      Width           =   5325
      _ExtentX        =   9393
      _ExtentY        =   1482
      ButtonWidth     =   1720
      ButtonHeight    =   1429
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   3
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "New Game"
            Object.ToolTipText     =   "New Game"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Exit"
            Object.ToolTipText     =   "Exit"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Help"
            Key             =   "Help Topics"
            ImageIndex      =   3
         EndProperty
      EndProperty
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   3
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   6
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   7
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   1
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   7
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   4
      Left            =   4920
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   2
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2040
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   5
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   3720
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   8
      Left            =   6840
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   5400
      Width           =   1455
   End
   Begin VB.CommandButton cmd 
      BackColor       =   &H00C0E0FF&
      Height          =   1215
      Index           =   0
      Left            =   3000
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2040
      Width           =   1455
   End
   Begin VB.Label Label4 
      BackColor       =   &H00C0E0FF&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   24
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   615
      Left            =   3120
      TabIndex        =   12
      Top             =   7080
      Width           =   4935
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   11160
      Picture         =   "Form1.frx":0FFC
      Top             =   3120
      Width           =   480
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   11160
      Picture         =   "Form1.frx":1306
      Top             =   2400
      Width           =   480
   End
   Begin VB.Label Label3 
      BackColor       =   &H00C0E0FF&
      Caption         =   "player 2"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   8880
      TabIndex        =   11
      Top             =   3240
      Width           =   2055
   End
   Begin VB.Label Label2 
      BackColor       =   &H00C0E0FF&
      Caption         =   "player 1"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008080&
      Height          =   375
      Left            =   8880
      TabIndex        =   10
      Top             =   2520
      Width           =   2055
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   8
      Left            =   7440
      Picture         =   "Form1.frx":1748
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   7
      Left            =   5280
      Picture         =   "Form1.frx":1B8A
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   6
      Left            =   3360
      Picture         =   "Form1.frx":1FCC
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   5
      Left            =   7320
      Picture         =   "Form1.frx":240E
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   4
      Left            =   5280
      Picture         =   "Form1.frx":2850
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   3
      Left            =   3240
      Picture         =   "Form1.frx":2C92
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   2
      Left            =   7200
      Picture         =   "Form1.frx":30D4
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":3516
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image2 
      Height          =   735
      Index           =   0
      Left            =   3240
      Picture         =   "Form1.frx":3958
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   8
      Left            =   7440
      Picture         =   "Form1.frx":3D9A
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   7
      Left            =   5400
      Picture         =   "Form1.frx":41DC
      Stretch         =   -1  'True
      Top             =   5640
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   6
      Left            =   3360
      Picture         =   "Form1.frx":461E
      Stretch         =   -1  'True
      Top             =   5520
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   5
      Left            =   7320
      Picture         =   "Form1.frx":4A60
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   4
      Left            =   5400
      Picture         =   "Form1.frx":4EA2
      Stretch         =   -1  'True
      Top             =   4080
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   3
      Left            =   3360
      Picture         =   "Form1.frx":52E4
      Stretch         =   -1  'True
      Top             =   3960
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   2
      Left            =   7200
      Picture         =   "Form1.frx":5726
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   1
      Left            =   5280
      Picture         =   "Form1.frx":5B68
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Image Image1 
      Height          =   735
      Index           =   0
      Left            =   3240
      Picture         =   "Form1.frx":5FAA
      Stretch         =   -1  'True
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
   End
   Begin VB.Shape Shape3 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   6480
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape2 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   5055
      Left            =   4560
      Top             =   1800
      Width           =   255
   End
   Begin VB.Shape Shape4 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2760
      Top             =   5040
      Width           =   5775
   End
   Begin VB.Shape Shape1 
      BorderStyle     =   0  'Transparent
      FillColor       =   &H0080C0FF&
      FillStyle       =   0  'Solid
      Height          =   255
      Left            =   2760
      Top             =   3360
      Width           =   5775
   End
   Begin VB.Menu game 
      Caption         =   "Game"
      Begin VB.Menu new 
         Caption         =   "New Game"
      End
      Begin VB.Menu exit 
         Caption         =   "Exit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "Help"
      Begin VB.Menu topics 
         Caption         =   "Help Topics"
      End
      Begin VB.Menu about 
         Caption         =   "About"
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public tr As Boolean
Private Sub cmd_Click(Index As Integer)
turn tr, Index
PlayerWin
TrueFalse
drawn
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
EndGame
End Sub

Private Sub Label1_Click()

End Sub

Private Sub new_Click()
form2.Show (1)
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
If Button.Index = 1 Then
form2.Show (1)
ElseIf Button.Index = 2 Then
End
ElseIf Button.Index = 3 Then
End If
End Sub
