VERSION 5.00
Begin VB.Form form2 
   BackColor       =   &H008080FF&
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Enter Player's Names"
   ClientHeight    =   1545
   ClientLeft      =   2835
   ClientTop       =   3480
   ClientWidth     =   3750
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   912.837
   ScaleMode       =   0  'User
   ScaleWidth      =   3521.047
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox text1 
      Height          =   345
      Left            =   1290
      TabIndex        =   1
      Top             =   135
      Width           =   2325
   End
   Begin VB.CommandButton cmdOK 
      BackColor       =   &H00C0C0FF&
      Caption         =   "OK"
      Default         =   -1  'True
      Enabled         =   0   'False
      Height          =   390
      Left            =   495
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   1020
      Width           =   1140
   End
   Begin VB.CommandButton cmdCancel 
      BackColor       =   &H00C0C0FF&
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   390
      Left            =   2100
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   1020
      Width           =   1140
   End
   Begin VB.TextBox text2 
      Height          =   345
      IMEMode         =   3  'DISABLE
      Left            =   1290
      TabIndex        =   3
      Top             =   525
      Width           =   2325
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H008080FF&
      Caption         =   "Player 1 "
      Height          =   270
      Index           =   0
      Left            =   105
      TabIndex        =   0
      Top             =   150
      Width           =   1080
   End
   Begin VB.Label lblLabels 
      BackColor       =   &H008080FF&
      Caption         =   "Player 2"
      Height          =   270
      Index           =   1
      Left            =   105
      TabIndex        =   2
      Top             =   540
      Width           =   1080
   End
End
Attribute VB_Name = "form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public LoginSucceeded As Boolean

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
 pl1 = Trim(text1.Text)
 pl2 = Trim(text2.Text)
 Form1.Label2.Caption = pl1
 Form1.Label3.Caption = pl2
 Unload Me
 NewGame
End Sub

Private Sub text1_Change()
If text1.Text <> "" And text2.Text <> "" Then
cmdOK.Enabled = True
Else
cmdOK.Enabled = False
End If

End Sub

Private Sub text2_Change()
If text1.Text <> "" And text2.Text <> "" Then
cmdOK.Enabled = True
Else
cmdOK.Enabled = False
End If
End Sub
