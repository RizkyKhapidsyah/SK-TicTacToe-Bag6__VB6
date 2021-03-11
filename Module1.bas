Attribute VB_Name = "Module1"
Public pl1, pl2, winner As String
Public Sub turn(tr As Boolean, ar As Integer)
Form1.cmd(ar).Visible = False
If tr = True Then
Form1.Image1(ar).Visible = True
Else
Form1.Image2(ar).Visible = True
End If
End Sub

Public Sub TrueFalse()
If Form1.tr = True Then
Form1.tr = False
ElseIf Form1.tr = False Then
Form1.tr = True
End If
End Sub

Public Sub PlayerWin()
With Form1

'Check if player 1 wins
    
    If (.Image1(0).Visible = True And .Image1(1).Visible = True And .Image1(2).Visible = True) Or _
    (.Image1(0).Visible = True And .Image1(3).Visible = True And .Image1(6).Visible = True) Or _
    (.Image1(0).Visible = True And .Image1(4).Visible = True And .Image1(8).Visible = True) Or _
    (.Image1(1).Visible = True And .Image1(4).Visible = True And .Image1(7).Visible = True) Or _
    (.Image1(2).Visible = True And .Image1(5).Visible = True And .Image1(8).Visible = True) Or _
    (.Image1(3).Visible = True And .Image1(4).Visible = True And .Image1(5).Visible = True) Or _
    (.Image1(6).Visible = True And .Image1(7).Visible = True And .Image1(8).Visible = True) Or _
    (.Image1(6).Visible = True And .Image1(4).Visible = True And .Image1(2).Visible = True) Then
     MsgBox "player1 wins the game"
     Form1.Label4.Caption = Form1.Label2.Caption & " Wins the game"
     EndGame
'Check if player 2 wins

 ElseIf (.Image2(0).Visible = True And .Image2(1).Visible = True And .Image2(2).Visible = True) Or _
    (.Image2(0).Visible = True And .Image2(3).Visible = True And .Image2(6).Visible = True) Or _
    (.Image2(0).Visible = True And .Image2(4).Visible = True And .Image2(8).Visible = True) Or _
    (.Image2(1).Visible = True And .Image2(4).Visible = True And .Image2(7).Visible = True) Or _
    (.Image2(2).Visible = True And .Image2(5).Visible = True And .Image2(8).Visible = True) Or _
    (.Image2(3).Visible = True And .Image2(4).Visible = True And .Image2(5).Visible = True) Or _
    (.Image2(6).Visible = True And .Image2(7).Visible = True And .Image2(8).Visible = True) Or _
    (.Image2(6).Visible = True And .Image2(4).Visible = True And .Image2(2).Visible = True) Then
    MsgBox "player 2 wins the game  "
    Form1.Label4.Caption = Form1.Label3.Caption & " Wins the game"
    EndGame
    End If


End With
End Sub

Public Sub NewGame()
' Reset all variables
   pl1 = ""
   pl2 = ""
   Form1.tr = True
      
 'Disable images,buttons
 Dim ctrl As Control
   For Each ctrl In Form1
     If TypeOf ctrl Is CommandButton Then
      ctrl.Visible = True
      ctrl.Enabled = True
     End If
     If TypeOf ctrl Is Image Then
      ctrl.Visible = False
    End If
   Next
 Form1.Image3.Visible = True
 Form1.Image4.Visible = True
End Sub

Public Sub EndGame()
' Disable all buttons
   Dim ctrl As Control
   For Each ctrl In Form1
   If TypeOf ctrl Is CommandButton Then
    ctrl.Enabled = False
   End If
   Next
End Sub

Public Sub drawn()
Dim dec As Boolean
Dim i As Integer
For i = 0 To 8
If Form1.cmd(i).Visible = True Then
dec = True
End If
Next i
If dec = False Then
MsgBox "Game Drawn"
Form1.Label4.Caption = " Game Drawn"
EndGame
End If
End Sub
