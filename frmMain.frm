VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "SMURF WARS Demo"
   ClientHeight    =   6795
   ClientLeft      =   165
   ClientTop       =   450
   ClientWidth     =   9480
   LinkTopic       =   "Form1"
   ScaleHeight     =   453
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   632
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdExit 
      Caption         =   "E&xit"
      Height          =   255
      Left            =   5520
      TabIndex        =   1
      Top             =   7680
      Width           =   1215
   End
   Begin VB.PictureBox picGameField 
      BorderStyle     =   0  'None
      Height          =   7200
      Left            =   480
      MousePointer    =   2  'Cross
      Picture         =   "frmMain.frx":0000
      ScaleHeight     =   480
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   736
      TabIndex        =   0
      Top             =   240
      Width           =   11040
   End
   Begin VB.Timer tmrAnimation 
      Interval        =   75
      Left            =   8040
      Top             =   120
   End
   Begin VB.Label lblBodyCount 
      Alignment       =   2  'Center
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   8280
      TabIndex        =   5
      Top             =   7800
      Width           =   1095
   End
   Begin VB.Label lblCountSign 
      Alignment       =   2  'Center
      Caption         =   "BODY COUNT"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   7920
      TabIndex        =   4
      Top             =   7440
      Width           =   1815
   End
   Begin VB.Label lblHealthSign 
      Alignment       =   2  'Center
      Caption         =   "HEALTH"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3120
      TabIndex        =   3
      Top             =   7440
      Width           =   1215
   End
   Begin VB.Label lblHealth 
      Alignment       =   2  'Center
      Caption         =   "100"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   2
      Top             =   7800
      Width           =   975
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim SmurfCounter As Integer
Dim Player As tBitmap
Dim MouseX, MouseY As Integer
Dim Health As Integer
Private Sub cmdExit_Click()
  End
End Sub
Private Sub cmdMask_Click()
  Dim RC As Integer
  RC = BitBlt(picGameField.hDC, 0, 0, picGameField.ScaleWidth, _
              picGameField.ScaleHeight, _
              frmGraphics.picSmurf(1).hDC, 0, 0, SRCAND)
End Sub
Private Sub cmdSprite_Click()
  Dim RC As Integer
  RC = BitBlt(picGameField.hDC, 0, 0, picGameField.ScaleWidth, _
              picGameField.ScaleHeight, _
              frmGraphics.picSmurf(0).hDC, 0, 0, SRCPAINT)
End Sub
Private Sub Form_Load()
  ScaleMode = vbPixels
  DoEvents
  Dim X As Integer
  Dim GameField As tBitmap
  Player.Height = frmGraphics.picPlayer.ScaleHeight
  Player.Width = frmGraphics.picPlayer.ScaleWidth
  Player.Bottom = picGameField.ScaleHeight
  Player.Top = Player.Bottom - Player.Height
  Player.Left = 200
  Player.Right = Player.Left + (Player.Width - 1)
  Randomize
  'picGameField.ScaleWidth = frmMain.ScaleWidth - 50
  'picGameField.Left = (frmMain.ScaleWidth / 2) - (picGameField.ScaleWidth / 2)
  'picGameField.ScaleHeight = frmMain.ScaleHeight - (frmMain.ScaleHeight / 9)
  Load frmGraphics
  For X = 0 To 2
    Smurf(X).Height = frmGraphics.picSmurf(0).ScaleHeight
    Smurf(X).Width = frmGraphics.picSmurf(0).ScaleWidth
    NewSmurf (X)
  Next X
  'Smurf(0).Left = picGameField.Width
  'Smurf(0).Right = Smurf(0).Left + (Smurf(0).Width - 1)
  'Smurf(0).Top = picGameField.Height / 2
  'Smurf(0).Bottom = Smurf(0).Top + Smurf(0).Height
  'Smurf(0).Frame = 0
  SmurfCounter = 0
  'lblGFSW.Caption = Str(picGameField.Width)
  'lblSmurfLeft.Caption = Str(frmMain.Width)
  Health = 100
  BodyCount = 0
End Sub

Private Sub picGameField_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim RC As Integer
  If Button = 1 Then
    For RC = 0 To 2
        If Smurf(RC).Alive = True Then
            'Hit check
            If X > Smurf(RC).Left And X < Smurf(RC).Right Then
                If Y > Smurf(RC).Top And Y < Smurf(RC).Bottom Then
                    Smurf(RC).Alive = False
                    Smurf(RC).Frame = 10
                    BodyCount = BodyCount + 1
                    lblBodyCount.Caption = Str(BodyCount)
                    If BodyCount > 99 Then
                        If Smurf(0).Alive = False And _
                            Smurf(1).Alive = False And _
                            Smurf(2).Alive = False Then
                            WinGame
                        End If
                    End If
                End If
            End If
        End If
    Next RC
  End If
End Sub

Private Sub picGameField_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
  Dim RC As Integer
  Player.Left = Int(X) - (Player.Width / 2)
  Player.Right = Player.Left + (Player.Width - 1)
End Sub

Private Sub tmrAnimation_Timer()
  Dim RC As Integer
  Dim X As Integer
  Dim randomfire As Integer
  Dim NewMove As Integer
  'Copy unblemished background to workarea
  RC = BitBlt(frmGraphics.picWorkArea.hDC, 0, 0, picGameField.ScaleWidth, _
              picGameField.ScaleHeight, frmGraphics.picBackground.hDC, _
              0, 0, SRCCOPY)
  For X = 0 To 2
    'Copy Smurf Mask to work area
    RC = BitBlt(frmGraphics.picWorkArea.hDC, Smurf(X).Left, _
                  Smurf(X).Top, Smurf(X).Width, _
                  Smurf(X).Height, _
                  frmGraphics.picSmurf(Smurf(X).Frame + 1).hDC, _
                  0, 0, SRCAND)
    'Copy Smurf onto mask on work area
    RC = BitBlt(frmGraphics.picWorkArea.hDC, Smurf(X).Left, _
                  Smurf(X).Top, Smurf(X).Width, _
                  Smurf(X).Height, _
                  frmGraphics.picSmurf(Smurf(X).Frame).hDC, _
                  0, 0, SRCPAINT)
  Next X
  'Player.Left = frmMain.MouseIcon
  'Copy Player Mask onto work area
  RC = BitBlt(frmGraphics.picWorkArea.hDC, Player.Left, _
                Player.Top, Player.Width, _
                Player.Height, _
                frmGraphics.picPlayerMask.hDC, _
                0, 0, SRCAND)
  'Copy Player onto mask on work area
  RC = BitBlt(frmGraphics.picWorkArea.hDC, Player.Left, _
                Player.Top, Player.Width, _
                Player.Height, _
                frmGraphics.picPlayer.hDC, _
                0, 0, SRCPAINT)
  'Copy workarea to main game screen
  RC = BitBlt(picGameField.hDC, 0, 0, picGameField.ScaleWidth, _
              picGameField.ScaleHeight, frmGraphics.picWorkArea.hDC, _
              0, 0, SRCCOPY)
  'Increment/Decrement the Smurf's position
  For X = 0 To 2
    If Smurf(X).Frame >= 0 And Smurf(X).Frame < 4 Then
        Smurf(X).Left = Smurf(X).Left - SmurfSpeed
    End If
    If Smurf(X).Frame > 3 And Smurf(X).Frame < 7 Then
        Smurf(X).Left = Smurf(X).Left + SmurfSpeed
    End If
    Smurf(X).Right = Smurf(X).Left + (Smurf(X).Width - 1)
    If Smurf(X).Right <= 0 Or Smurf(X).Left >= picGameField.Width Then
        NewSmurf (X)
    End If
  'MsgBox "Smurf Left = " + Str(Smurf(0).Left)
    Select Case Smurf(X).Frame
        Case 0
            Smurf(X).Frame = 2
        Case 2
            Smurf(X).Frame = 0
        Case 4
            Smurf(X).Frame = 6
        Case 6
            Smurf(X).Frame = 4
        Case 8
            NewMove = Int((10 * Rnd) + 1)
            If NewMove < 6 Then
                Smurf(X).Frame = 0
            Else
                Smurf(X).Frame = 4
            End If
        Case 10
            Smurf(X).Frame = 12
        Case 12
            Smurf(X).Frame = 14
        Case 14
            'Draw the Smurf's dead body onto the
            'prestine background
            RC = BitBlt(frmGraphics.picBackground.hDC, _
                        Smurf(X).Left, Smurf(X).Top, _
                        Smurf(X).Width, Smurf(X).Height, _
                        frmGraphics.picSmurf(15).hDC, _
                        0, 0, SRCAND)
            RC = BitBlt(frmGraphics.picBackground.hDC, _
                        Smurf(X).Left, Smurf(X).Top, _
                        Smurf(X).Width, Smurf(X).Height, _
                        frmGraphics.picSmurf(14).hDC, _
                        0, 0, SRCPAINT)
            If BodyCount < 100 Then
                NewSmurf (X)
            End If
    End Select
    'Random smurf firing -- if #1, smurf fires
    If Smurf(X).Alive = True Then
        randomfire = Int((10 * Rnd) + 1)
        If randomfire = 1 Then
            Smurf(X).Frame = 8
            'Check for hit
            If Smurf(X).Left > Player.Left Then
                If Smurf(X).Right < Player.Right Then
                    Health = Health - DamageFactor
                    lblHealth.Caption = Str(Health)
                    lblHealth.Refresh
                    If Health <= 0 Then
                        LoseGame
                    End If
                End If
            End If
        End If
    End If
  Next X
  'If Smurf(0).Frame = 2 Then
  '  SmurfCounter = 0
  'Else
  '  SmurfCounter = 2
  'End If
End Sub

Public Sub NewSmurf(SmurfNumber As Integer)
  Dim LeftorRight As Integer
  'Generate random number to determine if new smurf starts
  'on left or right side of screen
  LeftorRight = Int((2 * Rnd) + 1)
  If LeftorRight = 1 Then
    'Starts on left side
    Smurf(SmurfNumber).Frame = 4
    Smurf(SmurfNumber).Left = 0
  Else
    'Starts on right side
    Smurf(SmurfNumber).Frame = 0
    Smurf(SmurfNumber).Left = picGameField.Width
  End If
  Smurf(SmurfNumber).Right = Smurf(SmurfNumber).Left + (Smurf(SmurfNumber).Width - 1)
  'Assign Random number to Smurf.Top property
  Smurf(SmurfNumber).Top = Int((picGameField.Height * Rnd) + 1) + (picGameField.Height / 5)
  Smurf(SmurfNumber).Bottom = Smurf(SmurfNumber).Top + Smurf(SmurfNumber).Height
  Smurf(SmurfNumber).Alive = True
  If Smurf(SmurfNumber).Top > (picGameField.Height) - (picGameField.Height / 2) Then
    NewSmurf (SmurfNumber)
  End If
End Sub

Public Sub LoseGame()
  MsgBox "You have fallen to the Smurfs!", vbOKOnly, "GAME OVER"
  frmIntro.Show
  Unload Me
  Unload frmGraphics
End Sub

Public Sub WinGame()
  MsgBox "You have defeated the Smurfs! " + vbCrLf + _
         "In the full version of the demo, you will face" + vbCrLf + _
         "a boss character after the Smurf troops are defeated." + vbCrLf + _
         "Thanks for playing!", vbOKOnly, "YOU WIN!"
  frmIntro.Show
  Unload Me
  Unload frmGraphics
End Sub
