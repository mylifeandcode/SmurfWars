VERSION 5.00
Begin VB.Form frmIntro 
   BackColor       =   &H00FF8080&
   Caption         =   "Smurf Wars Pre Demo by Alan Hummel"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4620
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   ScaleHeight     =   386
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   308
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   0
      TabIndex        =   4
      Top             =   4440
      Width           =   4695
   End
   Begin VB.CommandButton cmdIntro 
      Caption         =   "I am Blue Death!  >:)"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   2
      Left            =   0
      TabIndex        =   3
      Top             =   3840
      Width           =   4695
   End
   Begin VB.CommandButton cmdIntro 
      Caption         =   "Smurf Killer!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   1
      Left            =   0
      TabIndex        =   2
      Top             =   3240
      Width           =   4695
   End
   Begin VB.CommandButton cmdIntro 
      Caption         =   "3 Apples tall? HA!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Index           =   0
      Left            =   0
      TabIndex        =   1
      Top             =   2640
      Width           =   4695
   End
   Begin VB.Image Image2 
      BorderStyle     =   1  'Fixed Single
      Height          =   810
      Left            =   960
      Picture         =   "frmIntro.frx":0000
      Stretch         =   -1  'True
      Top             =   5040
      Width           =   2700
   End
   Begin VB.Label lblSelect 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Please select a carnage level"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   495
      Left            =   0
      TabIndex        =   0
      Top             =   2160
      Width           =   4575
   End
   Begin VB.Image Image1 
      Height          =   2130
      Left            =   0
      Picture         =   "frmIntro.frx":B40A
      Stretch         =   -1  'True
      Top             =   0
      Width           =   4620
   End
End
Attribute VB_Name = "frmIntro"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
  End
End Sub

Private Sub cmdIntro_Click(Index As Integer)
  Select Case Index
    Case 0
        SmurfSpeed = 10
        SmurfAttack = 5
        DamageFactor = 2
        frmMain.Show
        'frmMain.tmrAnimation.Interval = 30
        Unload Me
    Case 1
        SmurfSpeed = 15
        SmurfAttack = 10
        DamageFactor = 4
        frmMain.Show
        'frmMain.tmrAnimation.Interval = 15
        Unload Me
    Case 2
        SmurfSpeed = 20
        SmurfAttack = 15
        DamageFactor = 10
        frmMain.Show
        'frmMain.tmrAnimation.Interval = 5
        Unload Me
  End Select
End Sub

