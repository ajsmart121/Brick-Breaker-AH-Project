VERSION 5.00
Begin VB.Form frmhome 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Home"
   ClientHeight    =   8205
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8280
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8205
   ScaleWidth      =   8280
   StartUpPosition =   3  'Windows Default
   Begin VB.Label lblcustomise 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Customisations"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   1800
      TabIndex        =   5
      Top             =   5880
      Width           =   5055
   End
   Begin VB.Label lblinstructions 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Instructions"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   2520
      TabIndex        =   4
      Top             =   4320
      Width           =   3375
   End
   Begin VB.Label lblscoreboard 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Scoreboard"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   615
      Left            =   2520
      TabIndex        =   3
      Top             =   5160
      Width           =   3375
   End
   Begin VB.Label lblquit 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Quit."
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   3360
      TabIndex        =   2
      Top             =   6720
      Width           =   1575
   End
   Begin VB.Label lblstart 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Start!"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   26.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C000&
      Height          =   735
      Left            =   3240
      TabIndex        =   1
      Top             =   3360
      Width           =   2055
   End
   Begin VB.Label lbltiitle 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Bricks"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   60
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   1455
      Left            =   1920
      TabIndex        =   0
      Top             =   1320
      Width           =   4575
   End
   Begin VB.Image imgbackground 
      Height          =   8295
      Left            =   0
      Picture         =   "frmhome.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmhome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()

imgbackground.Enabled = False
frmhome.Enabled = True
frmgame.Enabled = False

'** load the background image
End Sub


Private Sub lblcustomise_Click()

frmhome.Hide
frmhome.Enabled = False
frmcustomise.Enabled = True
frmcustomise.Show

'** if the customize button is pressed, then show the customize form and hide the home form
End Sub

Private Sub lblinstructions_Click()

frmhome.Hide
frmhome.Enabled = False
frminstructions.Show
frminstructions.Enabled = True

'** if the instructions button is pressed, then show the instructions form and hide the home form

End Sub

Private Sub lblquit_Click()

End

'** if the quit button is pressed then close the program

End Sub

Private Sub lblscoreboard_Click()

frmhome.Hide
frmhome.Enabled = False
frmscoreboard.Enabled = True
frmscoreboard.Show

End Sub

Private Sub lblstart_Click()

frmhome.Visible = False
frmhome.Enabled = False

frmgame.Visible = True
frmgame.Enabled = True

'** if the start button is pressed, then show the game form and hide the home form

End Sub
