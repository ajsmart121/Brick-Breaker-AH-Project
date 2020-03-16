VERSION 5.00
Begin VB.Form frmcustomise 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Customisations"
   ClientHeight    =   7875
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8445
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7875
   ScaleWidth      =   8445
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame fraball 
      BackColor       =   &H00000000&
      Caption         =   "Ball Colour"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   4200
      TabIndex        =   1
      Top             =   3840
      Width           =   3855
      Begin VB.OptionButton optgrey 
         BackColor       =   &H00000000&
         Caption         =   "Grey"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00808080&
         Height          =   615
         Left            =   360
         TabIndex        =   11
         Top             =   2640
         Width           =   3255
      End
      Begin VB.OptionButton optpurple 
         BackColor       =   &H00000000&
         Caption         =   "Purple"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C000C0&
         Height          =   615
         Left            =   360
         TabIndex        =   10
         Top             =   1920
         Width           =   3255
      End
      Begin VB.OptionButton optorange 
         BackColor       =   &H00000000&
         Caption         =   "Orange"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080C0FF&
         Height          =   615
         Left            =   360
         TabIndex        =   9
         Top             =   1200
         Width           =   3375
      End
      Begin VB.OptionButton optblue 
         BackColor       =   &H00000000&
         Caption         =   "Blue"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFC0&
         Height          =   495
         Left            =   360
         TabIndex        =   8
         Top             =   600
         Width           =   3015
      End
   End
   Begin VB.Frame frapaddle 
      BackColor       =   &H00000000&
      Caption         =   "Paddle Colour"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   3615
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   3855
      Begin VB.OptionButton optpink 
         BackColor       =   &H00000000&
         Caption         =   "Pink"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF00FF&
         Height          =   495
         Left            =   360
         TabIndex        =   7
         Top             =   600
         Width           =   3015
      End
      Begin VB.OptionButton optgreen 
         BackColor       =   &H80000012&
         Caption         =   "Green"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00C0FFC0&
         Height          =   495
         Left            =   360
         TabIndex        =   6
         Top             =   1320
         Width           =   3255
      End
      Begin VB.OptionButton optyellow 
         BackColor       =   &H00000000&
         Caption         =   "Yellow"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H0080FFFF&
         Height          =   495
         Left            =   360
         TabIndex        =   5
         Top             =   2040
         Width           =   3255
      End
      Begin VB.OptionButton optwhite 
         BackColor       =   &H00000000&
         Caption         =   "White"
         BeginProperty Font 
            Name            =   "Showcard Gothic"
            Size            =   18
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   615
         Left            =   360
         TabIndex        =   4
         Top             =   2760
         Width           =   3015
      End
   End
   Begin VB.Label lblgame 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "Play Game"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0FFFF&
      Height          =   1575
      Left            =   840
      TabIndex        =   3
      Top             =   5160
      Width           =   2535
   End
   Begin VB.Label lblhome 
      BackColor       =   &H00FF80FF&
      BackStyle       =   0  'Transparent
      Caption         =   "BAck to home"
      BeginProperty Font 
         Name            =   "Showcard Gothic"
         Size            =   24
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFC0FF&
      Height          =   1335
      Left            =   5520
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Image imgbackground 
      Enabled         =   0   'False
      Height          =   8295
      Left            =   120
      Picture         =   "frmcustomise.frx":0000
      Stretch         =   -1  'True
      Top             =   0
      Width           =   8535
   End
End
Attribute VB_Name = "frmcustomise"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblgame_Click()

frmcustomise.Hide
frmgame.Enabled = True
frmgame.Show
End Sub

Private Sub lblhome_Click()
frmcustomise.Hide
frmhome.Enabled = True
frmhome.Show
End Sub

Private Sub optblue_Click()
frmgame.Ball.BackColor = optblue.ForeColor
End Sub

Private Sub optgreen_Click()
frmgame.Paddle.FillColor = optgreen.ForeColor
End Sub

Private Sub optgrey_Click()
frmgame.Ball.BackColor = optgrey.ForeColor
End Sub

Private Sub optorange_Click()
frmgame.Ball.BackColor = optorange.ForeColor
End Sub

Private Sub optpink_Click()
frmgame.Paddle.FillColor = optpink.ForeColor
End Sub

Private Sub optpurple_Click()
frmgame.Ball.BackColor = optpurple.ForeColor
End Sub

Private Sub optwhite_Click()
frmgame.Paddle.FillColor = optwhite.ForeColor
End Sub

Private Sub optyellow_Click()
frmgame.Paddle.FillColor = optyellow.ForeColor
End Sub

Private Sub Form_Load()

optblue.Value = False
optgreen.Value = False
optorange.Value = False
optpink.Value = False
optpurple.Value = False
optyellow.Value = False
optwhite.Value = False
optgrey.Value = False

End Sub
