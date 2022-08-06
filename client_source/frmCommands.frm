VERSION 5.00
Begin VB.Form frmCommands 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Commands"
   ClientHeight    =   4620
   ClientLeft      =   2010
   ClientTop       =   1995
   ClientWidth     =   2970
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4620
   ScaleWidth      =   2970
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Change gang = @<newgang>"
      Top             =   4215
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Change name = :<newname>"
      Top             =   3975
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "Withdraw money = withdraw<amount>"
      Top             =   3735
      Width           =   2955
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Deposit money = deposit<amount>"
      Top             =   3495
      Width           =   2865
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Bank:"
      Top             =   3255
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Gamble = bet<amount>"
      Top             =   3015
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Hide = hide"
      Top             =   2775
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Track = track<playername>"
      Top             =   2535
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Snoop = snoop<playername>"
      Top             =   2295
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "$Transfer = $name:amount"
      Top             =   2055
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Aim = aim<name>"
      Top             =   1815
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Square talk = say Msg"
      Top             =   1560
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "PM = 'name-Msg"
      Top             =   1320
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Talk = ;"
      Top             =   1080
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Look = F4"
      Top             =   840
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Shoot = F3"
      Top             =   600
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Strike = F2"
      Top             =   360
      Width           =   2415
   End
   Begin VB.TextBox Text1 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Punch = F1"
      Top             =   120
      Width           =   2415
   End
End
Attribute VB_Name = "frmCommands"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
