VERSION 5.00
Begin VB.Form frmNpcRanks 
   BackColor       =   &H00808080&
   BorderStyle     =   1  'Fixed Single
   ClientHeight    =   4935
   ClientLeft      =   2010
   ClientTop       =   1800
   ClientWidth     =   3150
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4935
   ScaleWidth      =   3150
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   34
      Text            =   "70"
      Top             =   4580
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   18
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   18
      Text            =   "V.I.P"
      Top             =   4580
      Width           =   570
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   35
      Text            =   "67"
      Top             =   4340
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   1
      Text            =   "Scientist"
      Top             =   4340
      Width           =   840
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   36
      Text            =   "60"
      Top             =   4100
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   3
      Text            =   "Pimp"
      Top             =   4100
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   25
      Text            =   "55"
      Top             =   3860
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   16
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Swat Agent"
      Top             =   3860
      Width           =   1065
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   26
      Text            =   "50"
      Top             =   3620
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   2
      Text            =   "Silent Assasin"
      Top             =   3620
      Width           =   1200
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   8
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   27
      Text            =   "50"
      Top             =   3380
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   7
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   15
      Text            =   "Iraqi Terrorist"
      Top             =   3380
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   28
      Text            =   "45"
      Top             =   3140
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   17
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   8
      Text            =   "Body Guard"
      Top             =   3140
      Width           =   1155
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   10
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   29
      Text            =   "30"
      Top             =   2900
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "Yardie"
      Top             =   2900
      Width           =   660
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   11
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   30
      Text            =   "25"
      Top             =   2660
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   13
      Text            =   "Police Chief"
      Top             =   2660
      Width           =   1020
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   12
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   31
      Text            =   "20"
      Top             =   2420
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   12
      Text            =   "Police Officer"
      Top             =   2420
      Width           =   1110
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   32
      Text            =   "12"
      Top             =   2180
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   9
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   16
      Text            =   "Slut"
      Top             =   2180
      Width           =   525
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   33
      Text            =   "12"
      Top             =   1940
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   15
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   7
      Text            =   "Rapist"
      Top             =   1940
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   5
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   24
      Text            =   "10"
      Top             =   1700
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   14
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   6
      Text            =   "Burglar"
      Top             =   1700
      Width           =   750
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   4
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   23
      Text            =   "5"
      Top             =   1460
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   22
      Text            =   "5"
      Top             =   1220
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   20
      Text            =   "0"
      Top             =   980
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   21
      Text            =   "-250"
      Top             =   740
      Width           =   870
   End
   Begin VB.TextBox Text1 
      Alignment       =   2  'Center
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   1890
      Locked          =   -1  'True
      TabIndex        =   19
      Text            =   "-250"
      Top             =   500
      Width           =   870
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   13
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   5
      Text            =   "Goth"
      Top             =   1460
      Width           =   660
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   3
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   11
      Text            =   "Street Bum"
      Top             =   1220
      Width           =   1200
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   6
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   14
      Text            =   "News Reporter"
      Top             =   980
      Width           =   1245
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   2
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   10
      Text            =   "Drug Taker"
      Top             =   740
      Width           =   1155
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   1
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   9
      Text            =   "Dealer"
      Top             =   500
      Width           =   750
   End
   Begin VB.TextBox Text0 
      Appearance      =   0  'Flat
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Index           =   0
      Left            =   120
      Locked          =   -1  'True
      TabIndex        =   0
      Text            =   "Npc:                            Rank points u get:"
      Top             =   120
      Width           =   3000
   End
End
Attribute VB_Name = "frmNpcRanks"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
