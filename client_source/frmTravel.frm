VERSION 5.00
Begin VB.Form frmTravel 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Airport Checkout"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5595
   Icon            =   "frmTravel.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   384
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   373
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command1 
      Caption         =   "Washigton DC"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   1560
      TabIndex        =   14
      Top             =   4680
      Width           =   2415
   End
   Begin VB.PictureBox Picture1 
      Height          =   1575
      Left            =   120
      Picture         =   "frmTravel.frx":0442
      ScaleHeight     =   1515
      ScaleWidth      =   5355
      TabIndex        =   13
      Top             =   0
      Width           =   5415
   End
   Begin VB.CommandButton cmdForgetIt 
      Caption         =   "Forget It"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   6
      Top             =   5280
      Width           =   5370
   End
   Begin VB.CommandButton cmdNewJersey 
      Caption         =   "New Jersey"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   5
      Top             =   3720
      Width           =   1890
   End
   Begin VB.CommandButton cmdChicago 
      Caption         =   "Chicago"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   4
      Top             =   3720
      Width           =   1890
   End
   Begin VB.CommandButton cmdMiami 
      Caption         =   "Miami"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   3
      Top             =   2880
      Width           =   1890
   End
   Begin VB.CommandButton cmdHouston 
      Caption         =   "Houston"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   2
      Top             =   2880
      Width           =   1890
   End
   Begin VB.CommandButton cmdLosAngeles 
      Caption         =   "Los Angeles"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   2040
      Width           =   1890
   End
   Begin VB.CommandButton cmdNewYork 
      Caption         =   "New York"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   2040
      Width           =   1890
   End
   Begin VB.Label lblWashington 
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   1680
      TabIndex        =   15
      Top             =   4320
      Width           =   2175
   End
   Begin VB.Label lblNewJersey 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   12
      Top             =   3360
      Width           =   1890
   End
   Begin VB.Label lblChicago 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   3360
      Width           =   1890
   End
   Begin VB.Label lblMiami 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   10
      Top             =   2520
      Width           =   1890
   End
   Begin VB.Label lblHouston 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   1890
   End
   Begin VB.Label lblLosAngeles 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   3600
      TabIndex        =   8
      Top             =   1680
      Width           =   1890
   End
   Begin VB.Label lblNewYork 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   1890
   End
End
Attribute VB_Name = "frmTravel"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'*****************************************************
'  Streetwars Online 2 Version 1.00
'  Copyright 2000 - B.Smith aka (Wuzzbent)
'  All Rights Reserved
'  wuzzbent@swbell.net
'
'  By using this source code, you agree to the following
'  terms and conditions.
'
'  You may use this source code for your own personal
'  pleasure and use.  You may freely distribute it along with
'  any modification(s) made to it.  You may NOT remove, modify,
'  or adjust this copyright information.  You may NOT attempt
'  to charge for the use of this software under any conditions.
'
'  Support Free Software....
'
'******************************************************

Option Explicit

Private Sub cmdChicago_Click()

'Fly to Chicago
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "chicago" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdForgetIt_Click()

Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdHouston_Click()

'Fly to Houston
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "houston" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdLosAngeles_Click()

'Fly to Los Angeles
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "los angeles" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdMiami_Click()

'Fly to Miami
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "miami" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewJersey_Click()

'Fly to New Jersey
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new jersey" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdNewYork_Click()

'Fly to New York
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "new york" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdWashington_Click()

'Fly to Washington
frmMain.wsk.SendData Chr$(255) & Chr$(7) & "washington" & Chr$(0)
DoEvents
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_Load()

Dim hMenu As Long
Dim menuItemCount As Long
'Obtain the handle to the form's system menu
hMenu = GetSystemMenu(Me.hWnd, 0)
If hMenu Then
'Obtain the number of items in the menu
menuItemCount = GetMenuItemCount(hMenu)
'Remove the system menu Close menu item.
'The menu item is 0-based, so the last
'item on the menu is menuItemCount - 1
Call RemoveMenu(hMenu, menuItemCount - 1, _
MF_REMOVE Or MF_BYPOSITION)
'Remove the system menu separator line
Call RemoveMenu(hMenu, menuItemCount - 2, _
MF_REMOVE Or MF_BYPOSITION)
'Force a redraw of the menu. This
'refreshes the titlebar, dimming the X
Call DrawMenuBar(Me.hWnd)
End If

End Sub


