VERSION 5.00
Begin VB.Form frmNewAccount 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "New Account"
   ClientHeight    =   4935
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5430
   ControlBox      =   0   'False
   Icon            =   "frmNewAccount.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   329
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   362
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      Height          =   1695
      Left            =   0
      Picture         =   "frmNewAccount.frx":08CA
      ScaleHeight     =   1635
      ScaleWidth      =   5355
      TabIndex        =   13
      Top             =   0
      Width           =   5415
   End
   Begin VB.ComboBox cboCity 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   315
      Left            =   2400
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   2820
      Width           =   2055
   End
   Begin VB.TextBox txtPassTwo 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   2
      Top             =   2460
      Width           =   2055
   End
   Begin VB.TextBox txtPassOne 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      IMEMode         =   3  'DISABLE
      Left            =   2400
      MaxLength       =   10
      PasswordChar    =   "*"
      TabIndex        =   1
      Top             =   2100
      Width           =   2055
   End
   Begin VB.TextBox txtName 
      BackColor       =   &H00000000&
      ForeColor       =   &H0000FF00&
      Height          =   285
      Left            =   2400
      MaxLength       =   10
      TabIndex        =   0
      Top             =   1740
      Width           =   2055
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "Exit"
      Height          =   375
      Left            =   2400
      TabIndex        =   5
      Top             =   4440
      Width           =   1815
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "Create"
      Height          =   375
      Left            =   240
      TabIndex        =   4
      Top             =   4440
      Width           =   1815
   End
   Begin VB.Label lblIP 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H0000FF00&
      Height          =   255
      Left            =   2040
      TabIndex        =   12
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblIPAddress 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Your IP Address:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   11
      Top             =   3240
      Width           =   1695
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   $"frmNewAccount.frx":1E66C
      ForeColor       =   &H000000FF&
      Height          =   615
      Left            =   240
      TabIndex        =   10
      Top             =   3600
      Width           =   3495
   End
   Begin VB.Label lblCity 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "City Claimed:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   2880
      Width           =   1695
   End
   Begin VB.Label lblPasswordConfirm 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Confirm Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.Label lblPassword 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Password:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   7
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label lblName 
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1800
      Width           =   1695
   End
   Begin VB.Shape shpNewAccount 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2715
      Left            =   120
      Top             =   1620
      Width           =   3810
   End
End
Attribute VB_Name = "frmNewAccount"
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

Private Sub cmdCreate_Click()

'Make sure no info fields are blank
If txtName.Text = "" Or _
   txtPassOne.Text = "" Or _
   txtPassTwo.Text = "" Or _
   cboCity.Text = "" Then
     lblMessage.Caption = "You must complete all fields before you can continue."
     Exit Sub
End If
   
'Make sure name and password are at least four
'characters in lenght
If Len(Trim$(txtName.Text)) < 4 Or _
   Len(Trim$(txtPassOne.Text)) < 4 Or _
   Len(Trim$(txtPassTwo.Text)) < 4 Then
     lblMessage.Caption = "Your name and password must be four or more characters in length to continue."
     Exit Sub
End If

'Check passwords for a match
If txtPassOne.Text <> txtPassTwo.Text Then
   lblMessage.Caption = "Your passwords do not match."
   Exit Sub
End If

txtName.Enabled = False
txtPassOne.Enabled = False
txtPassTwo.Enabled = False
cboCity.Enabled = False
cmdCreate.Enabled = False

frmMain.wsk.SendData Trim$(txtName.Text) & Chr$(1) & Trim$(txtPassOne) & Chr$(1) & Trim$(cboCity.Text) & Chr$(1) & Chr$(0)
DoEvents

End Sub
Private Sub cmdExit_Click()

'unload the new account form and enable the main form
'and enable the disabled menus
frmMain.wsk.Close
Call ShowText("Your connection to the server has been reset." & vbCrLf & vbCrLf)
frmMain.mnuFileConnect.Enabled = True
frmMain.mnuFileExit.Enabled = True
Unload Me
frmMain.Enabled = True
frmMain.txtInput.SetFocus

End Sub

Private Sub Form_Load()

'Display users IP address
lblIP.Caption = frmMain.wsk.LocalIP

'Add citys to the combo box
cboCity.AddItem ("New York"), 0
cboCity.AddItem ("Los Angeles"), 1
cboCity.AddItem ("Chicago"), 2
cboCity.AddItem ("Houston"), 3
cboCity.AddItem ("Miami"), 4
cboCity.AddItem ("New Jersey"), 5
cboCity.AddItem ("Washington DC"), 6

End Sub
Private Sub Image1_Click()

End Sub


