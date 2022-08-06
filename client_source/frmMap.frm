VERSION 5.00
Begin VB.Form frmMap 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "City Map"
   ClientHeight    =   4170
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5400
   Icon            =   "frmMap.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   278
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1455
      Left            =   120
      Picture         =   "frmMap.frx":058A
      ScaleHeight     =   1455
      ScaleWidth      =   5175
      TabIndex        =   2
      Top             =   120
      Width           =   5175
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   ">>>  Ok  <<<"
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
      TabIndex        =   1
      Top             =   3720
      Width           =   5175
   End
   Begin VB.ListBox lstCity 
      BackColor       =   &H00000000&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   1740
      Left            =   120
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1920
      Width           =   5190
   End
End
Attribute VB_Name = "frmMap"
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

Private Sub cmdOk_Click()

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


Private Sub lstCity_GotFocus()

cmdOk.SetFocus

End Sub


