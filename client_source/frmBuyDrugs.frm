VERSION 5.00
Begin VB.Form frmBuyDrugs 
   BackColor       =   &H80000012&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Buy Drugs"
   ClientHeight    =   5580
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6000
   Icon            =   "frmBuyDrugs.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   372
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Left            =   3960
      TabIndex        =   10
      Top             =   5040
      Width           =   1695
   End
   Begin VB.CommandButton cmdBuyDrug 
      Caption         =   ">>> Buy >>>"
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
      Left            =   2280
      TabIndex        =   2
      Top             =   5040
      Width           =   1575
   End
   Begin VB.CommandButton cmdExit 
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
      Index           =   0
      Left            =   4200
      TabIndex        =   1
      Top             =   6840
      Width           =   1575
   End
   Begin VB.ListBox lstBuyDrug 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   4155
      Left            =   120
      TabIndex        =   0
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Line lineMain 
      BorderColor     =   &H0000FFFF&
      BorderWidth     =   2
      X1              =   168
      X2              =   376
      Y1              =   256
      Y2              =   256
   End
   Begin VB.Label lblCash 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      Caption         =   "Select Item"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   4200
      TabIndex        =   9
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblCashDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      ForeColor       =   &H000000FF&
      Height          =   255
      Left            =   2520
      TabIndex        =   8
      Top             =   3960
      Width           =   1455
   End
   Begin VB.Label lblMessage 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ok dude, hurry up and get your shit before the cops come around..."
      ForeColor       =   &H00FFFFFF&
      Height          =   615
      Left            =   2520
      TabIndex        =   7
      Top             =   4320
      Width           =   3135
   End
   Begin VB.Label lblPrice 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   6
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblDrug 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Select Item"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   5
      Top             =   3480
      Width           =   1455
   End
   Begin VB.Label lblPriceDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Price"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   4200
      TabIndex        =   4
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Label lblDrugDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Drug"
      ForeColor       =   &H0000FFFF&
      Height          =   255
      Left            =   2520
      TabIndex        =   3
      Top             =   3120
      Width           =   1455
   End
   Begin VB.Shape shpMain 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderWidth     =   2
      Height          =   2055
      Left            =   2400
      Top             =   3000
      Width           =   3375
   End
   Begin VB.Image imgDeal 
      BorderStyle     =   1  'Fixed Single
      Height          =   960
      Left            =   120
      Picture         =   "frmBuyDrugs.frx":030A
      Top             =   120
      Width           =   5535
   End
End
Attribute VB_Name = "frmBuyDrugs"
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

Private Sub cmdBuyDrug_Click()

If frmBuyDrugs.lstBuyDrug.ListIndex < 0 Or _
   frmBuyDrugs.lstBuyDrug.ListIndex > 19 Then
   Exit Sub
End If

cmdBuyDrug.Enabled = False
frmMain.wsk.SendData Chr$(253) & Chr$(2) & lstBuyDrug.ListIndex & Chr$(0)
DoEvents
cmdForgetIt.SetFocus

End Sub




Private Sub cmdForgetIt_Click()

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


Private Sub lstBuyDrug_Click()

cmdBuyDrug.Enabled = False

End Sub


Private Sub lstBuyDrug_DblClick()

If frmBuyDrugs.lstBuyDrug.Text = "<Empty>" Then
   cmdBuyDrug.Enabled = False
ElseIf frmBuyDrugs.lstBuyDrug.Text <> "<Empty>" Then
   frmMain.wsk.SendData Chr$(254) & Chr$(7) & lstBuyDrug.ListIndex & Chr$(0)
   DoEvents
   cmdBuyDrug.Enabled = True
End If

End Sub
