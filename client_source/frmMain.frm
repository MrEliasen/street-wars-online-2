VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "mswinsck.ocx"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H80000010&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Street Wars Online II-Death Vendetta-UNLEASHED"
   ClientHeight    =   8220
   ClientLeft      =   1815
   ClientTop       =   1905
   ClientWidth     =   11910
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   OLEDropMode     =   1  'Manual
   ScaleHeight     =   548
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   794
   WhatsThisButton =   -1  'True
   WhatsThisHelp   =   -1  'True
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   1095
      Left            =   0
      Picture         =   "frmMain.frx":030A
      ScaleHeight     =   1095
      ScaleWidth      =   9375
      TabIndex        =   38
      Top             =   0
      Width           =   9375
   End
   Begin VB.CommandButton cmdSkills 
      Caption         =   "Skills"
      Height          =   495
      Left            =   3720
      TabIndex        =   37
      Top             =   7800
      Width           =   1095
   End
   Begin VB.TextBox txtOutput2 
      BackColor       =   &H80000008&
      ForeColor       =   &H80000005&
      Height          =   1215
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   34
      TabStop         =   0   'False
      Top             =   2880
      Width           =   9375
   End
   Begin VB.Timer tmrMain 
      Enabled         =   0   'False
      Left            =   480
      Top             =   0
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   120
      Top             =   0
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
   End
   Begin VB.ListBox lstInventory 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   9360
      TabIndex        =   9
      Top             =   3600
      Width           =   2535
   End
   Begin VB.ListBox lstUsers 
      BackColor       =   &H00000000&
      ForeColor       =   &H00E0E0E0&
      Height          =   2790
      Left            =   9360
      Sorted          =   -1  'True
      TabIndex        =   7
      Top             =   360
      Width           =   2535
   End
   Begin VB.CommandButton cmdTravel 
      Caption         =   "Travel"
      Height          =   435
      Left            =   2520
      TabIndex        =   5
      Top             =   7800
      Width           =   1140
   End
   Begin VB.CommandButton cmdPawnShop 
      Caption         =   "Pawn Shop"
      Height          =   435
      Left            =   1320
      TabIndex        =   4
      Top             =   7800
      Width           =   1140
   End
   Begin VB.CommandButton cmdMap 
      BackColor       =   &H00E0E0E0&
      Caption         =   "Map"
      Height          =   435
      Left            =   135
      TabIndex        =   3
      Top             =   7800
      Width           =   1140
   End
   Begin VB.TextBox txtInput 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   285
      Left            =   0
      MaxLength       =   200
      TabIndex        =   0
      Top             =   7440
      Width           =   9375
   End
   Begin VB.TextBox txtOutput 
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
      Height          =   3375
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   2
      TabStop         =   0   'False
      Top             =   4080
      Width           =   9375
   End
   Begin VB.TextBox txtNews 
      BackColor       =   &H00000000&
      ForeColor       =   &H00FFFFFF&
      Height          =   645
      Left            =   0
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      TabIndex        =   1
      TabStop         =   0   'False
      Top             =   2280
      Width           =   9375
   End
   Begin VB.Label lblrankpoints 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6480
      TabIndex        =   40
      Top             =   1920
      Width           =   2775
   End
   Begin VB.Label lblpoints 
      BackColor       =   &H00808080&
      Caption         =   "Points:"
      ForeColor       =   &H80000014&
      Height          =   255
      Left            =   5760
      TabIndex        =   39
      Top             =   1920
      Width           =   615
   End
   Begin VB.Line Line2 
      BorderWidth     =   3
      X1              =   408
      X2              =   560
      Y1              =   464
      Y2              =   464
   End
   Begin VB.Label lblLastSell 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   36
      Top             =   7800
      Width           =   2055
   End
   Begin VB.Label lblLastSellDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Sell Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   9600
      TabIndex        =   35
      Top             =   7440
      Width           =   2055
   End
   Begin VB.Label lblLastBuy 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   9600
      TabIndex        =   33
      Top             =   7080
      Width           =   2055
   End
   Begin VB.Label lblLastBuyDisplay 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BackStyle       =   0  'Transparent
      Caption         =   "Last Buy Location"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C0C0C0&
      Height          =   255
      Left            =   9600
      TabIndex        =   32
      Top             =   6720
      Width           =   2055
   End
   Begin VB.Shape shpNavigation 
      BackColor       =   &H00000000&
      BackStyle       =   1  'Opaque
      BorderColor     =   &H00FFFFFF&
      Height          =   1860
      Left            =   9360
      Top             =   6360
      Width           =   2535
   End
   Begin VB.Label lblAmmo 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6480
      TabIndex        =   31
      Top             =   1680
      Width           =   2775
   End
   Begin VB.Label lblArmor 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6480
      TabIndex        =   30
      Top             =   1440
      Width           =   2775
   End
   Begin VB.Label lblAmmoDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Ammo:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   29
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblArmorDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Armor:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   28
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblWeapon 
      Alignment       =   2  'Center
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   6480
      TabIndex        =   27
      Top             =   1200
      Width           =   2775
   End
   Begin VB.Label lblWeaponDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Weapon:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   5760
      TabIndex        =   26
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblKills 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   25
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label lblRank 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   24
      Top             =   1680
      Width           =   2055
   End
   Begin VB.Label lblLocation 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   23
      Top             =   1440
      Width           =   2055
   End
   Begin VB.Label lblHomeTown 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   3720
      TabIndex        =   22
      Top             =   1200
      Width           =   2055
   End
   Begin VB.Label lblKillsDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Kills:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   21
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblRankDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Rank:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   20
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblLocationDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Location:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   19
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblHomeTownDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Home Town:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2640
      TabIndex        =   18
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblBank 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   17
      Top             =   1920
      Width           =   1695
   End
   Begin VB.Label lblCash 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   16
      Top             =   1680
      Width           =   1695
   End
   Begin VB.Label lblHealth 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   15
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label lblName 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "----"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   960
      TabIndex        =   14
      Top             =   1200
      Width           =   1695
   End
   Begin VB.Label lblBankDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Bank:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   13
      Top             =   1920
      Width           =   1095
   End
   Begin VB.Label lblCashDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Cash:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   1680
      Width           =   1095
   End
   Begin VB.Label lblHealthDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Health:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label lblNameDisplay 
      BackColor       =   &H00800000&
      BackStyle       =   0  'Transparent
      Caption         =   "Name:"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   1200
      Width           =   1095
   End
   Begin VB.Label lblInventory 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Inventory"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   3240
      Width           =   2535
   End
   Begin VB.Label lblDealers 
      Alignment       =   2  'Center
      BackColor       =   &H00000000&
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Dealers Online"
      ForeColor       =   &H00E0E0E0&
      Height          =   255
      Left            =   9360
      TabIndex        =   6
      Top             =   0
      Width           =   2535
   End
   Begin VB.Shape shpMain 
      BackColor       =   &H00808080&
      BackStyle       =   1  'Opaque
      BorderWidth     =   3
      Height          =   1335
      Left            =   0
      Top             =   960
      Width           =   9360
   End
   Begin VB.Menu mnuFileConnect 
      Caption         =   "&Connect"
   End
   Begin VB.Menu mnuFileDisconnect 
      Caption         =   "&Disconnect"
      Visible         =   0   'False
   End
   Begin VB.Menu mnuCommands 
      Caption         =   "&Commands"
   End
   Begin VB.Menu mnuRanks 
      Caption         =   "&Ranks"
   End
   Begin VB.Menu mnuNpcRanks 
      Caption         =   "&Npc Ranks"
   End
   Begin VB.Menu mnuVote 
      Caption         =   "&Vote For DVU"
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuHelpGuide 
         Caption         =   "Street Wars Online II Help Guide"
      End
      Begin VB.Menu mnuHelpVisitSite 
         Caption         =   "Visit Street Wars Online II Website"
      End
   End
   Begin VB.Menu mnuInventory 
      Caption         =   "Inventory"
      Visible         =   0   'False
      Begin VB.Menu mnuInventoryEquip 
         Caption         =   "Equip"
      End
      Begin VB.Menu mnuInventoryUnequip 
         Caption         =   "Un-Equip"
      End
      Begin VB.Menu mnuInventoryExamine 
         Caption         =   "Examine"
      End
      Begin VB.Menu mnuInventoryUse 
         Caption         =   "Use"
      End
      Begin VB.Menu mnuInventoryDrop 
         Caption         =   "Drop"
      End
   End
   Begin VB.Menu mnuFileExit 
      Caption         =   "&Exit"
   End
End
Attribute VB_Name = "frmMain"
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

Private Sub cmdMap_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(253) & Chr$(5) & Chr$(0)
DoEvents

End Sub


Private Sub cmdPawnShop_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(254) & Chr$(2) & Chr$(0)
DoEvents

End Sub
Private Sub cmdSkills_Click()

wsk.SendData Trim$("skills") & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub cmdTravel_Click()
On Error Resume Next

frmMain.wsk.SendData Chr$(255) & Chr$(6) & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub




Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
On Error Resume Next

If MoveDelay = True Then
   Exit Sub
End If

If KeyUsed = False Then
If KeyCode = vbKeyUp Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyRight Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyDown Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyLeft Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF1 Then
   frmMain.wsk.SendData "punch" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF2 Then
   frmMain.wsk.SendData "strike" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF3 Then
   frmMain.wsk.SendData "fire" & Chr$(0)
   DoEvents
   KeyUsed = True
ElseIf KeyCode = vbKeyF4 Then
   frmMain.wsk.SendData "look" & Chr$(0)
   DoEvents
   KeyUsed = True
End If
End If

End Sub
Private Sub Form_KeyUp(KeyCode As Integer, Shift As Integer)

If KeyCode = vbKeyUp Then
   KeyUsed = False
ElseIf KeyCode = vbKeyRight Then
   KeyUsed = False
ElseIf KeyCode = vbKeyDown Then
   KeyUsed = False
ElseIf KeyCode = vbKeyLeft Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF1 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF2 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF3 Then
   KeyUsed = False
ElseIf KeyCode = vbKeyF4 Then
   KeyUsed = False
End If

End Sub
Private Sub Form_Load()
Dim a As Integer 'Counter

'Setup initial inventory slots
For a = 0 To 19
  lstInventory.AddItem "<Empty>"
Next a

txtNews.BackColor = vbBlack
txtNews.ForeColor = vbWhite

End Sub
Private Sub imgEast_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "e" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgNorth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "n" & Chr$(0)
   DoEvents
End If


End Sub

Private Sub imgSouth_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "s" & Chr$(0)
   DoEvents
End If

End Sub

Private Sub imgWest_Click()

If frmMain.wsk.State <> sckClosed Then
   frmMain.wsk.SendData "w" & Chr$(0)
   DoEvents
End If

End Sub


Private Sub Form_Unload(Cancel As Integer)

wsk.Close

End Sub

Private Sub lstInventory_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

If Button = 2 Then
   PopupMenu mnuInventory
End If

End Sub

Private Sub mnuCommands_Click()

frmMain.Enabled = True
frmCommands.Show
DoEvents

End Sub

Private Sub mnuFileConnect_Click()
On Error Resume Next
Dim iServ As String

iServ = "80.161.55.60"

If iServ = "" Then Exit Sub

'Disable menus
frmMain.mnuFileConnect.Enabled = False
frmMain.mnuFileExit.Enabled = False
frmMain.mnuFileConnect.Visible = False
frmMain.mnuFileDisconnect.Visible = True

'Connect to the server
With wsk
  .Close
  .Protocol = sckTCPProtocol
  .RemotePort = ServerPort
  .RemoteHost = iServ
  .Connect
End With

Call ShowText("Connecting to the Street Wars Online II central server, please stand by...If it does no connect we may be lagging due to the new day problem..." & vbCrLf & vbCrLf)

End Sub
Private Sub mnuFileDisconnect_Click()
'Disconnect and enable menus

wsk.Close
frmMain.mnuFileConnect.Enabled = True
frmMain.mnuFileExit.Enabled = True
frmMain.mnuFileDisconnect.Visible = False
frmMain.mnuFileConnect.Visible = True

End Sub

Private Sub mnuFileExit_Click()

   'Close winsock and shut down the game
   wsk.Close
   Unload Me
   End

End Sub

Private Sub mnuHelpGuide_Click()

Call OpenLocation("http://streetwars.8m.com/street_wars_online_ii_online_hel.htm", SW_SHOWNORMAL)

End Sub

Private Sub mnuHelpVisitSite_Click()

Call OpenLocation("http://deathvendetta.2ya.com", SW_SHOWNORMAL)

End Sub
Private Sub mnuInventoryDrop_Click()

frmMain.wsk.SendData Chr$(7) & lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryEquip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(3) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryExamine_Click()

'Examine the item
frmMain.wsk.SendData Chr$(255) & Chr$(2) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUnequip_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(4) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuInventoryUse_Click()

frmMain.wsk.SendData Chr$(255) & Chr$(5) & frmMain.lstInventory.ListIndex & Chr$(0)
DoEvents
frmMain.txtInput.SetFocus

End Sub

Private Sub mnuNpcRanks_Click()

frmMain.Enabled = True
frmNpcRanks.Show
DoEvents

End Sub

Private Sub mnuRanks_Click()

frmMain.Enabled = True
frmRanks.Show
DoEvents

End Sub

Private Sub mnuVote_Click()

  Call OpenLocation("http://www.topwebgames.com/in.asp?id=864", SW_SHOWNORMAL)

End Sub

Private Sub txtInput_KeyPress(KeyAscii As Integer)
On Error GoTo Failed 'Error Handler

'Send textbox text to the server
If (KeyAscii = 13) And (txtInput.Text <> "") Then
  KeyAscii = 0
  If wsk.State <> sckClosed Then
     If InputDelay = True Then
        Exit Sub
     End If
    wsk.SendData Trim$(txtInput.Text) & Chr$(0)
    DoEvents
    txtInput.Text = ""
  End If
End If
Exit Sub

'If an error occurs,  close the socket and reset
'everything
Failed:
wsk.Close
With txtOutput
  .Text = .Text & "An error has occured while sending data to the server, your connection has been reset." & vbCrLf & vbCrLf
  .SelStart = Len(.Text)
End With
txtInput.Text = ""
tmrMain.Enabled = False
mnuFileConnect.Enabled = True
mnuFileExit.Enabled = True

End Sub
Private Sub txtNews_GotFocus()
  'Don't allow textbox to have focus
  txtInput.SetFocus
End Sub


Private Sub txtOutput_GotFocus()
  'Don't allow textbox to get focus
  txtInput.SetFocus
End Sub
Private Sub wsk_Connect()

frmMain.wsk.SendData ClientVer & Chr$(0)
DoEvents

End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
Dim a As Integer 'Counter
Dim Msg As String 'String to hold data off the wire
Dim SplitMsg() As String 'String array to parse data

'Pull data off the wire
wsk.GetData Msg, vbString

'Split the string array
SplitMsg = Split(Msg, Chr$(0))

'Loop through data and process accordingly
For a = 0 To UBound(SplitMsg) - 1
   
   Select Case Left$(SplitMsg(a), 2)
      Case Chr$(255) & Chr$(2)
         Call TravelMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(3)
         Call PawnShopMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(4)
         Call UpdateCashRank(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(5)
         Call PawnShopItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(6)
         Call PawnShopPlayerInventoryUpdate(Mid$(SplitMsg(a), 3))
      Case Chr$(255) & Chr$(7)
         Call UpdateGeneralInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(2)
         Call UpdateGearInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(3)
         Call UpdatePlayerList(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(4)
         Call BuyDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(5)
         Call CloseDrugDealMenu
      Case Chr$(254) & Chr$(6)
         Call DrugDealItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(254) & Chr$(7)
         Call DrugDealMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(2)
         Call UpdateDealerInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(3)
         Call SellDrugMenu(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(4)
         Call CloseDruggieMenu
      Case Chr$(253) & Chr$(5)
         Call DruggieMenuMessage(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(6)
         Call DruggieMenuItemInfo(Mid$(SplitMsg(a), 3))
      Case Chr$(253) & Chr$(7)
         Call ReUpdateDruggieInventory(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(2)
         Call ShowMap(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(3)
         Call UpdateNews(Mid$(SplitMsg(a), 3))
      Case Chr$(252) & Chr$(4)
         frmMain.lblLastBuy.Caption = Mid$(SplitMsg(a), 3)
      Case Chr$(252) & Chr$(5)
         frmMain.lblLastSell.Caption = Mid$(SplitMsg(a), 3)
   End Select
   
   Select Case Left$(SplitMsg(a), 1)
      Case Chr$(2)
         Call ShowText(Mid$(SplitMsg(a), 2))
      Case Chr$(3)
         Call NewAccount
      Case Chr$(4)
         Call DupeName
      Case Chr$(5)
         Call AccountCreated
      Case Chr$(6)
         Call UpdateFullInventory(Mid$(SplitMsg(a), 2))
      Case Chr$(7)
         Call UpdateSingleItem(Mid$(SplitMsg(a), 2))
      Case Chr$(8)
         Call ShowText2(Mid$(SplitMsg(a), 2))
   End Select
Next a

End Sub
