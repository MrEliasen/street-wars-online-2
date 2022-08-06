Attribute VB_Name = "Commands"
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




Public Sub DoCommand(Index As Integer, msg As String)
On Error Resume Next

If LCase$(msg) = "look" Then
   Call ShowCity(Index)
   Exit Sub
ElseIf LCase$(msg) = "dick" Then
   Call DickSmack(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "stats" Then
   Call stats(Index, Trim$(Mid$(msg, 6)))
   Exit Sub
   ElseIf LCase$(Left$(msg, 4)) = "rank" Then
   Call Rankage(Index, Trim$(Mid$(msg, 5)))
   Exit Sub
ElseIf LCase$(msg) = "n" Then
   Call North(Index)
   Exit Sub
ElseIf LCase$(msg) = "e" Then
   Call East(Index)
   Exit Sub
ElseIf LCase$(msg) = "s" Then
   Call South(Index)
   Exit Sub
ElseIf LCase$(msg) = "w" Then
   Call West(Index)
   Exit Sub
ElseIf Left$(msg, 1) = ";" Then
   Call SendChat(Index, Trim$(Mid$(msg, 2)))
   Exit Sub
ElseIf Left$(msg, 1) = "'" Then
   Call PrivateChat(Index, Trim$(Mid$(msg, 2)))
   Exit Sub
ElseIf Left$(msg, 3) = "say" Then
   Call SquareChat(Index, Trim$(Mid$(msg, 4)))
   Exit Sub
ElseIf LCase$(msg) = "mute" Then
   Call Mute(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 8)) = "/additem" Then
   Call AddItemGM(Index, Trim$(Mid$(msg, 9)))
   Exit Sub
ElseIf LCase$(msg) = "/listitems" Then
   Call ListItemGM(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 3)) = "get" Then
   Call GetItem(Index, Trim$(Mid$(msg, 4)))
   Exit Sub
ElseIf LCase$(Left$(msg, 4)) = "nuke" Then
   Call NukeCity(Index, Trim$(Mid$(msg, 5)))
   Exit Sub
ElseIf Left$(msg, 1) = Chr$(7) Then
   Call DropItem(Index, Mid$(msg, 2))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(255) & Chr$(2) Then
   Call ExamineItem(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(255) & Chr$(3) Then
   Call EquipItem(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(255) & Chr$(4) Then
   Call UnEquip(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(255) & Chr$(5) Then
   Call UseItem(Index, Mid$(msg, 3))
   Exit Sub
ElseIf msg = Chr$(255) & Chr$(6) Then
   Call TravelMenu(Index)
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(255) & Chr$(7) Then
   Call Travel(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(254) & Chr$(2) Then
   Call PawnShopMenu(Index)
   Exit Sub
ElseIf Left(msg, 2) = Chr$(254) & Chr$(3) Then
   Call PlayerItemInfo(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left(msg, 2) = Chr$(254) & Chr$(4) Then
   Call ShopItemInfo(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left(msg, 2) = Chr$(254) & Chr$(5) Then
   Call BuyItem(Index, Mid$(msg, 3))
   Exit Sub
ElseIf Left(msg, 2) = Chr$(254) & Chr$(6) Then
   Call SellItem(Index, Mid$(msg, 3))
   Exit Sub
ElseIf LCase$(Left$(msg, 7)) = "/addmob" Then
   Call AddNpcGM(Index, Trim$(Mid$(msg, 8)))
   Exit Sub
ElseIf LCase$(Left$(msg, 3)) = "buy" Then
   Call BuyDrugMenu(Index, Trim$(Mid$(msg, 4)))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(254) & Chr$(7) Then
   Call DrugDealItemInfo(Index, Trim$(Mid$(msg, 3)))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(253) & Chr$(2) Then
   Call BuyNpcDrug(Index, Mid$(msg, 3))
   Exit Sub
ElseIf LCase(Left$(msg, 4)) = "sell" Then
   Call DruggieMenu(Index, Trim$(Mid$(msg, 5)))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(253) & Chr$(3) Then
   Call DruggieItemInfo(Index, Trim$(Mid$(msg, 3)))
   Exit Sub
ElseIf LCase$(msg) = "/listnpcs" Then
   Call ListNPCs(Index)
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(253) & Chr$(4) Then
   Call SellDruggieItem(Index, Trim$(Mid$(msg, 3)))
   Exit Sub
ElseIf Left$(msg, 2) = Chr$(253) & Chr$(5) Then
   Call SendMap(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 3)) = "aim" Then
   Call Aim(Index, Trim$(Mid$(msg, 4)))
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "/goto" Then
   Call GotoPlayer(Index, Trim$(Mid$(msg, 6)))
   Exit Sub

ElseIf LCase$(msg) = "punch" Then
   Call Punch(Index)
   Exit Sub
ElseIf LCase$(msg) = "fire" Then
   Call Fire(Index)
   Exit Sub
ElseIf LCase$(msg) = "strike" Then
   Call Strike(Index)
   Exit Sub
ElseIf LCase$(msg) = "hide" Then
   Call Hide(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 4)) = "flee" Then
   Call Flee(Index, Trim$(Mid$(msg, 5)))
   Exit Sub
ElseIf LCase$(msg) = "healinfo" Then
   Call HealInfo(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 6)) = "healme" Then
   Call HealMe(Index, Trim$(Mid$(msg, 7)))
   Exit Sub
ElseIf LCase$(msg) = "skills" Then
   Call ShowSkills(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 7)) = "deposit" Then
   Call Deposit(Index, Trim$(Mid$(msg, 8)))
   Exit Sub
ElseIf LCase$(Left$(msg, 8)) = "withdraw" Then
   Call Withdraw(Index, Trim$(Mid$(msg, 9)))
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "track" Then
   Call TrackPlayer(Index, Trim$(Mid$(msg, 6)))
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "snipe" Then
   Call SnipePlayer(Index, Trim$(Mid$(msg, 6)))
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "snoop" Then
   Call SnoopPlayer(Index, Trim$(Mid$(msg, 6)))
   Exit Sub
ElseIf LCase$(Left$(msg, 3)) = "rob" Then
   Call RobBank(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 6)) = "gamble" Then
   Call Gamble(Index, Trim$(Mid$(msg, 7)))
   Exit Sub
ElseIf LCase$(Left$(msg, 6)) = "search" Then
   Call Search(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "steal" Then
   Call StealPlayer(Index, Trim$(Mid$(msg, 6)))
   Exit Sub
ElseIf LCase$(Left$(msg, 4)) = "/ban" Then
   Call BanUser(Index, Trim$(Mid$(msg, 5)))
   Exit Sub
ElseIf LCase$(Left$(msg, 6)) = "/getip" Then
    Call GetIP(Index, Trim$(Mid$(msg, 7)))
    Exit Sub
ElseIf Left$(msg, 1) = "*" Then
   Call News(Index, Trim$(Mid$(msg, 2)))
   Exit Sub
ElseIf LCase$(Left$(msg, 5)) = "/kick" Then
    Call KickUser(Index, Trim$(Mid$(msg, 6)))
    Exit Sub
ElseIf LCase$(Left$(msg, 1)) = "@" Then
    Call Gang(Index, Trim$(Mid$(msg, 2)))
    Exit Sub
ElseIf LCase$(Left$(msg, 1)) = "~" Then
    Call AdminName(Index, Trim$(Mid$(msg, 2)))
    Exit Sub
ElseIf LCase$(Left$(msg, 6)) = "/force" Then
    Call ForceUser(Index, Trim$(Mid$(msg, 7)))
    Exit Sub
ElseIf LCase$(msg) = "drink" Then
   Call drink(Index)
   Exit Sub
ElseIf LCase$(Left$(msg, 7)) = "newpass" Then
   Call ChangePassword(Index, Trim$(Mid$(msg, 8)))
   Exit Sub
ElseIf Left$(msg, 1) = "$" Then
   Call Transfer(Index, Trim$(Mid$(msg, 2)))
   Exit Sub
ElseIf Left$(msg, 6) = "bounty" Then
   Call Bounty(Index, Trim$(Mid$(msg, 7)))
   Exit Sub
Else
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

End Sub
Public Sub North(Index As Integer)
On Error Resume Next

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player north
If City(User(Index).Location).North <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the north." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).North
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from south." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).North = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the North." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub East(Index As Integer)
On Error Resume Next

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player east
If City(User(Index).Location).East <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the east." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).East
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from west." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).East = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the East." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub South(Index As Integer)
On Error Resume Next

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player south
If City(User(Index).Location).South <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the south." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).South
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from north." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).South = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the South." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub West(Index As Integer)
On Error Resume Next

'Check to see if player is hiding
Call NoHiding(Index)

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

'Move the player west
If City(User(Index).Location).West <> -1 Then
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll off to the west." & vbCrLf & vbCrLf & Chr$(0))
   User(Index).Location = City(User(Index).Location).West
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " stroll in from east." & vbCrLf & vbCrLf & Chr$(0))
   Call ShowCity(Index)
ElseIf City(User(Index).Location).West = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You find no exit leading to the West." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub
Public Sub ShowWatchers(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter

'Show message to all players exept Index Player
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = User(Index).Location And _
      Index <> a Then
         frmMain.wsk(a).SendData msg
         DoEvents
   End If
Next a

End Sub

Public Sub SendChat(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter

Call ChatLog(Index, msg)

'Dont allow player to send a global msg with mute on
If User(Index).Mute = True Then
   frmMain.wsk(Index).SendData Chr$(2) & "You cannot transmit a global message while you're in mute mode." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

'Display global yell message
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Mute = False Then
         frmMain.wsk(a).SendData Chr$(8) & vbCrLf & "[" & User(Index).UName & "] - " & msg & Chr$(0)
         DoEvents
   End If
Next a

End Sub


Public Sub Mute(Index As Integer)
On Error Resume Next

If User(Index).Mute = False Then
   User(Index).Mute = True
   frmMain.wsk(Index).SendData Chr$(2) & "You have selected to mute all global messages.  You can still transmit and recieve private messages while in mute mode." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
ElseIf User(Index).Mute = True Then
   User(Index).Mute = False
   frmMain.wsk(Index).SendData Chr$(2) & "You have selected to recieve all global messages." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub NoHiding(Index As Integer)
On Error Resume Next

'Check to see if the player is hiding
If User(Index).IsHiding = True Then
   User(Index).IsHiding = False
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " the " & User(Index).Rank & " emerge from the shadows." & vbCrLf & vbCrLf & Chr$(0))
   frmMain.wsk(Index).SendData Chr$(2) & "You are no longer in hiding." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub AddItemGM(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer, b As Integer 'Counters

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To UBound(ItemDB)
   If LCase$(msg) = LCase$(Left$(ItemDB(a).IName, Len(msg))) Then
      For b = 0 To UBound(City(User(Index).Location).CItem)
         If City(User(Index).Location).CItem(b) = -1 Then
            ReDim Preserve Item(UBound(Item) + 1)
            Item(UBound(Item)) = ItemDB(a)
            Item(UBound(Item)).ItemGUID = "" 'city(user(index).Location).CityGUID
            Item(UBound(Item)).ILocation = User(Index).Location
            Item(UBound(Item)).Decay = GetTickCount()
            City(User(Index).Location).CItem(b) = UBound(Item)
            frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
            Exit Sub
         ElseIf b = UBound(City(User(Index).Location).CItem) Then
            With City(User(Index).Location)
            ReDim Preserve .CItem(UBound(.CItem) + 1)
            ReDim Preserve Item(UBound(Item) + 1)
            Item(UBound(Item)) = ItemDB(a)
            Item(UBound(Item)).ItemGUID = "" 'City(User(Index).Location).CityGUID
            Item(UBound(Item)).ILocation = User(Index).Location
            Item(UBound(Item)).Decay = GetTickCount()
            .CItem(UBound(.CItem)) = UBound(Item)
            End With
            frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(UBound(Item)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
            Exit Sub
         End If
      Next b
   End If
Next a

End Sub

Public Sub ListItemGM(Index As Integer)
On Error Resume Next
Dim a As Integer 'Counter
Dim msg As String 'String
msg = Chr$(2)

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

'List Items to GM's Only
For a = 0 To UBound(ItemDB)
   msg = msg & ItemDB(a).IName & "     "
Next a

msg = msg & vbCrLf & vbCrLf & Chr$(0)
frmMain.wsk(Index).SendData msg
DoEvents

End Sub

Public Function InventoryFull(Index As Integer) As Boolean
Dim a As Integer 'Counter

'Check to see if a players inventory is full
For a = 0 To 19
   If User(Index).Item(a) = -1 Then
      InventoryFull = False
      Exit Function
   ElseIf a = 19 Then
      InventoryFull = True
      frmMain.wsk(Index).SendData Chr$(2) & "You have no room in your inventory, try selling something." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Function
   End If
Next a

End Function

Public Sub GetItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter

'Check for full inventory
If InventoryFull(Index) = True Then
   Exit Sub
End If

'Pick the item up off the ground
For a = 0 To UBound(City(User(Index).Location).CItem)
   If City(User(Index).Location).CItem(a) <> -1 Then
   If LCase$(msg) = LCase$(Left$(Item(City(User(Index).Location).CItem(a)).IName, Len(msg))) Then
      For b = 0 To 19
         If User(Index).Item(b) = -1 Then
            User(Index).Item(b) = City(User(Index).Location).CItem(a)
            City(User(Index).Location).CItem(a) = -1
            Item(User(Index).Item(b)).ItemGUID = User(Index).UserGUID
            Item(User(Index).Item(b)).OnPlayer = True
            Item(User(Index).Item(b)).Decay = -1
            Item(User(Index).Item(b)).ILocation = -1
            frmMain.wsk(Index).SendData Chr$(2) & "You pick up a " & Item(User(Index).Item(b)).IName & " and put it in your pack." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " pick up a " & Item(User(Index).Item(b)).IName & "." & vbCrLf & vbCrLf & Chr$(0))
            Call UpdateSingleItem(Index, b)
            Exit Sub
         End If
      Next b
   End If
   End If
Next a
           
'Runs if no item in room matches get message
frmMain.wsk(Index).SendData Chr$(2) & "You can't pick up what isn't there." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

           
           
End Sub

Public Sub DropItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter

If IsNumeric(msg) = True Then
   b = msg
Else
   Exit Sub
End If

If b > 19 Or b < 0 Then
   Exit Sub
End If

If User(Index).Item(b) = -1 Then
   Exit Sub
End If

'Drop the item the players chooses
For a = 0 To UBound(City(User(Index).Location).CItem)
   If City(User(Index).Location).CItem(a) = -1 Then
      
      If User(Index).Item(b) = User(Index).Weapon Then
         User(Index).Weapon = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Armor Then
         User(Index).Armor = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Ammo Then
         User(Index).Ammo = -1
         Call UpdateGearInfo(Index)
      End If
      
      City(User(Index).Location).CItem(a) = User(Index).Item(b)
      Item(User(Index).Item(b)).OnPlayer = False
      Item(User(Index).Item(b)).Equip = False
      Item(User(Index).Item(b)).Decay = GetTickCount()
      Item(User(Index).Item(b)).ItemGUID = ""
      Item(User(Index).Item(b)).ILocation = User(Index).Location
      User(Index).Item(b) = -1
      frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(City(User(Index).Location).CItem(a)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(City(User(Index).Location).CItem(a)).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
      Call UpdateSingleItem(Index, b)
      Exit Sub
   ElseIf a = UBound(City(User(Index).Location).CItem) Then
      With City(User(Index).Location)
      ReDim Preserve .CItem(UBound(.CItem) + 1)
      
      If User(Index).Item(b) = User(Index).Weapon Then
         User(Index).Weapon = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Armor Then
         User(Index).Armor = -1
         Call UpdateGearInfo(Index)
      ElseIf User(Index).Item(b) = User(Index).Ammo Then
         User(Index).Ammo = -1
         Call UpdateGearInfo(Index)
      End If

      .CItem(UBound(.CItem)) = User(Index).Item(b)
      Item(User(Index).Item(b)).OnPlayer = False
      Item(User(Index).Item(b)).Equip = False
      Item(User(Index).Item(b)).Decay = GetTickCount()
      Item(User(Index).Item(b)).ItemGUID = ""
      Item(User(Index).Item(b)).ILocation = User(Index).Location
      User(Index).Item(b) = -1
      frmMain.wsk(Index).SendData Chr$(2) & "You toss a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
      Call UpdateSingleItem(Index, b)
      End With
      Exit Sub
   End If
Next a

End Sub
Public Sub ExamineItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer

If IsNumeric(msg) = True Then
   a = msg
Else
   Exit Sub
End If

If a < 0 Or a > 19 Then
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & Item(User(Index).Item(a)).IDesc & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub EquipItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter

'Make sure the index is correct
If IsNumeric(msg) = True Then
   b = msg
Else
   Exit Sub
End If

If msg < 0 Or msg > 19 Then
   Exit Sub
End If

If User(Index).Item(b) = -1 Then
   Exit Sub
End If

If User(Index).Item(b) <> -1 Then
   If Item(User(Index).Item(b)).IType = C_Gun Or _
      Item(User(Index).Item(b)).IType = C_Armor Or _
      Item(User(Index).Item(b)).IType = C_Ammo Or _
      Item(User(Index).Item(b)).IType = C_Melee Then
      For a = 0 To 19
      If User(Index).Item(a) <> -1 Then
         If Item(User(Index).Item(a)).IType = _
            Item(User(Index).Item(b)).IType Then
            Item(User(Index).Item(a)).Equip = False
         End If
         'Make sure guns/melee are unquipted if opposite
         'No dual weapon weilding
            If Item(User(Index).Item(b)).IType = C_Melee And _
               Item(User(Index).Item(a)).IType = C_Gun And _
               Item(User(Index).Item(a)).Equip = True Then
                  Item(User(Index).Item(a)).Equip = False
            ElseIf Item(User(Index).Item(b)).IType = C_Gun And _
               Item(User(Index).Item(a)).IType = C_Melee And _
               Item(User(Index).Item(a)).Equip = True Then
                  Item(User(Index).Item(a)).Equip = False
            End If
      End If
      Next a
         If Item(User(Index).Item(b)).IType = C_Gun Then
            User(Index).Weapon = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Armor Then
            User(Index).Armor = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Ammo Then
            User(Index).Ammo = User(Index).Item(b)
         ElseIf Item(User(Index).Item(b)).IType = C_Melee Then
            User(Index).Weapon = User(Index).Item(b)
         End If
   Item(User(Index).Item(b)).Equip = True
   Call FullInventoryUpdate(Index)
   Call UpdateGearInfo(Index)
   End If
End If

End Sub
Public Sub UpdateSingleItem(Index As Integer, ItemNo As Integer)
On Error Resume Next
Dim msg As String
msg = Chr$(7) & ItemNo & Chr$(1)

   If User(Index).Item(ItemNo) = -1 Then
      msg = msg & "<Empty>"
   ElseIf User(Index).Item(ItemNo) <> -1 Then
            
      'Check to see if item is multiple
      If Item(User(Index).Item(ItemNo)).IType = C_Ammo And _
         Item(User(Index).Item(ItemNo)).Amount > 0 And _
         Item(User(Index).Item(ItemNo)).Multiple = True Then
            msg = msg & "(" & Item(User(Index).Item(ItemNo)).Amount & ") "
      End If
            
      'Check to see if the item is equipted
      If Item(User(Index).Item(ItemNo)).Equip = True Then
            msg = msg & "<E> "
      End If
      
      'Add Item Name
      msg = msg & "<" & Item(User(Index).Item(ItemNo)).IName & ">"
   End If

msg = msg & Chr$(0)
frmMain.wsk(Index).SendData msg
DoEvents

End Sub

Public Sub UnEquip(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer

If IsNumeric(msg) = True Then
   a = msg
Else
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).Equip = False Then
   Exit Sub
End If

'Remove Gear Index
If User(Index).Item(a) <> -1 Then
   If Item(User(Index).Item(a)).IType = C_Melee Or _
      Item(User(Index).Item(a)).IType = C_Gun Then
         User(Index).Weapon = -1
   ElseIf Item(User(Index).Item(a)).IType = C_Armor Then
         User(Index).Armor = -1
   ElseIf Item(User(Index).Item(a)).IType = C_Ammo Then
         User(Index).Ammo = -1
   End If
Call UpdateGearInfo(Index)
End If

'UnEquip users item
If User(Index).Item(a) <> -1 Then
   Item(User(Index).Item(a)).Equip = False
   Call UpdateSingleItem(Index, a)
End If

End Sub

Public Sub UseItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer

If IsNumeric(msg) = True Then
   a = msg
ElseIf IsNumeric(msg) = False Then
   Exit Sub
End If

If a < LBound(User(Index).Item) Or _
   a > UBound(User(Index).Item) Then
      Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

Select Case Item(User(Index).Item(a)).IType
   Case C_Phone
      Call UsePhone(Index)
      Exit Sub
   Case C_MedStick
      Call UseMedStick(Index, a)
      Exit Sub
    Case C_Dope
      Call UseDrugs(Index, a)
      Exit Sub
    Case C_Pager
      Call UsePager(Index)
      Exit Sub
    Case C_RedBull
        Call UseRedBull(Index, a)
        Exit Sub
    Case C_Steroids
        Call UseSteroids(Index, a)
        Exit Sub
    'Case C_Rank2
        'Call UseRank2(Index, a)
        'Exit Sub
    'Case C_Acc
        'Call UseAcc(Index, a)
        'Exit Sub
    Case C_Pepper
        Call UsePepper(Index, a)
        Exit Sub
    Case C_Horn
        Call UseHorn(Index)
        Exit Sub
    'Case C_Mine
        'Call UseMine(Index, a)
        'Exit Sub
    Case C_Grenade
        Call UseGrenade(Index, a)
        Exit Sub
    Case C_Adrenalin
        Call UseAdrenalin(Index, a)
        Exit Sub
    Case C_Admin
        Call Admin(Index, a)
        Exit Sub
End Select


frmMain.wsk(Index).SendData Chr$(2) & "You don't see any specific way you could use this item." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub TravelMenu(Index As Integer)
On Error Resume Next
Dim msg As String

'Check to see if the player is at an airport first
If City(User(Index).Location).AirPort = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find an Airport before you can travel anywhere." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

msg = Chr$(255) & Chr$(2) & User(Index).CurrTown & Chr$(1)
msg = msg & NY_Price & Chr$(1) & LA_Price & Chr$(1) & _
HO_Price & Chr$(1) & MI_Price & Chr$(1) & _
CH_Price & Chr$(1) & NJ_Price & Chr$(1) & _
SYD_Price & Chr$(1) & UK_Price & Chr$(1) & Chr$(0)

frmMain.wsk(Index).SendData msg
DoEvents

End Sub
Public Sub Travel(Index As Integer, msg As String)
On Error Resume Next

'Check to see if the player is in combat
If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Select Case LCase$(msg)
   
   'Fly to New York
   Case "new york"
      If User(Index).Cash < NY_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= NY_Price Then
         User(Index).Cash = User(Index).Cash - NY_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = NY_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to New York." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If
   
   'Fly to Los Angeles
   Case "los angeles"
      If User(Index).Cash < LA_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= LA_Price Then
         User(Index).Cash = User(Index).Cash - LA_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = LA_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Los Angeles." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Houston
   Case "houston"
      If User(Index).Cash < HO_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= HO_Price Then
         User(Index).Cash = User(Index).Cash - HO_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = HO_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Houston." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Miami
   Case "miami"
      If User(Index).Cash < MI_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= MI_Price Then
         User(Index).Cash = User(Index).Cash - MI_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = MI_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Miami." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to Chicago
   Case "chicago"
      If User(Index).Cash < CH_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= CH_Price Then
         User(Index).Cash = User(Index).Cash - CH_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = CH_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Chicago." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If

   'Fly to New Jersey
   Case "new jersey"
      If User(Index).Cash < NJ_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= NJ_Price Then
         User(Index).Cash = User(Index).Cash - NJ_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = NJ_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to New Jersey." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If
    'Fly to Sydney
   Case "sydney"
      If User(Index).Cash < SYD_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= SYD_Price Then
         User(Index).Cash = User(Index).Cash - SYD_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = SYD_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to Sydney." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If
   
      
    'Fly to London
   Case "london"
      If User(Index).Cash < UK_Price Then
         frmMain.wsk(Index).SendData Chr$(2) & "Start flapping your arms buddy, you don't have the cash to do fly anywhere." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf User(Index).Cash >= UK_Price Then
         User(Index).Cash = User(Index).Cash - UK_Price
         Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " jump on a plane and jet out of the city." & vbCrLf & vbCrLf & Chr$(0))
         User(Index).Location = UK_Location
         User(Index).CurrTown = City(User(Index).Location).CName
         Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " just arrived on a plane from another city." & vbCrLf & vbCrLf & Chr$(0))
         frmMain.wsk(Index).SendData Chr$(2) & "You jump on a plane and jet to London." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Call UpdateGeneralInfo(Index)
         Call ShowCity(Index)
         Exit Sub
      End If
        

'On Data Error Run This
frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Select

End Sub
Public Function PlayerIsTarget(Index As Integer) As Boolean
Dim a As Integer 'Counter

'Check to see if the index user is a player/npc target
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
         Index <> a And _
         User(a).TargetNum = Index And _
         User(a).TargetGUID = User(Index).UserGUID And _
         User(a).Location = User(Index).Location Then
         PlayerIsTarget = True
         frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " has taken aim on you, the only way you execute this action is to kill " & User(a).UName & " or flee the area.  If you choose to flee, you will lose a fair amount of rank and possibly drop an item or two in the scramble to get away." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Function
   End If
Next a

For a = 0 To 15
   If City(User(Index).Location).CNpc(a) <> -1 Then
      If Npc(City(User(Index).Location).CNpc(a)).NTargetID = Index And _
         Npc(City(User(Index).Location).CNpc(a)).NTargetGUID = User(Index).UserGUID And _
         Npc(City(User(Index).Location).CNpc(a)).NLocation = User(Index).Location Then
         Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()
         PlayerIsTarget = True
         frmMain.wsk(Index).SendData Chr$(2) & Npc(City(User(Index).Location).CNpc(a)).NName & " has taken aim on you, the only way you can execute this actions is to kill " & Npc(City(User(Index).Location).CNpc(a)).NName & " or flee the area.  If you choose to flee, you will lose a fair amount of rank and possibly drop an item or two in the scramble to get away." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Function
      End If
   End If
Next a
      
PlayerIsTarget = False
User(Index).TargetNum = -1
User(Index).TargetGUID = ""

End Function
Public Sub PawnShopMenu(Index As Integer)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter
Dim msg As String 'String

If City(User(Index).Location).PawnShop = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a pawn shop first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

Call NoHiding(Index)

msg = Chr$(255) & Chr$(3)

For a = 0 To UBound(ItemDB)
   If ItemDB(a).ForSale = True Then
      msg = msg & ItemDB(a).IName & Chr$(1)
   End If
Next a

msg = msg & Chr$(2)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      msg = msg & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      msg = msg & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData msg & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

End Sub

Public Sub ShopItemInfo(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim YesOrNo As String

'Check to see if the message is a number
If IsNumeric(msg) = False Then
   Exit Sub
End If

a = msg

If a < LBound(SlotID) Or a > UBound(SlotID) Then
   Exit Sub
End If

'Check to see if the item fits in the scope
'If User(Index).Item(a) = -1 Then
'   Exit Sub
'End If

'check to see if the item is dope
'If Item(User(Index).Item(a)).IType = C_Dope Then
'   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & Chr$(0)
'   DoEvents
'   Exit Sub
'End If

'Send the item info to the Pawn Shop Menu
If User(Index).Reputation < ItemDB(SlotID(a)).CanBuy Then
   YesOrNo = "No"
ElseIf User(Index).Reputation >= ItemDB(SlotID(a)).CanBuy Then
   YesOrNo = "Yes"
End If

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & ItemDB(SlotID(a)).Price & Chr$(1) & YesOrNo & Chr$(1) & ItemDB(SlotID(a)).IName & Chr$(1) & Chr$(0)
DoEvents

End Sub
Public Sub PlayerItemInfo(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer

If IsNumeric(msg) = False Then
   Exit Sub
End If

a = msg

If msg < 0 Or msg > 19 Then
   Exit Sub
End If

If User(Index).Item(a) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).IType = C_Ammo And _
   Item(User(Index).Item(a)).Amount <> 10 Then
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & Int(Item(User(Index).Item(a)).Price / 2) & Chr$(1) & "N/A" & Chr$(1) & Item(User(Index).Item(a)).IName & Chr$(1) & Chr$(0)
DoEvents

End Sub

Public Sub BuyItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer
Dim c As Integer
Dim MsgX As String

For a = 0 To 19
   If User(Index).Item(a) = -1 Then
      Exit For
   ElseIf a = 19 Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Inventory Full" & Chr$(1) & "Inventory Full" & Chr$(1) & "Inventory Full" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next a

If IsNumeric(msg) = False Then
   Exit Sub
End If

b = msg

If User(Index).Cash < ItemDB(SlotID(b)).Price Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Lack Of Cash" & Chr$(1) & "Lack Of Cash" & Chr$(1) & "Lack Of Cash" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
End If

If User(Index).Reputation < ItemDB(SlotID(b)).CanBuy Then
      frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Lack Of Rank" & Chr$(1) & "Lack Of Rank" & Chr$(1) & "Lack Of Rank" & Chr$(1) & Chr$(0)
      DoEvents
      Exit Sub
End If

User(Index).Cash = User(Index).Cash - ItemDB(SlotID(b)).Price
ReDim Preserve Item(UBound(Item) + 1)
Item(UBound(Item)) = ItemDB(SlotID(b))
Item(UBound(Item)).Decay = -1
Item(UBound(Item)).Equip = False
Item(UBound(Item)).ForSale = False
Item(UBound(Item)).ILocation = -1
Item(UBound(Item)).ItemGUID = User(Index).UserGUID
Item(UBound(Item)).OnPlayer = True
User(Index).Item(a) = UBound(Item)
Call UpdateSingleItem(Index, a)
frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

MsgX = Chr$(255) & Chr$(6)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)

End Sub
Public Sub SellItem(Index As Integer, msg As String)
On Error GoTo BadItemNo
Dim a As Integer
Dim MsgX As String

If IsNumeric(msg) = False Then
   Exit Sub
End If

a = msg

If a < 0 Or a > 19 Then
   Exit Sub
End If

If Item(User(Index).Item(a)).ItemGUID <> User(Index).UserGUID Then
   Exit Sub
End If

If Item(User(Index).Item(msg)).IType = C_Dope Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & "Get Lost" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Item(a)).Equip = True Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Item Equipted" & Chr$(1) & "Item Equipted" & Chr$(1) & "Item Equipted" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Item(msg)).IType = C_Ammo And _
   Item(User(Index).Item(msg)).Amount < 10 Then
   frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Used Ammo" & Chr$(1) & "Used Ammo" & Chr$(1) & "Used Ammo" & Chr$(1) & Chr$(0)
   DoEvents
   Exit Sub
End If

User(Index).Cash = User(Index).Cash + Int((Item(User(Index).Item(a)).Price / 2))
Call ResetItem(User(Index).Item(a))
User(Index).Item(a) = -1
Call UpdateSingleItem(Index, a)

MsgX = Chr$(255) & Chr$(6)

For a = 0 To 19
   If User(Index).Item(a) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(a)).IName & Chr$(1)
   ElseIf User(Index).Item(a) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData Chr$(255) & Chr$(5) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(255) & Chr$(4) & User(Index).Cash & Chr$(1) & User(Index).Rank & Chr$(1) & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)
Exit Sub

BadItemNo:
Dim ff As Integer
ff = FreeFile
Open App.Path & "\error.log" For Append As ff
Print #ff, "[BOE]"
Print #ff, "Bad Item Number In Sell Menu"
Print #ff, User(Index).UName & " | " & msg
Print #ff, "[EOE]"
Close ff

End Sub
Public Function AddItem(ItemNo As Integer) As Integer
Dim a As Integer

'This adds items to NPCs who are just spawned

For a = 0 To UBound(Item)
   If Item(a).IName = "" And _
      Item(a).ItemGUID = "" Then
         Item(a) = ItemDB(ItemNo)
         Item(a).OnPlayer = True
         AddItem = a
         Exit Function
   ElseIf a = UBound(Item) Then
      ReDim Preserve Item(UBound(Item) + 1)
         Item(UBound(Item)) = ItemDB(ItemNo)
         Item(UBound(Item)).OnPlayer = True
         AddItem = UBound(Item)
         Exit Function
   End If
Next a

End Function

Public Sub AddNpcGM(Index As Integer, NpcType As String)
On Error Resume Next
Dim a As Integer

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If IsNumeric(NpcType) = True Then
   a = NpcType
ElseIf IsNumeric(NpcType) = False Then
   Exit Sub
End If

Call AddNpc(a, User(Index).Location)

End Sub

Public Sub BuyDrugMenu(Index As Integer, NpcName As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim msg As String

Dim c As Integer


If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)
If User(Index).Reputation > 400 Then
Randomize
c = Int(30 - 1) * Rnd + 1

Select Case c
Case 1
    
        Call AddNpc(N_Cop, User(Index).Location)
        frmMain.wsk(Index).SendData Chr$(2) & Npc(City(User(Index).Location).CNpc(c)).NName & " yells: Stop right there punk!" & vbCrLf & vbCrLf & Chr$(0)
    
        If City(User(Index).Location).CNpc(c) <> -1 Then
        Npc(City(User(Index).Location).CNpc(c)).NTargetID = Index
        Npc(City(User(Index).Location).CNpc(c)).NTargetGUID = User(Index).UserGUID
        Npc(City(User(Index).Location).CNpc(c)).CanMove = False
        DoEvents
        End If
    End Select
    End If
 
'Clear users dealer tag
User(Index).NpcTrade = -1

If NpcName = "" Then
   Exit Sub
End If

If InventoryFull(Index) = True Then
   Exit Sub
End If

'Check to see if NPC Name is in room
For a = 0 To 15
   If City(User(Index).Location).CNpc(a) <> -1 Then
   If LCase$(Left$(Npc(City(User(Index).Location).CNpc(a)).NName, Len(NpcName))) = _
      LCase$(NpcName) And _
         Npc(City(User(Index).Location).CNpc(a)).NpcType = N_Dealer Then
         Exit For
      End If
   ElseIf a = 15 Then
      frmMain.wsk(Index).SendData Chr$(2) & "There is no dealer by that name here with you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next a

'Set users NPC trading number
User(Index).NpcTrade = City(User(Index).Location).CNpc(a)
'Stop npc from walking out of room for 2 minutes
Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()

'Send npc's inventory to user
msg = Chr$(254) & Chr$(4)
For b = 0 To 19
      If Npc(City(User(Index).Location).CNpc(a)).NItem(b) = -1 Then
         msg = msg & "<Empty>" & Chr$(1)
      ElseIf Npc(City(User(Index).Location).CNpc(a)).NItem(b) <> -1 Then
         msg = msg & Item(Npc(City(User(Index).Location).CNpc(a)).NItem(b)).IName & Chr$(1)
      End If
Next b

msg = msg & Chr$(0)

frmMain.wsk(Index).SendData msg
DoEvents
      
frmMain.wsk(Index).SendData Chr$(252) & Chr$(4) & City(User(Index).Location).Compass & Chr$(0)
DoEvents
   
End Sub
Public Sub DrugDealItemInfo(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Single

If IsNumeric(msg) = False Then
   Exit Sub
ElseIf IsNumeric(msg) = True Then
   a = msg
End If

If a < 0 Or a > 19 Then
   Exit Sub
End If

If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(5) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That dealer has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

If Npc(User(Index).NpcTrade).NItem(msg) = -1 Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "Man, there is nothing in that pocket you can buy." & Chr$(0)
   DoEvents
   Call UpdateNPCInventory(Index)
   Exit Sub
End If

b = Item(Npc(User(Index).NpcTrade).NItem(a)).Price - (Item(Npc(User(Index).NpcTrade).NItem(a)).Price * 0.06)
b = Int(b)

frmMain.wsk(Index).SendData Chr$(254) & Chr$(6) & Item(Npc(User(Index).NpcTrade).NItem(msg)).IName & Chr$(1) & b & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

End Sub
Public Sub UpdateNPCInventory(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim msg As String
msg = Chr$(253) & Chr$(2)

For a = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(a) = -1 Then
      msg = msg & "<Empty>" & Chr$(1)
   ElseIf Npc(User(Index).NpcTrade).NItem(a) <> -1 Then
      msg = msg & Item(Npc(User(Index).NpcTrade).NItem(a)).IName & Chr$(1)
   End If
Next a

frmMain.wsk(Index).SendData msg & Chr$(0)
DoEvents

End Sub

Public Sub BuyNpcDrug(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Single

'Make sure the item is Good
If IsNumeric(msg) = False Then
   Exit Sub
ElseIf IsNumeric(msg) = True Then
   a = msg
End If

'Make sure user has room in inventory
For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      Exit For
   ElseIf b = 19 Then
      frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You ain't got the room man, try selling something." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next b

'Make sure the NPC is still in the same room as the player
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(5) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That dealer has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'Make sure the dealer hasn't sold the item to another player
If Npc(User(Index).NpcTrade).NItem(msg) = -1 Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "Man, there is nothing in that pocket you can buy." & Chr$(0)
   DoEvents
   Call UpdateNPCInventory(Index)
   Exit Sub
End If

c = Item(Npc(User(Index).NpcTrade).NItem(msg)).Price - (Item(Npc(User(Index).NpcTrade).NItem(msg)).Price * 0.06)
c = Int(c)

If User(Index).Cash >= c Then
   User(Index).Cash = User(Index).Cash - c
   User(Index).Reputation = User(Index).Reputation + 0
   Call SetRank(Index)
   Call UpdateGeneralInfo(Index)
   User(Index).Item(b) = Npc(User(Index).NpcTrade).NItem(msg)
   Npc(User(Index).NpcTrade).NItem(msg) = -1
   Item(User(Index).Item(b)).ItemGUID = User(Index).UserGUID
   Call UpdateNPCInventory(Index)
   Call UpdateSingleItem(Index, b)
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You got it..." & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(6) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
   DoEvents
ElseIf User(Index).Cash < c Then
   frmMain.wsk(Index).SendData Chr$(254) & Chr$(7) & "You ain't got the cash to buy that dope from me fool." & Chr$(0)
   DoEvents
End If

End Sub

Public Sub DruggieMenu(Index As Integer, NpcName As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim msg As String
Dim c As Integer

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).Reputation > 400 Then
Randomize
c = Int(30 - 1) * Rnd + 1

Select Case c
Case 10
    
        Call AddNpc(N_Cop, User(Index).Location)
    
        If City(User(Index).Location).CNpc(c) <> -1 Then
        Npc(City(User(Index).Location).CNpc(c)).NTargetID = Index
        Npc(City(User(Index).Location).CNpc(c)).NTargetGUID = User(Index).UserGUID
        Npc(City(User(Index).Location).CNpc(c)).CanMove = False
        frmMain.wsk(Index).SendData Chr$(2) & Npc(City(User(Index).Location).CNpc(c)).NName & " Yells:  Stop right there punk, and pulls out his weapon." & vbCrLf & Chr$(0)
        DoEvents
        End If
        
    End Select

'Clear users druggies tag
User(Index).NpcTrade = -1

If NpcName = "" Then
   Exit Sub
End If

Call NoHiding(Index)

'Check to see if NPC Name is in room
For a = 0 To 15
   If City(User(Index).Location).CNpc(a) <> -1 Then
     
   If LCase$(Left$(Npc(City(User(Index).Location).CNpc(a)).NName, Len(NpcName))) = _
      LCase$(NpcName) And _
         Npc(City(User(Index).Location).CNpc(a)).NpcType = N_Druggie Then
         Exit For
      End If
   ElseIf a = 15 Then
      frmMain.wsk(Index).SendData Chr$(2) & "There is no druggie by that name here with you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
   
Next a

'Set users NPC trading number
User(Index).NpcTrade = City(User(Index).Location).CNpc(a)
'Stop npc from walking out of room for 2 minutes
Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()

msg = Chr$(253) & Chr$(3)

For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      msg = msg & "<Empty>" & Chr$(1)
   ElseIf User(Index).Item(b) <> -1 Then
      msg = msg & Item(User(Index).Item(b)).IName & Chr$(1)
   End If
Next b

frmMain.wsk(Index).SendData msg & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(252) & Chr$(5) & City(User(Index).Location).Compass & Chr$(0)
DoEvents
End If
End Sub
Public Sub DruggieItemInfo(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Single
Dim c As Integer

If IsNumeric(msg) = True Then
   a = msg
ElseIf IsNumeric(msg) = False Then
   Exit Sub
End If

'Make sure the item is not empty
If User(Index).Item(msg) = -1 Then
   Exit Sub
End If

If Item(User(Index).Item(msg)).IType <> C_Dope Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I only deal in dope, try the pawn shop if you want to unload that junk." & Chr$(0)
      DoEvents
      Exit Sub
End If

'Make sure npc is still in same room as player
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(4) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That druggie has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'Make sure npc has room for the item
For c = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(c) = -1 Then
      Exit For
   ElseIf c = 19 Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I can't afford anything else right now, try me later after I unload some of this dope." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next c

'Set Price at a % Mark
b = Item(User(Index).Item(msg)).Price + (Item(User(Index).Item(msg)).Price * 0.06)
b = Int(b)

'Send item information to the player
frmMain.wsk(Index).SendData Chr$(253) & Chr$(6) & Item(User(Index).Item(msg)).IName & Chr$(1) & b & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "Well?" & Chr$(0)
DoEvents

End Sub

Public Sub ListNPCs(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim msg As String


'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To UBound(Npc)
   If Npc(a).NpcGUID <> "" And _
      Npc(a).NName <> "" And _
      Npc(a).NCity = City(User(Index).Location).CName Then
         msg = msg & "   (" & Npc(a).NName & " " & Npc(a).NameTag & " | " & City(Npc(a).NLocation).Compass & ")   "
   End If
Next a

frmMain.wsk(Index).SendData Chr$(2) & msg & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub SellDruggieItem(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Single
Dim d As Integer
Dim MsgX As String

'Make sure index number is a real number
If IsNumeric(msg) = True Then
   a = msg
ElseIf IsNumeric(msg) = False Then
   Exit Sub
End If

'make sure index number is not out of scope
If a < 0 Or a > 19 Then
   Exit Sub
End If

'make sure the item exists
If User(Index).Item(a) = -1 Then
   Exit Sub
End If

'make sure the druggie is still in the area
If Npc(User(Index).NpcTrade).NLocation <> User(Index).Location Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(4) & Chr$(0)
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "That druggie has left the area." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).NpcTrade = -1
   Exit Sub
End If

'make sure the item is a DOPE type
If Item(User(Index).Item(a)).IType <> C_Dope Then
   frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I told you once bitch, I don't buy that kinda merchandise!  Dope Only!" & Chr$(0)
   DoEvents
   Exit Sub
End If

'Make sure the npc has room to buy the item
For d = 0 To 19
   If Npc(User(Index).NpcTrade).NItem(d) = -1 Then
      Exit For
   ElseIf d = 19 Then
      frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "I can't afford anything else right now, try me later after I unload some of this dope." & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next d

'Set Price at a % Mark
c = Item(User(Index).Item(a)).Price + (Item(User(Index).Item(a)).Price * 0.06)
c = Int(c)

'Do Transaction
User(Index).Cash = User(Index).Cash + c
Npc(User(Index).NpcTrade).NItem(d) = User(Index).Item(a)
User(Index).Item(a) = -1
Item(Npc(User(Index).NpcTrade).NItem(d)).ItemGUID = Npc(User(Index).NpcTrade).NpcGUID
Call UpdateSingleItem(Index, a)
User(Index).Reputation = User(Index).Reputation + 2
Call SetRank(Index)
Call UpdateGeneralInfo(Index)

'Update the druggie menu inventory list
MsgX = Chr$(253) & Chr$(7)
For b = 0 To 19
   If User(Index).Item(b) = -1 Then
      MsgX = MsgX & "<Empty>" & Chr$(1)
   ElseIf User(Index).Item(b) <> -1 Then
      MsgX = MsgX & Item(User(Index).Item(b)).IName & Chr$(1)
   End If
Next b
frmMain.wsk(Index).SendData MsgX & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(6) & "Sold" & Chr$(1) & "Sold" & Chr$(1) & User(Index).Cash & Chr$(1) & Chr$(0)
DoEvents

frmMain.wsk(Index).SendData Chr$(253) & Chr$(5) & "Ok, anything else?" & Chr$(0)
DoEvents

End Sub

Public Sub SendMap(Index As Integer)
On Error Resume Next

Select Case User(Index).CurrTown
   Case "New York"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & NYMap & Chr$(0)
      DoEvents
   Case "Houston"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & HOMap & Chr$(0)
      DoEvents
   Case "Miami"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & MIMap & Chr$(0)
      DoEvents
   Case "Los Angeles"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & LAMap & Chr$(0)
      DoEvents
   Case "New Jersey"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & NJMap & Chr$(0)
      DoEvents
   Case "Chicago"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & CHMap & Chr$(0)
      DoEvents
   Case "Sydney"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & SYDMap & Chr$(0)
      DoEvents
   Case "London"
      frmMain.wsk(Index).SendData Chr$(252) & Chr$(2) & UKMap & Chr$(0)
      DoEvents
End Select

End Sub

Public Sub Aim(Index As Integer, PlayerName As String)
On Error Resume Next
Dim a As Integer

For a = 1 To MaxUsers
   If LCase$(PlayerName) = LCase$(Left$(User(a).UName, Len(PlayerName))) And _
      User(Index).Location = User(a).Location And _
      a <> Index And User(a).Status = "Playing" And _
      User(a).IsHiding = False Then
      
      

      Call NoHiding(Index)
      User(Index).TargetNum = a
      User(Index).TargetGUID = User(a).UserGUID
      frmMain.wsk(Index).SendData Chr$(2) & "You take aim on " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(a).SendData Chr$(2) & User(Index).UName & " takes aim on you." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call CombatMessage(Index, a, Chr$(2) & "You see " & User(Index).UName & " take aim on " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0))
      Exit Sub
   End If
Next a

For a = 0 To 15
   If City(User(Index).Location).CNpc(a) <> -1 Then
   If LCase$(Left$(Npc(City(User(Index).Location).CNpc(a)).NName, Len(PlayerName))) = _
      LCase$(PlayerName) Then
      Call NoHiding(Index)
      User(Index).TargetNum = City(User(Index).Location).CNpc(a)
      User(Index).TargetGUID = Npc(City(User(Index).Location).CNpc(a)).NpcGUID
      Npc(City(User(Index).Location).CNpc(a)).NTargetID = Index
      Npc(City(User(Index).Location).CNpc(a)).NTargetGUID = User(Index).UserGUID
      Npc(City(User(Index).Location).CNpc(a)).CanMove = GetTickCount()
      frmMain.wsk(Index).SendData Chr$(2) & "You take aim on " & Npc(City(User(Index).Location).CNpc(a)).NName & "." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " take aim on " & Npc(City(User(Index).Location).CNpc(a)).NName & "." & vbCrLf & vbCrLf & Chr$(0))
      Exit Sub
   End If
   End If
Next a

User(Index).TargetGUID = ""
User(Index).TargetNum = -1
frmMain.wsk(Index).SendData Chr$(2) & "You look around but one matches that name..." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Function IsHiding(Index As Integer) As Boolean

If User(Index).IsHiding = True Then
   IsHiding = False
   User(Index).IsHiding = False
   frmMain.wsk(Index).SendData Chr$(2) & "You come out of hiding." & vbCrLf & vbCrLf & Chr$(2)
   DoEvents
   Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " come out from the shadows." & vbCrLf & vbCrLf & Chr$(0))
   Exit Function
ElseIf User(Index).IsHiding = False Then
   IsHiding = False
   Exit Function
End If

End Function

Public Sub GotoPlayer(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 1 To MaxUsers
   If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) Then
      User(Index).Location = User(a).Location
      User(Index).CurrTown = City(User(Index).Location).CName
      Call UpdateGeneralInfo(Index)
      Call ShowCity(Index)
   End If
Next a

End Sub

Public Sub Rankage(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer



'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

 User(Index).Reputation = User(Index).Reputation + msg
 UpdateGeneralInfo (Index)
   
   End Sub
   
   
   


'Public Sub Taxi(Index As Integer, msg As String)
'On Error Resume Next
'Dim a As Integer
'If msg = "" Then
'    Exit Sub
'End If
'For a = 1 To MaxUsers
   'If User(Index).Cash < 20000 Then
    '  frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash." & vbCrLf & vbCrLf & Chr$(0)
   '   DoEvents
  '    Exit Sub
 '  End If
'If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) Then
    'If User(Index).Tracking >= 75 Then
     ' Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " Jump into a private jet." & vbCrLf & vbCrLf & Chr$(0))
     ' User(Index).Location = User(a).Location
    '  User(Index).CurrTown = City(User(Index).Location).CName
    '  Call UpdateGeneralInfo(Index)
    ''  Call ShowWatchers(Index, Chr$(2) & "A small jet lands, you see " & User(Index).UName & " jump out." & vbCrLf & vbCrLf & Chr$(0))
    '  User(Index).Cash = User(Index).Cash - 20000
    '  frmMain.wsk(Index).SendData Chr$(2) & "You hop on a private jet, and it takes you straight to " & User(a).UName & vbCrLf & Chr$(0)
    '  Exit Sub
    'Else
    '  Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " Jump into a private jet." & vbCrLf & vbCrLf & Chr$(0))
    '  User(Index).Location = User(a).CurrTown
     ' User(Index).CurrTown = City(User(Index).Location).CName
     ' Call UpdateGeneralInfo(Index)
     ' Call ShowWatchers(Index, Chr$(2) & "A small jet lands, you see " & User(a).UName & " jump out." & vbCrLf & vbCrLf & Chr$(0))
    '  User(Index).Cash = User(Index).Cash - 20000
   '   frmMain.wsk(Index).SendData Chr$(2) & "You hop on a private jet, and it takes you to " & User(a).UName & "'s hometown" & vbCrLf & vbCrLf & Chr$(0)
  '    Exit Sub
 '   End If
'End If
'Next a
'Exit Sub
'End Sub

Public Sub Punch(Index As Integer)
On Error Resume Next
Dim a As Integer

If Skilldelay(Index) = True Then
    Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            If RunAccuracy(Index) = True Then
               User(User(Index).TargetNum).Health = User(User(Index).TargetNum).Health - 2
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " throw a hard punch hitting " & User(User(Index).TargetNum).UName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You land a solid punch on " & User(User(Index).TargetNum).UName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " lands a damaging blow on you." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " takes a swing at you but misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - 2
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " throw a hard punch hitting " & Npc(User(Index).TargetNum).NName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You land a solid punch on " & Npc(User(Index).TargetNum).NName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
         Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub Fire(Index As Integer)
On Error Resume Next
Dim a As Integer

If Skilldelay(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can shoot them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Weapon = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have a weapon equipped." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Weapon).IType <> C_Gun Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your equipped weapon is not a firearm." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Ammo = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have any ammunition loaded." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

For a = 0 To 19
   If User(Index).Item(a) = User(Index).Ammo Then
      Exit For
   ElseIf a = 19 Then
      Exit Sub
   End If
Next a

If Item(User(Index).Ammo).Amount <= 0 Then
   User(Index).Item(a) = -1
   Call UpdateSingleItem(Index, a)
   Call ResetItem(User(Index).Ammo)
   User(Index).Ammo = -1
   Call UpdateGearInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "Click, Click...  Sounds like your out of ammunition." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            'subtract ammo
            Item(User(Index).Ammo).Amount = Item(User(Index).Ammo).Amount - 1
            Call UpdateSingleItem(Index, a)
            If RunAccuracy(Index) = True Then
               Call GunDamage(Index)
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " and it's a direct hit." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " fires a " & Item(User(Index).Weapon).IName & " at you, it's a direct hit!" & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " fires a " & Item(User(Index).Weapon).IName & " at you and misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         Item(User(Index).Ammo).Amount = Item(User(Index).Ammo).Amount - 1
         Call UpdateSingleItem(Index, a)
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - (Item(User(Index).Weapon).Damage + Item(User(Index).Ammo).Damage)
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and it's a direct hit." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You fire your " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " fire a " & Item(User(Index).Weapon).IName & " at " & Npc(User(Index).TargetNum).NName & " and miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub



Public Sub GunDamage(Index As Integer)
On Error Resume Next
Dim a As Integer
a = 0

If User(User(Index).TargetNum).Armor <> -1 Then
   a = Item(User(User(Index).TargetNum).Armor).Armor
End If

User(User(Index).TargetNum).Health = (User(User(Index).TargetNum).Health - (Item(User(Index).Weapon).Damage + Item(User(Index).Ammo).Damage)) + a

End Sub

Public Sub Strike(Index As Integer)
On Error Resume Next

If Skilldelay(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can strike them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Weapon = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have a weapon equipped." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Item(User(Index).Weapon).IType <> C_Melee Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your equipped weapon can not be used to strike someone." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            If RunAccuracy(Index) = True Then
               User(User(Index).TargetNum).Health = User(User(Index).TargetNum).Health - Item(User(Index).Weapon).Damage
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " strike " & User(User(Index).TargetNum).UName & " with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & User(User(Index).TargetNum).UName & " with your " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " strikes you with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You strike at " & User(User(Index).TargetNum).UName & " with your " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " strikes at you with a " & Item(User(Index).Weapon).IName & " but misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " take a strike at " & User(User(Index).TargetNum).UName & " with a " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - Item(User(Index).Weapon).Damage
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " strike at " & Npc(User(Index).TargetNum).NName & " with a " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & Npc(User(Index).TargetNum).NName & " with your " & Item(User(Index).Weapon).IName & " doing massive damage." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You strike " & Npc(User(Index).TargetNum).NName & " with your " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " strike at " & Npc(User(Index).TargetNum).NName & " with a " & Item(User(Index).Weapon).IName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub Hide(Index As Integer)
On Error Resume Next

If Skilldelay(Index) = True Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

User(Index).IsHiding = False
If RunHiding(Index) = True Then
   User(Index).IsHiding = True
   Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " slips into the shadows." & vbCrLf & vbCrLf & Chr$(0))
   frmMain.wsk(Index).SendData Chr$(2) & "You manage to slip into the shadows." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
Else
   User(Index).IsHiding = False
   frmMain.wsk(Index).SendData Chr$(2) & "You failed to slip into the shadows." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
End Sub

Public Sub Flee(Index As Integer, msg As String)
On Error Resume Next

Dim DropItems As Boolean
Dim a As Integer
Dim b As Integer
DropItems = False

Dim i(2) As Integer

For a = 0 To 2
   i(a) = -1
Next a

'Flee North
Select Case msg
   Case "n"
      If City(User(Index).Location).North = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the north." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).North <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the north from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).North
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the north from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).North
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee East
   Case "e"
      If City(User(Index).Location).East = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the east." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).East <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the east from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).East
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the west." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the east from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).East
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the west." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee South
   Case "s"
      If City(User(Index).Location).South = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the south." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).South <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the south from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).South
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the north." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the south from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).South
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the north." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

'Flee West
   Case "w"
      If City(User(Index).Location).West = -1 Then
         frmMain.wsk(Index).SendData Chr$(2) & "You can't flee to the west." & vbCrLf & vbCrLf & Chr$(0)
         DoEvents
         Exit Sub
      ElseIf City(User(Index).Location).West <> 1 Then
      Randomize
      b = Int(19 - 0) * Rnd
      i(0) = b
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(1) = b
      Loop Until i(1) <> i(0)
      Do
      Randomize
      b = Int(19 - 0) * Rnd
      i(2) = b
      Loop Until i(2) <> i(0) And i(2) <> i(1)
      For a = 0 To 2
         If User(Index).Item(i(a)) <> -1 Then
            If Item(User(Index).Item(i(a))).Equip = False Then
            DropItems = True
            For b = 0 To UBound(City(User(Index).Location).CItem)
               If City(User(Index).Location).CItem(b) = -1 Then
                  City(User(Index).Location).CItem(b) = User(Index).Item(i(a))
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               ElseIf b = UBound(City(User(Index).Location).CItem) Then
                  With City(User(Index).Location)
                  ReDim Preserve .CItem(UBound(.CItem) + 1)
                  .CItem(UBound(.CItem)) = User(Index).Item(i(a))
                  End With
                  Item(User(Index).Item(i(a))).Equip = False
                  Item(User(Index).Item(i(a))).Decay = GetTickCount()
                  Item(User(Index).Item(i(a))).OnPlayer = False
                  Item(User(Index).Item(i(a))).ItemGUID = ""
                  Item(User(Index).Item(i(a))).ILocation = User(Index).Location
                  User(Index).Item(i(a)) = -1
                  Exit For
               End If
            Next b
          End If
         End If
      Next a
         If DropItems = False Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the west from battle.  Such a cowardly act costs you quite a bit of fame and reputation." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).West
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the east." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         ElseIf DropItems = True Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " has managed to escape the area from battle but dropped some items in the process of escaping." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Reputation = User(Index).Reputation - 100
            frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape to the west from battle.  Such a cowardly act costs you quite a bit of fame and reputation.  In your scuttle to escape, you drop some items." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
            User(Index).Location = City(User(Index).Location).West
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the east." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
            Exit Sub
         End If
      End If

End Select

End Sub

Public Sub HealInfo(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim b As Integer

If City(User(Index).Location).Hospital = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a hospital first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).Cash < HealPrice Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your current financial situation won't do you any good here." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health >= 100 Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need our services, you're in perfect health." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf User(Index).Health < 100 Then
   a = (100 - User(Index).Health) * HealPrice
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your health needs some attention,  It will take you " & 100 - User(Index).Health & " days in the hospital and $" & a & " dollars to have perfect health." & vbCrLf & vbCrLf & "To use our services, type  healme <amount>,  You currently can afford " & Int(User(Index).Cash / HealPrice) & " days in the hospital." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

End Sub

Public Sub HealMe(Index As Integer, msg As String)
On Error Resume Next
Dim a As Single


If IsNumeric(msg) = False Then
   Exit Sub
ElseIf IsNumeric(msg) = True Then
   a = Int(msg)
End If

If Int(a) < 1 Or Int(a) > 99 Then
   Exit Sub
End If

If City(User(Index).Location).Hospital = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a hospital first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If User(Index).Cash < HealPrice Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: Your current financial situation won't do you any good here." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health >= 100 Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need our services, you're in perfect health." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If Int(a) > (100 - User(Index).Health) Then
   frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: You don't need to stay here that long." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Health < 100 Then
   If (Int(a) * HealPrice) > User(Index).Cash Then
      frmMain.wsk(Index).SendData Chr$(2) & "A Nurse says: you do not have enough money to stay that long." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   ElseIf (Int(a) * HealPrice) <= User(Index).Cash Then
      User(Index).Cash = User(Index).Cash - Int(a) * HealPrice
      User(Index).Health = User(Index).Health + Int(a)
      Call UpdateGeneralInfo(Index)
      frmMain.wsk(Index).SendData Chr$(2) & "You hand over $" & a * HealPrice & " and the doctors fix you right up." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   End If
End If

frmMain.wsk(Index).SendData Chr$(2) & "The Nurse looks at you strangley." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub ShowSkills(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim msg As String

msg = "Your current skills:" & vbCrLf
msg = msg & "Accuracy:    " & Format$(User(Index).Accuracy, "#0.0") & vbCrLf
msg = msg & "Hiding:     " & Format$(User(Index).Hiding, "#0.0") & vbCrLf
msg = msg & "Searching:     " & Format$(User(Index).Search, "#0.0") & vbCrLf
msg = msg & "Tracking:     " & Format$(User(Index).Tracking, "#0.0") & vbCrLf
msg = msg & "Sniping:     " & Format$(User(Index).Sniping, "#0.0") & vbCrLf
'Msg = Msg & "Bounty:     " & Format$(User(Index).Bounty, "#0.0") & vbCrLf
'Msg = Msg & "Chemistry:     " & Format$(User(Index).Chemistry, "#0.0") & vbCrLf
msg = msg & "Snooping:     " & Format$(User(Index).Snooping, "#0.0") & vbCrLf
msg = msg & "Stealing:     " & Format$(User(Index).Stealing, "#0.0") & vbCrLf
'msg = msg & "Rank Points:     " & Format$(User(Index).Reputation) & vbCrLf

frmMain.wsk(Index).SendData Chr$(2) & msg & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub Deposit(Index As Integer, msg As String)
On Error Resume Next
Dim a As Single

If City(User(Index).Location).Bank = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bank before you can deposit any cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If City(User(Index).Location).CName <> User(Index).HomeTown Then
   frmMain.wsk(Index).SendData Chr$(2) & "You can only bank in your home town." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).Reputation <= 200 And _
   User(Index).Reputation >= -4000 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Your current rank doesn't allow you to open a bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
If IsNumeric(msg) = False Then
   Exit Sub
ElseIf IsNumeric(msg) = True Then
   a = Int(msg)
End If

If Int(a) < 1 Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If Int(a) > User(Index).Cash Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have that much cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf Int(a) <= User(Index).Cash Then
   User(Index).Cash = User(Index).Cash - Int(a)
   User(Index).Bank = User(Index).Bank + Int(a)
   Call UpdateGeneralInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "You deposit $" & Int(a) & " into your bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & "The bank teller looks strangley at you." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub Withdraw(Index As Integer, msg As String)
On Error Resume Next
Dim a As Single

If City(User(Index).Location).Bank = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bank before you can deposit any cash." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If City(User(Index).Location).CName <> User(Index).HomeTown Then
   frmMain.wsk(Index).SendData Chr$(2) & "You can only bank in your home town." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
   
If IsNumeric(msg) = False Then
   Exit Sub
ElseIf IsNumeric(msg) = True Then
   a = Int(msg)
End If

If Int(a) < 1 Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

Call NoHiding(Index)

If Int(a) > User(Index).Bank Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have that much cash in your bank." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf Int(a) <= User(Index).Bank Then
   User(Index).Bank = User(Index).Bank - Int(a)
   User(Index).Cash = User(Index).Cash + Int(a)
   Call UpdateGeneralInfo(Index)
   frmMain.wsk(Index).SendData Chr$(2) & "You withdraw $" & Int(a) & " from your bank account." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

frmMain.wsk(Index).SendData Chr$(2) & "The bank teller looks strangley at you." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub TrackPlayer(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
If msg = "" Then
    Exit Sub
End If

If Skilldelay(Index) = True Then
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

If RunTracking(Index) = True Then
    For a = 1 To MaxUsers
        If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) Then
            If User(Index).Tracking >= 75 Then
                For b = 0 To 7199
                    If User(a).Location = City(b).CLocation Then
                        frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " was last seen at " & City(b).CName & " " & City(b).Compass & vbCrLf & vbCrLf & Chr$(0)
                        DoEvents
                        Exit Sub
                    End If
                Next
            Else
                frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & "was last seen at " & City(b).CName & " " & vbCrLf & vbCrLf & Chr$(0)
                DoEvents
                Exit Sub
            End If
        End If
    Next
Else
    frmMain.wsk(Index).SendData Chr$(2) & "That player could not be located at this time." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
End Sub
 Public Sub SnipePlayer(Index As Integer, msg As String)
 On Error Resume Next
Dim a As Integer
Dim b As Integer
If msg = "" Then
    Exit Sub
End If
If Skilldelay(Index) = True Then
    Exit Sub
End If
If PlayerIsTarget(Index) = True Then
   Exit Sub
End If
If User(Index).Cash < 5000 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to snipe." & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If
User(Index).Cash = User(Index).Cash - 5000
If RunSniping(Index) = True Then
    For a = 1 To MaxUsers
        If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) Then
            If User(Index).CurrTown = User(a).CurrTown Then
                If User(a).IsHiding = True Then
                    frmMain.wsk(Index).SendData Chr$(2) & "This player is hiding and cannot be shot." & vbCrLf & vbCrLf & Chr$(0)
                    Exit Sub
                Else
                    User(a).Health = User(a).Health - 5
                    If PlayerKillPlayer(Index, a) = True Then
                        Exit Sub
                    End If
                    frmMain.wsk(Index).SendData Chr$(2) & "It's a direct hit!  " & User(a).UName & " never saw it coming." & vbCrLf & vbCrLf & Chr$(0)
                    frmMain.wsk(a).SendData Chr$(2) & "You've been shot by a sniper rifle!" & vbCrLf & vbCrLf & Chr$(0)
                    Call UpdateGeneralInfo(a)
                    Call UpdateGeneralInfo(Index)
                    Exit Sub
                    
                End If
            Else
                frmMain.wsk(Index).SendData Chr$(2) & "You can't snipe a player that's not in your town." & vbCrLf & vbCrLf & Chr$(0)
                Exit Sub
            End If
        End If
    Next
Else
    frmMain.wsk(Index).SendData Chr$(2) & "Your attempt to snipe has failed."
    Exit Sub
End If
frmMain.wsk(Index).SendData Chr$(2) & "Who?" & vbCrLf & vbCrLf & Chr$(0)

End Sub

Public Sub SnoopPlayer(Index As Integer, msg As String)
On Error Resume Next
    Dim a As Integer
 
    If Len(msg) <= 0 Then Exit Sub
 
    If Skilldelay(Index) = True Then Exit Sub
 
    For a = 1 To MaxUsers
    
    
        If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) Then

   
            If RunSnooping(Index) = True Then
                If User(Index).Snooping >= 70 Then
                    frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " the " & User(a).Rank & " has " & User(a).Health & " health, " & User(a).Cash & "$, and " & User(a).Accuracy & " Accuracy." & " Current bounty: $" & User(a).Bounty & vbCrLf & Chr$(0)
                ElseIf User(Index).Snooping < 70 Then
                    frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " the " & User(a).Rank & " has " & User(a).Health & " health, " & User(a).Cash & "$, " & User(a).Kills & " Kills." & " Current bounty: $" & User(a).Bounty & vbCrLf & vbCrLf & Chr$(0)
                End If
                      ElseIf a = Index And LCase$(Left$(User(a).UName, Len(msg))) = LCase$(msg) Then
        frmMain.wsk(Index).SendData Chr$(2) & "Why ya gunna try and snoop yourself?.. Skills aint that easy." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
            Else
                frmMain.wsk(Index).SendData Chr$(2) & "Unable to find out anything about " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0)
                
            End If
        
        
 End If
    Next a
 
    Exit Sub
 
End Sub

Public Sub RobBank(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
If City(User(Index).Location).Bank = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bank, before you can rob it." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
  Exit Sub
Else
    Call NoHiding(Index)
    User(Index).Cash = User(Index).Cash + 25000
    Call UpdateGeneralInfo(Index)
    For a = 1 To 2
        Call AddNpc(N_FBI, User(Index).Location)
    
        If City(User(Index).Location).CNpc(c) <> -1 Then
        Npc(City(User(Index).Location).CNpc(c)).NTargetID = Index
        Npc(City(User(Index).Location).CNpc(c)).NTargetGUID = User(Index).UserGUID
        Npc(City(User(Index).Location).CNpc(c)).CanMove = False
        DoEvents
        End If
    Next
    Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " rob the bank!  Cops appear out of nowhere and take aim on him." & vbCrLf & vbCrLf & Chr$(0))
    frmMain.wsk(Index).SendData Chr$(2) & "Just robbed the bank! You get $100k but you haven't gotten away yet.  Law enforcement officials flock to the scene and take aim on you." & vbCrLf & vbCrLf & Chr$(0)
    For c = 1 To MaxUsers
         If User(c).Status = "Playing" Then
            frmMain.wsk(c).SendData Chr$(252) & Chr$(3) & "<News Flash>" & User(Index).UName & " from " & User(Index).HomeTown & " has just robbed a bank in " & User(Index).CurrTown & Chr$(0)
            DoEvents
         End If
    Next
   Exit Sub
End If
Exit Sub
End Sub


Public Sub Gamble(Index As Integer, msg As String)
On Error Resume Next
Dim a As Double
Dim b As Integer
Dim c As Integer

If City(User(Index).Location).Casino = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a casino, before you can gamble." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If msg = "" Then
    Exit Sub
End If

If IsNumeric(msg) = False Then
        frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
        Exit Sub
ElseIf IsNumeric(msg) = True Then
        a = Int(msg)
End If

If Int(a) < 1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Ha ha, funny guy.  Got a real bet?" & vbCrLf & vbCrLf & Chr$(0)
   Exit Sub
End If

If Int(a) > 5000 Then
   frmMain.wsk(Index).SendData Chr$(2) & "We don't service that level of action.  Table limit is 5 grand." & vbCrLf & vbCrLf & Chr$(0)
   Exit Sub
End If

If PlayerIsTarget(Index) = True Then
   Exit Sub
End If

If User(Index).Cash < a Then
    frmMain.wsk(Index).SendData Chr$(2) & "You don't got that much cash." & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If

If Int(a) > User(Index).Cash Then
        frmMain.wsk(Index).SendData Chr$(2) & "You don't have that much cash." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
End If
Call NoHiding(Index)

Randomize
b = Int(1960 - 1) * Rnd + 1
'-----------------------------
Select Case b
Case 0 To 1000
    frmMain.wsk(Index).SendData Chr$(2) & "Sorry, you just lost $" & Int(a) & ". Better luck next time!" & vbCrLf & vbCrLf & Chr$(0)
    User(Index).Cash = User(Index).Cash - Int(a)
    Call UpdateGeneralInfo(Index)
    Exit Sub
Case 1001 To 1500
    frmMain.wsk(Index).SendData Chr$(2) & "You won back your $" & Int(a) & ". Try again?" & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
Case 1501 To 1800
    c = Int(a) * 2
Case 1801 To 1900
    c = Int(a) * 5
Case 1901 To 1955
    c = Int(a) * 10
Case 1956 To 1960
    c = Int(a) * 50
End Select
c = c - Int(a)
User(Index).Cash = User(Index).Cash + c
Call UpdateGeneralInfo(Index)
frmMain.wsk(Index).SendData Chr$(2) & "You just gained $" & c & ". Congratulations!" & vbCrLf & vbCrLf & Chr(0)
Exit Sub
End Sub

Public Sub PrivateChat(Index As Integer, xMsg As String)
On Error Resume Next
Dim a As Integer 'Counter
Dim name As String
Dim msg As String


If Len(xMsg) <= 0 Then Exit Sub

name = Split(xMsg, "-")(0)
msg = Split(xMsg, "-")(1)

If Len(msg) <= 0 Then Exit Sub

For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
   LCase$(name) = LCase$(Left$(User(a).UName, Len(name))) Then
         frmMain.wsk(a).SendData Chr$(8) & vbCrLf & "-" & User(Index).UName & " whispers- " & msg & vbCrLf & Chr$(0)
         DoEvents
         frmMain.wsk(Index).SendData Chr$(8) & vbCrLf & "* You say to " & User(a).UName & " - " & msg & vbCrLf & Chr$(0)
         DoEvents
   End If
Next a
frmMain.wsk(a).SendData Chr$(8) & vbCrLf & "-" & User(Index).UName & " whispers- " & msg & Chr$(0)


End Sub
Public Sub SquareChat(Index As Integer, msg As String)
On Error Resume Next

Call ShowWatchers(Index, Chr$(2) & "[" & User(Index).UName & " says] - " & msg & "." & vbCrLf & vbCrLf & Chr$(0))
frmMain.wsk(Index).SendData Chr$(2) & "[You say] - " & msg & vbCrLf & vbCrLf & Chr$(0)
End Sub
Public Sub Search(Index As Integer)
On Error Resume Next
Dim a As Integer
If Skilldelay(Index) = True Then
   Exit Sub
End If

If RunSearch(Index) = True Then
    For a = 1 To MaxUsers
        If User(a).Location = User(Index).Location Then
            Call NoHiding(a)
        End If
    Next
    frmMain.wsk(Index).SendData Chr$(2) & "You have successfully searched the area." & vbCrLf & Chr$(0)
Else
    frmMain.wsk(Index).SendData Chr$(2) & "Your attempt to search the aread has failed." & vbCrLf & Chr$(0)
End If

End Sub

Public Sub StealPlayer(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim c As Double
If msg = "" Then
    Exit Sub
End If

If Skilldelay(Index) = True Then
   Exit Sub
End If

For a = 1 To MaxUsers
If LCase$(msg) = LCase$(Left$(User(a).UName, Len(msg))) And a <> Index Then
        If User(Index).Location <> User(a).Location Then
            frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & " is not around to steal from." & vbCrLf & vbCrLf & Chr$(0)
            Exit Sub
        ElseIf runStealing(Index) = False Then
            frmMain.wsk(Index).SendData Chr$(2) & "Your attempt to steal from " & User(a).UName & " has failed." & vbCrLf & vbCrLf & Chr$(0)
            frmMain.wsk(a).SendData Chr$(2) & User(Index).UName & " has attempted to steal from you." & vbCrLf & vbCrLf & Chr$(0)
        ElseIf User(Index).Location = User(a).Location And runStealing(Index) = True Then
            c = Round(User(a).Cash * (User(Index).Stealing / 100), 0)
            frmMain.wsk(Index).SendData Chr$(2) & "You successfully stole $" & c & " from " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0)
            frmMain.wsk(Index).SendData Chr$(2) & "Someone just stole $" & c & " from you." & vbCrLf & vbCrLf & Chr$(0)
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " pick the pocket of " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0))
            User(Index).Cash = User(Index).Cash + c
            User(a).Cash = User(a).Cash - c
            Call UpdateGeneralInfo(Index)
            Call UpdateGeneralInfo(a)
        End If

End If
    
Next



End Sub
Public Sub BanUser(Index As Integer, ToKick As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer

If User(Index).AccessLevel <> 5 And User(Index).AccessLevel <> 3 Then
    frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

For b = 1 To MaxUsers
    If LCase$(Left$(User(b).UName, Len(ToKick))) = LCase$(ToKick) And User(b).Status = "Playing" Then
        Exit For
    ElseIf a = MaxUsers Then
        frmMain.wsk(Index).SendData Chr$(2) & "We cannot find that user." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
    End If
Next b

If User(b).AccessLevel = 5 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You cannot kick another admin." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

For a = 0 To UBound(IPBan)
   If IPBan(a) = "" Then
      IPBan(a) = frmMain.wsk(b).RemoteHostIP
      frmMain.wsk(b).SendData Chr$(2) & "You have been kicked and banned from the server." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(b).Close
      frmMain.lstUsers.List(b - 1) = "<Waiting>"
      If User(frmMain.lstUsers.ListIndex + 1).Status = "Playing" Then
         UserDB(User(b).DataBaseID) = User(b)
      End If
      Call ResetIndex(b)
      Call UpdatePlayerList
      Exit Sub
   ElseIf a = UBound(IPBan) Then
      ReDim Preserve IPBan(UBound(IPBan) + 1)
      IPBan(UBound(IPBan)) = frmMain.wsk(b).RemoteHostIP
      frmMain.wsk(b).SendData Chr$(2) & "You have been kicked and banned from the server." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(b).Close
      frmMain.lstUsers.List(b - 1) = "<Waiting>"
      If User(b).Status = "Playing" Then
         UserDB(User(b).DataBaseID) = User(b)
      End If
      Call ResetIndex(b)
      Call UpdatePlayerList
      Exit Sub
   End If
Next a

End Sub
Public Sub News(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer


If User(Index).AccessLevel = 0 Then Exit Sub

For a = 1 To MaxUsers
frmMain.wsk(a).SendData Chr$(252) & Chr$(3) & msg & Chr$(0)
DoEvents
Next a
End Sub
Public Sub KickUser(Index As Integer, ToKick As String)
On Error Resume Next
Dim a As Integer

If User(Index).AccessLevel <> 5 And User(Index).AccessLevel <> 3 Then
    frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

For a = 1 To MaxUsers
    If LCase$(Left$(User(a).UName, Len(ToKick))) = LCase$(ToKick) And User(a).Status = "Playing" Then
        Exit For
    ElseIf a = MaxUsers Then
        frmMain.wsk(Index).SendData Chr$(2) & "We cannot find that user." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
    End If
Next a

If User(a).AccessLevel = 5 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You cannot kick another admin." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

frmMain.wsk(a).SendData Chr$(2) & "You have been kicked from the server." & vbCrLf & vbCrLf & Chr$(0)
DoEvents
frmMain.wsk(Index).SendData Chr$(2) & "You have kicked " & User(a).UName & " from the server." & vbCrLf & vbCrLf & Chr$(0)
DoEvents
frmMain.wsk(a).Close
UserDB(User(a).DataBaseID) = User(a)
frmMain.lstUsers.List(a - 1) = "<Waiting>"
Call ResetIndex(a)
Call UpdatePlayerList
End Sub
Public Sub Gang(Index As Integer, msg As String)
On Error Resume Next

'Dim aGang As String
'Dim aTo As String
'Dim a As Integer

'aTo = Split(xmsg, ":")(0)
'aGang = Split(xmsg, ":")(1)

If Len(msg) > 10 Then Exit Sub

If User(Index).Cash < 30000 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to change your gang name." & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If
User(Index).Cash = User(Index).Cash - 30000
User(Index).HomeAbv = msg
DoEvents
frmMain.wsk(Index).SendData Chr$(2) & "Your Gang name is now " & msg & "." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)
Call UpdatePlayerList
End Sub
Public Sub ChangePassword(Index As Integer, NewPass As String)
On Error Resume Next
User(Index).UPass = NewPass
frmMain.wsk(Index).SendData Chr$(2) & "Your password has been changed to: " & NewPass & vbCrLf & vbCrLf & Chr$(0)
End Sub
Public Sub ForceUser(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer
Dim ToForce As String
Dim ToDo As String

a = InStr(1, msg, ":")
ToForce = Left$(msg, a - 1)
ToDo = Mid$(msg, a + 1)

If User(Index).AccessLevel <> 5 Then
    frmMain.wsk(Index).SendData Chr$(2) & "unknown command..." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

For a = 1 To MaxUsers
    If LCase$(Left$(User(a).UName, Len(ToForce))) = LCase$(ToForce) And User(a).Status = "Playing" And User(a).AccessLevel <> 5 And a <> Index Then
        Exit For
    ElseIf a = MaxUsers Then
        Exit Sub
    End If
Next a

Call DoCommand(a, ToDo)
End Sub
Public Sub drink(Index As Integer)
On Error Resume Next
If User(Index).Health <= 9 Then
  frmMain.wsk(Index).SendData Chr$(2) & "Dont Be Stupid That Will Kill U." & vbCrLf & vbCrLf & Chr$(0)
  DoEvents
  Exit Sub
If City(User(Index).Location).Bar = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bar first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents

  Exit Sub
End If
End If
If City(User(Index).Location).Bar = False Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to find a bar first." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
End If
If City(User(Index).Location).Bar = True Then
   If User(Index).Cash < 750 Then
      frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
   End If
If City(User(Index).Location).Bar = True Then
   If User(Index).Cash >= 750 Then
   User(Index).Cash = User(Index).Cash - 750
   User(Index).Reputation = User(Index).Reputation + 5
   User(Index).Health = User(Index).Health - 9
   User(Index).Accuracy = User(Index).Accuracy - 0.001
      frmMain.wsk(Index).SendData Chr$(2) & "You down a few cold ones.. And throw some quarters at the strippers.. In result your Reputation grows." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call UpdateGeneralInfo(Index)
   End If
   End If
   End If
   
End Sub
Public Sub GetIP(Index As Integer, ToBan As String)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim msg As String

If User(Index).AccessLevel <> 5 And User(Index).AccessLevel <> 3 Then
    frmMain.wsk(Index).SendData Chr$(2) & "unknown command!" & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

For a = 1 To MaxUsers
    If LCase$(Left$(User(a).UName, Len(ToBan))) = LCase$(ToBan) And a <> Index And User(a).Status = "Playing" Then
        Exit For
    ElseIf a = MaxUsers Then
        Exit Sub
    End If
Next a


Rem If User(Index).AccessLevel = 5 Or User(Index).AccessLevel = 3 Then
    frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & "'s IP is " & frmMain.wsk(a).RemoteHostIP & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
Rem End If

End Sub
Public Sub Transfer(Index As Integer, xMsg As String)

On Error Resume Next

Dim aCash As Long
Dim aTo As String
Dim a As Integer

aTo = Split(xMsg, ":")(0)
aCash = Split(xMsg, ":")(1)

For a = 1 To MaxUsers
    If LCase$(Left$(User(a).UName, Len(aTo))) = LCase$(aTo) And User(a).Status = "Playing" And a <> Index Then
        Exit For
    ElseIf a = MaxUsers Then
        frmMain.wsk(Index).SendData Chr$(2) & "User Not Found." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
    End If
Next a

If aCash < 1 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You cannot transfer less than $1." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

If User(Index).Cash < aCash Then
    frmMain.wsk(Index).SendData Chr$(2) & "You do not have that much cash to transfer." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

User(Index).Cash = User(Index).Cash - aCash
User(a).Cash = User(a).Cash + aCash

frmMain.wsk(Index).SendData Chr$(2) & "You just transfered $" & aCash & " to " & User(a).UName & "." & vbCrLf & vbCrLf & Chr$(0)
DoEvents
frmMain.wsk(a).SendData Chr$(2) & "You just received $" & aCash & " from " & User(Index).UName & "." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)
Call UpdateGeneralInfo(a)
End Sub


Public Sub stats(Index As Integer, xMsg As String)
On Error Resume Next
Dim a As Integer
Dim msg As Integer

'Check to make sure player access level is high enough
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If



For a = 1 To MaxUsers
    If LCase$(xMsg) = LCase$(Left$(User(a).UName, Len(xMsg))) Then
        frmMain.wsk(Index).SendData Chr$(2) & User(a).UName & "'s current skills:" & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Accuracy:    " & User(a).Accuracy & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Hiding:     " & User(a).Hiding & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Tracking:     " & User(a).Tracking & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Health:     " & User(a).Health & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Rank:     " & User(a).Reputation & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Money:     " & User(a).Cash & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Bank:     " & User(a).Bank & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Steal:     " & User(a).Stealing & vbCrLf & Chr$(0)
        frmMain.wsk(Index).SendData Chr$(2) & "Current town:     " & User(a).CurrTown & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
    
    End If
Next a

End Sub
Public Sub DickSmack(Index As Integer)
On Error Resume Next
Dim a As Integer

If Skilldelay(Index) = True Then
    Exit Sub
End If


If User(Index).AccessLevel <> 5 Then
    frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If
Call NoHiding(Index)

If User(Index).TargetNum = -1 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If

If User(Index).TargetNum >= 1 And _
   User(Index).TargetNum <= MaxUsers Then
      If User(Index).TargetGUID = User(User(Index).TargetNum).UserGUID And _
         User(User(Index).TargetNum).Status = "Playing" And _
         User(User(Index).TargetNum).Location = User(Index).Location Then
            If RunAccuracy(Index) = True Then
               User(User(Index).TargetNum).Health = 1
               If PlayerKillPlayer(Index, User(Index).TargetNum) = True Then
                  Exit Sub
               End If
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " dick smack " & User(User(Index).TargetNum).UName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You dick smack " & User(User(Index).TargetNum).UName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " dick smacks you." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call UpdateGeneralInfo(User(Index).TargetNum)
               Exit Sub
            Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               frmMain.wsk(User(Index).TargetNum).SendData Chr$(2) & User(Index).UName & " takes a swing at you but misses." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call CombatMessage(Index, User(Index).TargetNum, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & User(User(Index).TargetNum).UName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
            End If
      End If
End If

If Npc(User(Index).TargetNum).NLocation = User(Index).Location And _
   User(Index).TargetGUID = Npc(User(Index).TargetNum).NpcGUID And _
   Npc(User(Index).TargetNum).NpcGUID <> "" And _
   Npc(User(Index).TargetNum).NHealth > 0 Then
      Npc(User(Index).TargetNum).CanMove = GetTickCount()
      Npc(User(Index).TargetNum).NTargetID = Index
      Npc(User(Index).TargetNum).NTargetGUID = User(Index).UserGUID
         If RunAccuracy(Index) = True Then
            Npc(User(Index).TargetNum).NHealth = Npc(User(Index).TargetNum).NHealth - 2
               If PlayerKillNpc(Index) = True Then
                  Exit Sub
               End If
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " throw a hard punch hitting " & Npc(User(Index).TargetNum).NName & " square in the head." & vbCrLf & vbCrLf & Chr$(0))
               frmMain.wsk(Index).SendData Chr$(2) & "You land a solid punch on " & Npc(User(Index).TargetNum).NName & "." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Exit Sub
         Else
               frmMain.wsk(Index).SendData Chr$(2) & "You take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0)
               DoEvents
               Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " take a swing at " & Npc(User(Index).TargetNum).NName & " but miss." & vbCrLf & vbCrLf & Chr$(0))
               Exit Sub
         End If
End If
            
User(Index).TargetNum = -1
User(Index).TargetGUID = ""
frmMain.wsk(Index).SendData Chr$(2) & "You need to take aim on someone before you can punch them." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub NukeWarning(Index As Integer, msg As String)


End Sub

Public Sub NukeCity(Index As Integer, msg As String)
Dim citynuked As String
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer

If User(Index).Cash < 2000000 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to perform this action." & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If

If LCase$(msg) = "la" Then
    citynuked = "Los Angeles"
ElseIf LCase$(msg) = "mi" Then
    citynuked = "Miami"
ElseIf LCase$(msg) = "ho" Then
    citynuked = "Houston"
ElseIf LCase$(msg) = "ny" Then
    citynuked = "New York"
ElseIf LCase$(msg) = "ch" Then
    citynuked = "Chicago"
ElseIf LCase$(msg) = "nj" Then
    citynuked = "New Jersey"
ElseIf LCase$(msg) = "syd" Then
    citynuked = "Sydney"
ElseIf LCase$(msg) = "uk" Then
    citynuked = "London"
Else
    Exit Sub
End If

User(Index).Cash = User(Index).Cash - 2000000
Call UpdateGeneralInfo(Index)

For a = 1 To MaxUsers
If User(a).Status = "Playing" Then
    frmMain.wsk(a).SendData Chr$(252) & Chr$(3) & "<NEWS FLASH>Air raid sirens go off all over " & citynuked & "....... Citizens flee the city as a missles cruises straight for the city!" & Chr$(0)
End If
Next a
For b = 0 To MaxUsers


If User(b).CurrTown = citynuked Then
    
    frmMain.wsk(b).SendData Chr$(252) & Chr$(7) & Chr$(0)

    If User(b).Armor = -1 Then
        User(b).Health = User(b).Health = 0
        
        frmMain.wsk(b).SendData Chr$(2) & "A NUCLEAR BOMBS DROPS ON THE CITY, LAYING WASTE TO EVERYTHING AND EVERYONE!" & vbCrLf & vbCrLf & Chr$(0)
        
        DoEvents
    Else
        If Item(User(b).Armor).IName = "Biohazard Suit" Then
            User(b).Health = 1
            frmMain.wsk(b).SendData Chr$(2) & "A NUCLEAR BOMBS DROPS ON THE CITY, LAYING WASTE TO EVERYTHING AND EVERYONE!" & vbCrLf & vbCrLf & Chr$(0)
            frmMain.wsk(b).SendData Chr$(2) & "Phew!! Good thing you had your Bio Suit on." & vbCrLf & vbCrLf & Chr$(0)
            DoEvents

        Else
            User(b).Health = 0
            frmMain.wsk(b).SendData Chr$(2) & "A NUCLEAR BOMBS DROPS ON THE CITY, LAYING WASTE TO EVERYTHING AND EVERYONE!" & vbCrLf & vbCrLf & Chr$(0)
            DoEvents
       
End If
        End If
    End If

Call UpdateGeneralInfo(b)

If User(b).Health <= 0 Then
   For c = 0 To 19
      If User(b).Item(c) <> -1 Then
         For d = 0 To UBound(City(User(b).Location).CItem)
            If City(User(b).Location).CItem(d) = -1 Then
               City(User(b).Location).CItem(d) = User(b).Item(c)
               Item(User(b).Item(c)).OnPlayer = False
               Item(User(b).Item(c)).Equip = False
               Item(User(b).Item(c)).Decay = GetTickCount()
               Item(User(b).Item(c)).ItemGUID = ""
               Item(User(b).Item(c)).ILocation = User(b).Location
               User(b).Item(c) = -1
               Exit For
            ElseIf d = UBound(City(User(b).Location).CItem) Then
               With City(User(b).Location)
               ReDim Preserve .CItem(UBound(.CItem) + 1)
               .CItem(UBound(.CItem)) = User(b).Item(c)
               Item(User(b).Item(c)).OnPlayer = False
               Item(User(b).Item(c)).Equip = False
               Item(User(b).Item(c)).Decay = GetTickCount()
               Item(User(b).Item(c)).ItemGUID = ""
               Item(User(b).Item(c)).ILocation = User(b).Location
               User(b).Item(c) = -1
               End With
            End If
         Next d
      End If
   Next c
   Call FullInventoryUpdate(b)
   User(b).Reputation = User(b).Reputation - 50
   User(b).Cash = 50
   User(b).Health = 50
   frmMain.wsk(b).SendData Chr$(2) & "You die in a fireball and mushroom cloud. " & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Call PlaceOnDeath(b)
   Call UpdateGeneralInfo(b)
   User(b).Weapon = -1
   User(b).Armor = -1
   User(b).Ammo = -1
   Call UpdateGearInfo(b)
   User(b).TargetNum = -1
   User(b).TargetGUID = ""
End If

Next b

End Sub



Public Sub Bounty(Index As Integer, xMsg As String)

On Error Resume Next

Dim aCash As Long
Dim aTo As String
Dim a As Integer
Dim b As Integer
aTo = Split(xMsg, ":")(0)
aCash = Split(xMsg, ":")(1)

For a = 1 To MaxUsers
    If LCase$(Left$(User(a).UName, Len(aTo))) = LCase$(aTo) And User(a).Status = "Playing" And a <> Index Then
        Exit For
    ElseIf a = MaxUsers Then
        frmMain.wsk(Index).SendData Chr$(2) & "User Not Found." & vbCrLf & vbCrLf & Chr$(0)
        DoEvents
        Exit Sub
    End If
Next a

If aCash < 1 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You cannot bounty $1." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

If User(Index).Cash < aCash Then
    frmMain.wsk(Index).SendData Chr$(2) & "You do not have that much cash to bounty." & vbCrLf & vbCrLf & Chr$(0)
    DoEvents
    Exit Sub
End If

User(Index).Cash = User(Index).Cash - aCash
User(a).Bounty = User(a).Bounty + aCash

frmMain.wsk(Index).SendData Chr$(2) & "You just put a bounty of $" & aCash & " on " & User(a).UName & ". He now has a bounty worth $" & User(a).Bounty & vbCrLf & vbCrLf & Chr$(0)
DoEvents
 For b = 1 To MaxUsers
     If User(b).Status = "Playing" Then
        frmMain.wsk(b).SendData Chr$(252) & Chr$(3) & "<News Flash>" & User(a).UName & " from " & User(a).HomeTown & " bounty has just went up, his current bounty is at $" & User(a).Bounty & Chr$(0)
        DoEvents
    End If
    Next b

Call UpdateGeneralInfo(Index)
Call UpdateGeneralInfo(a)
End Sub


Public Sub AdminName(Index As Integer, msg As String)
On Error Resume Next

'Dim aGang As String
'Dim aTo As String
'Dim a As Integer

'aTo = Split(xmsg, ":")(0)
'aGang = Split(xmsg, ":")(1)
If User(Index).AccessLevel <> 5 Then
   frmMain.wsk(Index).SendData Chr$(2) & "Huh?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If




User(Index).UName = msg
DoEvents
frmMain.wsk(Index).SendData Chr$(2) & "Your Gang name is now " & msg & "." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

Call UpdateGeneralInfo(Index)
Call UpdatePlayerList
End Sub
