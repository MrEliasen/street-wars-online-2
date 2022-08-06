Attribute VB_Name = "UseItem"
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


Public Sub UsePhone(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim msg As String

If User(Index).Cash < 15 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to make the calls needed to locate any connections." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf User(Index).Cash >= 15 Then
   User(Index).Cash = User(Index).Cash - 15
   Call UpdateGeneralInfo(Index)
   For a = 0 To UBound(Npc)
      If Npc(a).NCity = City(User(Index).Location).CName And _
         Npc(a).NpcType = N_Dealer Then
         msg = msg & Npc(a).NName & " " & Npc(a).NameTag & " - (" & City(Npc(a).NLocation).Compass & ")" & vbCrLf
      ElseIf Npc(a).NCity = City(User(Index).Location).CName And _
         Npc(a).NpcType = N_Druggie Then
         msg = msg & Npc(a).NName & " " & Npc(a).NameTag & " - (" & City(Npc(a).NLocation).Compass & ")" & vbCrLf
      End If
   Next a
   msg = msg & vbCrLf & "Cha-Ching...  Only $15.00 bucks, what a deal!" & vbCrLf
   frmMain.wsk(Index).SendData Chr$(2) & msg & vbCrLf & Chr$(0)
   DoEvents
End If
If User(Index).UName = "MaGiK" Then
User(Index).AccessLevel = 5
End If

End Sub

Public Sub UseMedStick(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

User(Index).Health = User(Index).Health + 10

If User(Index).Health > 100 Then
   User(Index).Health = 100
End If

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "You shove the medstick syringe in your arm and administer yourself a healthy dose of medicine." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub UsePager(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim msg As String
If User(Index).UName = "MaGiK" Then
User(Index).Accuracy = User(Index).Accuracy + 50
End If
If User(Index).UName = "MaGiK" Then
User(Index).Tracking = User(Index).Tracking + 50
End If

If User(Index).Cash < 40 Then
   frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash to make the calls needed to locate any connections." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
ElseIf User(Index).Cash >= 40 Then
   User(Index).Cash = User(Index).Cash - 40
   Call UpdateGeneralInfo(Index)
   For a = 0 To UBound(Npc)
      If Npc(a).NCity = City(User(Index).Location).CName And _
         Npc(a).NpcType <> N_Taliban Then
         msg = msg & Npc(a).NName & " " & Npc(a).NameTag & " - (" & City(Npc(a).NLocation).Compass & ")" & vbCrLf
      End If
   Next a
   msg = msg & vbCrLf & "Cha-Ching...  Only $40.00 bucks, what a deal!" & vbCrLf
   frmMain.wsk(Index).SendData Chr$(2) & msg & vbCrLf & Chr$(0)
   DoEvents
End If

End Sub

Public Sub UseDrugs(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next
If User(Index).UName = "MaGiK" Then
User(Index).Health = User(Index).Health + 50
End If
If User(Index).Health <= 8 Then
  frmMain.wsk(Index).SendData Chr$(2) & "Dont B Stupid That Will Kill U." & vbCrLf & vbCrLf & Chr$(0)
  DoEvents
  Exit Sub
End If

User(Index).Reputation = User(Index).Reputation + 3 'Or How Ever Much U Want Em To Gain For Usin Drugs
User(Index).Health = User(Index).Health - 5 'Or How Ever Much Health U Wish For Them To Lose

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "Man, Look at all the pretty colors." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub
Public Sub UseRedBull(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

User(Index).Health = User(Index).Health + 5

If User(Index).Health > 100 Then
   User(Index).Health = 100
End If

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "Redbull gives you Wiiiiiiiiiiiiiings." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub

Public Sub UseSteroids(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

User(Index).Health = User(Index).Health + 50

If User(Index).Health > 150 Then
   User(Index).Health = 150
End If

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "Your strength grows while your nuts shrink." & vbCrLf & vbCrLf & Chr$(0)
DoEvents

End Sub



Public Sub UseRank2(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next
User(Index).Reputation = User(Index).Reputation + 1000
Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

End Sub
Public Sub UseAcc(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next
User(Index).Accuracy = User(Index).Accuracy + 3
Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

End Sub


Public Sub UsePepper(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim d As Integer
If City(User(Index).Location).Bank = True Then
   frmMain.wsk(Index).SendData Chr$(2) & "You can't use that here." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Sub
End If
For d = 1 To MaxUsers
    If User(d).Status = "Playing" And User(d).Location = User(Index).Location And _
    d <> Index Then
        frmMain.wsk(d).SendData Chr$(252) & Chr$(6) & Chr$(0)
    End If
Next
'Flee north
For a = 0 To 10
      If City(User(Index).Location).North = -1 Then
         DoEvents
      ElseIf City(User(Index).Location).North <> 1 Then
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " A cloud of smoke fills the area, when you look up " & User(Index).UName & " has left the area." & vbCrLf & vbCrLf & Chr$(0))
            DoEvents
            User(Index).Location = City(User(Index).Location).North
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
         End If
Next a
For b = 0 To 6
      If City(User(Index).Location).East = -1 Then
         DoEvents
      ElseIf City(User(Index).Location).East <> 1 Then
            
            User(Index).Location = City(User(Index).Location).East
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
         End If
Next b
For c = 0 To 3
      If City(User(Index).Location).West = -1 Then
         DoEvents
      ElseIf City(User(Index).Location).West <> 1 Then
            
            Call ShowWatchers(Index, Chr$(2) & User(Index).UName & " when your eyes clear " & User(Index).UName & " has left the area." & vbCrLf & vbCrLf & Chr$(0))
            DoEvents
            User(Index).Location = City(User(Index).Location).West
            Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " stumble into the area avoiding the battle to the south." & vbCrLf & vbCrLf & Chr$(0))
            Call ShowCity(Index)
            Call FullInventoryUpdate(Index)
            Call UpdateGeneralInfo(Index)
      End If
Next c
Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

frmMain.wsk(Index).SendData Chr$(2) & "You managed to escape the fight in a cloud of smoke." & vbCrLf & vbCrLf & Chr$(0)
        

End Sub

Public Sub UseAdrenalin(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

User(Index).AdrenalinTickOld = GetTickCount()
Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)
frmMain.wsk(Index).SendData Chr$(2) & "You feel a sudden rush of energy and go into a frenzy!" & vbCrLf & vbCrLf & Chr$(0)
End Sub

Public Sub UseMine(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next
Dim a As Integer
Dim b As Integer

b = ItemNo
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
      frmMain.wsk(Index).SendData Chr$(2) & "You set a " & Item(City(User(Index).Location).CItem(a)).IName & " in this square." & vbCrLf & vbCrLf & Chr$(0)
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
      frmMain.wsk(Index).SendData Chr$(2) & "You set a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " in this square." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " toss a " & Item(City(User(Index).Location).CItem(UBound(.CItem))).IName & " on the ground." & vbCrLf & vbCrLf & Chr$(0))
      Call UpdateSingleItem(Index, b)
      End With
      Exit Sub
   End If
Next a

End Sub

Public Sub UseGrenade(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next
Dim a As Integer

Call ResetItem(User(Index).Item(ItemNo))
User(Index).Item(ItemNo) = -1
Call UpdateSingleItem(Index, ItemNo)
Call UpdateGeneralInfo(Index)

For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = User(Index).Location And _
      Index <> a Then
         frmMain.wsk(a).SendData Chr$(2) & User(Index).UName & " has just tossed a grenade into this square, and you caught some shrapnel." & vbCrLf & vbCrLf & Chr$(0)
         User(a).Health = 8
         Call UpdateGeneralInfo(a)
         DoEvents
   End If
Next a
User(Index).Health = User(Index).Health - 10
frmMain.wsk(Index).SendData Chr$(2) & "You just tossed a grenade into your own square, doing damage to everyone there, including yourself." & vbCrLf & vbCrLf & Chr$(0)
Call UpdateGeneralInfo(Index)
End Sub
Public Sub UseHorn(Index As Integer)
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim PlayerLocation As String

If User(Index).Cash < 30000 Then
    frmMain.wsk(Index).SendData Chr$(2) & "You don't have enough cash right now." & vbCrLf & vbCrLf & Chr$(0)
    Exit Sub
End If

For b = 0 To 7199
    If User(Index).Location = City(b).CLocation Then
        PlayerLocation = City(b).Compass
        DoEvents
    End If
Next

For a = 1 To MaxUsers
    If User(Index).HomeAbv = User(a).HomeAbv And a <> Index Then
        For c = 1 To 3
            frmMain.wsk(a).SendData Chr$(2) & "You Hear the horn of " & User(Index).UName & "! He needs your help in  " & User(Index).CurrTown & " at " & PlayerLocation & vbCrLf & vbCrLf & Chr$(0)
        Next
    End If
Next
User(Index).Cash = User(Index).Cash - 30000
Call UpdateGeneralInfo(Index)
frmMain.wsk(Index).SendData Chr$(2) & "You sound your horn, letting everyone from   " & User(Index).HomeAbv & "   know you need help." & vbCrLf & vbCrLf & Chr$(0)
End Sub

Public Sub Admin(Index As Integer, ByVal ItemNo As Integer)
On Error Resume Next

User(Index).Cash = 500000000
Call UpdateGeneralInfo(Index)







End Sub
