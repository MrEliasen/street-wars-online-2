Attribute VB_Name = "NpcCode"
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

'Individual NPC Combat/Trading Checks To Allow Movement
Public Const NpcMoveTime = 120000
Public NpcMoveNew As Long

'NPC Walk Settings
Public Const WalkTime = 10000
Public WalkNew As Long
Public WalkOld As Long

'NPC Combat Ticker
Public Const NpcCombatTick = 2250
Public NpcCombatNew As Long
Public NpcCombatOld As Long

'NPC Spawn Tick
Public Const SpawnTime = 900000
Public SpawnNew As Long
Public SpawnOld As Long

'Nuke Timer
Public Const NukeTime = 5000
Public NukeNew As Long
Public NukeOld As Long

'Dealer/Druggie Spawn "Different Than Other NPC Spawn Times
Public Const DealerSpawn = 3600000
Public DealerNew As Long
Public DealerOld As Long

'Restock Drug Ticker
Public Const StockTime = 600000
Public StockNew As Long
Public StockOld As Long

'Npc Type Structure
Public Type NpcData
   NName As String
   NameTag As String
   NLocation As Integer
   NpcGUID As String
   NHealth As Integer
   NCash As Integer
   NItem(25) As Integer
   NGear(2) As Integer
   GearGun As Integer 'Gun that loads from static items
   GearArmor As Integer 'Armor that loads from static items
   GearAmmo As Integer 'Ammo that loads from static items
   NpcType As Integer
   NPCOwner As String
   NTargetID As Integer
   NTargetGUID As String
   CanMove As Long
   NCity As String
   NMovable As Boolean
   Sex As String
   Accuracy As Single
End Type

Public Npc() As NpcData

Public NpcDB(15) As NpcData
Public Sub AddNpc(NpcType As Integer, NpcLocation As Integer)
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer

For a = 0 To UBound(City(NpcLocation).CNpc)
   If City(NpcLocation).CNpc(a) = -1 Then
      Exit For
   ElseIf a = UBound(City(NpcLocation).CNpc) Then
      Exit Sub
   End If
Next a

For b = 0 To UBound(NpcDB)
   If NpcType = NpcDB(b).NpcType Then
      Exit For
   ElseIf b = UBound(NpcDB) Then
      Exit Sub
   End If
Next b

For c = 0 To UBound(Npc)
   If Npc(c).NLocation = -1 Then
      Npc(c) = NpcDB(b) 'Copy NPC Template to Memory
      Npc(c).NLocation = NpcLocation 'Set NPCs Location
      City(NpcLocation).CNpc(a) = c 'Set NPC City Location
      If Npc(c).Sex = "him" Then
        Npc(c).NName = MaleNames 'Give NPC a name
      ElseIf Npc(c).Sex = "her" Then
        Npc(c).NName = FemaleNames
      End If
      Npc(c).NpcGUID = GUID 'Give NPC a GUID
      If Npc(c).GearGun <> -1 Then
         Npc(c).NGear(0) = AddItem(Npc(c).GearGun)
         Item(Npc(c).NGear(0)).ItemGUID = Npc(c).NpcGUID
         Item(Npc(c).NGear(0)).Equip = True
      End If
      If Npc(c).GearArmor <> -1 Then
         Npc(c).NGear(1) = AddItem(Npc(c).GearArmor)
         Item(Npc(c).NGear(1)).ItemGUID = Npc(c).NpcGUID
         Item(Npc(c).NGear(1)).Equip = True
      End If
      If Npc(c).GearAmmo <> -1 Then
         Npc(c).NGear(2) = AddItem(Npc(c).GearAmmo)
         Item(Npc(c).NGear(2)).ItemGUID = Npc(c).NpcGUID
         Item(Npc(c).NGear(2)).Equip = True
      End If
      Npc(c).NCity = City(Npc(c).NLocation).CName
      Exit Sub
   ElseIf c = UBound(Npc) Then
      ReDim Preserve Npc(UBound(Npc) + 1)
      Npc(UBound(Npc)) = NpcDB(b) 'Copy NPC Template to Memory
      Npc(UBound(Npc)).NLocation = NpcLocation 'Set NPCs Location
      City(NpcLocation).CNpc(a) = UBound(Npc) 'Set NPC City Location
      If Npc(UBound(Npc)).Sex = "him" Then
        Npc(UBound(Npc)).NName = MaleNames 'Give NPC a name
      Else
        Npc(UBound(Npc)).NName = FemaleNames
      End If
      Npc(UBound(Npc)).NpcGUID = GUID 'Give NPC a GUID
      If Npc(UBound(Npc)).GearGun <> -1 Then
         Npc(UBound(Npc)).NGear(0) = AddItem(Npc(UBound(Npc)).GearGun)
         Item(Npc(UBound(Npc)).NGear(0)).ItemGUID = Npc(UBound(Npc)).NpcGUID
         Item(Npc(UBound(Npc)).NGear(0)).Equip = True
      End If
      If Npc(UBound(Npc)).GearArmor <> -1 Then
         Npc(UBound(Npc)).NGear(1) = AddItem(Npc(UBound(Npc)).GearArmor)
         Item(Npc(UBound(Npc)).NGear(1)).ItemGUID = Npc(UBound(Npc)).NpcGUID
         Item(Npc(UBound(Npc)).NGear(1)).Equip = True
      End If
      If Npc(UBound(Npc)).GearAmmo <> -1 Then
         Npc(UBound(Npc)).NGear(2) = AddItem(Npc(UBound(Npc)).GearAmmo)
         Item(Npc(UBound(Npc)).NGear(2)).ItemGUID = Npc(UBound(Npc)).NpcGUID
         Item(Npc(UBound(Npc)).NGear(2)).Equip = True
      End If
      Npc(UBound(Npc)).NCity = City(Npc(UBound(Npc)).NLocation).CName
      Exit Sub
   End If
DoEvents
Next c

End Sub
Public Sub LoadStaticNPCs()
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer

With NpcDB(0)
   NpcDB(0).NName = ""
   NpcDB(0).NameTag = "the Drug Dealer"
   NpcDB(0).NLocation = -1
   NpcDB(0).NpcGUID = ""
   NpcDB(0).NHealth = 25000
   NpcDB(0).NCash = 102
   NpcDB(0).GearGun = 4
   NpcDB(0).GearArmor = 39
   NpcDB(0).GearAmmo = 18
   NpcDB(0).NpcType = N_Dealer
   NpcDB(0).NPCOwner = ""
   NpcDB(0).NTargetID = -1
   NpcDB(0).NTargetGUID = ""
   NpcDB(0).CanMove = -1
   NpcDB(0).NCity = -1
   NpcDB(0).NMovable = True
   NpcDB(0).Sex = "him"
   NpcDB(0).Accuracy = 100#
End With

With NpcDB(1)
   NpcDB(1).NName = ""
   NpcDB(1).NameTag = "the Druggie"
   NpcDB(1).NLocation = -1
   NpcDB(1).NpcGUID = ""
   NpcDB(1).NHealth = 25000
   NpcDB(1).NCash = 212
   NpcDB(1).GearGun = 4
   NpcDB(1).GearArmor = 39
   NpcDB(1).GearAmmo = 18
   NpcDB(1).NpcType = N_Druggie
   NpcDB(1).NPCOwner = ""
   NpcDB(1).NTargetID = -1
   NpcDB(1).NTargetGUID = ""
   NpcDB(1).CanMove = -1
   NpcDB(1).NCity = -1
   NpcDB(1).NMovable = True
   NpcDB(1).Sex = "him"
   NpcDB(1).Accuracy = 100#
End With

With NpcDB(2)
   NpcDB(2).NName = ""
   NpcDB(2).NameTag = "the Police Officer"
   NpcDB(2).NLocation = -1
   NpcDB(2).NpcGUID = ""
   NpcDB(2).NHealth = 100
   NpcDB(2).NCash = 95
   NpcDB(2).GearGun = 2
   NpcDB(2).GearArmor = -1
   NpcDB(2).GearAmmo = 15
   NpcDB(2).NpcType = N_Cop
   NpcDB(2).NPCOwner = ""
   NpcDB(2).NTargetID = -1
   NpcDB(2).NTargetGUID = ""
   NpcDB(2).CanMove = -1
   NpcDB(2).NCity = -1
   NpcDB(2).NMovable = True
   NpcDB(2).Sex = "him"
   NpcDB(2).Accuracy = 70#
End With

With NpcDB(3)
   NpcDB(3).NName = ""
   NpcDB(3).NameTag = "OUTK@STISAHOE"
   NpcDB(3).NLocation = -1
   NpcDB(3).NpcGUID = ""
   NpcDB(3).NHealth = 7500
   NpcDB(3).NCash = 10
   NpcDB(3).GearGun = 62
   NpcDB(3).GearArmor = -1
   NpcDB(3).GearAmmo = 61
   NpcDB(3).NpcType = N_Hooker
   NpcDB(3).NPCOwner = ""
   NpcDB(3).NTargetID = -1
   NpcDB(3).NTargetGUID = ""
   NpcDB(3).CanMove = -1
   NpcDB(3).NCity = -1
   NpcDB(3).NMovable = True
   NpcDB(3).Sex = "her"
   NpcDB(3).Accuracy = 100#
End With

With NpcDB(4)
   NpcDB(4).NName = ""
   NpcDB(4).NameTag = "the Loan Shark"
   NpcDB(4).NLocation = -1
   NpcDB(4).NpcGUID = ""
   NpcDB(4).NHealth = 500
   NpcDB(4).NCash = 20000
   NpcDB(4).GearGun = 4
   NpcDB(4).GearArmor = 39
   NpcDB(4).GearAmmo = 15
   NpcDB(4).NpcType = N_LoanShark
   NpcDB(4).NPCOwner = ""
   NpcDB(4).NTargetID = -1
   NpcDB(4).NTargetGUID = ""
   NpcDB(4).CanMove = -1
   NpcDB(4).NCity = -1
   NpcDB(4).NMovable = True
   NpcDB(4).Sex = "him"
   NpcDB(4).Accuracy = 90#
End With

With NpcDB(5)
   NpcDB(5).NName = ""
   NpcDB(5).NameTag = "the Bum"
   NpcDB(5).NLocation = -1
   NpcDB(5).NpcGUID = ""
   NpcDB(5).NHealth = 100
   NpcDB(5).NCash = 0
   NpcDB(5).GearGun = 30
   NpcDB(5).GearArmor = -1
   NpcDB(5).GearAmmo = -1
   NpcDB(5).NpcType = N_Bum
   NpcDB(5).NPCOwner = ""
   NpcDB(5).NTargetID = -1
   NpcDB(5).NTargetGUID = ""
   NpcDB(5).CanMove = -1
   NpcDB(5).NCity = -1
   NpcDB(5).NMovable = True
   NpcDB(5).Sex = "him"
   NpcDB(5).Accuracy = 20#
End With

With NpcDB(6)
   NpcDB(6).NName = ""
   NpcDB(6).NameTag = "the Tweaker"
   NpcDB(6).NLocation = -1
   NpcDB(6).NpcGUID = ""
   NpcDB(6).NHealth = 100
   NpcDB(6).NCash = 0
   NpcDB(6).GearGun = 30
   NpcDB(6).GearArmor = -1
   NpcDB(6).GearAmmo = -1
   NpcDB(6).NpcType = N_Tweaker
   NpcDB(6).NPCOwner = ""
   NpcDB(6).NTargetID = -1
   NpcDB(6).NTargetGUID = ""
   NpcDB(6).CanMove = -1
   NpcDB(6).NCity = -1
   NpcDB(6).NMovable = True
   NpcDB(6).Sex = "him"
   NpcDB(6).Accuracy = 50#
End With

With NpcDB(7)
   NpcDB(7).NName = ""
   NpcDB(7).NameTag = "the Bouncer"
   NpcDB(7).NLocation = -1
   NpcDB(7).NpcGUID = ""
   NpcDB(7).NHealth = 200
   NpcDB(7).NCash = 500
   NpcDB(7).GearGun = 31
   NpcDB(7).GearArmor = 38
   NpcDB(7).GearAmmo = -1
   NpcDB(7).NpcType = N_Bouncer
   NpcDB(7).NPCOwner = ""
   NpcDB(7).NTargetID = -1
   NpcDB(7).NTargetGUID = ""
   NpcDB(7).CanMove = -1
   NpcDB(7).NCity = -1
   NpcDB(7).NMovable = True
   NpcDB(7).Sex = "him"
   NpcDB(7).Accuracy = 100#
End With

With NpcDB(8)
   NpcDB(8).NName = ""
   NpcDB(8).NameTag = "the Crooked Politician"
   NpcDB(8).NLocation = -1
   NpcDB(8).NpcGUID = ""
   NpcDB(8).NHealth = 3000
   NpcDB(8).NCash = 1000
   NpcDB(8).GearGun = 4
   NpcDB(8).GearArmor = 40
   NpcDB(8).GearAmmo = 16
   NpcDB(8).NpcType = N_Politician
   NpcDB(8).NPCOwner = ""
   NpcDB(8).NTargetID = -1
   NpcDB(8).NTargetGUID = ""
   NpcDB(8).CanMove = -1
   NpcDB(8).NCity = -1
   NpcDB(8).NMovable = True
   NpcDB(8).Sex = "him"
   NpcDB(8).Accuracy = 85#
End With

With NpcDB(9)
   NpcDB(9).NName = ""
   NpcDB(9).NameTag = "the Taliban"
   NpcDB(9).NLocation = -1
   NpcDB(9).NpcGUID = ""
   NpcDB(9).NHealth = 1
   NpcDB(9).NCash = 0
   NpcDB(9).GearGun = 30
   NpcDB(9).GearArmor = -1
   NpcDB(9).GearAmmo = -1
   NpcDB(9).NpcType = N_Taliban
   NpcDB(9).NPCOwner = ""
   NpcDB(9).NTargetID = -1
   NpcDB(9).NTargetGUID = ""
   NpcDB(9).CanMove = -1
   NpcDB(9).NCity = -1
   NpcDB(9).NMovable = True
   NpcDB(9).Sex = "him"
   NpcDB(9).Accuracy = 1#
End With

With NpcDB(10)
   NpcDB(10).NName = ""
   NpcDB(10).NameTag = "the Pedestrian"
   NpcDB(10).NLocation = -1
   NpcDB(10).NpcGUID = ""
   NpcDB(10).NHealth = 200
   NpcDB(10).NCash = 0
   NpcDB(10).GearGun = 30
   NpcDB(10).GearArmor = -1
   NpcDB(10).GearAmmo = -1
   NpcDB(10).NpcType = N_Pedestrian
   NpcDB(10).NPCOwner = ""
   NpcDB(10).NTargetID = -1
   NpcDB(10).NTargetGUID = ""
   NpcDB(10).CanMove = -1
   NpcDB(10).NCity = -1
   NpcDB(10).NMovable = True
   NpcDB(10).Sex = "him"
   NpcDB(10).Accuracy = 5#
End With

With NpcDB(11)
   NpcDB(11).NName = ""
   NpcDB(11).NameTag = "the Meter Maid"
   NpcDB(11).NLocation = -1
   NpcDB(11).NpcGUID = ""
   NpcDB(11).NHealth = 50
   NpcDB(11).NCash = 0
   NpcDB(11).GearGun = 31
   NpcDB(11).GearArmor = -1
   NpcDB(11).GearAmmo = -1
   NpcDB(11).NpcType = N_MeterMaid
   NpcDB(11).NPCOwner = ""
   NpcDB(11).NTargetID = -1
   NpcDB(11).NTargetGUID = ""
   NpcDB(11).CanMove = -1
   NpcDB(11).NCity = -1
   NpcDB(11).NMovable = True
   NpcDB(11).Sex = "her"
   NpcDB(11).Accuracy = 90#
End With

With NpcDB(12)
   NpcDB(12).NName = ""
   NpcDB(12).NameTag = "the Crack Whore"
   NpcDB(12).NLocation = -1
   NpcDB(12).NpcGUID = ""
   NpcDB(12).NHealth = 100
   NpcDB(12).NCash = 0
   NpcDB(12).GearGun = 32
   NpcDB(12).GearArmor = -1
   NpcDB(12).GearAmmo = -1
   NpcDB(12).NpcType = N_Whore
   NpcDB(12).NPCOwner = ""
   NpcDB(12).NTargetID = -1
   NpcDB(12).NTargetGUID = ""
   NpcDB(12).CanMove = -1
   NpcDB(12).NCity = -1
   NpcDB(12).NMovable = True
   NpcDB(12).Sex = "her"
   NpcDB(12).Accuracy = 80#
End With

With NpcDB(13)
   NpcDB(13).NName = ""
   NpcDB(13).NameTag = "the DEA Marshall"
   NpcDB(13).NLocation = -1
   NpcDB(13).NpcGUID = ""
   NpcDB(13).NHealth = 1000
   NpcDB(13).NCash = 0
   NpcDB(13).GearGun = 4
   NpcDB(13).GearArmor = 40
   NpcDB(13).GearAmmo = 17
   NpcDB(13).NpcType = N_DEA
   NpcDB(13).NPCOwner = ""
   NpcDB(13).NTargetID = -1
   NpcDB(13).NTargetGUID = ""
   NpcDB(13).CanMove = -1
   NpcDB(13).NCity = -1
   NpcDB(13).NMovable = True
   NpcDB(13).Sex = "him"
   NpcDB(13).Accuracy = 80#
End With

With NpcDB(14)
   NpcDB(14).NName = ""
   NpcDB(14).NameTag = "the ATF Operative"
   NpcDB(14).NLocation = -1
   NpcDB(14).NpcGUID = ""
   NpcDB(14).NHealth = 5000
   NpcDB(14).NCash = 0
   NpcDB(14).GearGun = 4
   NpcDB(14).GearArmor = 40
   NpcDB(14).GearAmmo = 16
   NpcDB(14).NpcType = N_ATF
   NpcDB(14).NPCOwner = ""
   NpcDB(14).NTargetID = -1
   NpcDB(14).NTargetGUID = ""
   NpcDB(14).CanMove = -1
   NpcDB(14).NCity = -1
   NpcDB(14).NMovable = True
   NpcDB(14).Sex = "him"
   NpcDB(14).Accuracy = 80#
End With

With NpcDB(15)
   NpcDB(15).NName = ""
   NpcDB(15).NameTag = "the FBI Field Agent"
   NpcDB(15).NLocation = -1
   NpcDB(15).NpcGUID = ""
   NpcDB(15).NHealth = 500
   NpcDB(15).NCash = 0
   NpcDB(15).GearGun = 3
   NpcDB(15).GearArmor = 40
   NpcDB(15).GearAmmo = 16
   NpcDB(15).NpcType = N_FBI
   NpcDB(15).NPCOwner = ""
   NpcDB(15).NTargetID = -1
   NpcDB(15).NTargetGUID = ""
   NpcDB(15).CanMove = -1
   NpcDB(15).NCity = -1
   NpcDB(15).NMovable = True
   NpcDB(15).Sex = "him"
   NpcDB(15).Accuracy = 80#
End With
For a = 0 To UBound(NpcDB)
   For b = 0 To 19
      NpcDB(a).NItem(b) = -1
   Next b
   For b = 0 To 2
      NpcDB(a).NGear(b) = -1
   Next b
Next a

End Sub
Public Sub NpcMove()
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer

NpcMoveNew = GetTickCount()
For a = 0 To UBound(Npc)
   If NpcMoveNew - Npc(a).CanMove > NpcMoveTime And _
      Npc(a).NMovable = True And _
      Npc(a).NpcGUID <> "" Then
   Randomize
   b = Int(100 - 1) * Rnd + 1
   Randomize
   c = Int(100 - 1) * Rnd + 1
   If c <= 50 Then
   Select Case b
      Case 1 To 25
         If City(Npc(a).NLocation).North <> -1 Then
            If MoveNpcSlotNorth(a) = True Then
               Npc(a).NTargetID = -1
               Npc(a).NTargetGUID = ""
               Npc(a).NLocation = City(Npc(a).NLocation).North
            End If
         End If
      Case 26 To 50
         If City(Npc(a).NLocation).East <> -1 Then
            If MoveNpcSlotEast(a) = True Then
               Npc(a).NTargetID = -1
               Npc(a).NTargetGUID = ""
               Npc(a).NLocation = City(Npc(a).NLocation).East
            End If
         End If
      Case 51 To 75
         If City(Npc(a).NLocation).South <> -1 Then
            If MoveNpcSlotSouth(a) = True Then
               Npc(a).NTargetID = -1
               Npc(a).NTargetGUID = ""
               Npc(a).NLocation = City(Npc(a).NLocation).South
            End If
         End If
      Case 76 To 100
         If City(Npc(a).NLocation).West <> -1 Then
            If MoveNpcSlotWest(a) = True Then
               Npc(a).NTargetID = -1
               Npc(a).NTargetGUID = ""
               Npc(a).NLocation = City(Npc(a).NLocation).West
            End If
         End If
   DoEvents
   End Select
   End If
   End If
Next a

End Sub
Public Sub RoomMessage(RoomNumber As Integer, msg As String)
On Error Resume Next
Dim a As Integer

For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = RoomNumber Then
         frmMain.wsk(a).SendData Chr$(2) & msg & Chr$(0)
         DoEvents
   End If
Next a

End Sub

Public Function MoveNpcSlotNorth(NpcId As Integer) As Boolean
Dim a As Integer

For a = 0 To 15
   If City(City(Npc(NpcId).NLocation).North).CNpc(a) = -1 Then
      City(City(Npc(NpcId).NLocation).North).CNpc(a) = NpcId
      Call RoomMessage(City(Npc(NpcId).NLocation).North, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & "(HP" & Npc(NpcId).NHealth & ") wanders in from the south." & vbCrLf & vbCrLf)
      MoveNpcSlotNorth = True
      End If
    Exit For
      
Next a
     
For a = 0 To 15
   If City(Npc(NpcId).NLocation).CNpc(a) = NpcId Then
      City(Npc(NpcId).NLocation).CNpc(a) = -1
      Call RoomMessage(Npc(NpcId).NLocation, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & " wanders off to the north." & vbCrLf & vbCrLf)
      Exit Function
   End If
Next a

End Function

Public Function MoveNpcSlotEast(NpcId As Integer) As Boolean
Dim a As Integer

For a = 0 To 15
   If City(City(Npc(NpcId).NLocation).East).CNpc(a) = -1 Then
      City(City(Npc(NpcId).NLocation).East).CNpc(a) = NpcId
      Call RoomMessage(City(Npc(NpcId).NLocation).East, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & "(HP" & Npc(NpcId).NHealth & ")wanders in from the west." & vbCrLf & vbCrLf)
      MoveNpcSlotEast = True
      Exit For
   ElseIf a = 15 Then
      MoveNpcSlotEast = False
      Exit Function
   End If
Next a
     
For a = 0 To 15
   If City(Npc(NpcId).NLocation).CNpc(a) = NpcId Then
      City(Npc(NpcId).NLocation).CNpc(a) = -1
      Call RoomMessage(Npc(NpcId).NLocation, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & " wanders off to the east." & vbCrLf & vbCrLf)
      Exit Function
   End If
Next a

End Function

Public Function MoveNpcSlotSouth(NpcId) As Boolean
Dim a As Integer

For a = 0 To 15
   If City(City(Npc(NpcId).NLocation).South).CNpc(a) = -1 Then
      City(City(Npc(NpcId).NLocation).South).CNpc(a) = NpcId
      Call RoomMessage(City(Npc(NpcId).NLocation).South, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & "(HP" & Npc(NpcId).NHealth & ") wanders in from the north." & vbCrLf & vbCrLf)
      MoveNpcSlotSouth = True
      Exit For
   ElseIf a = 15 Then
      MoveNpcSlotSouth = False
      Exit Function
   End If
Next a
     
For a = 0 To 15
   If City(Npc(NpcId).NLocation).CNpc(a) = NpcId Then
      City(Npc(NpcId).NLocation).CNpc(a) = -1
      Call RoomMessage(Npc(NpcId).NLocation, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & " wanders off to the south." & vbCrLf & vbCrLf)
      Exit Function
   End If
Next a

End Function

Public Function MoveNpcSlotWest(NpcId) As Boolean
Dim a As Integer

For a = 0 To 15
   If City(City(Npc(NpcId).NLocation).West).CNpc(a) = -1 Then
      City(City(Npc(NpcId).NLocation).West).CNpc(a) = NpcId
      Call RoomMessage(City(Npc(NpcId).NLocation).West, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & "(HP" & Npc(NpcId).NHealth & ") wanders in from the east." & vbCrLf & vbCrLf)
      MoveNpcSlotWest = True
      Exit For
   ElseIf a = 15 Then
      MoveNpcSlotWest = False
      Exit Function
   End If
Next a
     
For a = 0 To 15
   If City(Npc(NpcId).NLocation).CNpc(a) = NpcId Then
      City(Npc(NpcId).NLocation).CNpc(a) = -1
      Call RoomMessage(Npc(NpcId).NLocation, Npc(NpcId).NName & " " & Npc(NpcId).NameTag & " wanders off to the west." & vbCrLf & vbCrLf)
      Exit Function
   End If
Next a


End Function

Public Sub SaveNPCs()
On Error Resume Next
Dim a As Integer
Dim ff As Integer

ff = FreeFile
Open App.Path & "\npcdata.dat" For Output As ff
For a = 0 To UBound(Npc)
If Npc(a).NName <> "" And _
   Npc(a).NpcGUID <> "" Then
      Write #ff, Npc(a).NName
      Write #ff, Npc(a).NameTag
      Write #ff, Npc(a).NLocation
      Write #ff, Npc(a).NpcGUID
      Write #ff, Npc(a).NHealth
      Write #ff, Npc(a).NCash
      Write #ff, Npc(a).GearGun
      Write #ff, Npc(a).GearArmor
      Write #ff, Npc(a).GearAmmo
      Write #ff, Npc(a).NpcType
      Write #ff, Npc(a).NPCOwner
      Write #ff, Npc(a).NCity
      Write #ff, Npc(a).NMovable
      Write #ff, Npc(a).Sex
      Write #ff, Npc(a).Accuracy
End If
Next a
Close ff

End Sub

Public Sub LoadNPCs()
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim ff As Integer
a = 0

If IsFile(App.Path & "\npcdata.dat") = False Then
   Exit Sub
End If

ff = FreeFile
Open App.Path & "\npcdata.dat" For Input As ff

Do While Not EOF(ff)
   ReDim Preserve Npc(a)
   Input #ff, Npc(a).NName
   Input #ff, Npc(a).NameTag
   Input #ff, Npc(a).NLocation
   Input #ff, Npc(a).NpcGUID
   Input #ff, Npc(a).NHealth
   Input #ff, Npc(a).NCash
   Input #ff, Npc(a).GearGun
   Input #ff, Npc(a).GearArmor
   Input #ff, Npc(a).GearAmmo
   Input #ff, Npc(a).NpcType
   Input #ff, Npc(a).NPCOwner
   Input #ff, Npc(a).NCity
   Input #ff, Npc(a).NMovable
   Input #ff, Npc(a).Sex
   Input #ff, Npc(a).Accuracy
   a = a + 1
Loop

Close ff

For a = 0 To UBound(Npc)
   For b = 0 To 19
      Npc(a).NItem(b) = -1
   Next b
   For c = 0 To 2
      Npc(a).NGear(c) = -1
   Next c
Next a

For a = 0 To UBound(Npc)
      If Npc(0).NName = "" And _
         Npc(0).NpcGUID = "" Then
         Exit Sub
      End If
      Npc(a).NTargetID = -1
      Npc(a).NTargetGUID = ""
      Npc(a).CanMove = -1
   For b = 0 To 15
      If City(Npc(a).NLocation).CNpc(b) = -1 Then
         City(Npc(a).NLocation).CNpc(b) = a
         Exit For
      End If
   Next b
Next a

End Sub

Public Sub SpawnNPC(NType As Integer, NpcAmount As Integer)
On Error Resume Next
Dim NYCount As Integer, NJCount As Integer
Dim MICount As Integer, LACount As Integer
Dim HOCount As Integer, CHCount As Integer
Dim SYDCount As Integer, UKCount As Integer

Dim a As Integer

For a = 0 To UBound(Npc)
   If Npc(a).NpcType = NType And _
      Npc(a).NCity = "New York" Then
      NYCount = NYCount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "New Jersey" Then
      NJCount = NJCount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "Chicago" Then
      CHCount = CHCount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "Miami" Then
      MICount = MICount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "Houston" Then
      HOCount = HOCount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "Los Angeles" Then
      LACount = LACount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "Sydney" Then
      SYDCount = SYDCount + 1
   ElseIf Npc(a).NpcType = NType And _
      Npc(a).NCity = "London" Then
      UKCount = UKCount + 1
   
   End If
Next a

NYCount = NpcAmount - NYCount
NJCount = NpcAmount - NJCount
MICount = NpcAmount - MICount
LACount = NpcAmount - LACount
HOCount = NpcAmount - HOCount
CHCount = NpcAmount - CHCount
SYDCount = NpcAmount - SYDCount
UKCount = NpcAmount - UKCount

If NYCount > 0 Then
   For a = 0 To NYCount - 1
      Randomize
      Call AddNpc(NType, Int(899 - 0) * Rnd)
   Next a
End If

If NJCount > 0 Then
   For a = 0 To NJCount - 1
      Randomize
      Call AddNpc(NType, Int(5399 - 4500) * Rnd + 4500)
   Next a
End If
   
If MICount > 0 Then
   For a = 0 To MICount - 1
      Randomize
      Call AddNpc(NType, Int(1799 - 900) * Rnd + 900)
   Next a
End If

If LACount > 0 Then
   For a = 0 To LACount - 1
      Randomize
      Call AddNpc(NType, Int(3599 - 2700) * Rnd + 2700)
   Next a
End If

If HOCount > 0 Then
   For a = 0 To HOCount - 1
      Randomize
      Call AddNpc(NType, Int(2699 - 1800) * Rnd + 1800)
   Next a
End If

If CHCount > 0 Then
   For a = 0 To CHCount - 1
      Randomize
      Call AddNpc(NType, Int(4499 - 3600) * Rnd + 3600)
   Next a
End If
If SYDCount > 0 Then
   For a = 0 To SYDCount - 1
      Randomize
      Call AddNpc(NType, Int(6299 - 5400) * Rnd + 5400)
   Next a
End If


If UKCount > 0 Then
   For a = 0 To UKCount - 1
      Randomize
      Call AddNpc(NType, Int(7199 - 6300) * Rnd + 6300)
   Next a
End If

On Error Resume Next
For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      frmMain.wsk(a).SendData Chr$(252) & Chr$(3) & "A new day begins..." & Chr$(0)
      DoEvents
   End If
Next a

End Sub

Public Sub RestockDrugs()
On Error Resume Next
Dim a As Integer
Dim b As Integer
Dim c As Integer
Dim D As Integer

'Restock NPC Dealers and Remove Drugs from Buyers

For a = 0 To UBound(Npc)
   If Npc(a).NpcType = N_Dealer Then
      For b = 0 To 25
         If Npc(a).NItem(b) = -1 Then
            Randomize
            c = Int(110 - 1) * Rnd + 1
            '-----------------------------
            Select Case c
               Case 1 To 10 '10
                  Npc(a).NItem(b) = AddItem(19)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 101 To 110
                  Npc(a).NItem(b) = AddItem(20)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 11 To 20 '11
                  Npc(a).NItem(b) = AddItem(21)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 21 To 30 '12
                  Npc(a).NItem(b) = AddItem(22)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 31 To 40 '13
                  Npc(a).NItem(b) = AddItem(23)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 41 To 50 '14
                  Npc(a).NItem(b) = AddItem(24)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 51 To 60 '19
                  Npc(a).NItem(b) = AddItem(25)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 61 To 70 '21
                  Npc(a).NItem(b) = AddItem(26)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 71 To 80 '21
                  Npc(a).NItem(b) = AddItem(27)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 81 To 90
                  Npc(a).NItem(b) = AddItem(28)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
               Case 91 To 100
                  Npc(a).NItem(b) = AddItem(29)
                  Item(Npc(a).NItem(b)).ItemGUID = Npc(a).NpcGUID
                  Item(Npc(a).NItem(b)).Decay = -1
                  Item(Npc(a).NItem(b)).ILocation = -1
            End Select
         End If
      Next b
   ElseIf Npc(a).NpcType = N_Druggie Then
      For D = 0 To 19
         If Npc(a).NItem(D) <> -1 Then
            Call ResetItem(Npc(a).NItem(D))
         End If
         Npc(a).NItem(D) = -1
      Next D
   End If
Next a
   
On Error Resume Next
For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      frmMain.wsk(a).SendData Chr$(252) & Chr$(3) & "<News Flash>  Large shipments of illegal drugs have hit the streets of all major U.S. citys across America..." & Chr$(0)
      DoEvents
   End If
Next a
   
   
End Sub

Public Sub GenNpcInventory()
On Error Resume Next
Dim a As Integer
Dim b As Integer

For a = 0 To UBound(Npc)
   If Npc(a).NpcType = N_Cop Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(52)
         Npc(a).NItem(1) = AddItem(54)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_Bum Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(49)
         Npc(a).NItem(1) = AddItem(50)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_Tweaker Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(49)
         Npc(a).NItem(1) = AddItem(51)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_Bouncer Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(43)
         Npc(a).NItem(1) = AddItem(48)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_Hooker Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(53)
         Npc(a).NItem(1) = AddItem(24)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_LoanShark Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(44)
         Npc(a).NItem(1) = AddItem(22)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_Politician Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(44)
         Npc(a).NItem(1) = AddItem(56)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_Pedestrian Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(46)
         Npc(a).NItem(1) = AddItem(37)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
    ElseIf Npc(a).NpcType = N_MeterMaid Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(52)
         Npc(a).NItem(1) = AddItem(21)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_Whore Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(25)
         Npc(a).NItem(1) = AddItem(53)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_DEA Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(20)
         Npc(a).NItem(1) = AddItem(47)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_ATF Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(34)
         Npc(a).NItem(1) = AddItem(44)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
   ElseIf Npc(a).NpcType = N_FBI Then
      If Npc(a).NItem(0) = -1 And _
         Npc(a).NItem(1) = -1 Then
         Npc(a).NItem(0) = AddItem(44)
         Npc(a).NItem(1) = AddItem(52)
         Item(Npc(a).NItem(0)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(0)).ILocation = -1
         Item(Npc(a).NItem(0)).Decay = -1
         Item(Npc(a).NItem(1)).ItemGUID = Npc(a).NpcGUID
         Item(Npc(a).NItem(1)).ILocation = -1
         Item(Npc(a).NItem(1)).Decay = -1
      End If
End If
Next a
   
End Sub

Public Sub NpcCombat()
On Error Resume Next
Dim a As Integer
Dim b As Integer

For a = 0 To UBound(Npc)
   If Npc(a).NTargetID <> -1 And _
      Npc(a).NTargetGUID <> "" And _
      Npc(a).NTargetID >= 1 And _
      Npc(a).NTargetID <= MaxUsers Then
         With Npc(a)
            If .NTargetGUID = User(.NTargetID).UserGUID And _
               .NLocation = User(.NTargetID).Location And _
               User(.NTargetID).Status = "Playing" Then
               Randomize
               b = Int(100 - 1) * Rnd + 1
               .CanMove = GetTickCount()
               If b <= Npc(a).Accuracy Then
                  Call NpcCombatDamage(a)
                  If NpcKillPlayer(a, .NTargetID) = True Then
                     GoTo DoNextLoop
                  End If
                  If Item(Npc(a).NGear(0)).IType = C_Gun Then
                     frmMain.wsk(.NTargetID).SendData Chr$(2) & .NName & " fires " & " a " & Item(.NGear(0)).IName & " at you, It's a direct hit." & vbCrLf & vbCrLf & Chr$(0)
                     DoEvents
                     Call ShowWatchers(.NTargetID, Chr$(2) & "You see " & .NName & " fire " & " a " & " " & Item(.NGear(0)).IName & " at " & User(.NTargetID).UName & " and it's a direct hit." & vbCrLf & vbCrLf & Chr$(0))
                  ElseIf Item(Npc(a).NGear(0)).IType = C_Melee Then
                     frmMain.wsk(.NTargetID).SendData Chr$(2) & .NName & " strikes at you with a " & Item(.NGear(0)).IName & ", It's a direct hit." & vbCrLf & vbCrLf & Chr$(0)
                     DoEvents
                     Call ShowWatchers(.NTargetID, Chr$(2) & "You see " & .NName & " strike at " & User(.NTargetID).UName & " with a " & Item(.NGear(0)).IName & ", it's a direct hit." & vbCrLf & vbCrLf & Chr$(0))
                  End If
               ElseIf b > Npc(a).Accuracy Then
                  If Item(Npc(a).NGear(0)).IType = C_Gun Then
                     frmMain.wsk(.NTargetID).SendData Chr$(2) & .NName & " fires " & " a " & Item(.NGear(0)).IName & " at you and misses." & vbCrLf & vbCrLf & Chr$(0)
                     DoEvents
                     Call ShowWatchers(.NTargetID, Chr$(2) & "You see " & .NName & " fire " & " a " & " " & Item(.NGear(0)).IName & " at " & User(.NTargetID).UName & " and miss." & vbCrLf & vbCrLf & Chr$(0))
                  ElseIf Item(Npc(a).NGear(0)).IType = C_Melee Then
                     frmMain.wsk(.NTargetID).SendData Chr$(2) & .NName & " strikes at you with a " & Item(.NGear(0)).IName & " and misses." & vbCrLf & vbCrLf & Chr$(0)
                     DoEvents
                     Call ShowWatchers(.NTargetID, Chr$(2) & "You see " & .NName & " strike at " & User(.NTargetID).UName & " with a " & Item(.NGear(0)).IName & " and miss." & vbCrLf & vbCrLf & Chr$(0))
                  End If
               End If
               End If
         End With
   End If
DoNextLoop:
Next a

End Sub

Public Sub NpcCombatDamage(NpcIndex As Integer)
On Error Resume Next
Dim a As Integer
a = 0

With Npc(NpcIndex)

If User(.NTargetID).Armor <> -1 Then
   a = Item(User(.NTargetID).Armor).Armor
End If

If Item(Npc(NpcIndex).NGear(0)).IType = C_Gun Then
   User(.NTargetID).Health = (User(.NTargetID).Health - (Item(.NGear(0)).Damage + 1)) + a
   Call UpdateGeneralInfo(.NTargetID)
ElseIf Item(Npc(NpcIndex).NGear(0)).IType = C_Melee Then
   User(.NTargetID).Health = (User(.NTargetID).Health - Item(.NGear(0)).Damage)
   Call UpdateGeneralInfo(.NTargetID)
End If

End With

End Sub

Public Function NpcKillPlayer(NpcIndex As Integer, Index As Integer) As Boolean
Dim a As Integer
Dim b As Integer

'If a player dies,  drop all his items to the ground
If User(Index).Health > 0 Then
   NpcKillPlayer = False
   Exit Function
ElseIf User(Index).Health <= 0 Then
   NpcKillPlayer = True
   For a = 0 To 19
      If User(Index).Item(a) <> -1 Then
         For b = 0 To UBound(City(User(Index).Location).CItem)
            If City(User(Index).Location).CItem(b) = -1 Then
               City(User(Index).Location).CItem(b) = User(Index).Item(a)
               Item(User(Index).Item(a)).OnPlayer = False
               Item(User(Index).Item(a)).Equip = False
               Item(User(Index).Item(a)).Decay = GetTickCount()
               Item(User(Index).Item(a)).ItemGUID = ""
               Item(User(Index).Item(a)).ILocation = User(Index).Location
               User(Index).Item(a) = -1
               Exit For
            ElseIf b = UBound(City(User(Index).Location).CItem) Then
               With City(User(Index).Location)
               ReDim Preserve .CItem(UBound(.CItem) + 1)
               .CItem(UBound(.CItem)) = User(Index).Item(a)
               Item(User(Index).Item(a)).OnPlayer = False
               Item(User(Index).Item(a)).Equip = False
               Item(User(Index).Item(a)).Decay = GetTickCount()
               Item(User(Index).Item(a)).ItemGUID = ""
               Item(User(Index).Item(a)).ILocation = User(Index).Location
               User(Index).Item(a) = -1
               End With
            End If
         Next b
      End If
   Next a
   Call FullInventoryUpdate(Index)
   User(Index).Cash = 50
   User(Index).Health = 50
      User(Index).Accuracy = User(Index).Accuracy * 0.96
   User(Index).Search = User(Index).Search * 0.96
   User(Index).Tracking = User(Index).Tracking * 0.96
   User(Index).Hiding = User(Index).Hiding * 0.96
   User(Index).Snooping = User(Index).Snooping * 0.96
   User(Index).Sniping = User(Index).Sniping * 0.96
   User(Index).Stealing = User(Index).Stealing * 0.96

   frmMain.wsk(Index).SendData Chr$(2) & Npc(NpcIndex).NName & " has just wasted you!  You should be more carefull next time..." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Call ShowWatchers(Index, Chr$(2) & "You just witnessed " & Npc(NpcIndex).NName & " kill " & User(Index).UName & " right before your eyes." & vbCrLf & "You see " & User(Index).UName & " 's items fall to the ground." & vbCrLf & vbCrLf & Chr$(0))
   Call PlaceOnDeath(Index)
   Call UpdateGeneralInfo(Index)
   User(Index).Weapon = -1
   User(Index).Armor = -1
   User(Index).Ammo = -1
   Call UpdateGearInfo(Index)
   User(Index).TargetNum = -1
   User(Index).TargetGUID = ""
End If


End Function
