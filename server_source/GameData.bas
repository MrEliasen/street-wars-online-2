Attribute VB_Name = "GameData"
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

'City Type Constants
Public Const NewYorkCity = 0
Public Const MiamiCity = 1
Public Const LosAngelesCity = 2
Public Const HoustonCity = 3
Public Const ChicagoCity = 4
Public Const NewJerseyCity = 5
Public Const SydneyCity = 6
Public Const LondonCity = 7


'Item Type Constants
Public Const C_General = 0
Public Const C_Gun = 1
Public Const C_Armor = 2
Public Const C_Melee = 3
Public Const C_Ammo = 4
Public Const C_Dope = 5
Public Const C_Phone = 6
Public Const C_MedStick = 7
Public Const C_Paint = 8
Public Const C_Wall = 9
Public Const C_Beer = 10
Public Const C_Pager = 11
Public Const C_RedBull = 12
Public Const C_Steroids = 13
Public Const C_Rank = 14
Public Const C_Rank2 = 15
Public Const C_Acc = 16
Public Const C_Pepper = 17
Public Const C_Adrenalin = 18
Public Const C_Mine = 19
Public Const C_Grenade = 20
Public Const C_Horn = 21
Public Const C_Admin = 22

'NPC Type Constants
Public Const N_Dealer = 0
Public Const N_Druggie = 1
Public Const N_Cop = 2
Public Const N_Hooker = 3
Public Const N_LoanShark = 4
Public Const N_Bum = 5
Public Const N_Tweaker = 6
Public Const N_Bouncer = 7
Public Const N_Politician = 8
Public Const N_Taliban = 9
Public Const N_Pedestrian = 10
Public Const N_MeterMaid = 11
Public Const N_Whore = 12
Public Const N_FBI = 13
Public Const N_DEA = 14
Public Const N_ATF = 15

Public SlotID() As Integer

'User Data Type Structure
Public Type UserData
   UName As String
   UPass As String
   UserGUID As String
   AccessLevel As Integer
   Purge As Date
   Status As String
   DataBaseID As Integer
   Location As Integer
   Health As Long
   HomeTown As String
   CurrTown As String
   HomeAbv As String
   Reputation As Long
   Rank As String
   Kills As Integer
   Weapon As Integer
   Armor As Long
   Ammo As Integer
   Cash As Long
   Bank As Long
   TargetNum As Integer
   TargetGUID As String
   NpcTrade As Integer
   Accuracy As Single
   Hiding As Single
   Search As Single
   Tracking As Single
   Sniping As Single
   Searching As Single
   Chemistry As Single
   Bounty As Long
   Snooping As Single
   Stealing As Single
   Item(19) As Integer
   SkillTickNew As Long
   SkillTickOld As Long
   AdrenalinTickNew As Long
   AdrenalinTickOld As Long
   IsHiding As Boolean
   Mute As Boolean
End Type

Public User(MaxUsers) As UserData

Public UserDB() As UserData

'--------------------------------------

'City Data Type Structure
Public Type CityData
  CLocation As Integer
  CName As String
  CDesc As String
  CityGUID As String
  OwnerGUID As String
  North As Integer
  East As Integer
  South As Integer
  West As Integer
  Compass As String
  AirPort As Boolean
  Hospital As Boolean
  PawnShop As Boolean
  WhoreHouse As Boolean
  Alley As Boolean
  Bank As Boolean
  Bar As Boolean
  Casino As Boolean
  CItem() As Integer
  CNpc(15) As Integer
  Storage(49) As Integer
End Type

Public City(7199) As CityData

'-----------------------------------

'Item Data Type Sructure
Public Type ItemData
   IName As String
   IDesc As String
   ItemGUID As String
   OnPlayer As Boolean
   Equip As Boolean
   Amount As Integer
   Damage As Integer
   Armor As Integer
   Condition As Integer
   IType As Integer
   Price As Long
   Multiple As Boolean
   Movable As Boolean
   CanBuy As Long
   Decay As Long
   ILocation As Integer
   ForSale As Boolean
End Type

Public Item() As ItemData

Public ItemDB(71) As ItemData

Public Sub SendWelcome(Index As Integer, msg As String)
On Error Resume Next

If CheckIPBan(Index) = True Then
   Exit Sub
End If

'Check for correct client version and send welcome screen
If msg = SVersion Then
   frmMain.wsk(Index).SendData Chr$(2) & "Welcome to Street Wars Online: The New Era(Rebirth):  Please report all bugs/problems to JimmyJames at jimmyjames@writeme.com .  Duplicate accounts created in SWO will not be tolerated.  IP address will be banned for creating multiple accounts." & vbCrLf & vbCrLf & "Alright, what is your name?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).Status = "GetName"
   Exit Sub
Else
   frmMain.wsk(Index).SendData Chr$(2) & "Your client version is outdated, please visit the Street Wars Online II website and get the latest version at http://www.angelfire.com/oh5/swo/index.html" & vbCrLf & vbCrLf & "Closing Connection..." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   frmMain.wsk(Index).Close
   Call ResetIndex(Index)
   frmMain.lstUsers.List(Index - 1) = "<Waiting>"
   Exit Sub
End If

End Sub
Public Sub VerifyName(Index As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter

For a = 1 To MaxUsers
   If Trim$(LCase$(msg)) = LCase$(User(a).UName) Then
      frmMain.wsk(Index).SendData Chr$(2) & "That player name shows to be logged in the server at this time.  Only one login at any given time is allowed." & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(Index).Close
      Call ResetIndex(Index)
      Exit Sub
   End If
Next

For a = 1 To MaxUsers
   If a <> Index And User(a).Status = "Playing" And _
      frmMain.wsk(a).RemoteHostIP = frmMain.wsk(Index).RemoteHostIP Then
      frmMain.wsk(Index).SendData Chr$(2) & "Multiple login from a single IP detected.  <Event Logged>" & vbCrLf & Chr$(0)
      DoEvents
      frmMain.wsk(Index).Close
      Call ResetIndex(Index)
      Exit Sub
   End If
Next

'Find the user in the database
For a = 0 To UBound(UserDB)
   If LCase$(msg) = LCase$(UserDB(a).UName) Then
      User(Index) = UserDB(a)
      User(Index).Status = "GetPass"
      User(Index).DataBaseID = a
      frmMain.wsk(Index).SendData Chr$(2) & "Ok " & UserDB(a).UName & ", what is your password?" & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      Exit Sub
   ElseIf a = UBound(UserDB) Then
      frmMain.wsk(Index).SendData Chr$(2) & "We don't know any " & Trim$(msg) & " around here, would you like to join the hood?" & vbCrLf & "Yes/No" & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      User(Index).Status = "YesNo"
      Exit Sub
   End If
Next a

End Sub

Public Sub YesNo(Index As Integer, msg As String)
On Error Resume Next

'Check to see if the player wants to create a new account
If LCase$(msg) = "yes" Then
   frmMain.wsk(Index).SendData Chr$(3) & Chr$(0)
   DoEvents
   User(Index).Status = "NewAccount"
   frmMain.lstUsers.List(Index - 1) = "<New Account>"
   Exit Sub
ElseIf LCase$(msg) = "no" Then
   frmMain.wsk(Index).SendData Chr$(2) & "Ok fool, quit jerking me around!  What is your name?" & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   User(Index).Status = "GetName"
   Exit Sub
End If

End Sub

Public Sub NewAccount(Index As Integer, msg As String)
On Error GoTo Failed
Dim a As Integer, b As Integer
Dim c As Integer, D As Integer
Dim TempName As String

a = InStr(1, msg, Chr$(1))
TempName = Left$(msg, a - 1)

For D = 0 To UBound(UserDB)
   If LCase$(Trim$(TempName)) = LCase$(Trim$(UserDB(D).UName)) Then
      frmMain.wsk(Index).SendData Chr$(4) & Chr$(0)
      DoEvents
      Exit Sub
   End If
Next D

ReDim Preserve UserDB(UBound(UserDB) + 1)

User(Index).UName = Left$(msg, a - 1)
b = InStr(a + 1, msg, Chr$(1))
User(Index).UPass = Mid$(msg, a + 1, b - a - 1)
c = InStr(b + 1, msg, Chr$(1))
User(Index).HomeTown = Mid$(msg, b + 1, c - b - 1)
User(Index).UserGUID = GUID
User(Index).Purge = Date
User(Index).DataBaseID = UBound(UserDB)
User(Index).Status = "Playing"
User(Index).CurrTown = User(Index).HomeTown
Call SetTownAbv(Index)
User(Index).Reputation = 0 'Change Before Going Beta
User(Index).Kills = 0
Call SetRank(Index)
User(Index).Cash = 100 'Change Before Going Beta
User(Index).Bank = 0
User(Index).TargetGUID = ""
User(Index).TargetNum = -1
User(Index).NpcTrade = -1
User(Index).Accuracy = 20#
User(Index).Hiding = 20#
User(Index).Search = 20#
User(Index).Tracking = 20#
User(Index).Sniping = 20#
User(Index).Bounty = 0
User(Index).Chemistry = 10#
User(Index).Snooping = 20#
User(Index).Stealing = 20#
Call SetCity(Index)
User(Index).Health = 100
User(Index).AccessLevel = 0 'Change Before Going Beta
User(Index).Weapon = -1
User(Index).Armor = -1
User(Index).Ammo = -1

'No Items For Newbies
For D = 0 To 19
   User(Index).Item(D) = -1
Next D

'Set Users Name On User List Box
frmMain.lstUsers.List(Index - 1) = "<" & User(Index).UName & ">"

'Close New Account Window
frmMain.wsk(Index).SendData Chr$(5) & Chr$(0)
DoEvents

'Update Left Side Client Information
Call UpdateGeneralInfo(Index)
Call UpdateGearInfo(Index)

'Show room description
Call ShowCity(Index)

'update players inventory
Call FullInventoryUpdate(Index)

'Copy player data to User DataBase Memory
UserDB(UBound(UserDB)) = User(Index)

'Update Clients Player List
Call UpdatePlayerList

Exit Sub

'Error Handler - Kill Connection
Failed:
Dim ff As Integer 'Free File
ff = FreeFile
Open App.Path & "\error.log" For Append As ff
Print #ff, "[BOE]"
Print #ff, "Player Creation Error - UserIP/Message"
Print #ff, frmMain.wsk(Index).RemoteHostIP & " | " & msg
Print #ff, "[EOE]"
Close ff
frmMain.wsk(Index).Close
Call ResetIndex(Index)
frmMain.lstUsers.List(Index - 1) = "<Waiting>"
End Sub
Public Sub SetTownAbv(Index As Integer)
On Error Resume Next

'Set players town abbreviation upon account creation
Select Case User(Index).HomeTown
   Case "Miami"
      User(Index).HomeAbv = "MI - "
   Case "New York"
      User(Index).HomeAbv = "NY - "
   Case "New Jersey"
      User(Index).HomeAbv = "NJ - "
   Case "Los Angeles"
      User(Index).HomeAbv = "LA - "
   Case "Chicago"
      User(Index).HomeAbv = "CH - "
   Case "Houston"
      User(Index).HomeAbv = "HO - "
   Case "Sydney"
      User(Index).HomeAbv = "SYD - "
   Case "London"
      User(Index).HomeAbv = "UK - "
End Select

End Sub

Public Sub SetRank(Index As Integer)
On Error Resume Next

Select Case User(Index).Reputation
   Case Is <= -999
      User(Index).Rank = "Bitch"
   Case -1000 To -99
      User(Index).Rank = "Pathetic Loser"
   Case -100 To -1
      User(Index).Rank = "Disrespectable Punk"
   Case 0 To 99
      User(Index).Rank = "Wannabee"
   Case 100 To 999
      User(Index).Rank = "Slacker"
   Case 1000 To 1999
      User(Index).Rank = "Street Punk"
   Case 2000 To 4999
      User(Index).Rank = "Junior Thug"
   Case 5000 To 7499
      User(Index).Rank = "Thug"
   Case 7500 To 9999
      User(Index).Rank = "Theif"
   Case 10000 To 14999
      User(Index).Rank = "Wanksta"
   Case 15000 To 24999
      User(Index).Rank = "Gangsta"
   Case 25000 To 39999
      User(Index).Rank = "Soldier"
   Case 40000 To 49999
      User(Index).Rank = "Playa"
   Case 50000 To 74999
      User(Index).Rank = "Pimp"
   Case 75000 To 99999
      User(Index).Rank = "Drug Pusher"
   Case 100000 To 124999
      User(Index).Rank = "Smuggler"
   Case 125000 To 149999
      User(Index).Rank = "Drug Runner"
   Case 150000 To 199999
      User(Index).Rank = "Ballah"
   Case 200000 To 224999
      User(Index).Rank = "Shot Callah"
   Case 225000 To 249999
      User(Index).Rank = "Druglord"
   Case 250000 To 274999
      User(Index).Rank = "Hustler"
   Case 275000 To 299999
      User(Index).Rank = "Supreme Druglord"
   Case 300000 To 349999
      User(Index).Rank = "Mafia Boss"
   Case 350000 To 374999
      User(Index).Rank = "Don"
   Case 375000 To 399999
      User(Index).Rank = "Godfather"
   Case 400000 To 449999
      User(Index).Rank = "High Roller"
   Case 450000 To 999999
      User(Index).Rank = "Overlord"
   Case Is >= 1000000
      User(Index).Rank = "Stud"
End Select


End Sub

Public Sub LoadCitys()
On Error Resume Next
Dim a As Integer 'Counter
Dim ff As Integer 'Free File
a = 0

If IsFile(App.Path & "\citys.dat") = False Then
   Unload frmMain
   End
End If

'Load the world
ff = FreeFile
Open App.Path & "\citys.dat" For Input As ff

Do While Not EOF(ff)
  Input #ff, City(a).CLocation
  Input #ff, City(a).CName
  Input #ff, City(a).CDesc
  Input #ff, City(a).CityGUID
  Input #ff, City(a).OwnerGUID
  Input #ff, City(a).North
  Input #ff, City(a).East
  Input #ff, City(a).South
  Input #ff, City(a).West
  Input #ff, City(a).AirPort
  Input #ff, City(a).Hospital
  Input #ff, City(a).PawnShop
  Input #ff, City(a).WhoreHouse
  Input #ff, City(a).Alley
  Input #ff, City(a).Bank
  Input #ff, City(a).Bar
  Input #ff, City(a).Casino
  Input #ff, City(a).Compass
  a = a + 1
  DoEvents
Loop

Close ff

End Sub

Public Sub SaveCitys()
On Error Resume Next
Dim a As Integer 'Counter
Dim ff As Integer 'Free File

ff = FreeFile
Open App.Path & "\citys.dat" For Output As ff
For a = 0 To UBound(City)
  Write #ff, City(a).CLocation
  Write #ff, City(a).CName
  Write #ff, City(a).CDesc
  Write #ff, City(a).CityGUID
  Write #ff, City(a).OwnerGUID
  Write #ff, City(a).North
  Write #ff, City(a).East
  Write #ff, City(a).South
  Write #ff, City(a).West
  Write #ff, City(a).AirPort
  Write #ff, City(a).Hospital
  Write #ff, City(a).PawnShop
  Write #ff, City(a).WhoreHouse
  Write #ff, City(a).Alley
  Write #ff, City(a).Bank
  Write #ff, City(a).Bar
  Write #ff, City(a).Casino
  Write #ff, City(a).Compass
Next a
Close ff

End Sub

Public Sub SetCity(Index As Integer)
On Error Resume Next
Dim a As Integer

For a = 0 To UBound(City)
   If User(Index).HomeTown = City(a).CName And _
      City(a).Bank = True Then
         User(Index).Location = City(a).CLocation
   End If
Next a

End Sub

Public Sub ShowCity(Index As Integer)
On Error Resume Next
Dim a As Integer 'Counter
Dim msg As String 'Send String


msg = Chr$(2) & "[" & City(User(Index).Location).CName & "]" & vbCrLf & City(User(Index).Location).Compass & vbCrLf & City(User(Index).Location).CDesc & vbCrLf & "* * * * *" & vbCrLf

For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = User(Index).Location And _
      a <> Index And _
      User(a).IsHiding = False Then
         msg = msg & User(a).UName & " the " & User(a).Rank & " is here with you." & vbCrLf
   End If
Next a

For a = 0 To UBound(City(User(Index).Location).CNpc)
   If City(User(Index).Location).CNpc(a) <> -1 Then
      msg = msg & "You see " & Npc(City(User(Index).Location).CNpc(a)).NName & " " & Npc(City(User(Index).Location).CNpc(a)).NameTag & "(HP" & Npc(City(User(Index).Location).CNpc(a)).NHealth & ") standing here with you." & vbCrLf
   End If
Next a

For a = 0 To UBound(City(User(Index).Location).CItem)
   If City(User(Index).Location).CItem(a) <> -1 Then
    If Item(City(User(Index).Location).CItem(a)).IName = "Land Mine" Then
        frmMain.wsk(Index).SendData Chr$(2) & "---BOOM---YOU JUST STEPPED ON A LAND MINE." & vbCrLf & Chr$(0)
        Call ShowWatchers(Index, Chr$(2) & "You see " & User(Index).UName & " wander into the area and step on a land mine.  BOOM!" & vbCrLf & vbCrLf & Chr$(0))
        User(Index).Health = 3
        City(User(Index).Location).CItem(a) = -1
        Call ResetItem(a)
        Call UpdateGeneralInfo(Index)
        DoEvents
        Exit Sub
    End If
    msg = msg & "You see a " & Item(City(User(Index).Location).CItem(a)).IName & " on the ground." & vbCrLf
      
   End If
Next a

msg = msg & vbCrLf & Chr$(0)
frmMain.wsk(Index).SendData msg
DoEvents

End Sub
Public Sub SavePlayerData()
On Error Resume Next
Dim a As Integer 'Counter
Dim ff As Integer 'Free File

'Copy index data to database area
For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      UserDB(User(a).DataBaseID) = User(a)
   End If
Next a

'Save player information
ff = FreeFile
Open App.Path & "\pdata.dat" For Output As ff
For a = 0 To UBound(UserDB)
   If UserDB(a).UName <> "" And _
      UserDB(a).UserGUID <> "" Then
         Write #ff, UserDB(a).UName
         Write #ff, UserDB(a).UPass
         Write #ff, UserDB(a).UserGUID
         Write #ff, UserDB(a).Purge
         Write #ff, UserDB(a).Location
         Write #ff, UserDB(a).HomeTown
         Write #ff, UserDB(a).HomeAbv
         Write #ff, UserDB(a).Reputation
         Write #ff, UserDB(a).Rank
         Write #ff, UserDB(a).Kills
         Write #ff, UserDB(a).Weapon
         Write #ff, UserDB(a).Armor
         Write #ff, UserDB(a).Ammo
         Write #ff, UserDB(a).Cash
         Write #ff, UserDB(a).Bank
         Write #ff, UserDB(a).Accuracy
         Write #ff, UserDB(a).Hiding
         Write #ff, UserDB(a).Search
         Write #ff, UserDB(a).Tracking
         Write #ff, UserDB(a).Sniping
         Write #ff, UserDB(a).Chemistry
         Write #ff, UserDB(a).Bounty
         Write #ff, UserDB(a).Snooping
         Write #ff, UserDB(a).Stealing
         Write #ff, UserDB(a).Health
         Write #ff, UserDB(a).AccessLevel
   End If
Next a

Close ff

End Sub

Public Sub LoadPlayerData()
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter
Dim ff As Integer 'Free File
a = 0

'Check to see if the player file exists
If IsFile(App.Path & "\pdata.dat") = False Then
   Exit Sub
End If

'Load Player Data
ff = FreeFile
Open App.Path & "\pdata.dat" For Input As ff
Do While Not EOF(ff)
   ReDim Preserve UserDB(a)
         Input #ff, UserDB(a).UName
         Input #ff, UserDB(a).UPass
         Input #ff, UserDB(a).UserGUID
         Input #ff, UserDB(a).Purge
         Input #ff, UserDB(a).Location
         Input #ff, UserDB(a).HomeTown
         Input #ff, UserDB(a).HomeAbv
         Input #ff, UserDB(a).Reputation
         Input #ff, UserDB(a).Rank
         Input #ff, UserDB(a).Kills
         Input #ff, UserDB(a).Weapon
         Input #ff, UserDB(a).Armor
         Input #ff, UserDB(a).Ammo
         Input #ff, UserDB(a).Cash
         Input #ff, UserDB(a).Bank
         Input #ff, UserDB(a).Accuracy
         Input #ff, UserDB(a).Hiding
         Input #ff, UserDB(a).Search
         Input #ff, UserDB(a).Tracking
         Input #ff, UserDB(a).Sniping
         Input #ff, UserDB(a).Chemistry
         Input #ff, UserDB(a).Bounty
         Input #ff, UserDB(a).Snooping
         Input #ff, UserDB(a).Stealing
         Input #ff, UserDB(a).Health
         Input #ff, UserDB(a).AccessLevel
         a = a + 1
Loop

Close ff

For a = 0 To UBound(UserDB)
   For b = 0 To 19
      UserDB(a).Item(b) = -1
   Next b
Next a

End Sub

Public Sub GetPassword(Index As Integer, msg As String)
On Error Resume Next

'Check password of user
If msg = User(Index).UPass Then
   User(Index).Status = "Playing"
   User(Index).CurrTown = City(User(Index).Location).CName
   Call RunGMCheck(Index)
   User(Index).Purge = Date
   User(Index).TargetGUID = ""
   User(Index).TargetNum = -1
   User(Index).NpcTrade = -1
   Call SetRank(Index)
   Call UpdateGeneralInfo(Index)
   Call UpdateGearInfo(Index)
   frmMain.lstUsers.List(Index - 1) = "<" & User(Index).UName & ">"
   'update players inventory
   Call FullInventoryUpdate(Index)
   'show city
   Call ShowCity(Index)
   'update player list
   Call UpdatePlayerList
   Exit Sub
Else
   frmMain.wsk(Index).SendData Chr$(2) & "Wrong answer, try again..." & vbCrLf & vbCrLf & "Disconnecting..." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   frmMain.wsk(Index).Close
   Call ResetIndex(Index)
   frmMain.lstUsers.List(Index - 1) = "<Waiting>"
   Exit Sub
End If

End Sub
Public Sub LoadStaticItems()
On Error Resume Next

With ItemDB(0)
   ItemDB(0).IName = "BB Gun"
   ItemDB(0).IDesc = "Your typical BB gun, not a very potent weapon."
   ItemDB(0).ItemGUID = ""
   ItemDB(0).OnPlayer = False
   ItemDB(0).Equip = False
   ItemDB(0).Amount = -1
   ItemDB(0).Damage = 8
   ItemDB(0).Armor = -1
   ItemDB(0).Condition = 2000
   ItemDB(0).IType = C_Gun
   ItemDB(0).Price = 125
   ItemDB(0).Multiple = False
   ItemDB(0).Movable = True
   ItemDB(0).CanBuy = 100
   ItemDB(0).Decay = -1
   ItemDB(0).ForSale = True
End With

With ItemDB(1)
   ItemDB(1).IName = "Tech 9"
   ItemDB(1).IDesc = "A well assembled weapon, although not very potent."
   ItemDB(1).ItemGUID = ""
   ItemDB(1).OnPlayer = False
   ItemDB(1).Equip = False
   ItemDB(1).Amount = -1
   ItemDB(1).Damage = 8
   ItemDB(1).Armor = -1
   ItemDB(1).Condition = 2000
   ItemDB(1).IType = C_Gun
   ItemDB(1).Price = 1000
   ItemDB(1).Multiple = False
   ItemDB(1).Movable = True
   ItemDB(1).CanBuy = 25000
   ItemDB(1).Decay = -1
   ItemDB(1).ForSale = True
End With

With ItemDB(2)
   ItemDB(2).IName = "Mack 10"
   ItemDB(2).IDesc = "A clip loading, semi automatic."
   ItemDB(2).ItemGUID = ""
   ItemDB(2).OnPlayer = False
   ItemDB(2).Equip = False
   ItemDB(2).Amount = -1
   ItemDB(2).Damage = 9
   ItemDB(2).Armor = -1
   ItemDB(2).Condition = 2000
   ItemDB(2).IType = C_Gun
   ItemDB(2).Price = 5000
   ItemDB(2).Multiple = False
   ItemDB(2).Movable = True
   ItemDB(2).CanBuy = 40000
   ItemDB(2).Decay = -1
   ItemDB(2).ForSale = True
End With

With ItemDB(3)
   ItemDB(3).IName = "Desert Eagle"
   ItemDB(3).IDesc = "This gun makes you want to say.. Do you feel lucky punk?"
   ItemDB(3).ItemGUID = ""
   ItemDB(3).OnPlayer = False
   ItemDB(3).Equip = False
   ItemDB(3).Amount = -1
   ItemDB(3).Damage = 10
   ItemDB(3).Armor = -1
   ItemDB(3).Condition = 2001
   ItemDB(3).IType = C_Gun
   ItemDB(3).Price = 15000
   ItemDB(3).Multiple = False
   ItemDB(3).Movable = True
   ItemDB(3).CanBuy = 50000
   ItemDB(3).Decay = -1
   ItemDB(3).ForSale = True
End With


With ItemDB(4)
   ItemDB(4).IName = "Uzi"
   ItemDB(4).IDesc = "You know it, you love it."
   ItemDB(4).ItemGUID = ""
   ItemDB(4).OnPlayer = False
   ItemDB(4).Equip = False
   ItemDB(4).Amount = -1
   ItemDB(4).Damage = 11
   ItemDB(4).Armor = -1
   ItemDB(4).Condition = 2000
   ItemDB(4).IType = C_Gun
   ItemDB(4).Price = 25000
   ItemDB(4).Multiple = False
   ItemDB(4).Movable = True
   ItemDB(4).CanBuy = 75000
   ItemDB(4).Decay = -1
   ItemDB(4).ForSale = True
End With

With ItemDB(5)
   ItemDB(5).IName = "AK-47"
   ItemDB(5).IDesc = "The fabled assault rifle, holy thou art."
   ItemDB(5).ItemGUID = ""
   ItemDB(5).OnPlayer = False
   ItemDB(5).Equip = False
   ItemDB(5).Amount = -1
   ItemDB(5).Damage = 15
   ItemDB(5).Armor = -1
   ItemDB(5).Condition = 2000
   ItemDB(5).IType = C_Gun
   ItemDB(5).Price = 50000
   ItemDB(5).Multiple = False
   ItemDB(5).Movable = True
   ItemDB(5).CanBuy = 150000
   ItemDB(5).Decay = -1
   ItemDB(5).ForSale = True
End With

With ItemDB(6)
   ItemDB(6).IName = ".12 Guage Shotgun"
   ItemDB(6).IDesc = "A typical Shotgun."
   ItemDB(6).ItemGUID = ""
   ItemDB(6).OnPlayer = False
   ItemDB(6).Equip = False
   ItemDB(6).Amount = -1
   ItemDB(6).Damage = 16
   ItemDB(6).Armor = -1
   ItemDB(6).Condition = 2000
   ItemDB(6).IType = C_Gun
   ItemDB(6).Price = 200000
   ItemDB(6).Multiple = False
   ItemDB(6).Movable = True
   ItemDB(6).CanBuy = 190001
   ItemDB(6).Decay = -1
   ItemDB(6).ForSale = True
End With

With ItemDB(7)
   ItemDB(7).IName = "Glock .60 Special"
   ItemDB(7).IDesc = ".60 Caliber Glock.  Great for the 4th of July."
   ItemDB(7).ItemGUID = ""
   ItemDB(7).OnPlayer = False
   ItemDB(7).Equip = False
   ItemDB(7).Amount = -1
   ItemDB(7).Damage = 17
   ItemDB(7).Armor = -1
   ItemDB(7).Condition = 2000
   ItemDB(7).IType = C_Gun
   ItemDB(7).Price = 75000
   ItemDB(7).Multiple = False
   ItemDB(7).Movable = True
   ItemDB(7).CanBuy = 225000
   ItemDB(7).Decay = -1
   ItemDB(7).ForSale = True
End With

With ItemDB(8)
   ItemDB(8).IName = "Double Barrel Shotgun"
   ItemDB(8).IDesc = "Like the .12 guage, but twice the kick."
   ItemDB(8).ItemGUID = ""
   ItemDB(8).OnPlayer = False
   ItemDB(8).Equip = False
   ItemDB(8).Amount = -1
   ItemDB(8).Damage = 18
   ItemDB(8).Armor = -1
   ItemDB(8).Condition = 2000
   ItemDB(8).IType = C_Gun
   ItemDB(8).Price = 100000
   ItemDB(8).Multiple = False
   ItemDB(8).Movable = True
   ItemDB(8).CanBuy = 250000
   ItemDB(8).Decay = -1
   ItemDB(8).ForSale = True
End With

With ItemDB(9)
   ItemDB(9).IName = "Mini Gun"
   ItemDB(9).IDesc = "Spews several thousand round per minute."
   ItemDB(9).ItemGUID = ""
   ItemDB(9).OnPlayer = False
   ItemDB(9).Equip = False
   ItemDB(9).Amount = -1
   ItemDB(9).Damage = 19
   ItemDB(9).Armor = -1
   ItemDB(9).Condition = 2000
   ItemDB(9).IType = C_Gun
   ItemDB(9).Price = 200000
   ItemDB(9).Multiple = False
   ItemDB(9).Movable = True
   ItemDB(9).CanBuy = 275000
   ItemDB(9).Decay = -1
   ItemDB(9).ForSale = True
End With

With ItemDB(10)
   ItemDB(10).IName = "M-60"
   ItemDB(10).IDesc = "The best assault gun money can buy."
   ItemDB(10).ItemGUID = ""
   ItemDB(10).OnPlayer = False
   ItemDB(10).Equip = False
   ItemDB(10).Amount = -1
   ItemDB(10).Damage = 20
   ItemDB(10).Armor = -1
   ItemDB(10).Condition = 2000
   ItemDB(10).IType = C_Gun
   ItemDB(10).Price = 250000
   ItemDB(10).Multiple = False
   ItemDB(10).Movable = True
   ItemDB(10).CanBuy = 300000
   ItemDB(10).Decay = -1
   ItemDB(10).ForSale = True
End With

With ItemDB(11)
   ItemDB(11).IName = "Stinger"
   ItemDB(11).IDesc = "A shoulder mounted, surface-to-surface missle launcher."
   ItemDB(11).ItemGUID = ""
   ItemDB(11).OnPlayer = False
   ItemDB(11).Equip = False
   ItemDB(11).Amount = -1
   ItemDB(11).Damage = 21
   ItemDB(11).Armor = -1
   ItemDB(11).Condition = 2000
   ItemDB(11).IType = C_Gun
   ItemDB(11).Price = 500000
   ItemDB(11).Multiple = False
   ItemDB(11).Movable = True
   ItemDB(11).CanBuy = 350000
   ItemDB(11).Decay = -1
   ItemDB(11).ForSale = True
End With

With ItemDB(12)
   ItemDB(12).IName = "Rail Gun"
   ItemDB(12).IDesc = "A high tech gun using magnetic propulsion."
   ItemDB(12).ItemGUID = ""
   ItemDB(12).OnPlayer = False
   ItemDB(12).Equip = False
   ItemDB(12).Amount = -1
   ItemDB(12).Damage = 22
   ItemDB(12).Armor = -1
   ItemDB(12).Condition = 2000
   ItemDB(12).IType = C_Gun
   ItemDB(12).Price = 1000000
   ItemDB(12).Multiple = False
   ItemDB(12).Movable = True
   ItemDB(12).CanBuy = 375000
   ItemDB(12).Decay = -1
   ItemDB(12).ForSale = True
End With

With ItemDB(13)
   ItemDB(13).IName = "Darts"
   ItemDB(13).IDesc = "Good for target practice, but not much else."
   ItemDB(13).ItemGUID = ""
   ItemDB(13).OnPlayer = False
   ItemDB(13).Equip = False
   ItemDB(13).Amount = 10
   ItemDB(13).Damage = 0
   ItemDB(13).Armor = -1
   ItemDB(13).Condition = 2000
   ItemDB(13).IType = C_Ammo
   ItemDB(13).Price = 10
   ItemDB(13).Multiple = True
   ItemDB(13).Movable = True
   ItemDB(13).CanBuy = 1000
   ItemDB(13).Decay = -1
   ItemDB(13).ForSale = True
End With

With ItemDB(14)
   ItemDB(14).IName = "Basic Ammo"
   ItemDB(14).IDesc = "Your standard bulletpack, nothing special."
   ItemDB(14).ItemGUID = ""
   ItemDB(14).OnPlayer = False
   ItemDB(14).Equip = False
   ItemDB(14).Amount = 10
   ItemDB(14).Damage = 1
   ItemDB(14).Armor = -1
   ItemDB(14).Condition = 2000
   ItemDB(14).IType = C_Ammo
   ItemDB(14).Price = 50
   ItemDB(14).Multiple = True
   ItemDB(14).Movable = True
   ItemDB(14).CanBuy = 5000
   ItemDB(14).Decay = -1
   ItemDB(14).ForSale = True
End With

With ItemDB(15)
   ItemDB(15).IName = "HPT Ammo"
   ItemDB(15).IDesc = "Upgraded Bullets, more charge, more power."
   ItemDB(15).ItemGUID = ""
   ItemDB(15).OnPlayer = False
   ItemDB(15).Equip = False
   ItemDB(15).Amount = 10
   ItemDB(15).Damage = 2
   ItemDB(15).Armor = -1
   ItemDB(15).Condition = 2000
   ItemDB(15).IType = C_Ammo
   ItemDB(15).Price = 200
   ItemDB(15).Multiple = True
   ItemDB(15).Movable = True
   ItemDB(15).CanBuy = 15000
   ItemDB(15).Decay = -1
   ItemDB(15).ForSale = True
End With


With ItemDB(16)
   ItemDB(16).IName = "ET Ammo"
   ItemDB(16).IDesc = "Exploding Tip Ammunition... Fun Fun"
   ItemDB(16).ItemGUID = ""
   ItemDB(16).OnPlayer = False
   ItemDB(16).Equip = False
   ItemDB(16).Amount = 10
   ItemDB(16).Damage = 3
   ItemDB(16).Armor = -1
   ItemDB(16).Condition = 2000
   ItemDB(16).IType = C_Ammo
   ItemDB(16).Price = 2000
   ItemDB(16).Multiple = True
   ItemDB(16).Movable = True
   ItemDB(16).CanBuy = 50000
   ItemDB(16).Decay = -1
   ItemDB(16).ForSale = True
End With

With ItemDB(17)
   ItemDB(17).IName = "AP Ammo"
   ItemDB(17).IDesc = "Armor Piercing."
   ItemDB(17).ItemGUID = ""
   ItemDB(17).OnPlayer = False
   ItemDB(17).Equip = False
   ItemDB(17).Amount = 10
   ItemDB(17).Damage = 4
   ItemDB(17).Armor = -1
   ItemDB(17).Condition = 2000
   ItemDB(17).IType = C_Ammo
   ItemDB(17).Price = 5000
   ItemDB(17).Multiple = True
   ItemDB(17).Movable = True
   ItemDB(17).CanBuy = 100000
   ItemDB(17).Decay = -1
   ItemDB(17).ForSale = True
End With

With ItemDB(18)
   ItemDB(18).IName = "Mortar Rounds"
   ItemDB(18).IDesc = "High caliber, explosive rounds."
   ItemDB(18).ItemGUID = ""
   ItemDB(18).OnPlayer = False
   ItemDB(18).Equip = False
   ItemDB(18).Amount = 10
   ItemDB(18).Damage = 5
   ItemDB(18).Armor = -1
   ItemDB(18).Condition = 2000
   ItemDB(18).IType = C_Ammo
   ItemDB(18).Price = 10000
   ItemDB(18).Multiple = True
   ItemDB(18).Movable = True
   ItemDB(18).CanBuy = 400000
   ItemDB(18).Decay = -1
   ItemDB(18).ForSale = True
End With

With ItemDB(19)
   ItemDB(19).IName = "Special K"
   ItemDB(19).IDesc = "A large bag of Special K."
   ItemDB(19).ItemGUID = ""
   ItemDB(19).OnPlayer = False
   ItemDB(19).Equip = False
   ItemDB(19).Amount = -1
   ItemDB(19).Damage = -1
   ItemDB(19).Armor = -1
   ItemDB(19).Condition = 2000
   ItemDB(19).IType = C_Dope
   ItemDB(19).Price = 20000
   ItemDB(19).Multiple = False
   ItemDB(19).Movable = True
   ItemDB(19).CanBuy = -1
   ItemDB(19).Decay = -1
   ItemDB(19).ForSale = False
End With

With ItemDB(20)
   ItemDB(20).IName = "Kilo of Cocaine"
   ItemDB(20).IDesc = "A large brick of coke."
   ItemDB(20).ItemGUID = ""
   ItemDB(20).OnPlayer = False
   ItemDB(20).Equip = False
   ItemDB(20).Amount = -1
   ItemDB(20).Damage = -1
   ItemDB(20).Armor = -1
   ItemDB(20).Condition = 2000
   ItemDB(20).IType = C_Dope
   ItemDB(20).Price = 30000
   ItemDB(20).Multiple = False
   ItemDB(20).Movable = True
   ItemDB(20).CanBuy = -1
   ItemDB(20).Decay = -1
   ItemDB(20).ForSale = False
End With

With ItemDB(21)
   ItemDB(21).IName = "Stash of Ludes"
   ItemDB(21).IDesc = "A unit of Ludes."
   ItemDB(21).ItemGUID = ""
   ItemDB(21).OnPlayer = False
   ItemDB(21).Equip = False
   ItemDB(21).Amount = -1
   ItemDB(21).Damage = -1
   ItemDB(21).Armor = -1
   ItemDB(21).Condition = 2000
   ItemDB(21).IType = C_Dope
   ItemDB(21).Price = 18
   ItemDB(21).Multiple = False
   ItemDB(21).Movable = True
   ItemDB(21).CanBuy = -1
   ItemDB(21).Decay = -1
   ItemDB(21).ForSale = False
End With

With ItemDB(22)
   ItemDB(22).IName = "Bag of Weed"
   ItemDB(22).IDesc = "A bag -o- weed."
   ItemDB(22).ItemGUID = ""
   ItemDB(22).OnPlayer = False
   ItemDB(22).Equip = False
   ItemDB(22).Amount = -1
   ItemDB(22).Damage = -1
   ItemDB(22).Armor = -1
   ItemDB(22).Condition = 2000
   ItemDB(22).IType = C_Dope
   ItemDB(22).Price = 35
   ItemDB(22).Multiple = False
   ItemDB(22).Movable = True
   ItemDB(22).CanBuy = -1
   ItemDB(22).Decay = -1
   ItemDB(22).ForSale = False
End With

With ItemDB(23)
   ItemDB(23).IName = "Oxycotton"
   ItemDB(23).IDesc = "Oxycotton."
   ItemDB(23).ItemGUID = ""
   ItemDB(23).OnPlayer = False
   ItemDB(23).Equip = False
   ItemDB(23).Amount = -1
   ItemDB(23).Damage = -1
   ItemDB(23).Armor = -1
   ItemDB(23).Condition = 2000
   ItemDB(23).IType = C_Dope
   ItemDB(23).Price = 50000
   ItemDB(23).Multiple = False
   ItemDB(23).Movable = True
   ItemDB(23).CanBuy = -1
   ItemDB(23).Decay = -1
   ItemDB(23).ForSale = False
End With

With ItemDB(24)
   ItemDB(24).IName = "Plate of Crank"
   ItemDB(24).IDesc = "A unit of a noxious concotion."
   ItemDB(24).ItemGUID = ""
   ItemDB(24).OnPlayer = False
   ItemDB(24).Equip = False
   ItemDB(24).Amount = -1
   ItemDB(24).Damage = -1
   ItemDB(24).Armor = -1
   ItemDB(24).Condition = 2000
   ItemDB(24).IType = C_Dope
   ItemDB(24).Price = 100
   ItemDB(24).Multiple = False
   ItemDB(24).Movable = True
   ItemDB(24).CanBuy = -1
   ItemDB(24).Decay = -1
   ItemDB(24).ForSale = False
End With

With ItemDB(25)
   ItemDB(25).IName = "Rock of Crack"
   ItemDB(25).IDesc = "A unit of Crack."
   ItemDB(25).ItemGUID = ""
   ItemDB(25).OnPlayer = False
   ItemDB(25).Equip = False
   ItemDB(25).Amount = -1
   ItemDB(25).Damage = -1
   ItemDB(25).Armor = -1
   ItemDB(25).Condition = 2000
   ItemDB(25).IType = C_Dope
   ItemDB(25).Price = 500
   ItemDB(25).Multiple = False
   ItemDB(25).Movable = True
   ItemDB(25).CanBuy = -1
   ItemDB(25).Decay = -1
   ItemDB(25).ForSale = False
End With

With ItemDB(26)
   ItemDB(26).IName = "Foiled Shrooms"
   ItemDB(26).IDesc = "A couple of giant mushrooms."
   ItemDB(26).ItemGUID = ""
   ItemDB(26).OnPlayer = False
   ItemDB(26).Equip = False
   ItemDB(26).Amount = -1
   ItemDB(26).Damage = -1
   ItemDB(26).Armor = -1
   ItemDB(26).Condition = 2000
   ItemDB(26).IType = C_Dope
   ItemDB(26).Price = 900
   ItemDB(26).Multiple = False
   ItemDB(26).Movable = True
   ItemDB(26).CanBuy = -1
   ItemDB(26).Decay = -1
   ItemDB(26).ForSale = False
End With

With ItemDB(27)
   ItemDB(27).IName = "OZ of Cocaine"
   ItemDB(27).IDesc = "An Ounce."
   ItemDB(27).ItemGUID = ""
   ItemDB(27).OnPlayer = False
   ItemDB(27).Equip = False
   ItemDB(27).Amount = -1
   ItemDB(27).Damage = -1
   ItemDB(27).Armor = -1
   ItemDB(27).Condition = 2000
   ItemDB(27).IType = C_Dope
   ItemDB(27).Price = 7500
   ItemDB(27).Multiple = False
   ItemDB(27).Movable = True
   ItemDB(27).CanBuy = -1
   ItemDB(27).Decay = -1
   ItemDB(27).ForSale = False
End With

With ItemDB(28)
   ItemDB(28).IName = "Balloon of Heroin"
   ItemDB(28).IDesc = "A unit of Heroin."
   ItemDB(28).ItemGUID = ""
   ItemDB(28).OnPlayer = False
   ItemDB(28).Equip = False
   ItemDB(28).Amount = -1
   ItemDB(28).Damage = -1
   ItemDB(28).Armor = -1
   ItemDB(28).Condition = 2000
   ItemDB(28).IType = C_Dope
   ItemDB(28).Price = 3000
   ItemDB(28).Multiple = False
   ItemDB(28).Movable = True
   ItemDB(28).CanBuy = -1
   ItemDB(28).Decay = -1
   ItemDB(28).ForSale = False
End With

With ItemDB(29)
   ItemDB(29).IName = "Gram of Opium"
   ItemDB(29).IDesc = "A Gram of Opium"
   ItemDB(29).ItemGUID = ""
   ItemDB(29).OnPlayer = False
   ItemDB(29).Equip = False
   ItemDB(29).Amount = -1
   ItemDB(29).Damage = -1
   ItemDB(29).Armor = -1
   ItemDB(29).Condition = 2000
   ItemDB(29).IType = C_Dope
   ItemDB(29).Price = 15000
   ItemDB(29).Multiple = False
   ItemDB(29).Movable = True
   ItemDB(29).CanBuy = -1
   ItemDB(29).Decay = -1
   ItemDB(29).ForSale = False
End With

With ItemDB(30)
   ItemDB(30).IName = "Steel Pipe"
   ItemDB(30).IDesc = "It's a hollow steel pipe, great for bashing skulls."
   ItemDB(30).ItemGUID = ""
   ItemDB(30).OnPlayer = False
   ItemDB(30).Equip = False
   ItemDB(30).Amount = -1
   ItemDB(30).Damage = 3
   ItemDB(30).Armor = -1
   ItemDB(30).Condition = 2000
   ItemDB(30).IType = C_Melee
   ItemDB(30).Price = 20
   ItemDB(30).Multiple = False
   ItemDB(30).Movable = True
   ItemDB(30).CanBuy = 2000
   ItemDB(30).Decay = -1
   ItemDB(30).ForSale = True
End With

With ItemDB(31)
   ItemDB(31).IName = "Brass Knuckle"
   ItemDB(31).IDesc = ""
   ItemDB(31).ItemGUID = ""
   ItemDB(31).OnPlayer = False
   ItemDB(31).Equip = False
   ItemDB(31).Amount = -1
   ItemDB(31).Damage = 4
   ItemDB(31).Armor = -1
   ItemDB(31).Condition = 2000
   ItemDB(31).IType = C_Melee
   ItemDB(31).Price = 500
   ItemDB(31).Multiple = False
   ItemDB(31).Movable = False
   ItemDB(31).CanBuy = 7500
   ItemDB(31).Decay = -1
   ItemDB(31).ForSale = True
End With

With ItemDB(32)
   ItemDB(32).IName = "Baseball Bat"
   ItemDB(32).IDesc = "It's a standard police issued PR-24 Baton"
   ItemDB(32).ItemGUID = ""
   ItemDB(32).OnPlayer = False
   ItemDB(32).Equip = False
   ItemDB(32).Amount = -1
   ItemDB(32).Damage = 5
   ItemDB(32).Armor = -1
   ItemDB(32).Condition = 2000
   ItemDB(32).IType = C_Melee
   ItemDB(32).Price = 500
   ItemDB(32).Multiple = False
   ItemDB(32).Movable = True
   ItemDB(32).CanBuy = 10000
   ItemDB(32).Decay = -1
   ItemDB(32).ForSale = True
End With

With ItemDB(33)
   ItemDB(33).IName = "Crowbar"
   ItemDB(33).IDesc = "It's a standard tire iron."
   ItemDB(33).ItemGUID = ""
   ItemDB(33).OnPlayer = False
   ItemDB(33).Equip = False
   ItemDB(33).Amount = -1
   ItemDB(33).Damage = 6
   ItemDB(33).Armor = -1
   ItemDB(33).Condition = 2000
   ItemDB(33).IType = C_Melee
   ItemDB(33).Price = 1000
   ItemDB(33).Multiple = False
   ItemDB(33).Movable = True
   ItemDB(33).CanBuy = 15000
   ItemDB(33).Decay = -1
   ItemDB(33).ForSale = True
End With

With ItemDB(34)
   ItemDB(34).IName = "Tazer"
   ItemDB(34).IDesc = "Shockingly refreshing."
   ItemDB(34).ItemGUID = ""
   ItemDB(34).OnPlayer = False
   ItemDB(34).Equip = False
   ItemDB(34).Amount = -1
   ItemDB(34).Damage = 7
   ItemDB(34).Armor = -1
   ItemDB(34).Condition = 2000
   ItemDB(34).IType = C_Melee
   ItemDB(34).Price = 1000
   ItemDB(34).Multiple = False
   ItemDB(34).Movable = True
   ItemDB(34).CanBuy = 25000
   ItemDB(34).Decay = -1
   ItemDB(34).ForSale = True
End With

With ItemDB(35)
   ItemDB(35).IName = "Broad Sword"
   ItemDB(35).IDesc = "Get Medieval on their ass."
   ItemDB(35).ItemGUID = ""
   ItemDB(35).OnPlayer = False
   ItemDB(35).Equip = False
   ItemDB(35).Amount = -1
   ItemDB(35).Damage = 12
   ItemDB(35).Armor = -1
   ItemDB(35).Condition = 2000
   ItemDB(35).IType = C_Melee
   ItemDB(35).Price = 40000
   ItemDB(35).Multiple = False
   ItemDB(35).Movable = True
   ItemDB(35).CanBuy = 100000
   ItemDB(35).Decay = -1
   ItemDB(35).ForSale = True
End With

With ItemDB(36)
   ItemDB(36).IName = "Chainsaw"
   ItemDB(36).IDesc = "Recommended by lumberjacks and psycho killer alike."
   ItemDB(36).ItemGUID = ""
   ItemDB(36).OnPlayer = False
   ItemDB(36).Equip = False
   ItemDB(36).Amount = -1
   ItemDB(36).Damage = 13
   ItemDB(36).Armor = -1
   ItemDB(36).Condition = 2000
   ItemDB(36).IType = C_Melee
   ItemDB(36).Price = 50000
   ItemDB(36).Multiple = False
   ItemDB(36).Movable = True
   ItemDB(36).CanBuy = 125000
   ItemDB(36).Decay = -1
   ItemDB(36).ForSale = True
End With

With ItemDB(37)
   ItemDB(37).IName = "Leather Jacket"
   ItemDB(37).IDesc = "A tough leather jacket, and a bad fashion statement."
   ItemDB(37).ItemGUID = ""
   ItemDB(37).OnPlayer = False
   ItemDB(37).Equip = False
   ItemDB(37).Amount = -1
   ItemDB(37).Damage = -1
   ItemDB(37).Armor = 1
   ItemDB(37).Condition = 2000
   ItemDB(37).IType = C_Armor
   ItemDB(37).Price = 900
   ItemDB(37).Multiple = False
   ItemDB(37).Movable = True
   ItemDB(37).CanBuy = 10000
   ItemDB(37).Decay = -1
   ItemDB(37).ForSale = True
End With

With ItemDB(38)
   ItemDB(38).IName = "Chain Mail"
   ItemDB(38).IDesc = "Midevil Chain Mail."
   ItemDB(38).ItemGUID = ""
   ItemDB(38).OnPlayer = False
   ItemDB(38).Equip = False
   ItemDB(38).Amount = -1
   ItemDB(38).Damage = -1
   ItemDB(38).Armor = 2
   ItemDB(38).Condition = 2000
   ItemDB(38).IType = C_Armor
   ItemDB(38).Price = 6500
   ItemDB(38).Multiple = False
   ItemDB(38).Movable = True
   ItemDB(38).CanBuy = 40000
   ItemDB(38).Decay = -1
   ItemDB(38).ForSale = True
End With

With ItemDB(39)
   ItemDB(39).IName = "Kevlar Jacket"
   ItemDB(39).IDesc = "Cop issue bullet proof vest."
   ItemDB(39).ItemGUID = ""
   ItemDB(39).OnPlayer = False
   ItemDB(39).Equip = False
   ItemDB(39).Amount = -1
   ItemDB(39).Damage = -1
   ItemDB(39).Armor = 3
   ItemDB(39).Condition = 2000
   ItemDB(39).IType = C_Armor
   ItemDB(39).Price = 10000
   ItemDB(39).Multiple = False
   ItemDB(39).Movable = True
   ItemDB(39).CanBuy = 75000
   ItemDB(39).Decay = -1
   ItemDB(39).ForSale = True
End With

With ItemDB(40)
   ItemDB(40).IName = "Flak Jacket"
   ItemDB(40).IDesc = "Military style flack jacket."
   ItemDB(40).ItemGUID = ""
   ItemDB(40).OnPlayer = False
   ItemDB(40).Equip = False
   ItemDB(40).Amount = -1
   ItemDB(40).Damage = -1
   ItemDB(40).Armor = 4
   ItemDB(40).Condition = 2000
   ItemDB(40).IType = C_Armor
   ItemDB(40).Price = 16500
   ItemDB(40).Multiple = False
   ItemDB(40).Movable = True
   ItemDB(40).CanBuy = 150000
   ItemDB(40).Decay = -1
   ItemDB(40).ForSale = True
End With

With ItemDB(41)
   ItemDB(41).IName = "Cermet Suit"
   ItemDB(41).IDesc = "Walking tank armor."
   ItemDB(41).ItemGUID = ""
   ItemDB(41).OnPlayer = False
   ItemDB(41).Equip = False
   ItemDB(41).Amount = -1
   ItemDB(41).Damage = -1
   ItemDB(41).Armor = 5
   ItemDB(41).Condition = 2000
   ItemDB(41).IType = C_Armor
   ItemDB(41).Price = 106500
   ItemDB(41).Multiple = False
   ItemDB(41).Movable = True
   ItemDB(41).CanBuy = 250000
   ItemDB(41).Decay = -1
   ItemDB(41).ForSale = True
End With

With ItemDB(42)
   ItemDB(42).IName = "Depleted Uranium Suit"
   ItemDB(42).IDesc = "Made of depleted uranium, the hardest substance this side of diamonds."
   ItemDB(42).ItemGUID = ""
   ItemDB(42).OnPlayer = False
   ItemDB(42).Equip = False
   ItemDB(42).Amount = -1
   ItemDB(42).Damage = -1
   ItemDB(42).Armor = 6
   ItemDB(42).Condition = 2000
   ItemDB(42).IType = C_Armor
   ItemDB(42).Price = 250000
   ItemDB(42).Multiple = False
   ItemDB(42).Movable = True
   ItemDB(42).CanBuy = 300000
   ItemDB(42).Decay = -1
   ItemDB(42).ForSale = True
End With

With ItemDB(43)
   ItemDB(43).IName = "Cell Phone"
   ItemDB(43).IDesc = "It's a cellular phone to locate drug dealers and druggies.  Only $15.00 per use, what a deal."
   ItemDB(43).ItemGUID = ""
   ItemDB(43).OnPlayer = False
   ItemDB(43).Equip = False
   ItemDB(43).Amount = -1
   ItemDB(43).Damage = -1
   ItemDB(43).Armor = -1
   ItemDB(43).Condition = 2000
   ItemDB(43).IType = C_Phone
   ItemDB(43).Price = 310
   ItemDB(43).Multiple = False
   ItemDB(43).Movable = True
   ItemDB(43).CanBuy = 1000
   ItemDB(43).Decay = -1
   ItemDB(43).ForSale = True
End With

With ItemDB(44)
   ItemDB(44).IName = "-Pager-"
   ItemDB(44).IDesc = "A pager, Used to locate NPCs in your city."
   ItemDB(44).ItemGUID = ""
   ItemDB(44).OnPlayer = False
   ItemDB(44).Equip = False
   ItemDB(44).Amount = -1
   ItemDB(44).Damage = -1
   ItemDB(44).Armor = -1
   ItemDB(44).Condition = 2000
   ItemDB(44).IType = C_Pager
   ItemDB(44).Price = 400
   ItemDB(44).Multiple = False
   ItemDB(44).Movable = True
   ItemDB(44).CanBuy = 7500
   ItemDB(44).Decay = -1
   ItemDB(44).ForSale = True
End With


With ItemDB(45)
   ItemDB(45).IName = "Steroids"
   ItemDB(45).IDesc = "Give in to the peer pressure."
   ItemDB(45).ItemGUID = ""
   ItemDB(45).OnPlayer = False
   ItemDB(45).Equip = False
   ItemDB(45).Amount = -1
   ItemDB(45).Damage = -1
   ItemDB(45).Armor = -1
   ItemDB(45).Condition = 2000
   ItemDB(45).IType = C_Steroids
   ItemDB(45).Price = 250000
   ItemDB(45).Multiple = False
   ItemDB(45).Movable = True
   ItemDB(45).CanBuy = 225000
   ItemDB(45).Decay = -1
   ItemDB(45).ForSale = True
End With

With ItemDB(46)
   ItemDB(46).IName = "Red Bull"
   ItemDB(46).IDesc = "The Yuppie energy drink."
   ItemDB(46).ItemGUID = ""
   ItemDB(46).OnPlayer = False
   ItemDB(46).Equip = False
   ItemDB(46).Amount = -1
   ItemDB(46).Damage = -1
   ItemDB(46).Armor = -1
   ItemDB(46).Condition = 2000
   ItemDB(46).IType = C_RedBull
   ItemDB(46).Price = 1000
   ItemDB(46).Multiple = False
   ItemDB(46).Movable = True
   ItemDB(46).CanBuy = -1
   ItemDB(46).Decay = -1
   ItemDB(46).ForSale = False
End With

With ItemDB(47)
   ItemDB(47).IName = "Medstick"
   ItemDB(47).IDesc = ""
   ItemDB(47).ItemGUID = ""
   ItemDB(47).OnPlayer = False
   ItemDB(47).Equip = False
   ItemDB(47).Amount = -1
   ItemDB(47).Damage = -1
   ItemDB(47).Armor = -1
   ItemDB(47).Condition = 2000
   ItemDB(47).IType = C_MedStick
   ItemDB(47).Price = 25000
   ItemDB(47).Multiple = False
   ItemDB(47).Movable = True
   ItemDB(47).CanBuy = 5000
   ItemDB(47).Decay = -1
   ItemDB(47).ForSale = True
End With

With ItemDB(48)
   ItemDB(48).IName = "Gold Chain"
   ItemDB(48).IDesc = "It's a 18k gold chain, Mr.T style."
   ItemDB(48).ItemGUID = ""
   ItemDB(48).OnPlayer = False
   ItemDB(48).Equip = False
   ItemDB(48).Amount = -1
   ItemDB(48).Damage = -1
   ItemDB(48).Armor = -1
   ItemDB(48).Condition = 2000
   ItemDB(48).IType = C_General
   ItemDB(48).Price = 600
   ItemDB(48).Multiple = False
   ItemDB(48).Movable = True
   ItemDB(48).CanBuy = -2000
   ItemDB(48).Decay = -1
   ItemDB(48).ForSale = False
End With

With ItemDB(49)
   ItemDB(49).IName = "Cheap Watch"
   ItemDB(49).IDesc = "It's a cheap watch.  Upon examining it even closer, you realize the watch doesn't even work."
   ItemDB(49).ItemGUID = ""
   ItemDB(49).OnPlayer = False
   ItemDB(49).Equip = False
   ItemDB(49).Amount = -1
   ItemDB(49).Damage = -1
   ItemDB(49).Armor = -1
   ItemDB(49).Condition = 2000
   ItemDB(49).IType = C_General
   ItemDB(49).Price = 2
   ItemDB(49).Multiple = False
   ItemDB(49).Movable = True
   ItemDB(49).CanBuy = 15000
   ItemDB(49).Decay = -1
   ItemDB(49).ForSale = False
End With

With ItemDB(50)
   ItemDB(50).IName = "Wine Bottle"
   ItemDB(50).IDesc = "It's a half empty wine bottle."
   ItemDB(50).ItemGUID = ""
   ItemDB(50).OnPlayer = False
   ItemDB(50).Equip = False
   ItemDB(50).Amount = -1
   ItemDB(50).Damage = -1
   ItemDB(50).Armor = -1
   ItemDB(50).Condition = 2000
   ItemDB(50).IType = C_General
   ItemDB(50).Price = 2
   ItemDB(50).Multiple = False
   ItemDB(50).Movable = True
   ItemDB(50).CanBuy = 15000
   ItemDB(50).Decay = -1
   ItemDB(50).ForSale = False
End With

With ItemDB(51)
   ItemDB(51).IName = "Blowdoll"
   ItemDB(51).IDesc = "Hmm.. You need a girlfriend."
   ItemDB(51).ItemGUID = ""
   ItemDB(51).OnPlayer = False
   ItemDB(51).Equip = False
   ItemDB(51).Amount = -1
   ItemDB(51).Damage = -1
   ItemDB(51).Armor = -1
   ItemDB(51).Condition = 2000
   ItemDB(51).IType = C_General
   ItemDB(51).Price = 4
   ItemDB(51).Multiple = False
   ItemDB(51).Movable = True
   ItemDB(51).CanBuy = -1
   ItemDB(51).Decay = -1
   ItemDB(51).ForSale = False
End With

With ItemDB(52)
   ItemDB(52).IName = "Police Badge"
   ItemDB(52).IDesc = "It's a standard issued cops badge."
   ItemDB(52).ItemGUID = ""
   ItemDB(52).OnPlayer = False
   ItemDB(52).Equip = False
   ItemDB(52).Amount = -1
   ItemDB(52).Damage = -1
   ItemDB(52).Armor = -1
   ItemDB(52).Condition = 2000
   ItemDB(52).IType = C_General
   ItemDB(52).Price = 115
   ItemDB(52).Multiple = False
   ItemDB(52).Movable = True
   ItemDB(52).CanBuy = -1
   ItemDB(52).Decay = -1
   ItemDB(52).ForSale = False
End With

With ItemDB(53)
   ItemDB(53).IName = "Large Dildo"
   ItemDB(53).IDesc = "A large bloody slime covered dildo, why don't you shove it up your ass."
   ItemDB(53).ItemGUID = ""
   ItemDB(53).OnPlayer = False
   ItemDB(53).Equip = False
   ItemDB(53).Amount = -1
   ItemDB(53).Damage = 4
   ItemDB(53).Armor = -1
   ItemDB(53).Condition = 2000
   ItemDB(53).IType = C_Melee
   ItemDB(53).Price = 3000
   ItemDB(53).Multiple = False
   ItemDB(53).Movable = True
   ItemDB(53).CanBuy = -1
   ItemDB(53).Decay = -1
   ItemDB(53).ForSale = False
End With

With ItemDB(54)
   ItemDB(54).IName = "PR-24 Baton"
   ItemDB(54).IDesc = "It's a standard police issued PR-24 Baton"
   ItemDB(54).ItemGUID = ""
   ItemDB(54).OnPlayer = False
   ItemDB(54).Equip = False
   ItemDB(54).Amount = -1
   ItemDB(54).Damage = 4
   ItemDB(54).Armor = -1
   ItemDB(54).Condition = 2000
   ItemDB(54).IType = C_Melee
   ItemDB(54).Price = 89
   ItemDB(54).Multiple = False
   ItemDB(54).Movable = True
   ItemDB(54).CanBuy = -1
   ItemDB(54).Decay = -1
   ItemDB(54).ForSale = False
End With

With ItemDB(55)
   ItemDB(55).IName = "Bank Roll"
   ItemDB(55).IDesc = "The newbie starter package."
   ItemDB(55).ItemGUID = ""
   ItemDB(55).OnPlayer = False
   ItemDB(55).Equip = False
   ItemDB(55).Amount = -1
   ItemDB(55).Damage = -1
   ItemDB(55).Armor = -1
   ItemDB(55).Condition = 2000
   ItemDB(55).IType = C_General
   ItemDB(55).Price = 10000
   ItemDB(55).Multiple = False
   ItemDB(55).Movable = True
   ItemDB(55).CanBuy = -1
   ItemDB(55).Decay = -1
   ItemDB(55).ForSale = False
End With

With ItemDB(56)
   ItemDB(56).IName = "Briefcase Full of Cash"
   ItemDB(56).IDesc = "Looks suspicious.  Where'd you get this?"
   ItemDB(56).ItemGUID = ""
   ItemDB(56).OnPlayer = False
   ItemDB(56).Equip = False
   ItemDB(56).Amount = -1
   ItemDB(56).Damage = -1
   ItemDB(56).Armor = -1
   ItemDB(56).Condition = 2000
   ItemDB(56).IType = C_General
   ItemDB(56).Price = 200000
   ItemDB(56).Multiple = False
   ItemDB(56).Movable = True
   ItemDB(56).CanBuy = 0
   ItemDB(56).Decay = -1
   ItemDB(56).ForSale = False
End With

With ItemDB(57)
   ItemDB(57).IName = "Winning Lottery Ticket"
   ItemDB(57).IDesc = "Lucky you!"
   ItemDB(57).ItemGUID = ""
   ItemDB(57).OnPlayer = False
   ItemDB(57).Equip = False
   ItemDB(57).Amount = -1
   ItemDB(57).Damage = -1
   ItemDB(57).Armor = -1
   ItemDB(57).Condition = 2000
   ItemDB(57).IType = C_General
   ItemDB(57).Price = 2000000
   ItemDB(57).Multiple = False
   ItemDB(57).Movable = True
   ItemDB(57).CanBuy = -1
   ItemDB(57).Decay = -1
   ItemDB(57).ForSale = False
End With

With ItemDB(58)
   ItemDB(58).IName = "FireWall"
   ItemDB(58).IDesc = "Give it back, or be banned."
   ItemDB(58).ItemGUID = ""
   ItemDB(58).OnPlayer = False
   ItemDB(58).Equip = False
   ItemDB(58).Amount = -1
   ItemDB(58).Damage = -1
   ItemDB(58).Armor = 20000
   ItemDB(58).Condition = 2000
   ItemDB(58).IType = C_Armor
   ItemDB(58).Price = 0
   ItemDB(58).Multiple = False
   ItemDB(58).Movable = True
   ItemDB(58).CanBuy = -1
   ItemDB(58).Decay = -1
   ItemDB(58).ForSale = False
End With

With ItemDB(59)
   ItemDB(59).IName = "Right hook"
   ItemDB(59).IDesc = "Give it back, or be banned."
   ItemDB(59).ItemGUID = ""
   ItemDB(59).OnPlayer = False
   ItemDB(59).Equip = False
   ItemDB(59).Amount = -1
   ItemDB(59).Damage = 1000
   ItemDB(59).Armor = -1
   ItemDB(59).Condition = 2000
   ItemDB(59).IType = C_Melee
   ItemDB(59).Price = 0
   ItemDB(59).Multiple = False
   ItemDB(59).Movable = True
   ItemDB(59).CanBuy = -1
   ItemDB(59).Decay = -1
   ItemDB(59).ForSale = False
End With

With ItemDB(60)
   ItemDB(60).IName = "Baton"
   ItemDB(60).IDesc = "A Baton"
   ItemDB(60).ItemGUID = ""
   ItemDB(60).OnPlayer = False
   ItemDB(60).Equip = False
   ItemDB(60).Amount = -1
   ItemDB(60).Damage = -1
   ItemDB(60).Armor = -1
   ItemDB(60).Condition = 2000
   ItemDB(60).IType = C_Rank2
   ItemDB(60).Price = 0.5
   ItemDB(60).Multiple = False
   ItemDB(60).Movable = True
   ItemDB(60).CanBuy = -1
   ItemDB(60).Decay = -1
   ItemDB(60).ForSale = False
End With

With ItemDB(61)
   ItemDB(61).IName = "Super Ammo"
   ItemDB(61).IDesc = "Powerful rounds, 1000 per clip."
   ItemDB(61).ItemGUID = ""
   ItemDB(61).OnPlayer = False
   ItemDB(61).Equip = False
   ItemDB(61).Amount = 1000
   ItemDB(61).Damage = 10
   ItemDB(61).Armor = -1
   ItemDB(61).Condition = 2000
   ItemDB(61).IType = C_Ammo
   ItemDB(61).Price = 10000
   ItemDB(61).Multiple = True
   ItemDB(61).Movable = True
   ItemDB(61).CanBuy = 150001
   ItemDB(61).Decay = -1
   ItemDB(61).ForSale = False
End With

With ItemDB(62)
   ItemDB(62).IName = "Knife"
   ItemDB(62).IDesc = "A switchblade"
   ItemDB(62).ItemGUID = ""
   ItemDB(62).OnPlayer = False
   ItemDB(62).Equip = False
   ItemDB(62).Amount = -1
   ItemDB(62).Damage = 45
   ItemDB(62).Armor = -1
   ItemDB(62).Condition = 2000
   ItemDB(62).IType = C_Gun
   ItemDB(62).Price = 2
   ItemDB(62).Multiple = False
   ItemDB(62).Movable = True
   ItemDB(62).CanBuy = -1
   ItemDB(62).Decay = -1
   ItemDB(62).ForSale = False
End With

With ItemDB(63)
   ItemDB(63).IName = "Pepper Spray"
   ItemDB(63).IDesc = "Used for blinding your opponent"
   ItemDB(63).ItemGUID = ""
   ItemDB(63).OnPlayer = False
   ItemDB(63).Equip = False
   ItemDB(63).Amount = -1
   ItemDB(63).Damage = -1
   ItemDB(63).Armor = -1
   ItemDB(63).Condition = 2000
   ItemDB(63).IType = C_Pepper
   ItemDB(63).Price = 20000
   ItemDB(63).Multiple = False
   ItemDB(63).Movable = True
   ItemDB(63).CanBuy = 0
   ItemDB(63).Decay = -1
   ItemDB(63).ForSale = True
End With

With ItemDB(64)
   ItemDB(64).IName = "Adrenalin"
   ItemDB(64).IDesc = "Gives you a burst of energy."
   ItemDB(64).ItemGUID = ""
   ItemDB(64).OnPlayer = False
   ItemDB(64).Equip = False
   ItemDB(64).Amount = -1
   ItemDB(64).Damage = -1
   ItemDB(64).Armor = -1
   ItemDB(64).Condition = 2000
   ItemDB(64).IType = C_Adrenalin
   ItemDB(64).Price = 20000
   ItemDB(64).Multiple = False
   ItemDB(64).Movable = True
   ItemDB(64).CanBuy = 1000
   ItemDB(64).Decay = -1
   ItemDB(64).ForSale = False
End With
With ItemDB(65)
   ItemDB(65).IName = "Land Mine"
   ItemDB(65).IDesc = "A Land Mine."
   ItemDB(65).ItemGUID = ""
   ItemDB(65).OnPlayer = False
   ItemDB(65).Equip = False
   ItemDB(65).Amount = -1
   ItemDB(65).Damage = -1
   ItemDB(65).Armor = -1
   ItemDB(65).Condition = 2000
   ItemDB(65).IType = C_Mine
   ItemDB(65).Price = 100000
   ItemDB(65).Multiple = False
   ItemDB(65).Movable = True
   ItemDB(65).CanBuy = 125000
   ItemDB(65).Decay = -1
   ItemDB(65).ForSale = True
End With
With ItemDB(66)
   ItemDB(66).IName = "Grenade"
   ItemDB(66).IDesc = "You got the nutz to pull the pin?"
   ItemDB(66).ItemGUID = ""
   ItemDB(66).OnPlayer = False
   ItemDB(66).Equip = False
   ItemDB(66).Amount = -1
   ItemDB(66).Damage = -1
   ItemDB(66).Armor = -1
   ItemDB(66).Condition = 2000
   ItemDB(66).IType = C_Grenade
   ItemDB(66).Price = 50000
   ItemDB(66).Multiple = False
   ItemDB(66).Movable = True
   ItemDB(66).CanBuy = 375000
   ItemDB(66).Decay = -1
   ItemDB(66).ForSale = True
 End With
 With ItemDB(67)
   ItemDB(67).IName = "Super Ammo"
   ItemDB(67).IDesc = "Perfect for Killing Anything Anywhere."
   ItemDB(67).ItemGUID = ""
   ItemDB(67).OnPlayer = False
   ItemDB(67).Equip = False
   ItemDB(67).Amount = 1000
   ItemDB(67).Damage = 0
   ItemDB(67).Armor = 7
   ItemDB(67).Condition = 2000
   ItemDB(67).IType = C_Ammo
   ItemDB(67).Price = 1000
   ItemDB(67).Multiple = True
   ItemDB(67).Movable = True
   ItemDB(67).CanBuy = 40001
   ItemDB(67).Decay = -1
   ItemDB(67).ForSale = False
   End With
With ItemDB(68)
   ItemDB(68).IName = "Ugly Stick"
   ItemDB(68).IDesc = "Really Really Ugly"
   ItemDB(68).ItemGUID = ""
   ItemDB(68).OnPlayer = False
   ItemDB(68).Equip = False
   ItemDB(68).Amount = -1
   ItemDB(68).Damage = 25000
   ItemDB(68).Armor = -1
   ItemDB(68).Condition = 2000
   ItemDB(68).IType = C_Melee
   ItemDB(68).Price = 1
   ItemDB(68).Multiple = False
   ItemDB(68).Movable = True
   ItemDB(68).CanBuy = 10000
   ItemDB(68).Decay = -1
   ItemDB(68).ForSale = False
End With
With ItemDB(69)
   ItemDB(69).IName = "Horn"
   ItemDB(69).IDesc = "Use it if you get in trouble"
   ItemDB(69).ItemGUID = ""
   ItemDB(69).OnPlayer = False
   ItemDB(69).Equip = False
   ItemDB(69).Amount = -1
   ItemDB(69).Damage = -1
   ItemDB(69).Armor = -1
   ItemDB(69).Condition = 2000
   ItemDB(69).IType = C_Horn
   ItemDB(69).Price = 5000
   ItemDB(69).Multiple = False
   ItemDB(69).Movable = True
   ItemDB(69).CanBuy = -999999
   ItemDB(69).Decay = -1
   ItemDB(69).ForSale = True
End With
With ItemDB(70)
   ItemDB(70).IName = "Biohazard Suit"
   ItemDB(70).IDesc = "Use it if you get in trouble"
   ItemDB(70).ItemGUID = ""
   ItemDB(70).OnPlayer = False
   ItemDB(70).Equip = False
   ItemDB(70).Amount = -1
   ItemDB(70).Damage = -1
   ItemDB(70).Armor = -1
   ItemDB(70).Condition = 2000
   ItemDB(70).IType = C_Armor
   ItemDB(70).Price = 100000
   ItemDB(70).Multiple = False
   ItemDB(70).Movable = True
   ItemDB(70).CanBuy = -999999
   ItemDB(70).Decay = -1
   ItemDB(70).ForSale = True
End With
With ItemDB(71)
   ItemDB(71).IName = "admin quick money"
   ItemDB(71).IDesc = "ehhh"
   ItemDB(71).ItemGUID = ""
   ItemDB(71).OnPlayer = False
   ItemDB(71).Equip = False
   ItemDB(71).Amount = -1
   ItemDB(71).Damage = -1
   ItemDB(71).Armor = -1
   ItemDB(71).Condition = 2000
   ItemDB(71).IType = C_Admin
   ItemDB(71).Price = 100000
   ItemDB(71).Multiple = False
   ItemDB(71).Movable = True
   ItemDB(71).CanBuy = -999999
   ItemDB(71).Decay = -1
   ItemDB(71).ForSale = False
End With

End Sub
Public Sub FullInventoryUpdate(Index As Integer)
On Error Resume Next
Dim a As Integer 'Counter
Dim b As Integer 'Counter
Dim c As Integer 'Counter
Dim msg As String
msg = Chr$(6)

'Update all 20 spots in players inventory
For a = 0 To 19
   If User(Index).Item(a) = -1 Then
      msg = msg & "<Empty>" & Chr$(1)
   ElseIf User(Index).Item(a) <> -1 Then
      
      'Check if item is a multiple item
      If Item(User(Index).Item(a)).IType = C_Ammo And _
         Item(User(Index).Item(a)).Amount > 0 And _
         Item(User(Index).Item(a)).Multiple = True Then
            msg = msg & "(" & Item(User(Index).Item(a)).Amount & ") "
      End If
            
      'Check to see if the item is equipted
      If Item(User(Index).Item(a)).Equip = True Then
            msg = msg & "<E> "
      End If
      
      'Add Item Name
      msg = msg & "<" & Item(User(Index).Item(a)).IName & ">" & Chr$(1)
   End If
Next a

msg = msg & Chr$(0)
frmMain.wsk(Index).SendData msg
DoEvents

End Sub

Public Sub SaveItems()
On Error Resume Next
Dim a As Integer 'Counter
Dim ff As Integer 'Free File

'Save all active world items
ff = FreeFile
Open App.Path & "\itemdata.dat" For Output As ff
For a = 0 To UBound(Item)
   If Item(a).IName <> "" And _
      Item(a).ItemGUID <> "" And _
      Item(a).Decay = -1 Then
         Write #ff, Item(a).IName
         Write #ff, Item(a).IDesc
         Write #ff, Item(a).ItemGUID
         Write #ff, Item(a).OnPlayer
         Write #ff, Item(a).Equip
         Write #ff, Item(a).Amount
         Write #ff, Item(a).Damage
         Write #ff, Item(a).Armor
         Write #ff, Item(a).Condition
         Write #ff, Item(a).IType
         Write #ff, Item(a).Price
         Write #ff, Item(a).Multiple
         Write #ff, Item(a).Movable
         Write #ff, Item(a).CanBuy
         Write #ff, Item(a).Decay
         Write #ff, Item(a).ILocation
   End If
Next a

Close ff
End Sub

Public Sub LoadItems()
On Error Resume Next
Dim a As Integer
Dim ff As Integer
a = 0

If IsFile(App.Path & "\itemdata.dat") = False Then
   Exit Sub
End If

ff = FreeFile
Open App.Path & "\itemdata.dat" For Input As ff
   Do While Not EOF(ff)
         ReDim Preserve Item(a)
         Input #ff, Item(a).IName
         Input #ff, Item(a).IDesc
         Input #ff, Item(a).ItemGUID
         Input #ff, Item(a).OnPlayer
         Input #ff, Item(a).Equip
         Input #ff, Item(a).Amount
         Input #ff, Item(a).Damage
         Input #ff, Item(a).Armor
         Input #ff, Item(a).Condition
         Input #ff, Item(a).IType
         Input #ff, Item(a).Price
         Input #ff, Item(a).Multiple
         Input #ff, Item(a).Movable
         Input #ff, Item(a).CanBuy
         Input #ff, Item(a).Decay
         Input #ff, Item(a).ILocation
         a = a + 1
   Loop

Close ff

End Sub

Public Sub UpdateGeneralInfo(Index As Integer)
On Error Resume Next
Call SetRank(Index)

frmMain.wsk(Index).SendData Chr$(255) & Chr$(7) & _
User(Index).UName & Chr$(1) & User(Index).Health & Chr$(1) & _
User(Index).Cash & Chr$(1) & User(Index).Bank & Chr$(1) & _
User(Index).HomeTown & Chr$(1) & User(Index).CurrTown & Chr$(1) & _
User(Index).Rank & Chr$(1) & User(Index).Kills & Chr$(1) & Chr$(0)
DoEvents

End Sub
Public Sub UpdateGearInfo(Index As Integer)
On Error GoTo ReSetGI
Dim a As String, b As String, c As String

If User(Index).Weapon <> -1 Then
   a = Item(User(Index).Weapon).IName
ElseIf User(Index).Weapon = -1 Then
   a = "Nothing"
End If

If User(Index).Armor <> -1 Then
   b = Item(User(Index).Armor).IName
ElseIf User(Index).Armor = -1 Then
   b = "Nothing"
End If

If User(Index).Ammo <> -1 Then
   c = Item(User(Index).Ammo).IName
ElseIf User(Index).Ammo = -1 Then
   c = "Nothing"
End If

frmMain.wsk(Index).SendData Chr$(254) & Chr$(2) & a & Chr$(1) & b & Chr$(1) & c & Chr$(1) & Chr$(0)
DoEvents
Exit Sub

'A common error that will reset the players Gear if
'It's wrong
ReSetGI:
Dim D As Integer

User(Index).Weapon = -1
User(Index).Armor = -1
User(Index).Ammo = -1

For D = 0 To UBound(Item)
   If Item(D).ItemGUID = User(Index).UserGUID And _
      Item(D).Equip = True Then
         Item(D).Equip = False
   End If
Next D

End Sub
Public Sub UpdatePlayerList()
On Error Resume Next
Dim a As Integer
Dim msg As String
msg = Chr$(254) & Chr$(3)

For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).UserGUID <> "" Then
         msg = msg & User(a).HomeAbv & "<" & User(a).UName & ">" & Chr$(1)
   End If
Next a

For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      frmMain.wsk(a).SendData msg & Chr$(0)
      DoEvents
   End If
Next a

End Sub

Public Sub SetGearValues()
On Error Resume Next
Dim a As Integer
Dim b As Integer

'Link weapons, armor, ammo to player when server loads
For a = 0 To UBound(UserDB)
   For b = 0 To 19
      If UserDB(a).Item(b) <> -1 Then
         If Item(UserDB(a).Item(b)).Equip = True Then
            If Item(UserDB(a).Item(b)).IType = C_Gun Or _
               Item(UserDB(a).Item(b)).IType = C_Melee Then
               UserDB(a).Weapon = UserDB(a).Item(b)
            ElseIf Item(UserDB(a).Item(b)).IType = C_Armor Then
               UserDB(a).Armor = UserDB(a).Item(b)
            ElseIf Item(UserDB(a).Item(b)).IType = C_Ammo Then
               UserDB(a).Ammo = UserDB(a).Item(b)
            End If
         End If
      End If
   Next b
Next a

End Sub

Public Sub RunGMCheck(Index As Integer)
On Error Resume Next

If User(Index).AccessLevel = 5 Then
   User(Index).HomeAbv = "<A>"
End If


End Sub


Public Sub LoadMap()
On Error Resume Next
Dim a As Integer

For a = 0 To UBound(City)
If City(a).CName = "New York" And _
   City(a).AirPort = True Then
   NYMap = NYMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New York" And _
   City(a).Bank = True Then
   NYMap = NYMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New York" And _
   City(a).Hospital = True Then
   NYMap = NYMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New York" And _
   City(a).Casino = True Then
   NYMap = NYMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New York" And _
   City(a).PawnShop = True Then
   NYMap = NYMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New York" And _
   City(a).Bar = True Then
   NYMap = NYMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "New Jersey" And _
   City(a).AirPort = True Then
   NJMap = NJMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New Jersey" And _
   City(a).Bank = True Then
   NJMap = NJMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New Jersey" And _
   City(a).Hospital = True Then
   NJMap = NJMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New Jersey" And _
   City(a).Casino = True Then
   NJMap = NJMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New Jersey" And _
   City(a).PawnShop = True Then
   NJMap = NJMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "New Jersey" And _
   City(a).Bar = True Then
   NJMap = NJMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "Miami" And _
   City(a).AirPort = True Then
   MIMap = MIMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Miami" And _
   City(a).Bank = True Then
   MIMap = MIMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Miami" And _
   City(a).Hospital = True Then
   MIMap = MIMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Miami" And _
   City(a).Casino = True Then
   MIMap = MIMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Miami" And _
   City(a).PawnShop = True Then
   MIMap = MIMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Miami" And _
   City(a).Bar = True Then
   MIMap = MIMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "Chicago" And _
   City(a).AirPort = True Then
   CHMap = CHMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Chicago" And _
   City(a).Bank = True Then
   CHMap = CHMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Chicago" And _
   City(a).Hospital = True Then
   CHMap = CHMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Chicago" And _
   City(a).Casino = True Then
   CHMap = CHMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Chicago" And _
   City(a).PawnShop = True Then
   CHMap = CHMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Chicago" And _
   City(a).Bar = True Then
   CHMap = CHMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "Houston" And _
   City(a).AirPort = True Then
   HOMap = HOMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Houston" And _
   City(a).Bank = True Then
   HOMap = HOMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Houston" And _
   City(a).Hospital = True Then
   HOMap = HOMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Houston" And _
   City(a).Casino = True Then
   HOMap = HOMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Houston" And _
   City(a).PawnShop = True Then
   HOMap = HOMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Houston" And _
   City(a).Bar = True Then
   HOMap = HOMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "Los Angeles" And _
   City(a).AirPort = True Then
   LAMap = LAMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Los Angeles" And _
   City(a).Bank = True Then
   LAMap = LAMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Los Angeles" And _
   City(a).Hospital = True Then
   LAMap = LAMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Los Angeles" And _
   City(a).Casino = True Then
   LAMap = LAMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Los Angeles" And _
   City(a).PawnShop = True Then
   LAMap = LAMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Los Angeles" And _
   City(a).Bar = True Then
   LAMap = LAMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If
If City(a).CName = "Sydney" And _
   City(a).AirPort = True Then
   SYDMap = SYDMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Sydney" And _
   City(a).Bank = True Then
   SYDMap = SYDMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Sydney" And _
   City(a).Hospital = True Then
   SYDMap = SYDMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Sydney" And _
   City(a).Casino = True Then
   SYDMap = SYDMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Sydney" And _
   City(a).PawnShop = True Then
   SYDMap = SYDMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "Sydney" And _
   City(a).Bar = True Then
   SYDMap = SYDMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

If City(a).CName = "London" And _
   City(a).AirPort = True Then
   UKMap = UKMap & City(a).CName & " International Airport (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "London" And _
   City(a).Bank = True Then
   UKMap = UKMap & City(a).CName & " Bank (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "London" And _
   City(a).Hospital = True Then
   UKMap = UKMap & City(a).CName & " Memorial Hospital (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "London" And _
   City(a).Casino = True Then
   UKMap = UKMap & City(a).CName & " Luxury Casino (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "London" And _
   City(a).PawnShop = True Then
   UKMap = UKMap & City(a).CName & " Pawn Shop (" & City(a).Compass & ")" & Chr$(1)
ElseIf City(a).CName = "London" And _
   City(a).Bar = True Then
   UKMap = UKMap & City(a).CName & " Sports Bar (" & City(a).Compass & ")" & Chr$(1)
End If

Next a

End Sub

Public Sub CombatMessage(IndexOne As Integer, IndexTwo As Integer, msg As String)
On Error Resume Next
Dim a As Integer 'Counter

'Show message to all players exept Index Player
For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Location = User(IndexOne).Location And _
      IndexOne <> a And IndexTwo <> a Then
         frmMain.wsk(a).SendData msg
         DoEvents
   End If
Next a

End Sub


Public Function PlayerKillPlayer(Index As Integer, ByVal Killed As Integer) As Boolean
Dim a As Integer
Dim b As Integer

'If a player dies,  drop all his items to the ground
If User(Killed).Health > 0 Then
   PlayerKillPlayer = False
   Exit Function
ElseIf User(Killed).Health <= 0 Then
   For a = 0 To 19
      If User(Killed).Item(a) <> -1 Then
         For b = 0 To UBound(City(User(Killed).Location).CItem)
            If City(User(Index).Location).CItem(b) = -1 Then
               City(User(Index).Location).CItem(b) = User(Killed).Item(a)
               Item(User(Killed).Item(a)).OnPlayer = False
               Item(User(Killed).Item(a)).Equip = False
               Item(User(Killed).Item(a)).Decay = GetTickCount()
               Item(User(Killed).Item(a)).ItemGUID = "" 'City(User(Killed).Location).CityGUID
               Item(User(Killed).Item(a)).ILocation = User(Killed).Location
               User(Killed).Item(a) = -1
               Exit For
            ElseIf b = UBound(City(User(Killed).Location).CItem) Then
               With City(User(Killed).Location)
               ReDim Preserve .CItem(UBound(.CItem) + 1)
               .CItem(UBound(.CItem)) = User(Killed).Item(a)
               Item(User(Killed).Item(a)).OnPlayer = False
               Item(User(Killed).Item(a)).Equip = False
               Item(User(Killed).Item(a)).Decay = GetTickCount()
               Item(User(Killed).Item(a)).ItemGUID = "" 'City(User(Killed).Location).CityGUID
               Item(User(Killed).Item(a)).ILocation = User(Killed).Location
               User(Killed).Item(a) = -1
               End With
            End If
         Next b
      End If
   Next a

   Call FullInventoryUpdate(Killed)
   User(Index).Cash = User(Index).Cash + User(Killed).Cash
   User(Index).Cash = User(Index).Cash + User(Killed).Bounty
   User(Index).Kills = User(Index).Kills + 1
   User(Killed).Cash = 50
   User(Killed).Bounty = 0
   User(Killed).Health = 50
   User(Killed).Accuracy = User(Killed).Accuracy * 0.96
   User(Killed).Search = User(Killed).Search * 0.96
   User(Killed).Tracking = User(Killed).Tracking * 0.96
   User(Killed).Hiding = User(Killed).Hiding * 0.96
   User(Killed).Snooping = User(Killed).Snooping * 0.96
   User(Killed).Sniping = User(Killed).Sniping * 0.96
   User(Killed).Stealing = User(Killed).Stealing * 0.96
   Call PlaceOnDeath(Killed)
   For a = 1 To MaxUsers
   If User(a).Status = "Playing" And _
      User(a).Mute = False Then
   frmMain.wsk(a).SendData Chr$(8) & vbCrLf & "*System Blurb* - " & User(Index).UName & " has murdered " & User(Killed).UName & "." & Chr$(0)
   DoEvents
   End If
   Next a
   frmMain.wsk(Killed).SendData Chr$(2) & User(Index).UName & " has just punked you like the little bitch you are." & vbCrLf & vbCrLf & Chr$(0)
   
   DoEvents
   frmMain.wsk(Index).SendData Chr$(2) & "You just put " & User(Killed).UName & " in his place, six feet under and in a box..." & vbCrLf & User(Killed).UName & "'s items fall to the ground." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Call ShowWatchers(Index, Chr$(2) & "You just witnessed " & User(Index).UName & " murder " & User(Killed).UName & " in cold blood." & vbCrLf & "You see " & User(Killed).UName & " 's items fall to the ground." & vbCrLf & vbCrLf & Chr$(0))
   Call UpdateGeneralInfo(Index)
   Call UpdateGeneralInfo(Killed)
   User(Killed).Weapon = -1
   User(Killed).Armor = -1
   User(Killed).Ammo = -1
   Call UpdateGearInfo(Killed)
   User(Killed).TargetNum = -1
   User(Killed).TargetGUID = ""
   User(Index).TargetNum = -1
   User(Index).TargetGUID = ""
   PlayerKillPlayer = True
End If

End Function
Public Sub PlaceOnDeath(Index As Integer)
On Error Resume Next

Randomize

Select Case City(User(Index).Location).CName
   Case "New York"
      User(Index).Location = Int(899 - 0) * Rnd
   Case "Miami"
      User(Index).Location = Int(1799 - 800) * Rnd + 800
   Case "Houston"
      User(Index).Location = Int(2699 - 1800) * Rnd + 1800
   Case "Los Angeles"
      User(Index).Location = Int(3599 - 2700) * Rnd + 2700
   Case "Chicago"
      User(Index).Location = Int(4499 - 3600) * Rnd + 3600
   Case "New Jersey"
      User(Index).Location = Int(5399 - 4500) * Rnd + 4500
   Case "Sydney"
      User(Index).Location = Int(6299 - 5400) * Rnd + 5400
   Case "London"
      User(Index).Location = Int(7199 - 6300) * Rnd + 6300
      
      End Select

End Sub

Public Function PlayerKillNpc(Index As Integer) As Boolean
Dim a As Integer

If Npc(User(Index).TargetNum).NHealth > 0 Then
   PlayerKillNpc = False
   Exit Function
ElseIf Npc(User(Index).TargetNum).NHealth <= 0 Then
   PlayerKillNpc = True
   Call NpcDropGear(Index)
   For a = 0 To UBound(City(User(Index).Location).CNpc)
      If City(User(Index).Location).CNpc(a) = User(Index).TargetNum Then
         City(User(Index).Location).CNpc(a) = -1
         Exit For
      End If
   Next a
   'Check Npc type and set reputation accordingly
   If Npc(User(Index).TargetNum).NpcType = N_Druggie Or _
      Npc(User(Index).TargetNum).NpcType = N_Dealer Then
         User(Index).Reputation = User(Index).Reputation - 250
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_Cop Then
         User(Index).Reputation = User(Index).Reputation + 20
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_Bum Then
         User(Index).Reputation = User(Index).Reputation + 5
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_Tweaker Then
         User(Index).Reputation = User(Index).Reputation + 10
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_Bouncer Then
         User(Index).Reputation = User(Index).Reputation + 15
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
     ElseIf Npc(User(Index).TargetNum).NpcType = N_Hooker Then
         User(Index).Reputation = User(Index).Reputation + 10
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_LoanShark Then
         User(Index).Reputation = User(Index).Reputation + 25
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_Politician Then
         User(Index).Reputation = User(Index).Reputation + 200
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_Taliban Then
         User(Index).Reputation = User(Index).Reputation + 100
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_Pedestrian Then
         User(Index).Reputation = User(Index).Reputation + 0
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_MeterMaid Then
         User(Index).Reputation = User(Index).Reputation + 20
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
    ElseIf Npc(User(Index).TargetNum).NpcType = N_Whore Then
         User(Index).Reputation = User(Index).Reputation + 5
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_FBI Then
         User(Index).Reputation = User(Index).Reputation + 50
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_DEA Then
         User(Index).Reputation = User(Index).Reputation + 100
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   ElseIf Npc(User(Index).TargetNum).NpcType = N_ATF Then
         User(Index).Reputation = User(Index).Reputation + 250
         User(Index).Cash = User(Index).Cash + Npc(User(Index).TargetNum).NCash
         User(Index).Kills = User(Index).Kills + 1
         Call UpdateGeneralInfo(Index)
   End If
   'Set misc. settings to default and delete npc from world
   frmMain.wsk(Index).SendData Chr$(2) & "You just put " & Npc(User(Index).TargetNum).NName & " in his place, six feet under and in a box..." & vbCrLf & Npc(User(Index).TargetNum).NName & "'s items fall to the ground." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Call ShowWatchers(Index, Chr$(2) & "You just witnessed " & User(Index).UName & " murder " & Npc(User(Index).TargetNum).NName & " in cold blood." & vbCrLf & "You see " & Npc(User(Index).TargetNum).NName & " 's items fall to the ground." & vbCrLf & vbCrLf & Chr$(0))
   If Npc(User(Index).TargetNum).NpcType = N_Dealer Or _
      Npc(User(Index).TargetNum).NpcType = N_Druggie Then
      Dim b As Integer
      For b = 1 To MaxUsers
         If User(b).Status = "Playing" Then
            frmMain.wsk(b).SendData Chr$(252) & Chr$(3) & "<News Flash>  Rumor on the street is " & User(Index).UName & " from " & User(Index).HomeTown & " has been slaughtering druggies and dealers!  This street crime shouldn't go unpunished!" & Chr$(0)
            DoEvents
         End If
      Next b
   End If
   If Npc(User(Index).TargetNum).NpcType = N_Taliban Then
      Dim c As Integer
      For c = 1 To MaxUsers
         If User(c).Status = "Playing" Then
            frmMain.wsk(c).SendData Chr$(252) & Chr$(3) & "<News Flash>  Rumor on the street is " & User(Index).UName & " from " & User(Index).HomeTown & " has been killing Taliban scum!  His reputation skyrockets!" & Chr$(0)
            DoEvents
         End If
      Next c
   End If
   Call ResetNPC(User(Index).TargetNum)
   User(Index).TargetNum = -1
   User(Index).TargetGUID = ""
End If
      
End Function

Public Sub NpcDropGear(Index As Integer)
On Error Resume Next
Dim a As Integer
Dim b As Integer

'Drop NPC's gear and Items
For a = 0 To 2
   If Npc(User(Index).TargetNum).NGear(a) <> -1 Then
      For b = 0 To UBound(City(User(Index).Location).CItem)
         If City(User(Index).Location).CItem(b) = -1 Then
            City(User(Index).Location).CItem(b) = Npc(User(Index).TargetNum).NGear(a)
            Item(Npc(User(Index).TargetNum).NGear(a)).OnPlayer = False
            Item(Npc(User(Index).TargetNum).NGear(a)).Decay = GetTickCount()
            Item(Npc(User(Index).TargetNum).NGear(a)).ItemGUID = ""
            Item(Npc(User(Index).TargetNum).NGear(a)).Equip = False
            Item(Npc(User(Index).TargetNum).NGear(a)).ILocation = User(Index).Location
            Exit For
         ElseIf b = UBound(City(User(Index).Location).CItem) Then
            With City(User(Index).Location)
            ReDim Preserve .CItem(UBound(.CItem) + 1)
            .CItem(UBound(.CItem)) = Npc(User(Index).TargetNum).NGear(a)
            Item(Npc(User(Index).TargetNum).NGear(a)).OnPlayer = False
            Item(Npc(User(Index).TargetNum).NGear(a)).Decay = GetTickCount()
            Item(Npc(User(Index).TargetNum).NGear(a)).ItemGUID = ""
            Item(Npc(User(Index).TargetNum).NGear(a)).Equip = False
            Item(Npc(User(Index).TargetNum).NGear(a)).ILocation = User(Index).Location
            End With
            Exit For
         End If
     Next b
   End If
Next a

For a = 0 To 19
   If Npc(User(Index).TargetNum).NItem(a) <> -1 Then
      For b = 0 To UBound(City(User(Index).Location).CItem)
         If City(User(Index).Location).CItem(b) = -1 Then
            City(User(Index).Location).CItem(b) = Npc(User(Index).TargetNum).NItem(a)
            Item(Npc(User(Index).TargetNum).NItem(a)).OnPlayer = False
            Item(Npc(User(Index).TargetNum).NItem(a)).Decay = GetTickCount()
            Item(Npc(User(Index).TargetNum).NItem(a)).ItemGUID = ""
            Item(Npc(User(Index).TargetNum).NItem(a)).Equip = False
            Item(Npc(User(Index).TargetNum).NItem(a)).ILocation = User(Index).Location
            Exit For
         ElseIf b = UBound(City(User(Index).Location).CItem) Then
            With City(User(Index).Location)
            ReDim Preserve .CItem(UBound(.CItem) + 1)
            .CItem(UBound(.CItem)) = Npc(User(Index).TargetNum).NItem(a)
            Item(Npc(User(Index).TargetNum).NItem(a)).OnPlayer = False
            Item(Npc(User(Index).TargetNum).NItem(a)).Decay = GetTickCount()
            Item(Npc(User(Index).TargetNum).NItem(a)).ItemGUID = ""
            Item(Npc(User(Index).TargetNum).NItem(a)).Equip = False
            Item(Npc(User(Index).TargetNum).NItem(a)).ILocation = User(Index).Location
            End With
            Exit For
         End If
     Next b
   End If
Next a

End Sub

Public Function Skilldelay(Index As Integer) As Boolean

User(Index).SkillTickNew = GetTickCount()
If User(Index).AccessLevel = 5 Or Adrenalin(Index) = True Then
   Skilldelay = False
   Exit Function
End If

If User(Index).SkillTickNew - User(Index).SkillTickOld > SkillDelayTick Then
   Skilldelay = False
   User(Index).SkillTickOld = GetTickCount()
   Exit Function
Else
   Skilldelay = True
   frmMain.wsk(Index).SendData Chr$(2) & "You must wait a few seconds before using that skill." & vbCrLf & vbCrLf & Chr$(0)
   DoEvents
   Exit Function
End If

End Function
Public Function Adrenalin(Index As Integer) As Boolean

'adrenalin
User(Index).AdrenalinTickNew = GetTickCount()
If User(Index).AdrenalinTickNew - User(Index).AdrenalinTickOld > AdrenalinTick Then
    Adrenalin = False
    Exit Function
Else
    Adrenalin = True
    Exit Function
End If

End Function

Public Function CheckIPBan(Index As Integer) As Boolean
Dim a As Integer

For a = 0 To UBound(IPBan)
   If frmMain.wsk(Index).RemoteHostIP = IPBan(a) Then
      CheckIPBan = True
      frmMain.wsk(Index).SendData Chr$(2) & "Your IP address has been banned by the server administrator.  If you feel this action was unjust, contact the administrator by e-mail at jimmyjames@writeme.com to resolve the issue." & vbCrLf & vbCrLf & "Have a nice day...." & vbCrLf & vbCrLf & Chr$(0)
      DoEvents
      With frmMain.txtOutput
        .Text = .Text & "IP BAN LOG IN ATTEMPT: " & frmMain.wsk(Index).RemoteHostIP & vbCrLf
        .SelStart = Len(.Text)
      End With
      frmMain.wsk(Index).Close
      frmMain.lstUsers.List(Index - 1) = "<Waiting>"
      Exit Function
   End If
Next a

CheckIPBan = False

End Function

Public Sub PlayerHealth()
On Error Resume Next
Dim a As Integer

For a = 1 To MaxUsers
   If User(a).Status = "Playing" Then
      If User(a).Health < 100 Then
         User(a).Health = User(a).Health + 1
         Call UpdateGeneralInfo(a)
      End If
   End If
Next a

End Sub
