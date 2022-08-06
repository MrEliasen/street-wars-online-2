Attribute VB_Name = "SkillMod"
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


Public Function RunAccuracy(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Accuracy > 100# Then
  User(Index).Accuracy = 100#
End If

If User(Index).Accuracy < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.03
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 30# And _
  User(Index).Accuracy < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.025
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 40# And _
  User(Index).Accuracy < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.02
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 50# And _
  User(Index).Accuracy < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.008
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 60# And _
  User(Index).Accuracy < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.003
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 70# And _
  User(Index).Accuracy < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.0008
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 80# And _
  User(Index).Accuracy < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     User(Index).Accuracy = User(Index).Accuracy + 0.00009
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
ElseIf User(Index).Accuracy >= 90# And _
  User(Index).Accuracy <= 100# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Accuracy Then
     If User(Index).Accuracy < 100# Then
       User(Index).Accuracy = User(Index).Accuracy + 0.000005
     End If
     RunAccuracy = True
     Exit Function
  ElseIf a > User(Index).Accuracy Then
     RunAccuracy = False
     Exit Function
  End If
End If

End Function

Public Function RunHiding(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Hiding > 100# Then
  User(Index).Hiding = 100#
End If

If User(Index).Hiding < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.1
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 30# And _
  User(Index).Hiding < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.05
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 40# And _
  User(Index).Hiding < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.02
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 50# And _
  User(Index).Hiding < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.008
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 60# And _
  User(Index).Hiding < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.005
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 70# And _
  User(Index).Hiding < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.002
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 80# And _
  User(Index).Hiding < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     User(Index).Hiding = User(Index).Hiding + 0.0008
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
ElseIf User(Index).Hiding >= 90# And _
  User(Index).Hiding <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Hiding Then
     If User(Index).Hiding < 100# Then
       User(Index).Hiding = User(Index).Hiding + 0.0005
     End If
     RunHiding = True
     Exit Function
  ElseIf a > User(Index).Hiding Then
     RunHiding = False
     Exit Function
  End If
End If

End Function


Public Function RunTracking(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Tracking > 100# Then
  User(Index).Tracking = 100#
End If

If User(Index).Tracking < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.1
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 30# And _
  User(Index).Tracking < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.05
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 40# And _
  User(Index).Tracking < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.02
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 50# And _
  User(Index).Tracking < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.008
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 60# And _
  User(Index).Tracking < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.005
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 70# And _
  User(Index).Tracking < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.002
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 80# And _
  User(Index).Tracking < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     User(Index).Tracking = User(Index).Tracking + 0.0008
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
ElseIf User(Index).Tracking >= 90# And _
  User(Index).Tracking <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Tracking Then
     If User(Index).Tracking < 100# Then
       User(Index).Tracking = User(Index).Tracking + 0.0005
     End If
     RunTracking = True
     Exit Function
  ElseIf a > User(Index).Tracking Then
     RunTracking = False
     Exit Function
  End If
End If

End Function
Public Function RunSniping(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Sniping > 100# Then
  User(Index).Sniping = 100#
End If

If User(Index).Sniping < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.1
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 30# And _
  User(Index).Sniping < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.05
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 40# And _
  User(Index).Sniping < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.02
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 50# And _
  User(Index).Sniping < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.008
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 60# And _
  User(Index).Sniping < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.005
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 70# And _
  User(Index).Sniping < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.002
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 80# And _
  User(Index).Sniping < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     User(Index).Sniping = User(Index).Sniping + 0.0008
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
ElseIf User(Index).Sniping >= 90# And _
  User(Index).Sniping <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Sniping Then
     If User(Index).Sniping < 100# Then
       User(Index).Sniping = User(Index).Sniping + 0.0005
     End If
     RunSniping = True
     Exit Function
  ElseIf a > User(Index).Sniping Then
     RunSniping = False
     Exit Function
  End If
End If

End Function

Public Function RunSearch(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Search > 100# Then
  User(Index).Search = 100#
End If

If User(Index).Search < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.1
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 30# And _
  User(Index).Search < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.05
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 40# And _
  User(Index).Search < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.02
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 50# And _
  User(Index).Search < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.008
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 60# And _
  User(Index).Search < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.005
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 70# And _
  User(Index).Search < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.002
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 80# And _
  User(Index).Search < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     User(Index).Search = User(Index).Search + 0.0008
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
ElseIf User(Index).Search >= 90# And _
  User(Index).Search <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Search Then
     If User(Index).Search < 100# Then
       User(Index).Search = User(Index).Search + 0.0005
     End If
     RunSearch = True
     Exit Function
  ElseIf a > User(Index).Search Then
     RunSearch = False
     Exit Function
  End If
End If

End Function
Public Function RunSnooping(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Snooping > 100# Then
  User(Index).Snooping = 100#
End If

If User(Index).Snooping < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.1
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 30# And _
  User(Index).Snooping < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.05
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 40# And _
  User(Index).Snooping < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.02
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 50# And _
  User(Index).Snooping < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.008
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 60# And _
  User(Index).Snooping < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.005
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 70# And _
  User(Index).Snooping < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.002
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 80# And _
  User(Index).Snooping < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     User(Index).Snooping = User(Index).Snooping + 0.0008
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
ElseIf User(Index).Snooping >= 90# And _
  User(Index).Snooping <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Snooping Then
     If User(Index).Snooping < 100# Then
       User(Index).Snooping = User(Index).Snooping + 0.0005
     End If
     RunSnooping = True
     Exit Function
  ElseIf a > User(Index).Snooping Then
     RunSnooping = False
     Exit Function
  End If
End If

End Function
Public Function runStealing(Index As Integer) As Boolean
Dim a As Integer

If User(Index).Stealing > 100# Then
  User(Index).Stealing = 100#
End If

If User(Index).Stealing < 30# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.1
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 30# And _
  User(Index).Stealing < 40# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.05
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 40# And _
  User(Index).Stealing < 50# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.02
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 50# And _
  User(Index).Stealing < 60# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.008
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 60# And _
  User(Index).Stealing < 70# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.005
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 70# And _
  User(Index).Stealing < 80# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.002
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 80# And _
  User(Index).Stealing < 90# Then
  Randomize
  a = Int(100 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     User(Index).Stealing = User(Index).Stealing + 0.0008
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
ElseIf User(Index).Stealing >= 90# And _
  User(Index).Stealing <= 100# Then
  Randomize
  a = Int(112 - 1) * Rnd + 1
  If a <= User(Index).Stealing Then
     If User(Index).Stealing < 100# Then
       User(Index).Stealing = User(Index).Stealing + 0.0005
     End If
     runStealing = True
     Exit Function
  ElseIf a > User(Index).Stealing Then
     runStealing = False
     Exit Function
  End If
End If

End Function
