Attribute VB_Name = "modTestBaseballClass"
Option Explicit

Sub TestDecoupledBaseballClass()
    Dim game As New clsBaseballGame
    
    ' 1. Build In-Memory Lineups
    Dim awayRoster As New Collection
    Dim homeRoster As New Collection
    Dim p As clsPlayer
    
    ' Add Away Batter
    Set p = New clsPlayer
    p.Name = "Away Star Hitter"
    p.IsPitcher = False
    awayRoster.Add p
    
    ' Add Home Pitcher
    Set p = New clsPlayer
    p.Name = "Home Ace Pitcher"
    p.IsPitcher = True
    homeRoster.Add p
    
    ' 2. Inject Rosters into Game Engine
    Set game.LineUpAway = awayRoster
    Set game.LineUpHome = homeRoster
    
    ' 3. Assert Decisions
    game.Half = "Top" ' Away team batting, Home pitching
    game.BatterIndexAway = 1
    
    ' Test decoupled matching
    Debug.Assert game.CurrentBatter.Name = "Away Star Hitter"
    Debug.Assert game.CurrentPitcher.Name = "Home Ace Pitcher"
    
    Debug.Print "Success: Decoupled testing passed completely in memory!"
End Sub
