VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGameMemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------
' Flat snapshot of every mutable field in clsBaseballGame.
' Created by clsBaseballGame.CaptureMemento before each user action
' and consumed by clsBaseballGame.RestoreMemento on undo.
'
' Lineup collections (m_LineUpAway / m_LineUpHome) are intentionally
' excluded — they are injected once at startup and never mutated
' during a game, so they do not need to participate in undo.
' ----------------------------------------------------------------
Public Inning         As Long
Public Half           As String
Public Outs           As Long
Public BatterIndexAway As Long
Public BatterIndexHome As Long
Public Runner1B       As String
Public Runner2B       As String
Public Runner3B       As String
Public Balls          As Long
Public Strikes        As Long
Public PitchSequence  As String
Public ScoreAway      As Long
Public ScoreHome      As Long
Public Pitchcounter   As Long
Public OutsPitched    As Long

Public HasLogEntry As Boolean ' Marker to check if the GameHistoryLogger logged an event or not e.g in case of an foulball there is no event logged
Public LineupSnapshot As clsLineupMemento ' Contains Nothing for non-substitution actions
