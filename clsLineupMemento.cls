VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLineupMemento"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public LineUpAway As Collection
Public LineUpHome As Collection

' Deep-copies both lineup collections so the snapshot is fully
' independent of the live lineups. clsPlayer is a flat value
' object so field-by-field copy is sufficient — no nested refs.
Public Sub CaptureFrom(ByVal game As clsBaseballGame)
    Set LineUpAway = DeepCopyLineup(game.LineUpAway)
    Set LineUpHome = DeepCopyLineup(game.LineUpHome)
End Sub

Private Function DeepCopyLineup(ByVal source As Collection) As Collection
    Dim result As New Collection
    Dim original As clsPlayer
    Dim copy As clsPlayer
    For Each original In source
        Set copy = New clsPlayer
        copy.CopyFrom original
        result.Add copy
    Next original
    Set DeepCopyLineup = result
End Function
