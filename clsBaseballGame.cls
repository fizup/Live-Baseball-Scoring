VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsBaseballGame"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------
' Private backing fields
' ----------------------------------------------------------------
Private m_Inning As Long
Private m_Half As String
Private m_BatterIndexAway As Long
Private m_BatterIndexHome As Long
Private m_Runner1B As String
Private m_Runner2B As String
Private m_Runner3B As String
Private m_Balls As Long
Private m_Strikes As Long
Private m_PitchSequence As String
Private m_Outs As Long
Private m_ScoreAway As Long
Private m_ScoreHome As Long
Private m_EventToHappen As String

' Collections to hold clsPlayer objects instead of relying on worksheets
Private m_LineUpAway As Collection
Private m_LineUpHome As Collection

' ----------------------------------------------------------------
' Injecting dependencies via properties (No worksheet references!)
' ----------------------------------------------------------------
Public Property Get LineUpAway() As Collection
    Set LineUpAway = m_LineUpAway
End Property
Public Property Set LineUpAway(ByVal Value As Collection)
    Set m_LineUpAway = Value
End Property

Public Property Get LineUpHome() As Collection
    Set LineUpHome = m_LineUpHome
End Property
Public Property Set LineUpHome(ByVal Value As Collection)
    Set m_LineUpHome = Value
End Property

' ----------------------------------------------------------------
' Refactored Computed Properties
' ----------------------------------------------------------------
Public Property Get BattingTeam() As String
    If m_Half = "Top" Then
        BattingTeam = "Away"
    Else
        BattingTeam = "Home"
    End If
End Property

Public Property Get PitchingTeam() As String
    If m_Half = "Top" Then
        PitchingTeam = "Home"
    Else
        PitchingTeam = "Away"
    End If
End Property

Public Property Get CurrentBatter() As clsPlayer
    Dim currentLineUp As Collection
    Dim targetSpot As Long
    
    ' Select correct roster in memory
    If Me.BattingTeam = "Away" Then
        Set currentLineUp = m_LineUpAway
        targetSpot = m_BatterIndexAway
    Else
        Set currentLineUp = m_LineUpHome
        targetSpot = m_BatterIndexHome
    End If
    
    ' Find the active player currently occupying this batting spot.
    ' Spot is the authoritative batting-order position — substitutions
    ' change Spot on existing players rather than reordering the
    ' Collection, so the player at this Spot may not sit at index targetSpot.
    If currentLineUp Is Nothing Then Exit Property
    
    Dim p As clsPlayer
    For Each p In currentLineUp
        If p.Spot = targetSpot Then
            Set CurrentBatter = p
            Exit Property
        End If
    Next p
End Property

Public Property Get CurrentPitcher() As clsPlayer
    Dim currentLineUp As Collection
    Dim p As clsPlayer
    
    ' The pitching team is the team currently in the field
    If Me.PitchingTeam = "Away" Then
        Set currentLineUp = m_LineUpAway
    Else
        Set currentLineUp = m_LineUpHome
    End If
    
    ' Scan collection for the designated pitcher
    If Not currentLineUp Is Nothing Then
        For Each p In currentLineUp
            If p.IsPitcher Then
                Set CurrentPitcher = p
                Exit Property
            End If
        Next p
    End If
End Property

Public Property Get BaseState() As String
    BaseState = _
        IIf(m_Runner1B <> "", "1", "_") & _
        IIf(m_Runner2B <> "", "2", "_") & _
        IIf(m_Runner3B <> "", "3", "_")
End Property

' ----------------------------------------------------------------
' Standard Properties & Game-flow methods (Kept from original)
' ----------------------------------------------------------------
Public Property Get Inning() As Long: Inning = m_Inning: End Property
Public Property Let Inning(ByVal Value As Long): m_Inning = Value: End Property
Public Property Get Half() As String: Half = m_Half: End Property
Public Property Let Half(ByVal Value As String): m_Half = Value: End Property
Public Property Get BatterIndexAway() As Long: BatterIndexAway = m_BatterIndexAway: End Property
Public Property Let BatterIndexAway(ByVal Value As Long): m_BatterIndexAway = Value: End Property
Public Property Get BatterIndexHome() As Long: BatterIndexHome = m_BatterIndexHome: End Property
Public Property Let BatterIndexHome(ByVal Value As Long): m_BatterIndexHome = Value: End Property
Public Property Get Outs() As Long: Outs = m_Outs: End Property
Public Property Let Outs(ByVal Value As Long): m_Outs = Value: End Property
Public Property Get Runner1B() As String: Runner1B = m_Runner1B: End Property
Public Property Let Runner1B(ByVal Value As String): m_Runner1B = Value: End Property
Public Property Get Runner2B() As String: Runner2B = m_Runner2B: End Property
Public Property Let Runner2B(ByVal Value As String): m_Runner2B = Value: End Property
Public Property Get Runner3B() As String: Runner3B = m_Runner3B: End Property
Public Property Let Runner3B(ByVal Value As String): m_Runner3B = Value: End Property
Public Property Get Balls() As Long: Balls = m_Balls: End Property
Public Property Let Balls(ByVal Value As Long): m_Balls = Value: End Property
Public Property Get Strikes() As Long: Strikes = m_Strikes: End Property
Public Property Let Strikes(ByVal Value As Long): m_Strikes = Value: End Property
Public Property Get PitchSequence() As String: PitchSequence = m_PitchSequence: End Property
Public Property Let PitchSequence(ByVal Value As String): m_PitchSequence = Value: End Property
Public Property Get ScoreAway() As Long: ScoreAway = m_ScoreAway: End Property
Public Property Let ScoreAway(ByVal Value As Long): m_ScoreAway = Value: End Property
Public Property Get ScoreHome() As Long: ScoreHome = m_ScoreHome: End Property
Public Property Let ScoreHome(ByVal Value As Long): m_ScoreHome = Value: End Property
Public Property Get EventToHappen() As String: EventToHappen = m_EventToHappen: End Property
Public Property Let EventToHappen(ByVal Value As String): m_EventToHappen = Value: End Property

Public Sub AdvanceBatter()
    If m_Half = "Top" Then
        m_BatterIndexAway = (m_BatterIndexAway Mod 9) + 1
    Else
        m_BatterIndexHome = (m_BatterIndexHome Mod 9) + 1
    End If
End Sub

'Public Sub AddOuts(Optional inpValue As Long)
'    If inpValue = 0 Then inpValue = 1
' -> code fragile cause it can't distinguish "passed 0" from "not passed"
Public Sub AddOuts(Optional inpValue As Variant = 1)
    If IsMissing(inpValue) Then inpValue = 1
    
    Me.CurrentPitcher.OutsPitched = Me.CurrentPitcher.OutsPitched + inpValue
    m_Outs = m_Outs + inpValue
    If m_Outs >= 3 Then ChangeSides
End Sub

Public Sub AppendPitch(ByVal PitchCode As String)
    m_PitchSequence = m_PitchSequence & PitchCode
End Sub

Public Sub ResetCount()
    m_Balls = 0
    m_Strikes = 0
    m_PitchSequence = "" ' Clear the sequence string for the new at-bat
End Sub

'Public Sub ScoreRuns(Optional inpValue As Long)
'    If inpValue = 0 Then inpValue = 1
' -> code fragile cause it can't distinguish "passed 0" from "not passed"
Public Sub ScoreRuns(Optional inpValue As Variant = 1)
    If IsMissing(inpValue) Then inpValue = 1
    
    If m_Half = "Top" Then
        m_ScoreAway = m_ScoreAway + inpValue
    Else
        m_ScoreHome = m_ScoreHome + inpValue
    End If
End Sub

Private Sub ChangeSides()
    If m_Half = "Top" Then
        m_Half = "Bot"
    Else
        m_Half = "Top"
        m_Inning = m_Inning + 1
    End If
    m_Runner1B = "": m_Runner2B = "": m_Runner3B = "": m_Outs = 0
    ResetCount ' Resets Pitch Count
End Sub

' ----------------------------------------------------------------
' Memento support — snapshot and restore all mutable game state.
' Called by frmRecordPlay before and after every user action.
' ----------------------------------------------------------------
 
' Returns a complete snapshot of the current game state.
' The caller (frmRecordPlay) pushes this onto the undo stack.
Public Function CaptureMemento() As clsGameMemento
    Dim m As clsGameMemento
    Set m = New clsGameMemento
    
    m.Inning = m_Inning
    m.Half = m_Half
    m.Outs = m_Outs
    m.BatterIndexAway = m_BatterIndexAway
    m.BatterIndexHome = m_BatterIndexHome
    m.Runner1B = m_Runner1B
    m.Runner2B = m_Runner2B
    m.Runner3B = m_Runner3B
    m.Balls = m_Balls
    m.Strikes = m_Strikes
    m.PitchSequence = m_PitchSequence
    m.ScoreAway = m_ScoreAway
    m.ScoreHome = m_ScoreHome
    m.Pitchcounter = Me.CurrentPitcher.Pitchcounter
    m.OutsPitched = Me.CurrentPitcher.OutsPitched
    
    Set CaptureMemento = m
End Function
 
' Writes all fields from a memento back into this engine instance.
' After this call the game is in exactly the state it was when
' CaptureMemento was called — no other side effects.
Public Sub RestoreMemento(ByVal m As clsGameMemento)
    m_Inning = m.Inning
    m_Half = m.Half
    m_Outs = m.Outs
    m_BatterIndexAway = m.BatterIndexAway
    m_BatterIndexHome = m.BatterIndexHome
    m_Runner1B = m.Runner1B
    m_Runner2B = m.Runner2B
    m_Runner3B = m.Runner3B
    m_Balls = m.Balls
    m_Strikes = m.Strikes
    m_PitchSequence = m.PitchSequence
    m_ScoreAway = m.ScoreAway
    m_ScoreHome = m.ScoreHome
    Me.CurrentPitcher.Pitchcounter = m.Pitchcounter
    Me.CurrentPitcher.OutsPitched = m.OutsPitched
End Sub

' ----------------------------------------------------------------
' Additional Memento support — snapshot and restore Lineups
' Called before and after every substitution
' ----------------------------------------------------------------
Public Function CaptureLineupMemento() As clsLineupMemento
    Dim m As clsLineupMemento
    Set m = New clsLineupMemento
    
    m.CaptureFrom Me
    Set CaptureLineupMemento = m
End Function

Public Sub RestoreLineupMemento(ByVal m As clsLineupMemento)
    Set m_LineUpAway = m.LineUpAway
    Set m_LineUpHome = m.LineUpHome
End Sub

' Substitutes a bench player into an active batting spot.
' Both players remain in the same lineup Collection — only their
' Spot/Position/IsPitcher fields change. team: "Away" or "Home"
Public Sub SubstitutePlayer(ByVal team As String, ByVal Spot As Long, _
                            ByVal IncomingPlayerName As String, ByVal newPosition As String)
    Dim lineup As Collection
    Set lineup = IIf(team = "Away", m_LineUpAway, m_LineUpHome)
    
    Dim outgoing As clsPlayer
    Dim incoming As clsPlayer
    
    Dim p As clsPlayer
    For Each p In lineup
        If p.Spot = Spot And p.Position <> "" Then Set outgoing = p ' identify outgoing player by spot
        If p.Name = IncomingPlayerName And p.Spot = 0 Then Set incoming = p ' identify incoming player by name
    Next p
    
    If outgoing Is Nothing Or incoming Is Nothing Then
        Err.Raise vbObjectError + 1, "SubstitutePlayer", "Outgoing or incoming player not found in lineup."
    End If
    
    ' Incoming player takes the vacated spot and position
    incoming.Spot = Spot
    incoming.Position = newPosition
    incoming.IsPitcher = (UCase(newPosition) = "P")
    
    ' Outgoing player returns to bench state
    outgoing.Spot = 0
    outgoing.Position = ""
    outgoing.IsPitcher = False
End Sub


' ----------------------------------------------------------------
' Class Initialisation
' ----------------------------------------------------------------

Private Sub Class_Initialize()
    m_Inning = 1
    m_Half = "Top"
    m_BatterIndexAway = 1
    m_BatterIndexHome = 1
    Set m_LineUpAway = New Collection
    Set m_LineUpHome = New Collection
    ResetCount ' Resets Pitch Count
End Sub


