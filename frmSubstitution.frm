VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSubstitution 
   Caption         =   "Substitute Player"
   ClientHeight    =   4860
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7725
   OleObjectBlob   =   "frmSubstitution.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSubstitution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_GameRef        As clsBaseballGame
Private m_Confirmed      As Boolean
Private m_SelectedTeam   As String
Private m_PendingChanges As Collection ' clsLineupAssignment objects staged so far

Public Function ExecuteLineupChanges(ByRef GameEngine As clsBaseballGame) As Boolean
    Set m_GameRef = GameEngine
    Set m_PendingChanges = New Collection
    m_Confirmed = False
    
    Me.cboTeam.Clear
    Me.cboTeam.AddItem "Away"
    Me.cboTeam.AddItem "Home"
    Me.cboTeam.ListIndex = 0
    
    PopulatePlayerList
    PopulatePositionList
    PopulateSpotList
    Me.lstPendingChanges.Clear
    
    Me.Show vbModal
    
    ExecuteLineupChanges = m_Confirmed
End Function

Private Sub cboTeam_Change()
    PopulatePlayerList
End Sub

' Lists every player on the selected team — bench and active alike —
' so any Spot/Position combination can be assigned to anyone.
Private Sub PopulatePlayerList()
    Me.cboPlayer.Clear
    Dim lineup As Collection
    Set lineup = GetSelectedLineup()
    
    Dim p As clsPlayer
    For Each p In lineup
        Dim posDisplay As String
        posDisplay = IIf(p.Position = "", "-", p.Position)
        Me.cboPlayer.AddItem p.Spot & " | " & posDisplay & " | " & p.Jersey & " | " & p.Name
    Next p
    
    If Me.cboPlayer.ListCount > 0 Then Me.cboPlayer.ListIndex = 0
End Sub

' Pre-fills the Spot/Position controls with the selected player's
' CURRENT values — the user only changes what's actually different.
Private Sub cboPlayer_Change()
    If Me.cboPlayer.ListIndex = -1 Then Exit Sub
    
    Dim parts() As String
    parts = Split(Me.cboPlayer.Value, " | ")
    
    Me.cboNewSpot.Value = parts(0)
    Me.cboNewPosition.Value = IIf(parts(1) = "-", "", parts(1))
End Sub

Private Sub PopulatePositionList()
    With Me.cboNewPosition
        .Clear
        .AddItem "" ' bench / no position
        .AddItem "P": .AddItem "C": .AddItem "1B": .AddItem "2B": .AddItem "3B"
        .AddItem "SS": .AddItem "LF": .AddItem "CF": .AddItem "RF"
        .AddItem "DH": .AddItem "DP": .AddItem "PH": .AddItem "PR"
    End With
End Sub

Private Sub PopulateSpotList()
    With Me.cboNewSpot
        .Clear
        Dim i As Long
        For i = 0 To 10
            .AddItem i
        Next i
    End With
End Sub

Private Function GetSelectedLineup() As Collection
    Set GetSelectedLineup = IIf(Me.cboTeam.Value = "Away", m_GameRef.LineUpAway, m_GameRef.LineUpHome)
End Function

' Stages the current Player/Spot/Position selection as one pending
' change and resets the inputs for the next one. Nothing is applied
' to the game engine yet — that only happens on Confirm.
Private Sub cmdContinue_Click()
    If Me.cboPlayer.ListIndex = -1 Then
        MsgBox "Please select a player.", vbExclamation
        Exit Sub
    End If
    
    Dim parts() As String
    parts = Split(Me.cboPlayer.Value, " | ")
    Dim playerName As String: playerName = parts(3)
    
    Dim lineup As Collection
    Set lineup = GetSelectedLineup()
    
    Dim currentPlayer As clsPlayer
    Dim p As clsPlayer
    For Each p In lineup
        If p.Name = playerName Then Set currentPlayer = p
    Next p
    
    Dim chg As clsLineupAssignment
    Set chg = New clsLineupAssignment
    chg.playerName = playerName
    chg.OldSpot = currentPlayer.Spot
    chg.OldPosition = currentPlayer.Position
    chg.NewSpot = CLng(Me.cboNewSpot.Value)
    chg.NewPosition = Me.cboNewPosition.Value
    
    ' If this player already has a staged change, replace it rather
    ' than staging two conflicting entries for the same person.
    Dim i As Long
    For i = m_PendingChanges.Count To 1 Step -1
        If m_PendingChanges(i).playerName = playerName Then
            m_PendingChanges.Remove i
        End If
    Next i
    
    m_PendingChanges.Add chg
    RefreshPendingList
End Sub

Private Sub RefreshPendingList()
    Me.lstPendingChanges.Clear
    Dim chg As clsLineupAssignment
    For Each chg In m_PendingChanges
        Me.lstPendingChanges.AddItem BuildAssignmentDescription(chg)
    Next chg
End Sub

Private Function BuildAssignmentDescription(ByVal chg As clsLineupAssignment) As String
    Dim oldPosDisplay As String: oldPosDisplay = IIf(chg.OldPosition = "", "Bench", chg.OldPosition)
    Dim newPosDisplay As String: newPosDisplay = IIf(chg.NewPosition = "", "Bench", chg.NewPosition)
    
    BuildAssignmentDescription = chg.playerName & ":  Spot " & chg.OldSpot & " -> " & chg.NewSpot & _
                                  ",  " & oldPosDisplay & " -> " & newPosDisplay
End Function

' Removes the highlighted staged change, in case of a mistake —
' doesn't require cancelling the whole substitution.
Private Sub cmdRemoveSelected_Click()
    If Me.lstPendingChanges.ListIndex = -1 Then Exit Sub
    m_PendingChanges.Remove Me.lstPendingChanges.ListIndex + 1
    RefreshPendingList
End Sub

Private Sub cmdConfirm_Click()
    If m_PendingChanges.Count = 0 Then
        MsgBox "No changes staged. Add at least one, or Cancel.", vbExclamation
        Exit Sub
    End If
    
    m_SelectedTeam = Me.cboTeam.Value
    m_GameRef.ApplyLineupChanges m_SelectedTeam, m_PendingChanges
    
    m_Confirmed = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_Confirmed = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Confirmed = False
        Me.Hide
    End If
End Sub

Public Property Get SelectedTeam() As String
    SelectedTeam = m_SelectedTeam
End Property

' Builds the full log text listing every staged change with proper substitution phrasing.
' Used only for the exported log text — the pending-changes listbox
' still shows the literal Spot/Position values via BuildAssignmentDescription.
Public Property Get ChangesDescription() As String
    Dim result As String: result = ""
    Dim chg As clsLineupAssignment
    For Each chg In m_PendingChanges
        If Not IsBeingCreditedElsewhere(chg.playerName) Then
            result = result & "; " & BuildLogFragment(chg)
        End If
    Next chg
    If Left(result, 2) = "; " Then result = Mid(result, 3)
    ChangesDescription = result
End Property

Private Function BuildLogFragment(ByVal chg As clsLineupAssignment) As String
    Dim comingFromBench As Boolean
    comingFromBench = (chg.OldSpot = 0)
    
    If comingFromBench Then
        Dim replacedName As String
        replacedName = FindVacatingPlayer(chg.NewSpot, chg.playerName)
        
        If replacedName <> "" Then
            Select Case chg.NewPosition
                Case "PH": BuildLogFragment = chg.playerName & " pinch hit for " & replacedName
                Case "PR": BuildLogFragment = chg.playerName & " pinch ran for " & replacedName
                Case Else: BuildLogFragment = chg.playerName & " to " & chg.NewPosition & " for " & replacedName
            End Select
        Else
            BuildLogFragment = chg.playerName & " to " & chg.NewPosition
        End If
    Else
        BuildLogFragment = chg.playerName & " to " & IIf(chg.NewPosition = "", "Bench", chg.NewPosition)
    End If
End Function

' Finds who vacated targetSpot in this batch — matched purely by Spot,
' never by Position, since a baserunner being pinch-run for has no
' defensive Position at all while still occupying (vacating) a batting Spot.
Private Function FindVacatingPlayer(ByVal targetSpot As Long, ByVal excludeName As String) As String
    Dim chg As clsLineupAssignment
    For Each chg In m_PendingChanges
        If chg.playerName <> excludeName Then
            If chg.OldSpot = targetSpot And chg.NewSpot <> chg.OldSpot Then
                FindVacatingPlayer = chg.playerName
                Exit Function
            End If
        End If
    Next chg
    FindVacatingPlayer = ""
End Function

' True if some OTHER change in this batch already names this player
' as the one being replaced (i.e. FindVacatingPlayer resolves to her).
' If so, her own fragment would be redundant and is skipped.
Private Function IsBeingCreditedElsewhere(ByVal playerName As String) As Boolean
    Dim chg As clsLineupAssignment
    For Each chg In m_PendingChanges
        If chg.playerName <> playerName And chg.OldSpot = 0 Then
            If FindVacatingPlayer(chg.NewSpot, chg.playerName) = playerName Then
                IsBeingCreditedElsewhere = True
                Exit Function
            End If
        End If
    Next chg
    IsBeingCreditedElsewhere = False
End Function
