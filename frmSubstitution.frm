VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSubstitution 
   Caption         =   "Substitute Player"
   ClientHeight    =   3270
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6360
   OleObjectBlob   =   "frmSubstitution.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmSubstitution"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

Private m_GameRef As clsBaseballGame
Private m_Confirmed As Boolean
Private m_OutgoingName As String
Private m_IncomingName As String
Private m_SelectedTeam As String
Private m_SelectedSpot As Long
Private m_IncomingPosition As String

''
' Entry point called by frmRecordPlay.
' Returns True if a substitution was confirmed, False if cancelled.
''
Public Function ExecuteSubstitution(ByRef GameEngine As clsBaseballGame) As Boolean
    Set m_GameRef = GameEngine
    m_Confirmed = False
    
    Me.cboTeam.Clear
    Me.cboTeam.AddItem "Away"
    Me.cboTeam.AddItem "Home"
    Me.cboTeam.ListIndex = 0
    
    PopulateOutgoingPlayerList
    PopulateIncomingPlayerList
    PopulatePositionList
    
    Me.Show vbModal
    
    ExecuteSubstitution = m_Confirmed
End Function

Private Sub cboTeam_Change()
    PopulateOutgoingPlayerList
    PopulateIncomingPlayerList
End Sub

' Lists active players (on field) for the selected team —
' these are the candidates to be substituted OUT.
Private Sub PopulateOutgoingPlayerList()
    Me.cboOutgoingPlayer.Clear
    
    Dim lineup As Collection
    Set lineup = GetSelectedLineup()
    
    Dim p As clsPlayer
    For Each p In lineup
        If p.Spot > 0 And p.Position <> "" Then
            Me.cboOutgoingPlayer.AddItem p.Spot & " | " & p.Position & " | " & p.Name
        End If
    Next p
    
    If Me.cboOutgoingPlayer.ListCount > 0 Then Me.cboOutgoingPlayer.ListIndex = 0
End Sub

' Lists bench players (Spot = 0) for the selected team —
' these are the candidates to be substituted IN.
Private Sub PopulateIncomingPlayerList()
    Me.cboIncomingPlayer.Clear
    
    Dim lineup As Collection
    Set lineup = GetSelectedLineup()
    
    Dim p As clsPlayer
    For Each p In lineup
        If p.Spot = 0 Then
            Me.cboIncomingPlayer.AddItem p.Jersey & " | " & p.Name
        End If
    Next p
    
    If Me.cboIncomingPlayer.ListCount > 0 Then Me.cboIncomingPlayer.ListIndex = 0
End Sub

Private Sub PopulatePositionList()
    With Me.cboIncomingPosition
        .Clear
        .AddItem "P": .AddItem "C": .AddItem "1B": .AddItem "2B": .AddItem "3B"
        .AddItem "SS": .AddItem "LF": .AddItem "CF": .AddItem "RF"
        .AddItem "DH": .AddItem "DP": .AddItem "PH": .AddItem "PR"
    End With
End Sub

Private Function GetSelectedLineup() As Collection
    Set GetSelectedLineup = IIf(Me.cboTeam.Value = "Away", m_GameRef.LineUpAway, m_GameRef.LineUpHome)
End Function

Private Sub cmdConfirm_Click()
    If Me.cboOutgoingPlayer.ListIndex = -1 Then
        MsgBox "Please select a player to substitute out.", vbExclamation
        Exit Sub
    End If
    
    If Me.cboIncomingPlayer.ListIndex = -1 Then
        MsgBox "No bench players available for this team.", vbExclamation
        Exit Sub
    End If
    
    If Me.cboIncomingPosition.ListIndex = -1 Then
        MsgBox "Please select the incoming player's position.", vbExclamation
        Exit Sub
    End If
    
    ' Parse spot + names back out of the display strings
    Dim outgoingParts() As String: outgoingParts = Split(Me.cboOutgoingPlayer.Value, " | ")
    Dim incomingParts() As String: incomingParts = Split(Me.cboIncomingPlayer.Value, " | ")
    
    m_SelectedTeam = Me.cboTeam.Value
    m_SelectedSpot = CLng(outgoingParts(0))
    m_OutgoingName = outgoingParts(2)
    m_IncomingName = incomingParts(1)
    m_IncomingPosition = Me.cboIncomingPosition.Value
    
    m_GameRef.SubstitutePlayer m_SelectedTeam, m_SelectedSpot, m_IncomingName, Me.cboIncomingPosition.Value
    
    m_Confirmed = True
    Me.Hide
End Sub

' ----------------------------------------------------------------
' CANCEL FORM
' ----------------------------------------------------------------
Private Sub cmdCancel_Click()
    m_Confirmed = False
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' Treat X button exactly like clicking Cancel — hide rather than destroy,
    ' so the caller receives a clean Nothing collection and can abort safely.
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Confirmed = False
        Me.Hide
    End If
End Sub

''
' Read-only accessors so frmRecordPlay can build the log text
' and capture base-runner renames after ExecuteSubstitution returns True.
''
Public Property Get SelectedTeam() As String
    SelectedTeam = m_SelectedTeam
End Property

Public Property Get SelectedSpot() As Long
    SelectedSpot = m_SelectedSpot
End Property

Public Property Get OutgoingPlayerName() As String
    OutgoingPlayerName = m_OutgoingName
End Property

Public Property Get IncomingPlayerName() As String
    IncomingPlayerName = m_IncomingName
End Property

Public Property Get IncomingPosition() As String
    IncomingPosition = m_IncomingPosition
End Property
