VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmResolveRunners 
   Caption         =   "Resolve Runners"
   ClientHeight    =   6480
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7620
   OleObjectBlob   =   "frmResolveRunners.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmResolveRunners"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private m_GameRef        As clsBaseballGame
Private m_Outcomes       As Collection   ' cumulative across all segments
Private m_OutsAdded      As Long
Private m_ErrorCount     As Long
Private m_ContinueLoop   As Boolean
Private m_PendingSegmentIsError As Boolean ' True if the segment now being finalized resulted from an error declared on the previous click
Private m_Cancelled      As Boolean

''
' Entry point. Loops through one or more "segments" — the initial play,
' then optionally one or more error segments — until the user confirms
' the final result or cancels. Each segment's decisions are applied to
' GameEngine immediately so the next segment sees the updated base state.
''
Public Function GetSeparatedOutcomes(ByRef GameEngine As clsBaseballGame, _
                                     ByRef AdditionalOuts As Long, _
                                     ByRef ErrorCount As Long) As Collection
    If GameEngine.BaseState = "___" Then
        Set GetSeparatedOutcomes = Nothing
        Exit Function
    End If
    
    ' Reset block
    Set m_GameRef = GameEngine
    Set m_Outcomes = New Collection
    m_OutsAdded = 0
    m_ErrorCount = 0
    m_PendingSegmentIsError = False
    m_Cancelled = False

    Do
        RefreshForm
        Me.Show vbModal
        ' Execution resumes here after Me.Hide, called by cmdConfirm,
        ' cmdErrorContinue, cmdCancel, or UserForm_QueryClose.

        If m_Cancelled Then
            Set GetSeparatedOutcomes = Nothing
            Exit Function
        End If
    Loop While m_ContinueLoop

    AdditionalOuts = m_OutsAdded
    ErrorCount = m_ErrorCount
    Set GetSeparatedOutcomes = m_Outcomes
End Function

' Read-only accessor for the caller to check
Public Property Get WasCancelled() As Boolean
    WasCancelled = m_Cancelled
End Property

' Rebuilds the dropdowns from the CURRENT live base state (m_GameRef),
' which changes between segments as ApplySegmentToGame mutates it.
Private Sub RefreshForm()
    SetupDropdown Me.cboRunner1B, Me.lblRunner1B, m_GameRef.Runner1B, "1B"
    SetupDropdown Me.cboRunner2B, Me.lblRunner2B, m_GameRef.Runner2B, "2B"
    SetupDropdown Me.cboRunner3B, Me.lblRunner3B, m_GameRef.Runner3B, "3B"

    Me.Caption = IIf(m_ErrorCount > 0, "Resolve Runners — Error Segment " & (m_ErrorCount + 1), "Resolve Runners")
End Sub

Private Sub SetupDropdown(ByRef cbo As Control, ByRef lbl As Control, ByVal runnerName As String, ByVal baseCode As String)
    cbo.Clear
    If runnerName = "" Then
        cbo.Visible = False
        lbl.Visible = False
        Exit Sub
    End If

    lbl.Visible = True
    lbl.Caption = baseCode & " Runner: " & runnerName
    cbo.Visible = True

    cbo.AddItem "Stays/Holds"
    If baseCode = "1B" Then cbo.AddItem "Advances to 2B"
    If baseCode = "1B" Or baseCode = "2B" Then cbo.AddItem "Advances to 3B"
    cbo.AddItem "Scores"
    cbo.AddItem "Out on Play"

    cbo.ListIndex = 0
End Sub

' ----------------------------------------------------------------
' BUTTON HANDLERS
' ----------------------------------------------------------------
Private Sub cmdSaveRunners_Click()
    Dim segmentOutcomes As Collection
    Set segmentOutcomes = BuildSegmentOutcomes()
    If Not ValidateSegment(segmentOutcomes, WarnIfNothingMoved:=Not m_PendingSegmentIsError) Then Exit Sub

    ApplySegmentToGame segmentOutcomes, m_PendingSegmentIsError
    m_ContinueLoop = False
    Me.Hide
End Sub

Private Sub cmdError_Click()
    Dim segmentOutcomes As Collection
    Set segmentOutcomes = BuildSegmentOutcomes()
    If Not ValidateSegment(segmentOutcomes, WarnIfNothingMoved:=False) Then Exit Sub

    ' This segment (the one just decided) is NOT itself the error —
    ' it's whatever was flagged from the previous click (False on the first pass).
    ApplySegmentToGame segmentOutcomes, m_PendingSegmentIsError

    ' One distinct misplay is being declared now — count it once here,
    ' regardless of how many runners move because of it in the NEXT segment.
    m_ErrorCount = m_ErrorCount + 1

    ' The segment that comes after this click is the one caused by the error.
    m_PendingSegmentIsError = True

    m_ContinueLoop = True
    Me.Hide
End Sub

Private Sub cmdCancel_Click()
    m_Cancelled = True
    Me.Hide
End Sub

Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        m_Cancelled = True
        Me.Hide
    End If
End Sub

' ----------------------------------------------------------------
' SEGMENT PROCESSING
' ----------------------------------------------------------------
Private Sub FinalizeSegment(ByVal IsError As Boolean, ByVal KeepLooping As Boolean)
    Dim segmentOutcomes As Collection
    Set segmentOutcomes = BuildSegmentOutcomes()

    If Not ValidateSegment(segmentOutcomes, WarnIfNothingMoved:=Not IsError) Then Exit Sub

    ApplySegmentToGame segmentOutcomes, IsError

    If IsError Then m_ErrorCount = m_ErrorCount + 1
    m_ContinueLoop = KeepLooping
    Me.Hide
End Sub

' Reads the three dropdowns and builds one clsRunnerOutcome per visible
' (occupied) base. Does NOT touch m_GameRef — that happens in ApplySegmentToGame.
Private Function BuildSegmentOutcomes() As Collection
    Dim result As New Collection
    AddOutcomeIfVisible result, Me.cboRunner1B, m_GameRef.Runner1B, "1B"
    AddOutcomeIfVisible result, Me.cboRunner2B, m_GameRef.Runner2B, "2B"
    AddOutcomeIfVisible result, Me.cboRunner3B, m_GameRef.Runner3B, "3B"
    Set BuildSegmentOutcomes = result
End Function

Private Sub AddOutcomeIfVisible(ByVal target As Collection, ByRef cbo As Control, _
                                 ByVal runnerName As String, ByVal originalBase As String)
    If Not cbo.Visible Then Exit Sub

    Dim decision As clsRunnerOutcome
    Set decision = New clsRunnerOutcome
    decision.BaseSource = originalBase
    decision.PlayerName = runnerName

    Select Case cbo.Value
        Case "Stays/Holds":    decision.NewDestination = DEST_HOLDS
        Case "Advances to 2B": decision.NewDestination = DEST_SECOND
        Case "Advances to 3B": decision.NewDestination = DEST_THIRD
        Case "Scores":         decision.NewDestination = DEST_SCORES
        Case "Out on Play":    decision.NewDestination = DEST_OUT
    End Select

    target.Add decision
End Sub

' Applies a validated segment's outcomes to the live game state.
' Processed in DESCENDING base order (3B, 2B, 1B) so a runner moving
' into a base a moment before another runner vacates it is never
' overwritten — e.g. 1B?2B and 2B?3B in the same segment.
Private Sub ApplySegmentToGame(ByVal segmentOutcomes As Collection, ByVal IsError As Boolean)
    Dim source As Variant
    For Each source In Array("3B", "2B", "1B")
        Dim outcome As clsRunnerOutcome
        For Each outcome In segmentOutcomes
            If outcome.BaseSource = source Then
                outcome.AdvancedOnError = IsError
                ApplyOutcomeToField outcome
                m_Outcomes.Add outcome
                If outcome.NewDestination = DEST_OUT Then m_OutsAdded = m_OutsAdded + 1
            End If
        Next outcome
    Next source
End Sub

Private Sub ApplyOutcomeToField(ByVal outcome As clsRunnerOutcome)
    ' Vacate the source base unless the runner is holding there
    If outcome.NewDestination <> DEST_HOLDS Then
        Select Case outcome.BaseSource
            Case "1B": m_GameRef.Runner1B = ""
            Case "2B": m_GameRef.Runner2B = ""
            Case "3B": m_GameRef.Runner3B = ""
        End Select
    End If

    Select Case outcome.NewDestination
        Case DEST_SECOND: m_GameRef.Runner2B = outcome.PlayerName
        Case DEST_THIRD:  m_GameRef.Runner3B = outcome.PlayerName
        Case DEST_SCORES: m_GameRef.ScoreRuns 1
        ' DEST_HOLDS: no-op, base already correctly occupied
        ' DEST_OUT:   no-op beyond the vacate above
    End Select
End Sub

' ----------------------------------------------------------------
' VALIDATION — scoped to the CURRENT segment only
' ----------------------------------------------------------------
Private Function ValidateSegment(ByVal segmentOutcomes As Collection, ByVal WarnIfNothingMoved As Boolean) As Boolean
    Dim baseOccupied(1 To 3) As Boolean
    ' Pre-occupy the base the batter will eventually land on (TargetBase,
    ' stored in EventToHappen) so no runner in any segment can land there.
    If m_GameRef.EventToHappen = "1B" Then baseOccupied(1) = True
    If m_GameRef.EventToHappen = "2B" Then baseOccupied(2) = True

    Dim CheckSum As Long: CheckSum = 0
    Dim outcome As clsRunnerOutcome

    For Each outcome In segmentOutcomes
        CheckSum = CheckSum + outcome.NewDestination

        Dim finalBase As Long
        If outcome.NewDestination = DEST_HOLDS Then
            Select Case outcome.BaseSource
                Case "1B": finalBase = 1
                Case "2B": finalBase = 2
                Case "3B": finalBase = 3
            End Select
        Else
            finalBase = outcome.NewDestination
        End If

        ' --- FORCE PLAY RULES (only meaningful on the segment where the base is occupied) ---
        If outcome.BaseSource = "1B" Then
            If outcome.NewDestination = DEST_HOLDS And m_GameRef.EventToHappen = "1B" Then
                MsgBox "Rule Violation! The runner on 1st base is forced to advance on a Single.", vbExclamation
                ValidateSegment = False
                Exit Function
            End If
            If m_GameRef.EventToHappen = "2B" Then
                If outcome.NewDestination = DEST_HOLDS Or outcome.NewDestination = DEST_SECOND Then
                    MsgBox "Rule Violation! The runner on 1st base must advance to at least 3rd base on a Double.", vbExclamation
                    ValidateSegment = False
                    Exit Function
                End If
            End If
        End If

        If outcome.BaseSource = "2B" And outcome.NewDestination = DEST_HOLDS And Me.cboRunner1B.Visible Then
            If Me.cboRunner1B.Value = "Advances to 2B" Then
                MsgBox "Rule Violation! The runner on 2nd base is forced to move by the runner advancing behind them.", vbExclamation
                ValidateSegment = False
                Exit Function
            End If
        End If

        ' --- BASE COLLISION ---
        If finalBase >= 1 And finalBase <= 3 Then
            If baseOccupied(finalBase) Then
                MsgBox "Base Collision! Two runners cannot occupy base " & finalBase & " simultaneously.", vbExclamation
                ValidateSegment = False
                Exit Function
            End If
            baseOccupied(finalBase) = True
        End If
    Next outcome

    If WarnIfNothingMoved And CheckSum = 0 Then
        Dim userInput As Integer
        userInput = MsgBox("All runners stay on their bases! Proceed anyway?", vbQuestion + vbOKCancel)
        If userInput <> vbOK Then
            ValidateSegment = False
            Exit Function
        End If
    End If

    ValidateSegment = True
End Function

