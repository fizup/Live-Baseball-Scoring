VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmRecordPlay 
   Caption         =   "Live Game Input"
   ClientHeight    =   7485
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   8040
   OleObjectBlob   =   "frmRecordPlay.frx":0000
   StartUpPosition =   1  'Fenstermitte
End
Attribute VB_Name = "frmRecordPlay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit

' ----------------------------------------------------------------
' Module-level state
' ----------------------------------------------------------------
Private m_Game      As clsBaseballGame
Private m_Logger    As clsGameHistoryLogger
Private m_AllEvents As Collection
Private m_UndoStack As Collection

' ----------------------------------------------------------------
' INITIALISATION
' ----------------------------------------------------------------
Public Sub InitializeForm(ByVal TargetGame As clsBaseballGame, ByVal TargetLogger As clsGameHistoryLogger)
    Set m_Game = TargetGame
    Set m_Logger = TargetLogger
    Set m_UndoStack = New Collection

    Set m_AllEvents = GetEvents()

    With Me.cboPlayCategory
        .Clear
        .AddItem "GetOnBase"
        .AddItem "GetOut"
        .AddItem "AdvanceBase"
        .ListIndex = 0
    End With

    SetInterfaceStage InPlayMode:=False
    UpdateFormDisplay
End Sub

' ----------------------------------------------------------------
' INTERFACE STAGE CONTROL
' ----------------------------------------------------------------
Private Sub SetInterfaceStage(ByVal InPlayMode As Boolean)
    Me.cmdBall.Enabled = Not InPlayMode
    Me.cmdCalledStrike.Enabled = Not InPlayMode
    Me.cmdSwingMiss.Enabled = Not InPlayMode
    Me.cmdFoul.Enabled = Not InPlayMode
    Me.cmdInPlay.Enabled = Not InPlayMode

    Me.cboPlayCategory.Enabled = InPlayMode
    Me.lstDynamicPlays.Enabled = InPlayMode
    Me.cmdSubmit.Enabled = InPlayMode

    If InPlayMode Then
        PopulateDynamicPlaysList
        Me.lstDynamicPlays.SetFocus
    Else
        Me.lstDynamicPlays.Clear
    End If
End Sub

Private Sub PopulateDynamicPlaysList()
    Me.lstDynamicPlays.Clear
    Dim e As clsEvent
    For Each e In m_AllEvents
        If e.Section = Me.cboPlayCategory.Value Then Me.lstDynamicPlays.AddItem e.Name
    Next e
End Sub

Private Sub cboPlayCategory_Change()
    If Me.cboPlayCategory.Enabled Then PopulateDynamicPlaysList
End Sub

' ----------------------------------------------------------------
' STAGE 1: PITCH BUTTONS
' ----------------------------------------------------------------
Private Sub cmdBall_Click()
    ' Ball only writes a log entry when it completes a walk (4th ball).
    ' Capture before incrementing so the check sees the pre-increment value.
    PushUndo LogEntryWillBeWritten:=(m_Game.Balls = 3)
    
    m_Game.CurrentPitcher.Pitchcounter = m_Game.CurrentPitcher.Pitchcounter + 1

    m_Game.Balls = m_Game.Balls + 1
    m_Game.AppendPitch "B"

    If m_Game.Balls >= 4 Then
        Dim snapshot As clsPlayByPlayEvent
        Set snapshot = m_Logger.TakeSnapshot(m_Game)
        ExecuteAndLogForcedWalks snapshot, m_AllEvents("Base On Balls")
    End If

    UpdateFormDisplay
End Sub

Private Sub cmdCalledStrike_Click()
    RecordStrike "C"
End Sub

Private Sub cmdSwingMiss_Click()
    RecordStrike "S"
End Sub

Private Sub cmdFoul_Click()
    PushUndo LogEntryWillBeWritten:=False ' Fouls never produce a log entry
    m_Game.CurrentPitcher.Pitchcounter = m_Game.CurrentPitcher.Pitchcounter + 1
    m_Game.AppendPitch "F"
    If m_Game.Strikes < 2 Then m_Game.Strikes = m_Game.Strikes + 1
    UpdateFormDisplay
End Sub

Private Sub RecordStrike(ByVal Code As String)
    ' Strike only writes a log entry when it completes a strikeout (3rd strike).
    ' Capture before incrementing so the check sees the pre-increment value.
    PushUndo LogEntryWillBeWritten:=(m_Game.Strikes = 2)
    
    m_Game.CurrentPitcher.Pitchcounter = m_Game.CurrentPitcher.Pitchcounter + 1
    
    m_Game.Strikes = m_Game.Strikes + 1
    m_Game.AppendPitch Code

    If m_Game.Strikes >= 3 Then
        Dim StrikeOutType As String
        StrikeOutType = IIf(Code = "C", "Strikeout Looking", "Strikeout Swinging")

        Dim snapshot As clsPlayByPlayEvent
        Set snapshot = m_Logger.TakeSnapshot(m_Game)
        loadLogger snapshot, m_AllEvents(StrikeOutType)

        m_Game.AdvanceBatter ' AdvanceBatter before AddOuts ï¿½ if this is the 3rd out,
        m_Game.ResetCount   ' ChangeSides flips m_Half inside AddOuts, and AdvanceBatter
        m_Game.AddOuts 1    ' would then increment the wrong team's index.
    End If

    UpdateFormDisplay
End Sub

Private Sub cmdInPlay_Click()
    SetInterfaceStage InPlayMode:=True
End Sub

' ----------------------------------------------------------------
' STAGE 2: PLAY SUBMISSION
' ----------------------------------------------------------------
Private Sub cmdSubmit_Click()
    If Me.lstDynamicPlays.ListIndex = -1 Then
        MsgBox "Please choose a specific play outcome from the dynamic list.", vbExclamation
        Exit Sub
    End If

    PushUndo ' cmdSubmit always produces a log entry

    Dim choice As clsEvent
    Set choice = m_AllEvents(Me.lstDynamicPlays.Value)

    ' EventToHappen holds TargetBase, not Code — this is what frmResolveRunners
    ' uses to pre-occupy the batter's eventual base during collision validation.
    ' Blank for 3B/HR/forced-advance events, which never touch frmResolveRunners
    ' for batter placement.
    m_Game.EventToHappen = choice.targetBase

    If choice.BallIsHit Then
        m_Game.AppendPitch "X"
        m_Game.CurrentPitcher.Pitchcounter = m_Game.CurrentPitcher.Pitchcounter + 1
    End If

    Dim snapshot As clsPlayByPlayEvent
    Set snapshot = m_Logger.TakeSnapshot(m_Game)

    Dim isBatterTurnOver As Boolean: isBatterTurnOver = True
    Dim playCompleted As Boolean

    Select Case choice.Section
        Case "GetOnBase":   playCompleted = HandleGetOnBase(choice, snapshot, isBatterTurnOver)
        Case "GetOut":      playCompleted = HandleGetOut(choice, snapshot)
                            isBatterTurnOver = True
        Case "AdvanceBase": playCompleted = HandleAdvanceBase(choice, snapshot)
                            isBatterTurnOver = False
    End Select

    If Not playCompleted Then
        ' User cancelled a runner-resolution form mid-play — roll back
        ' everything this action changed, including any partial segments.
        m_Game.RestoreMemento m_UndoStack(m_UndoStack.Count)
        m_UndoStack.Remove m_UndoStack.Count
        SetInterfaceStage InPlayMode:=False
        UpdateFormDisplay
        Exit Sub
    End If

    If isBatterTurnOver Then m_Game.ResetCount
    m_Game.EventToHappen = ""
    SetInterfaceStage InPlayMode:=False
    UpdateFormDisplay
End Sub

' ----------------------------------------------------------------
' PLAY SECTION HANDLERS  (called only from cmdSubmit_Click)
' Each returns False if the user cancelled a resolution form mid-play.
' ----------------------------------------------------------------
Private Function HandleGetOnBase(ByVal choice As clsEvent, ByVal snapshot As clsPlayByPlayEvent, ByRef isBatterTurnOver As Boolean) As Boolean
    Dim batterName As String: batterName = m_Game.CurrentBatter.Name

    Select Case choice.Code
        Case "3B"
            Dim tripleRuns As Long
            Dim tripleText As String
            ScoreAllRunners m_Game.Runner3B, m_Game.Runner2B, m_Game.Runner1B, tripleRuns, tripleText
            m_Game.Runner1B = ""
            m_Game.Runner2B = ""
            m_Game.Runner3B = batterName
            If tripleRuns > 0 Then m_Game.ScoreRuns tripleRuns
            loadLogger snapshot, choice, tripleText
            m_Game.AdvanceBatter

        Case "HR"
            Dim hrRuns As Long
            Dim hrText As String
            ScoreAllRunners m_Game.Runner3B, m_Game.Runner2B, m_Game.Runner1B, hrRuns, hrText
            hrRuns = hrRuns + 1
            m_Game.Runner1B = ""
            m_Game.Runner2B = ""
            m_Game.Runner3B = ""
            m_Game.ScoreRuns hrRuns
            loadLogger snapshot, choice, hrText
            m_Game.AdvanceBatter

        Case "HBP", "IBB", "CI"
            ExecuteAndLogForcedWalks snapshot, choice
            isBatterTurnOver = False ' ExecuteAndLogForcedWalks handles AdvanceBatter/ResetCount internally

        Case Else
            ' 1B, 2B, FC, E#, ROE, GRD, K E2 — and any future GetOnBase event.
            ' Resolve existing runners FIRST (form pre-occupies TargetBase via
            ' EventToHappen so no runner can land where the batter is headed),
            ' THEN place the batter — base is guaranteed free at that point.
            Dim getonbaseForm As New frmResolveRunners
            Dim baserunningOuts As Long
            Dim ErrorCount As Long
            Dim decisions As Collection
            Set decisions = getonbaseForm.GetSeparatedOutcomes(m_Game, baserunningOuts, ErrorCount)

            If getonbaseForm.WasCancelled Then
                HandleGetOnBase = False
                Exit Function
            End If

            PlaceBatterOnBase batterName, choice.targetBase

            Dim runnerText As String
            runnerText = BuildRunnerText(decisions)

            loadLogger snapshot, choice, runnerText, ErrorCount
            m_Game.AdvanceBatter
            If baserunningOuts > 0 Then m_Game.AddOuts baserunningOuts
    End Select

    HandleGetOnBase = True
End Function

Private Function HandleGetOut(ByVal choice As clsEvent, ByVal snapshot As clsPlayByPlayEvent) As Boolean
    Dim getoutForm As New frmResolveRunners
    Dim extraOuts As Long
    Dim ErrorCount As Long
    Dim outDecisions As Collection
    Set outDecisions = getoutForm.GetSeparatedOutcomes(m_Game, extraOuts, ErrorCount)

    If getoutForm.WasCancelled Then
        HandleGetOut = False
        Exit Function
    End If

    Dim outText As String
    outText = BuildRunnerText(outDecisions)

    loadLogger snapshot, choice, outText, ErrorCount
    m_Game.AdvanceBatter          ' AdvanceBatter before AddOuts — same reason as RecordStrike
    m_Game.AddOuts (1 + extraOuts)

    HandleGetOut = True
End Function

Private Function HandleAdvanceBase(ByVal choice As clsEvent, ByVal snapshot As clsPlayByPlayEvent) As Boolean
    Dim advanceForm As New frmResolveRunners
    Dim baseOuts As Long
    Dim ErrorCount As Long
    Dim advDecisions As Collection
    Set advDecisions = advanceForm.GetSeparatedOutcomes(m_Game, baseOuts, ErrorCount)

    If advanceForm.WasCancelled Then
        HandleAdvanceBase = False
        Exit Function
    End If

    Dim advText As String
    advText = BuildRunnerText(advDecisions, choice.PlayText, choice.Name = "Caught Stealing")

    loadLogger snapshot, choice, advText, ErrorCount
    If baseOuts > 0 Then m_Game.AddOuts baseOuts

    HandleAdvanceBase = True
End Function

' ----------------------------------------------------------------
' PLAY HELPERS
' ----------------------------------------------------------------

' Places the batter on the given base. Called AFTER frmResolveRunners
' has resolved existing runners, so the target base is guaranteed free.
Private Sub PlaceBatterOnBase(ByVal batterName As String, ByVal targetBase As String)
    Select Case targetBase
        Case "1B": m_Game.Runner1B = batterName
        Case "2B": m_Game.Runner2B = batterName
        Case "3B": m_Game.Runner3B = batterName
        ' Blank targetBase (3B/HR/forced-advance paths) never reaches here
    End Select
End Sub

' Builds the human-readable play-by-play text from a resolved outcomes
' collection. Pure text formatting — does NOT touch game state, since
' frmResolveRunners already applied every outcome to m_Game internally.
Private Function BuildRunnerText(ByVal decisions As Collection, _
                                 Optional ByVal actionVerb As String = "advanced to", _
                                 Optional ByVal isCaughtStealing As Boolean = False) As String
    Dim summaryText As String: summaryText = ""
    If decisions Is Nothing Then
        BuildRunnerText = summaryText
        Exit Function
    End If

    Dim outcome As clsRunnerOutcome
    For Each outcome In decisions
        Select Case outcome.NewDestination
            Case DEST_HOLDS
                ' Not noteworthy — no text fragment
            Case DEST_SECOND
                summaryText = summaryText & "; " & outcome.PlayerName & " " & _
                              IIf(outcome.AdvancedOnError, "advanced to second on error", actionVerb & " second")
            Case DEST_THIRD
                summaryText = summaryText & "; " & outcome.PlayerName & " " & _
                              IIf(outcome.AdvancedOnError, "advanced to third on error", actionVerb & " third")
            Case DEST_SCORES
                summaryText = summaryText & "; " & outcome.PlayerName & _
                              IIf(outcome.AdvancedOnError, " scored on error", " scored")
            Case DEST_OUT
                summaryText = summaryText & "; " & outcome.PlayerName & _
                              IIf(isCaughtStealing, " caught stealing", " out on play")
        End Select
    Next outcome

    ' m_Outcomes is built in descending base order (3B, 2B, 1B) per segment
    ' for correct field mutation — but that reads backwards compared to
    ' ExecuteAndLogForcedWalks (ascending: 1B->2B->2B->3B->3B scores).
    ' Split on "; " and reverse the fragments so the final text always
    ' reads lowest-base-first, consistent across both code paths.
    BuildRunnerText = ReverseSemicolonFragments(summaryText)
End Function

' Splits a "; fragment; fragment; fragment" string and returns the
' fragments in reverse order, preserving the leading "; " on each.
Private Function ReverseSemicolonFragments(ByVal rawText As String) As String
    If rawText = "" Then
        ReverseSemicolonFragments = ""
        Exit Function
    End If

    Dim cleaned As String: cleaned = rawText
    If Left(cleaned, 2) = "; " Then cleaned = Mid(cleaned, 3)

    Dim parts() As String
    parts = Split(cleaned, "; ")

    Dim reversed As String: reversed = ""
    Dim i As Long
    For i = UBound(parts) To LBound(parts) Step -1
        reversed = reversed & "; " & parts(i)
    Next i

    ReverseSemicolonFragments = reversed
End Function

Private Sub ScoreAllRunners(ByVal r3B As String, ByVal r2B As String, ByVal r1B As String, _
                             ByRef runsScored As Long, ByRef summaryText As String)
    runsScored = 0
    summaryText = ""
    
    ' Build text in ascending base order (1B, 2B, 3B) for consistency with
    ' ExecuteAndLogForcedWalks and the reversed BuildRunnerText output.
    ' Parameter names still reflect the 3B/2B/1B call-site order; only the
    ' order of these three checks changed.
    If r1B <> "" Then: runsScored = runsScored + 1: summaryText = summaryText & "; " & r1B & " scored"
    If r2B <> "" Then: runsScored = runsScored + 1: summaryText = summaryText & "; " & r2B & " scored"
    If r3B <> "" Then: runsScored = runsScored + 1: summaryText = summaryText & "; " & r3B & " scored"
End Sub

Private Sub ExecuteAndLogForcedWalks(ByVal gameState As clsPlayByPlayEvent, ByVal eventType As clsEvent)
    Dim runnerText As String: runnerText = ""

    If m_Game.Runner1B <> "" Then
        runnerText = "; " & m_Game.Runner1B & " advanced to second"

        If m_Game.Runner2B <> "" Then
            runnerText = runnerText & "; " & m_Game.Runner2B & " advanced to third"

            If m_Game.Runner3B <> "" Then
                runnerText = runnerText & "; " & m_Game.Runner3B & " scored"
                m_Game.ScoreRuns 1
            End If

            m_Game.Runner3B = m_Game.Runner2B
        End If

        m_Game.Runner2B = m_Game.Runner1B
    End If

    m_Game.Runner1B = m_Game.CurrentBatter.Name
    loadLogger gameState, eventType, runnerText
    m_Game.AdvanceBatter
    m_Game.ResetCount
End Sub

Private Sub loadLogger(ByVal snapshot As clsPlayByPlayEvent, ByVal eventType As clsEvent, _
                       Optional ByVal baserunningSummary As String = "", _
                       Optional ByVal ErrorCount As Long = 0)
    Dim recordedBalls   As Long: recordedBalls = IIf(m_Game.Balls >= MAX_BALLS, MAX_BALLS, m_Game.Balls)
    Dim recordedStrikes As Long: recordedStrikes = IIf(m_Game.Strikes >= MAX_STRIKES, MAX_STRIKES, m_Game.Strikes)
    Dim pitchText As String
    pitchText = " ( " & recordedBalls & "-" & recordedStrikes & " " & m_Game.PitchSequence & " )"

    Dim detailedPlayText As String
    If eventType.Section = "AdvanceBase" Then
        Dim cleanSummary As String: cleanSummary = baserunningSummary
        If Left(cleanSummary, 2) = "; " Then cleanSummary = Mid(cleanSummary, 3)
        detailedPlayText = cleanSummary & pitchText
    Else
        detailedPlayText = m_Game.CurrentBatter.Name & " " & eventType.PlayText & pitchText
        If baserunningSummary <> "" Then detailedPlayText = detailedPlayText & " " & baserunningSummary
    End If

    m_Logger.RecordEvent snapshot, eventType.Code, detailedPlayText, eventType.IsHit, ErrorCount
End Sub

' ----------------------------------------------------------------
' SUBSTITUTION
' ----------------------------------------------------------------
Private Sub cmdSubstitution_Click()
    PushUndo LogEntryWillBeWritten:=True
    ' Attach a lineup snapshot to the memento just pushed, since
    ' RestoreMemento alone does not touch lineup state.
    Set m_UndoStack(m_UndoStack.Count).LineupSnapshot = m_Game.CaptureLineupMemento()
    
    Dim subForm As frmSubstitution
    Set subForm = New frmSubstitution
    Dim confirmed As Boolean
    confirmed = subForm.ExecuteSubstitution(m_Game)
    
    If Not confirmed Then
        ' Nothing happened ï¿½ pop the undo entry we just pushed
        ' so a cancelled substitution doesn't leave a dead stack entry.
        m_UndoStack.Remove m_UndoStack.Count
        Exit Sub
    End If
    
    ' Update base-runner strings if the outgoing player was on base.
    ' Substitutions normally only apply between batters / half-innings,
    ' but guard against it regardless (e.g. injury mid at-bat).
    Dim outName As String: outName = subForm.OutgoingPlayerName
    Dim inName As String: inName = subForm.IncomingPlayerName
    
    If m_Game.Runner1B = outName Then m_Game.Runner1B = inName
    If m_Game.Runner2B = outName Then m_Game.Runner2B = inName
    If m_Game.Runner3B = outName Then m_Game.Runner3B = inName
    
    ' Log the substitution as a play-by-play event
    Dim snapshot As clsPlayByPlayEvent
    Set snapshot = m_Logger.TakeSnapshot(m_Game)
    
    Dim inPos As String: inPos = subForm.IncomingPosition
    Dim midText As String
    Select Case inPos
        Case "PH": midText = " pinch hit for "
        Case "PR": midText = " pinch ran for "
        Case Else: midText = " to " & inPos & " for "
    End Select
    
    Dim subText As String
    subText = inName & midText & outName & _
              " (Spot " & subForm.SelectedSpot & ", " & subForm.SelectedTeam & ")"
    
    m_Logger.RecordEvent snapshot, "SUB", subText
    
    UpdateFormDisplay
End Sub

' ----------------------------------------------------------------
' UNDO
' ----------------------------------------------------------------
Private Sub PushUndo(Optional ByVal LogEntryWillBeWritten As Boolean = True)
    Dim m As clsGameMemento
    Set m = m_Game.CaptureMemento()
    m.HasLogEntry = LogEntryWillBeWritten
    m_UndoStack.Add m
End Sub

Private Sub cmdUndo_Click()
    If m_UndoStack.Count = 0 Then Exit Sub
    
    Dim lastMemento As clsGameMemento
    Set lastMemento = m_UndoStack(m_UndoStack.Count)
    m_UndoStack.Remove m_UndoStack.Count
    
    m_Game.RestoreMemento lastMemento
    
    ' Substitutions carry a lineup snapshot; ordinary actions don't.
    If Not lastMemento.LineupSnapshot Is Nothing Then
        m_Game.RestoreLineupMemento lastMemento.LineupSnapshot
    End If
    
    If lastMemento.HasLogEntry Then m_Logger.RemoveLastEvent
    
    SetInterfaceStage InPlayMode:=False
    UpdateFormDisplay
End Sub

' ----------------------------------------------------------------
' DISPLAY
' ----------------------------------------------------------------
Private Sub UpdateFormDisplay()
    If m_Game Is Nothing Then Exit Sub

    Me.lblScoreHeader.Caption = "AWAY: " & m_Game.ScoreAway & "   |   HOME: " & m_Game.ScoreHome
    Me.lblCountHeader.Caption = "Count: " & m_Game.Balls & "-" & m_Game.Strikes
    Me.lblBasesHeader.Caption = "Bases: [ " & IIf(m_Game.Runner3B <> "", "3", "_") & " " & _
                                               IIf(m_Game.Runner2B <> "", "2", "_") & " " & _
                                               IIf(m_Game.Runner1B <> "", "1", "_") & " ]"

    Dim halfArrow As String: halfArrow = IIf(m_Game.Half = "Top", "^ Top ", "v Bot ")
    Me.lblInningHeader.Caption = halfArrow & "of Inning " & m_Game.Inning & "  (" & m_Game.Outs & " Outs)"

    Me.txtCurrentBatter.Text = IIf(Not m_Game.CurrentBatter Is Nothing, m_Game.CurrentBatter.Spot & " |  " & m_Game.CurrentBatter.Jersey & " - " & m_Game.CurrentBatter.Name, "[No Lineup Loaded]")
    Me.txtCurrentPitcher.Text = IIf(Not m_Game.CurrentPitcher Is Nothing, m_Game.CurrentPitcher.Jersey & " - " & m_Game.CurrentPitcher.Name, "[No Lineup Loaded]")
    
    Me.lblPitchCounter.Caption = "P: " & m_Game.CurrentPitcher.Pitchcounter & " IP: " & m_Game.CurrentPitcher.InningsPitched
    
    Me.cmdInPlay.Caption = IIf(m_Game.BaseState = "___", "Ball Put In Play!", "Hit or Base Advancement...")
    Me.cmdUndo.Enabled = (m_UndoStack.Count > 0)
End Sub

' ----------------------------------------------------------------
' EXPORT
' ----------------------------------------------------------------
Private Sub cmdExport_Click()
    If m_Logger Is Nothing Then Exit Sub
    m_Logger.ExportToSheet "LiveGameLog"
    MsgBox "Game history sheet 'LiveGameLog' updated successfully!", vbInformation
End Sub

' ----------------------------------------------------------------
' CANCEL FORM
' ----------------------------------------------------------------
Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    If CloseMode = vbFormControlMenu Then
        Cancel = True
        Dim response As Integer
        response = MsgBox("Export the game log before closing?", vbYesNoCancel, "Close Tracker")
        Select Case response
            Case vbYes
                m_Logger.ExportToSheet "LiveGameLog"
                Cancel = False
                UpdateResumeButton ' Session ended ï¿½ disable Resume button
                Unload Me
            Case vbNo
                Cancel = False
                UpdateResumeButton ' Session ended ï¿½ disable Resume button
                Unload Me
            Case vbCancel
                ' Form stays open ï¿½ no button state change needed
        End Select
    End If
End Sub

' Fires after Unload Me completes ï¿½ guaranteed cleanup regardless
' of which code path triggered the unload.
Private Sub UserForm_Terminate()
    NotifySessionEnded
End Sub

Private Sub cmdPause_Click()
    Me.Hide
    UpdateResumeButton ' Session now paused ï¿½ enable Resume button
End Sub
