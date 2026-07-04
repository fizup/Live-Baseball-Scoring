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

Private m_Outcomes As Collection
Private m_GameRef As clsBaseballGame
Private m_OutsAdded As Long
Private m_Cancelled As Boolean

' Read-only accessor for the caller to check
Public Property Get WasCancelled() As Boolean
    WasCancelled = m_Cancelled
End Property

''
' The clean entry point called by your primary play-recording form.
''
Public Function GetSeparatedOutcomes(ByRef GameEngine As clsBaseballGame, ByRef AdditionalOuts As Long) As Collection
    
    If GameEngine.BaseState = "___" Then
        Set GetSeparatedOutcomes = Nothing
        Exit Function
    End If
    
    Set m_GameRef = GameEngine
    Set m_Outcomes = New Collection
    m_OutsAdded = 0
    
    ' Configure dropdown menus and hide rows where bases are empty
    SetupDropdown Me.cboRunner1B, Me.lblRunner1B, GameEngine.Runner1B, "1B"
    SetupDropdown Me.cboRunner2B, Me.lblRunner2B, GameEngine.Runner2B, "2B"
    SetupDropdown Me.cboRunner3B, Me.lblRunner3B, GameEngine.Runner3B, "3B"
    
    ' Display this window modally and halt parent execution until closed
    Me.Show vbModal
    
    ' Hand back the compiled collection and total outs to the caller
    AdditionalOuts = m_OutsAdded
    Set GetSeparatedOutcomes = m_Outcomes
End Function

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
    
    ' Available tactical base path moves
    cbo.AddItem "Stays/Holds"
    If baseCode = "1B" Then
        cbo.AddItem "Advances to 2B"
        cbo.AddItem "Advances to 3B"
    End If
    If baseCode = "2B" Then
        cbo.AddItem "Advances to 3B"
    End If
    cbo.AddItem "Scores"
    cbo.AddItem "Out on Play"
    
    cbo.ListIndex = 0 ' Default to hold position
End Sub


Private Sub cmdSaveRunners_Click()
    If m_GameRef Is Nothing Then
        MsgBox "Form not initialized via GetSeparatedOutcomes.", vbCritical
        Exit Sub
    End If
    
    ' Reset collection and out counter before every build attempt.
    ' This ensures a resubmit after a failed validation starts with a clean slate
    ' rather than accumulating stale entries or double-counting outs.
    Set m_Outcomes = New Collection
    m_OutsAdded = 0
    
    ' Build the outcomes collection from the visible dropdowns
    ProcessBaseChoice Me.cboRunner1B, m_GameRef.Runner1B, "1B"
    ProcessBaseChoice Me.cboRunner2B, m_GameRef.Runner2B, "2B"
    ProcessBaseChoice Me.cboRunner3B, m_GameRef.Runner3B, "3B"
    
    ' Track base occupations to block collisions.
    ' Pre-occupy the base where the batter is heading so runners can't land there.
    Dim baseOccupied(1 To 3) As Boolean
    If m_GameRef.EventToHappen = "1B" Then baseOccupied(1) = True
    If m_GameRef.EventToHappen = "2B" Then baseOccupied(2) = True
    
    Dim CheckSum As Long: CheckSum = 0
    Dim outcome As clsRunnerOutcome
    
    For Each outcome In m_Outcomes
        CheckSum = CheckSum + outcome.NewDestination
        
        ' Resolve the base number this runner will occupy (DEST_HOLDS means stay put)
        Dim finalBase As Long
        If outcome.NewDestination = DEST_HOLDS Then
            Select Case outcome.BaseSource
                Case "1B": finalBase = 1
                Case "2B": finalBase = 2
                Case "3B": finalBase = 3
            End Select
        Else
            finalBase = outcome.NewDestination ' DEST_SECOND=2, DEST_THIRD=3; DEST_SCORES/OUT don't land on a base
        End If
        
        ' --- VALIDATION 1: FORCE PLAY RULES ---
        ' Check before the collision guard so we only fire one message per violation.
        If outcome.BaseSource = "1B" Then
            ' On a Single: 1B runner cannot hold — batter is taking their base
            If outcome.NewDestination = DEST_HOLDS And m_GameRef.EventToHappen = "1B" Then
                MsgBox "Rule Violation! The runner on 1st base is forced to advance on a Single.", vbExclamation
                Exit Sub
            End If
            
            ' On a Double: 1B runner cannot stay on 1B or advance to 2B — batter is taking 2B
            If m_GameRef.EventToHappen = "2B" Then
                If outcome.NewDestination = DEST_HOLDS Or outcome.NewDestination = DEST_SECOND Then
                    MsgBox "Rule Violation! The runner on 1st base must advance to at least 3rd base on a Double.", vbExclamation
                    Exit Sub
                End If
            End If
        End If
        
        ' Runner on 2B cannot hold if the runner behind them is advancing to 2B (force play)
        If outcome.BaseSource = "2B" And outcome.NewDestination = DEST_HOLDS And Me.cboRunner1B.Visible Then
            If Me.cboRunner1B.Value = "Advances to 2B" Then
                MsgBox "Rule Violation! The runner on 2nd base is forced to move by the runner advancing behind them.", vbExclamation
                Exit Sub
            End If
        End If
        
        ' --- VALIDATION 2: BASE COLLISION ---
        ' Only applies to runners landing on an actual base (not scoring or getting out)
        If finalBase >= 1 And finalBase <= 3 Then
            If baseOccupied(finalBase) Then
                MsgBox "Base Collision! Two runners cannot occupy base " & finalBase & " simultaneously.", vbExclamation
                Exit Sub
            End If
            baseOccupied(finalBase) = True
        End If
    Next outcome
    
    ' Soft check: warn if all runners are holding (likely a user mistake)
    If CheckSum = 0 Then
        Dim userInput As Integer
        userInput = MsgBox("All runners stay on their bases! Proceed anyway?", vbQuestion + vbOKCancel)
        If userInput <> vbOK Then Exit Sub
    End If
    
    Me.Hide
End Sub

Private Sub ProcessBaseChoice(ByRef cbo As Control, ByVal runnerName As String, ByVal originalBase As String)
    If Not cbo.Visible Then Exit Sub
    
    Dim decision As clsRunnerOutcome
    Set decision = New clsRunnerOutcome
    
    decision.BaseSource = originalBase
    decision.PlayerName = runnerName
    
    Select Case cbo.Value
        Case "Stays/Holds":     decision.NewDestination = DEST_HOLDS
        Case "Advances to 2B":  decision.NewDestination = DEST_SECOND
        Case "Advances to 3B":  decision.NewDestination = DEST_THIRD
        Case "Scores":          decision.NewDestination = DEST_SCORES
        Case "Out on Play":
            decision.NewDestination = DEST_OUT
            m_OutsAdded = m_OutsAdded + 1
    End Select
    
    m_Outcomes.Add decision
End Sub

' ----------------------------------------------------------------
' CANCEL FORM
' ----------------------------------------------------------------
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
