Attribute VB_Name = "modMain"
Option Explicit

Private m_ActiveForm As frmRecordPlay
Private Const SHEET_NAME As String = "LiveGameLog"    ' adjust to match your sheet name
Private Const BTN_RESUME As String = "btnResume"    ' adjust to match your button name

Public Sub StartLiveGameTracker()
    Dim liveGame As New clsBaseballGame
    Dim liveLogger As New clsGameHistoryLogger
    
    Set liveGame.LineUpAway = GetRoster("LineUpAway")
    Set liveGame.LineUpHome = GetRoster("LineUpHome")
    liveGame.Half = "Top"
    liveGame.BatterIndexAway = 1

    Set m_ActiveForm = New frmRecordPlay
    m_ActiveForm.InitializeForm liveGame, liveLogger
    
    UpdateResumeButton
    m_ActiveForm.Show vbModal
End Sub

Public Sub ResumeGame()
    If m_ActiveForm Is Nothing Then
        MsgBox "No active game session found.", vbExclamation
        Exit Sub
    End If
    m_ActiveForm.Show vbModal
    UpdateResumeButton
End Sub

' Called by frmRecordPlay.cmdPause_Click and UserForm_QueryClose
' to keep the Resume button state in sync with form visibility.
Public Sub UpdateResumeButton()
    Dim btn As Object
    On Error Resume Next
    Set btn = ThisWorkbook.Sheets(SHEET_NAME).Buttons(BTN_RESUME)
    On Error GoTo 0
    
    If btn Is Nothing Then Exit Sub
        
    Dim sessionIsPaused As Boolean
    If Not m_ActiveForm Is Nothing Then
        On Error Resume Next  ' Guards against disconnected RPC reference
        sessionIsPaused = Not m_ActiveForm.Visible
        If Err.Number <> 0 Then
            ' Reference is stale — treat as no active session
            Set m_ActiveForm = Nothing
            sessionIsPaused = False
            Err.Clear
        End If
        On Error GoTo 0
    End If
    
    btn.Enabled = sessionIsPaused
End Sub

' A cleanup entry point called by the form on unload
Public Sub NotifySessionEnded()
    Set m_ActiveForm = Nothing
    UpdateResumeButton
End Sub
