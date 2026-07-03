VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsGameHistoryLogger"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_Events As Collection

Private Sub Class_Initialize()
    Set m_Events = New Collection
End Sub

' Read-only access to the event collection
Public Property Get Events() As Collection
    Set Events = m_Events
End Property

' Property to see total records logged
Public Property Get Count() As Long
    Count = m_Events.Count
End Property

' Clear history memory
Public Sub ClearLog()
    Set m_Events = New Collection
End Sub

Public Function TakeSnapshot(ByVal GameEngine As clsBaseballGame) As clsPlayByPlayEvent
    
    Dim snapshot As New clsPlayByPlayEvent
    
    ' Capture structural contextual states BEFORE engine progression executes
    ' Log data from the decoupled engine state
    snapshot.Inning = GameEngine.Inning
    snapshot.Half = GameEngine.Half
    snapshot.Outs = GameEngine.Outs
    snapshot.ScoreAway = GameEngine.ScoreAway
    snapshot.ScoreHome = GameEngine.ScoreHome
    snapshot.BaseState = GameEngine.BaseState
    
    If Not GameEngine.CurrentBatter Is Nothing Then
        snapshot.Batter = GameEngine.CurrentBatter.Name
    Else
        snapshot.Batter = "Unknown Batter"
    End If
    
    If Not GameEngine.CurrentPitcher Is Nothing Then
        snapshot.Pitcher = GameEngine.CurrentPitcher.Name
    Else
        snapshot.Pitcher = "Unknown Pitcher"
    End If
    
    Set TakeSnapshot = snapshot
    
End Function

' Record a snapshot using current details from your main game engine
Public Sub RecordEvent(ByVal snapshot As clsPlayByPlayEvent, _
                        ByVal eventCode As String, _
                        ByVal combinedPlayText As String)
    
    ' Log event details
    snapshot.Code = eventCode
    snapshot.PlayText = combinedPlayText
    
    ' Store into history memory
    m_Events.Add snapshot
End Sub

' Remove the last event
Public Sub RemoveLastEvent()
    If m_Events.Count > 0 Then m_Events.Remove m_Events.Count
End Sub

' Dump all play-by-play data out to an Excel worksheet
Public Sub ExportToSheet(ByVal targetSheetName As String)
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = Worksheets(targetSheetName)
    On Error GoTo 0
    
    If ws Is Nothing Then
        Set ws = Worksheets.Add(After:=Worksheets(Worksheets.Count))
        ws.Name = targetSheetName
    End If
    
    ws.Cells.Clear
    
  ' ---- CONSTANTS ----
    Const COL_COUNT     As Long = 11
    Dim COLOR_AWAY_ODD  As Long: COLOR_AWAY_ODD = RGB(255, 255, 255)  ' White
    Dim COLOR_AWAY_EVEN As Long: COLOR_AWAY_EVEN = RGB(235, 241, 250) ' Soft blue-grey
    Dim COLOR_HOME_ODD  As Long: COLOR_HOME_ODD = RGB(255, 252, 240)  ' Warm white
    Dim COLOR_HOME_EVEN As Long: COLOR_HOME_EVEN = RGB(255, 235, 205) ' Soft amber
    Dim COLOR_SUMMARY_AWAY As Long: COLOR_SUMMARY_AWAY = RGB(180, 198, 231) ' Deeper blue
    Dim COLOR_SUMMARY_HOME As Long: COLOR_SUMMARY_HOME = RGB(255, 195, 145) ' Deeper amber
    Dim COLOR_HEADER    As Long: COLOR_HEADER = RGB(68, 114, 196)
    
    ' ---- HEADERS ----
    Dim headers As Variant
    headers = Array("Event #", "Inning", "Half", "Code", "Play Text", "Batter", "Pitcher", "Outs", "Bases", "Away Score", "Home Score")
    
    With ws.Range("A1:K1")
        .Value = headers
        .Font.Bold = True
        .Font.Color = RGB(255, 255, 255)
        .Interior.Color = COLOR_HEADER
    End With
    
    ' ---- DATA ROWS ----
    Dim rowNum As Long: rowNum = 2
    Dim colorToggle As Boolean: colorToggle = False ' toggles each data row for alternating color
    
    Dim halfInningStartAway As Long: halfInningStartAway = 0
    Dim halfInningStartHome As Long: halfInningStartHome = 0
    
    Dim idx As Long
    Dim snapshot As clsPlayByPlayEvent
    Dim nextSnapshot As clsPlayByPlayEvent
    
    For idx = 1 To m_Events.Count
        Set snapshot = m_Events(idx)
        
        ' ---- DETECT HALF-INNING BOUNDARY ----
        ' A new half-inning starts when Inning or Half changes from the previous event,
        ' or on the very first event. Reset the starting score for the summary.
        Dim isNewHalf As Boolean
        isNewHalf = False
        If idx = 1 Then
            isNewHalf = True
        Else
            Set nextSnapshot = m_Events(idx - 1)
            If snapshot.Inning <> nextSnapshot.Inning Or snapshot.Half <> nextSnapshot.Half Then
                isNewHalf = True
            End If
        End If
        
        If isNewHalf Then
            halfInningStartAway = snapshot.ScoreAway
            halfInningStartHome = snapshot.ScoreHome
        End If
        
        ' ---- DETECT HALF-INNING END ----
        ' Insert a summary row after the last event of a half-inning —
        ' either when the next event is a different half, or on the last event overall.
        Dim isEndOfHalf As Boolean
        isEndOfHalf = False
        If idx = m_Events.Count Then
            isEndOfHalf = True
        Else
            Set nextSnapshot = m_Events(idx + 1)
            If snapshot.Inning <> nextSnapshot.Inning Or snapshot.Half <> nextSnapshot.Half Then
                isEndOfHalf = True
            End If
        End If
        
        ' ---- WRITE DATA ROW ----
        Dim rowColor As Long
        If snapshot.Half = "Top" Then
            rowColor = IIf(colorToggle, COLOR_AWAY_EVEN, COLOR_AWAY_ODD)
        Else
            rowColor = IIf(colorToggle, COLOR_HOME_EVEN, COLOR_HOME_ODD)
        End If
        colorToggle = Not colorToggle
        
        With ws.Rows(rowNum)
            .Interior.Color = rowColor
        End With
        
        ws.Cells(rowNum, 1).Value = idx
        ws.Cells(rowNum, 2).Value = snapshot.Inning
        ws.Cells(rowNum, 3).Value = snapshot.Half
        ws.Cells(rowNum, 4).Value = snapshot.Code
        ws.Cells(rowNum, 5).Value = snapshot.PlayText
        ws.Cells(rowNum, 6).Value = snapshot.Batter
        ws.Cells(rowNum, 7).Value = snapshot.Pitcher
        ws.Cells(rowNum, 8).Value = snapshot.Outs
        ws.Cells(rowNum, 9).Value = "'" & snapshot.BaseState
        ws.Cells(rowNum, 10).Value = snapshot.ScoreAway
        ws.Cells(rowNum, 11).Value = snapshot.ScoreHome
        
        
        ' ---- BOLD: run scored on this play ----
        If InStr(snapshot.PlayText, "scored") > 0 Or InStr(snapshot.PlayText, "homered") > 0 Then
            ws.Rows(rowNum).Font.Bold = True
        End If
        
        ' ---- ITALIC: substitution or future GameEvent codes ----
        If snapshot.Code = "SUB" Then
            ws.Rows(rowNum).Font.Italic = True
        End If
        
        rowNum = rowNum + 1
        
        ' ---- SUMMARY ROW ----
        If isEndOfHalf Then
            Dim runsAway As Long: runsAway = snapshot.ScoreAway - halfInningStartAway
            Dim runsHome As Long: runsHome = snapshot.ScoreHome - halfInningStartHome
            
            Dim summaryTeam As String
            Dim summaryRuns As Long
            If snapshot.Half = "Top" Then
                summaryTeam = "Away"
                summaryRuns = runsAway
            Else
                summaryTeam = "Home"
                summaryRuns = runsHome
            End If
            
            Dim summaryText As String
            summaryText = summaryTeam & " Team — " & snapshot.Half & " " & snapshot.Inning & ":  Runs: " & summaryRuns
            ' Placeholder columns for future stats — blank for now, easy to extend
            ' Hits: _ | RBI: _ | Errors: _ | LOB: _

            With ws.Range(ws.Cells(rowNum, 1), ws.Cells(rowNum, COL_COUNT))
                .Merge
                .Value = summaryText
                .Font.Bold = True
                .Font.Italic = True
                .Interior.Color = IIf(snapshot.Half = "Top", COLOR_SUMMARY_AWAY, COLOR_SUMMARY_HOME)
                .HorizontalAlignment = xlLeft
                .IndentLevel = 1
            End With
            
            ' Reset alternating color toggle for the next half-inning
            colorToggle = False
            rowNum = rowNum + 1
        End If
    Next idx
    
    ' ---- FINAL FORMATTING ----
    ws.Columns("A:K").AutoFit
    ws.Columns("E").ColumnWidth = 60 ' Play Text column — don't let AutoFit shrink it too much
    
    ' Freeze the header row
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub



