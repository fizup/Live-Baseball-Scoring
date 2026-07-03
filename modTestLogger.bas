Attribute VB_Name = "modTestLogger"
Option Explicit

Sub TestGameLoggerSystem()
    Dim game As New clsBaseballGame
    Dim logger As New clsGameHistoryLogger
    
    ' Simulate: Play 1 - Lead-off Batter hits a single
    game.Runner1B = "Away Player 1"
    logger.RecordEvent game, "Single to Left Field", "Away Player 1"
    
    ' Simulate: Play 2 - Next batter strikes out
    game.AddOuts 1
    logger.RecordEvent game, "Strikeout Looking", "Away Player 2"
    
    ' Simulate: Play 3 - Double play clears runners and forces a side change
    game.AddOuts 2 ' Triggers ChangeSides (clears bases, changes half to "Bot")
    logger.RecordEvent game, "Ground into Double Play (6-4-3)", "Away Player 3"
    
    ' Verify memory structure counts
    Debug.Assert logger.Count = 3
    
    ' Inspect a specific point in history
    Dim midGameRecord As clsPlayByPlayEvent
    Set midGameRecord = logger.Events(2)
    Debug.Assert midGameRecord.Outs = 1
    Debug.Assert midGameRecord.Half = "Top"
    
    ' Export everything cleanly to a logging sheet
    logger.ExportToSheet "GameLog_Output"
    
    Debug.Print "Log Test Completed: 3 game changes successfully processed into memory and saved!"
End Sub
