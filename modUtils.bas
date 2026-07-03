Attribute VB_Name = "modUtils"
Option Explicit

' Build objects out here to keep the game engine clean

' Returns a collection of clsPlayer objects loaded from
' the roster tables in the LineUp worksheets.
Public Function GetRoster(ByVal sheetName As String) As Collection
    Dim roster As New Collection
    Dim ws As Worksheet
    Dim i As Long, lastRow As Long
    Dim p As clsPlayer
    
    Set ws = Worksheets(sheetName)
    lastRow = ws.Cells(ws.Rows.Count, "D").End(xlUp).Row
        
    'Start in row 3 to skip title and headers row
    For i = 3 To lastRow
        Set p = New clsPlayer
        p.Spot = ws.Cells(i, 1).Value
        p.Position = ws.Cells(i, 2).Value
        p.Jersey = ws.Cells(i, 3).Value
        p.Name = ws.Cells(i, 4).Value
        p.IsPitcher = (UCase(ws.Cells(i, 2).Value) = "P")
        If p.IsPitcher Then
            p.Pitchcounter = 0
            p.OutsPitched = 0
        End If
        roster.Add p
    Next i
    
    Set GetRoster = roster

End Function

Sub testGetRoster()
    Dim awayRoster As Collection
    Set awayRoster = GetRoster("LineUpAway")
    
    Dim p As clsPlayer
    For Each p In awayRoster
        Debug.Print p.ShowPlayer
    Next
End Sub

' Retruns a list of event names to populate listbox in userform
Public Function GetList(ByVal eventType) As Variant
    Dim rng As Range
    Dim Section As Range
    
    Set rng = Tabelle3.Columns(1).Find(What:=eventType, LookAt:=xlWhole).CurrentRegion
    
    If rng Is Nothing Then Exit Function

    Set Section = rng.CurrentRegion

    ' Return values from column C, starting after the
    ' section title and column header rows.
    GetList = Section.Range("c3:c" & Section.Rows.Count).Value

End Function

' Returns a collection of clsEvent objects loaded from
' the event definition tables in the Glossary worksheet.
Public Function GetEvents() As Collection

    Dim Events As New Collection
    Dim ws As Worksheet
    Dim sectionName As Variant

    Set ws = Worksheets("Glossary")

    ' Load events from each glossary section.
    ' Each section is expected to be a separate table
    ' identified by its header in column A.
    For Each sectionName In Array("GetOnBase", "GetOut", "AdvanceBase")
        LoadEventsFromSection ws, CStr(sectionName), Events
    Next sectionName

    Set GetEvents = Events

End Function

' Reads all event definitions from a glossary section and
' adds them to the supplied collection.
'
' Expected table structure:
'   Row 1 = Section name (e.g. "GetOnBase")
'   Row 2 = Column headers
'   Row 3+ = Event data
'
' Columns:
'   A = Event Code
'   B = Play Text
'   C = Event Name
'   D = Ball is hit
'   E = Batter's Target Base
'   F = Is Hit
'   G = Is Error
Private Sub LoadEventsFromSection(ByVal ws As Worksheet, _
                                  ByVal sectionName As String, _
                                  ByRef Events As Collection)

    Dim rng As Range
    Dim e As clsEvent
    Dim i As Long
    Dim lastRow As Long

    ' Locate the section header in column A.
    Set rng = ws.Columns(1).Find(What:=sectionName, LookAt:=xlWhole)

    ' Skip processing if the section cannot be found.
    If rng Is Nothing Then Exit Sub

    ' The section is assumed to be a contiguous table.
    lastRow = rng.CurrentRegion.Rows.Count

    ' Start at row 3 to skip the section title and column headers.
    For i = 3 To lastRow

        Set e = New clsEvent

        ' Populate the event object from the current row.
        With rng.CurrentRegion
            e.Code = .Cells(i, 1).Value
            e.PlayText = .Cells(i, 2).Value
            e.Name = .Cells(i, 3).Value
            e.BallIsHit = IIf(.Cells(i, 4).Value <> "", True, False)
            e.targetBase = .Cells(i, 5).Value
            e.IsHit = IIf(.Cells(i, 6).Value <> "", True, False)
            e.IsError = IIf(.Cells(i, 7).Value <> "", True, False)
            e.Section = sectionName
        End With

        ' Store the event using its name as the collection key.
        Events.Add e, e.Name

    Next i

End Sub

Sub testGetEvents()
    Dim Events As Collection
    Set Events = GetEvents
    
    Dim e As clsEvent
    For Each e In Events
        Debug.Print e.Code, e.Name
    Next
    Debug.Print Events.Count
    Debug.Print Events("Base On Balls").Name
End Sub
