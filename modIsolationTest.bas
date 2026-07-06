Attribute VB_Name = "modIsolationTest"
Option Explicit

Sub RunIsolatedRunnerTest()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== STARTING ISOLATION TEST ==="
    
    ' 1. Set up a Mock Hitter and Pitcher
    Dim mockGame As New clsBaseballGame
    Dim mockLineup As New Collection
    Dim p1 As New clsPlayer: p1.Name = "Test Batter": p1.IsPitcher = False
    Dim p2 As New clsPlayer: p2.Name = "Test Pitcher": p2.IsPitcher = True
    mockLineup.Add p1
    mockLineup.Add p2
    Set mockGame.LineUpAway = mockLineup
    mockGame.Half = "Top"
    mockGame.BatterIndexAway = 1
    
    ' 2. Place a runner on 1st base manually to trigger our simulation scenario
    Debug.Print "Placing 'Runner One' on 1B..."
    mockGame.Runner1B = "Runner One"
    
    ' --- VERIFY BASELINE ---
    Debug.Print "Baseline Check - Runner1B value: " & mockGame.Runner1B
    
    ' 3. Simulate creating mock decision data (What the popup form normally generates)
    Debug.Print "Creating mock player decisions collection..."
    Dim mockDecisions As New Collection
    Dim d1 As New clsRunnerOutcome
    
    d1.BaseSource = "1B"
    d1.playerName = mockGame.Runner1B ' Calls Property Get
    d1.NewDestination = "2"           ' Operator wants him to advance to 2B
    mockDecisions.Add d1
    
    ' 4. Run the Transactional Reconstruction Loop (from cmdSubmit_Click)
    Debug.Print "Executing Transactional Reconstruction Loop..."
    
    Dim old1B As String: old1B = mockGame.Runner1B
    Dim old2B As String: old2B = mockGame.Runner2B
    Dim old3B As String: old3B = mockGame.Runner3B
    
    ' Wipe bases
    mockGame.Runner1B = ""
    mockGame.Runner2B = ""
    mockGame.Runner3B = ""
    
    ' Reassign based on choices
    Dim outcome As clsRunnerOutcome
    For Each outcome In mockDecisions
        Dim currentRunnerName As String
        Select Case outcome.BaseSource
            Case "1B": currentRunnerName = old1B
            Case "2B": currentRunnerName = old2B
            Case "3B": currentRunnerName = old3B
        End Select
        
        Select Case outcome.NewDestination
            Case "1": mockGame.Runner1B = currentRunnerName
            Case "2": mockGame.Runner2B = currentRunnerName ' <-- Watch if it breaks here!
            Case "3": mockGame.Runner3B = currentRunnerName
            Case "Scores": mockGame.ScoreRuns 1
        End Select
    Next outcome
    
    ' Place the new batter on 1B
    mockGame.Runner1B = "New Batter"
    
    Debug.Print "=== TEST COMPLETED SUCCESSFULLY ==="
    Debug.Print "New 1B: " & mockGame.Runner1B
    Debug.Print "New 2B: " & mockGame.Runner2B
    Exit Sub

ErrorHandler:
    Debug.Print "CRASH DETECTED!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Description: " & Err.description
    Debug.Print "Source: " & Err.source
End Sub


Sub TestUserFormInitialization()
    On Error GoTo ErrorHandler
    
    Debug.Print "=== DIAGNOSTIC: STARTING USERFORM TESTS ==="
    
    ' 1. Spin up a fresh engine instance
    Dim testGame As New clsBaseballGame
    testGame.Runner1B = "Test Runner 1B"
    testGame.Runner2B = "Test Runner 2B"
    testGame.Runner3B = ""
    
    Debug.Print "Step 1 Completed: Class initialized with test runners successfully."
    
    ' 2. Instantiating the form directly to see if instantiation cracks memory
    Dim testForm As frmResolveRunners
    Debug.Print "Step 2: Attempting Form allocation..."
    Set testForm = New frmResolveRunners
    Debug.Print "Step 2 Completed: Form instance successfully created in memory."
    
    ' 3. Execution of the entry point function
    Debug.Print "Step 3: Handing engine over to GetSeparatedOutcomes..."
    Dim sampleOutcomes As Collection
    Set sampleOutcomes = testForm.GetSeparatedOutcomes(testGame, 0)
    
    Debug.Print "Step 4: Form successfully closed and execution returned."
    If Not sampleOutcomes Is Nothing Then
        Debug.Print "Outcomes collection size: " & sampleOutcomes.Count
    End If
    
    Unload testForm
    Set testForm = Nothing
    
    Debug.Print "=== DIAGNOSTIC SUCCESSFUL: NO ERROR DETECTED ==="
    Exit Sub

ErrorHandler:
    Debug.Print "!!! DIAGNOSTIC FAILED !!!"
    Debug.Print "Error Number: " & Err.Number
    Debug.Print "Error Description: " & Err.description
    Debug.Print "Failing Module/Context: " & Err.source
    
    ' Inspect exactly what state caused the collapse
    On Error Resume Next
    If Not testGame Is Nothing Then
        Debug.Print "Game Engine Integrity - Runner1B: " & testGame.Runner1B
    Else
        Debug.Print "Game Engine variable was completely wiped out."
    End If
End Sub

