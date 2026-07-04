VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsRunnerOutcome"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public BaseSource As String      ' "1B", "2B", or "3B"
Public PlayerName As String      ' Name of the runner
Public NewDestination As Long   ' 0 = Hold Base, 2 = abvance to 2nd, 3 = advance to 3rd, 4 = Score Run, 5 = Out on play
Public AdvancedOnError As Boolean  ' True when the advance was due to a fielding error
