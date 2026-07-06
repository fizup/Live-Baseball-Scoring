VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsLineupAssignment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public playerName  As String
Public OldSpot     As Long    ' captured at staging time, for log text only
Public OldPosition As String  ' captured at staging time, for log text only
Public NewSpot     As Long
Public NewPosition As String

