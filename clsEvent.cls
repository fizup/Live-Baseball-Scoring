VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Code As String
Public PlayText As String
Public Name As String
Public BallIsHit As Boolean
Public targetBase As String
Public Section As String ' Stores "GetOnBase", "GetOut", or "AdvanceBase"
Public IsHit As Boolean
Public IsError As Boolean ' Checks if batter reaches on error
Public ErrorPosition As String  ' optional, for future fielder stats
