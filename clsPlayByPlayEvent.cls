VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayByPlayEvent"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' ----------------------------------------------------------------
' Explicit data container matching the target GameLog structure
' ----------------------------------------------------------------
' Game State Context
Public Inning As Long
Public Half As String
Public Outs As Long
Public ScoreAway As Long
Public ScoreHome As Long
Public BaseState As String

' Event Specifics
Public Code As String
Public PlayText As String
Public Pitcher As String
Public Batter As String
Public IsHit As Boolean
Public IsError As Boolean

