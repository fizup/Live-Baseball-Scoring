VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPlayer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Spot As Long
Public Position As String ' e.g., "P", "SS", "LF"
Public Jersey As String
Public Name As String
Public IsPitcher As Boolean
Public Pitchcounter As Long
Public OutsPitched As Long

Public Property Get InningsPitched() As Double
    InningsPitched = Int(OutsPitched / 3) + (OutsPitched Mod 3) / 10
End Property
Public Property Get ShowPlayer() As String
    ShowPlayer = Spot & " | " & Position & " | " & Jersey & " | " & Name & " | " & IsPitcher
End Property

Public Sub CopyFrom(ByVal source As clsPlayer)
    Spot = source.Spot
    Position = source.Position
    Jersey = source.Jersey
    Name = source.Name
    IsPitcher = source.IsPitcher
    Pitchcounter = source.Pitchcounter
    OutsPitched = source.OutsPitched
    ' Any future stat fields added here automatically apply to both
    ' GetRoster and DeepCopyLineup without touching either
End Sub
