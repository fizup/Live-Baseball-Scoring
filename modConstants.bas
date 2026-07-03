Attribute VB_Name = "modConstants"
Option Explicit

' ----------------------------------------------------------------
' Runner destination codes — shared between frmResolveRunners
' (where they are written) and frmRecordPlay (where they are read).
' Defined here so both forms always refer to the same values.
' ----------------------------------------------------------------
Public Const DEST_HOLDS  As Long = 0 ' Runner stays on their current base
Public Const DEST_SECOND As Long = 2 ' Runner advances to 2nd base
Public Const DEST_THIRD  As Long = 3 ' Runner advances to 3rd base
Public Const DEST_SCORES As Long = 4 ' Runner scores a run
Public Const DEST_OUT    As Long = 5 ' Runner is put out on the base paths

' ----------------------------------------------------------------
' Constants to compile a traditional scorecard count
' with loadLogger where they are read ( max count: 3-2 )
' ----------------------------------------------------------------
Public Const MAX_BALLS As Long = 3
Public Const MAX_STRIKES As Long = 2
