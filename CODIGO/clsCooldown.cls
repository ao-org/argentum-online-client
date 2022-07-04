VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCooldown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Public totalTime As Long
Public initialTime As Long
Public iconGrh As Long

Public Sub Cooldown_Initialize(ByVal duration As Long, ByVal grh As Long)
    initialTime = GetTickCount()
    totalTime = duration
    iconGrh = grh
End Sub


