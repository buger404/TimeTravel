VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "TimeSpace"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Private Ar(5, 5) As New Fighter
Public Right As TimeSpace
Public Up As TimeSpace
Public Down As TimeSpace
Public Failed As Boolean
Public ID As Integer
Public Distance As Integer

Public Property Get Area(X As Long, Y As Long) As Fighter
    Set Area = Ar(X, Y)
End Property

Private Sub Class_Initialize()
    ID = TotalId + 1
    TotalId = TotalId + 1
End Sub
