VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTimer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
' in clsTimer...
Dim start, finish
Private Declare Function GetTickCount Lib "kernel32" () As Long


Public Sub StopTimer()


    finish = GetTickCount()
End Sub



Public Sub StartTimer()

    start = GetTickCount()
    finish = 0
End Sub



Public Sub DebugTrace(v)

    'Debug.Print v & " " & Elapsed()
End Sub



Public Property Get Elapsed()



    If finish = 0 Then
        Elapsed = GetTickCount() - start
    Else
        Elapsed = finish - start
    End If

End Property

