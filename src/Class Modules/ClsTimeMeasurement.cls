VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsTimeMeasurement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private PrivateTimeMeasurementIsValid
Private StartTime As Double
Private LastTime As Double
Private EndTime As Double

Private Sub Class_Initialize()
    ResetTimer
End Sub

Private Sub Class_Terminate()
End Sub

Public Property Let TimeMeasurementIsValid( _
    SpecifiedTimeMeasurementIsValid As Boolean _
)
    PrivateTimeMeasurementIsValid = SpecifiedTimeMeasurementIsValid

    If PrivateTimeMeasurementIsValid Then ResetTimer
End Property

Public Property Get TimeMeasurementIsValid() As Boolean
    Set TimeMeasurementIsValid = PrivateTimeMeasurementIsValid
End Property

Public Sub ResetTimer()
    StartTime = Timer
    LastTime = Timer
End Sub

Public Function GetLapTime(Optional PrintTime As Boolean = True, Optional Header As String = "") As Double
    If Not PrivateTimeMeasurementIsValid Then Exit Function

    Dim LapTime As Double

    EndTime = Timer
    If EndTime < LastTime Then EndTime = EndTime + 24 * 60 * 60
    LapTime = EndTime - LastTime

    If PrintTime Then Debug.Print Header & Format(LapTime, "0.0000") & "•b"

    GetLapTime = LapTime

    LastTime = Timer
End Function

Public Function GetTotalTime(Optional PrintTime As Boolean = True, Optional Header As String = "") As Double
    If Not PrivateTimeMeasurementIsValid Then Exit Function

    Dim TotalTime As Double

    EndTime = Timer
    If EndTime < StartTime Then EndTime = EndTime + 24 * 60 * 60
    TotalTime = EndTime - StartTime

    If PrintTime Then Debug.Print Header & Format(TotalTime, "0.0000") & "•b"

    GetTotalTime = TotalTime
End Function
