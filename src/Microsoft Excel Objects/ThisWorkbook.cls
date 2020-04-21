VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private CsvOpener As ClsCsvOpener
Private WithEvents CsvOpenerControl As ClsCsvOpenerControl
Attribute CsvOpenerControl.VB_VarHelpID = -1

Private Sub ResetCsvOpener()
    Set CsvOpener = New ClsCsvOpener
    Set CsvOpenerControl = New ClsCsvOpenerControl

    CsvOpenerControl.StartEventObservation

    If CsvOpenerControl.CsvOpenerIsValid Then
        CsvOpener.StartEventObservation
    Else
        CsvOpener.StopEventObservation
    End If
End Sub

Private Sub UpdateBook()
    With Application
        .EnableEvents = False
        .ScreenUpdating = False
        .DisplayAlerts = False
    End With
    
    Me.Save

    With Application
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True
    End With
End Sub

Private Sub CsvOpenerControl_EventStartObservationRequest( _
    ByVal TargetRange As Range _
)
'    Debug.Print "EventCsvOpenerControl_EventStartObservationRequest()", TargetRange.Address

    If CsvOpener Is Nothing Then
        ResetCsvOpener
    Else
        CsvOpener.StartEventObservation
    End If

    UpdateBook
End Sub

Private Sub CsvOpenerControl_EventStopObservationRequest( _
    ByVal TargetRange As Range _
)
'    Debug.Print "EventCsvOpenerControl_EventStopObservationRequest()", TargetRange.Address

    If CsvOpener Is Nothing Then
        ResetCsvOpener
    Else
        CsvOpener.StopEventObservation
    End If

    UpdateBook
End Sub

Private Sub Workbook_Open()
    On Error Resume Next
    InstallAddIn
    ResetCsvOpener
End Sub

Private Sub Workbook_AfterSave( _
    ByVal Success As Boolean _
)
    ResetCsvOpener
End Sub

Private Sub Workbook_BeforeClose( _
    Cancel As Boolean _
)
    On Error Resume Next
    UpdateBook
    UninstallAddIn
End Sub

