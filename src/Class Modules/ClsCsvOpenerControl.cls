VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCsvOpenerControl"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Const CsvOpener_Parameters_Sheetname = "CsvOpener_Parameters"
Private Const IsValid_Cellname = "$B$1"

Private WithEvents TargetExcelApplication As Excel.Application
Attribute TargetExcelApplication.VB_VarHelpID = -1

Public Event EventStartObservationRequest( _
    ByVal TargetRange As Range _
)

Public Event EventStopObservationRequest( _
    ByVal TargetRange As Range _
)

Public Property Set ExcelApplication( _
    SpecifiedApplication As Excel.Application _
)
    Set TargetExcelApplication = SpecifiedApplication
End Property

Public Property Get ExcelApplication() As Excel.Application
    Set ExcelApplication = TargetExcelApplication
End Property

Public Property Let CsvOpenerIsValid( _
    SpecifiedFlag As Boolean _
)
    ThisWorkbook.Sheets(CsvOpener_Parameters_Sheetname).Range(IsValid_Cellname).Value = SpecifiedFlag
End Property

Public Property Get CsvOpenerIsValid() As Boolean
    CsvOpenerIsValid = ThisWorkbook.Sheets(CsvOpener_Parameters_Sheetname).Range(IsValid_Cellname).Value <> False
End Property

Public Property Get IsUnconvertedCSV() As Boolean
    If LCase(Right(ActiveWorkbook.Name, 4)) <> ".csv" Then
        IsUnconvertedCSV = False
        Exit Property
    End If

    Dim TargetQuery As WorkbookQuery
    
    On Error Resume Next
    Set TargetQuery = ActiveWorkbook.Queries("CSVTable")
    On Error GoTo 0

    IsUnconvertedCSV = TargetQuery Is Nothing
End Property

Private Sub Class_Initialize()
    Set Me.ExcelApplication = Nothing
End Sub

Private Sub Class_Terminate()
    Set Me.ExcelApplication = Nothing
End Sub

Public Sub StartEventObservation( _
    Optional SpecifiedApplication As Excel.Application _
)
    If SpecifiedApplication Is Nothing Then Set SpecifiedApplication = Application
    Set Me.ExcelApplication = SpecifiedApplication
End Sub

Public Sub StopEventObservation()
    Set Me.ExcelApplication = Nothing
End Sub

Private Sub TargetExcelApplication_SheetChange( _
    ByVal Sh As Object, _
    ByVal TargetRange As Range _
)
'    Debug.Print "TargetExcelApplication_SheetChange", TargetRange.Address
    
    Dim TargetSheet As Worksheet

    Set TargetSheet = TargetRange.Parent
    If TargetSheet.Name <> CsvOpener_Parameters_Sheetname Then Exit Sub
    If Application.Intersect(TargetRange, TargetRange.Parent.Range(IsValid_Cellname)) Is Nothing Then Exit Sub

    If CsvOpenerIsValid Then
        RaiseEvent EventStartObservationRequest(TargetRange)
    Else
        RaiseEvent EventStopObservationRequest(TargetRange)
    End If
End Sub
