Attribute VB_Name = "StdCsvOpenerControl"
Option Explicit

Private RibbonUI As IRibbonUI

Sub StartCsvOpener()
    Dim OpenerControl As New ClsCsvOpenerControl
    OpenerControl.CsvOpenerIsValid = True
    UpdateRiboonControl
End Sub

Sub StopCsvOpener()
    Dim OpenerControl As New ClsCsvOpenerControl
    OpenerControl.CsvOpenerIsValid = False
    UpdateRiboonControl
End Sub

Sub ConvertCsv()
    Dim CsvOpener As New ClsCsvOpener
    CsvOpener.ConvertCsv
End Sub

Sub ConvertSelectionTextToNumber()
    If Selection Is Nothing Then Exit Sub

    Dim TargetRange As Range
    Dim TempColumn As Range

    Set TargetRange = Selection

    On Error Resume Next
    For Each TempColumn In TargetRange.Columns
        TempColumn.TextToColumns _
            Destination:=TempColumn, _
            DataType:=xlDelimited, _
            TextQualifier:=xlDoubleQuote, _
            ConsecutiveDelimiter:=False, _
            Tab:=True, _
            Semicolon:=False, _
            Comma:=False, _
            Space:=False, _
            Other:=False, _
            FieldInfo:=Array(1, 1), _
            TrailingMinusNumbers:=True
    Next TempColumn
    On Error GoTo 0
End Sub

Sub UpdateRiboonControl()
    On Error Resume Next
    RibbonUI.InvalidateControl "StartAutomaticConversion"
    RibbonUI.InvalidateControl "StopAutomaticConversion"
    RibbonUI.InvalidateControl "ManualConversion"
End Sub

Sub RibbonControl_Onload(ByVal SpecifiedRibbonUI As IRibbonUI)
    Set RibbonUI = SpecifiedRibbonUI
    RibbonUI.Invalidate
End Sub

Sub RibbonControl_StartCsvOpener(ByVal RibbonControl As IRibbonControl)
    StartCsvOpener
End Sub

Sub RibbonControl_StopCsvOpener(ByVal RibbonControl As IRibbonControl)
    StopCsvOpener
End Sub

Sub RibbonControl_ConvertCsv(ByVal RibbonControl As IRibbonControl)
    ConvertCsv
End Sub

Sub RibbonControl_ConvertSelectionTextToNumber(ByVal control As IRibbonControl)
    ConvertSelectionTextToNumber
End Sub

Sub RibbonControl_StartCsvOpener_getEnabled(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Dim CsvOpenerControl As New ClsCsvOpenerControl
    
    ReturnedVal = Not CsvOpenerControl.CsvOpenerIsValid
End Sub

Sub RibbonControl_StopCsvOpener_getEnabled(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Dim CsvOpenerControl As New ClsCsvOpenerControl
    
    ReturnedVal = CsvOpenerControl.CsvOpenerIsValid
End Sub

Sub RibbonControl_ConvertCsv_getEnabled(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Dim CsvOpenerControl As New ClsCsvOpenerControl
    
    ReturnedVal = True
    'ReturnedVal = Not CsvOpenerControl.CsvOpenerIsValid
    'ReturnedVal = CsvOpenerControl.IsUnconvertedCSV
    ' TODO: 変換可能かどうかの判定と更新タイミングをはかるのが難しいので保留
End Sub

Sub RibbonControl_ApplicationName(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "CSV変換"
        Case Else: ReturnedVal = "CSV Conversion"
    End Select
End Sub

Sub RibbonControl_ControlLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "コントロール"
        Case Else: ReturnedVal = "Control"
    End Select
End Sub

Sub RibbonControl_ToolsLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "操作"
        Case Else: ReturnedVal = "Tools"
    End Select
End Sub

Sub RibbonControl_StartAutomaticConversionLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "自動変換開始"
        Case Else: ReturnedVal = "Start Automatic Conversion"
    End Select
End Sub

Sub RibbonControl_StopAutomaticConversionLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "自動変換停止"
        Case Else: ReturnedVal = "Stop Automatic Conversion"
    End Select
End Sub

Sub RibbonControl_ManualConversionLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "手動変換"
        Case Else: ReturnedVal = "Manual Conversion"
    End Select
End Sub

Sub RibbonControl_TextToNumberConversionLabel(ByVal RibbonControl As IRibbonControl, ByRef ReturnedVal)
    Select Case Application.International(xlCountryCode)
        Case 81: ReturnedVal = "文字を数値に変換"
        Case Else: ReturnedVal = "Text to Number Conversion"
    End Select
End Sub

