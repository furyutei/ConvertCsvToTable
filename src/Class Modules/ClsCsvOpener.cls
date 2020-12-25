VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ClsCsvOpener"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents TargetExcelApplication As Excel.Application
Attribute TargetExcelApplication.VB_VarHelpID = -1

Private Const DebugIsValid As Boolean = False
Private Const TimeMeasurementIsValid As Boolean = False

Public Property Set ExcelApplication( _
    SpecifiedApplication As Excel.Application _
)
    Set TargetExcelApplication = SpecifiedApplication
End Property

Public Property Get ExcelApplication() As Excel.Application
    Set ExcelApplication = TargetExcelApplication
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

Public Function IsCsvFile( _
    Filepath As String _
) As Boolean
    IsCsvFile = False
    If Filepath = vbNullString Then Exit Function
    If LCase(Right(Filepath, 4)) = ".csv" Then IsCsvFile = True
End Function

Public Function ReadCsvText( _
    Filepath As String, _
    Optional ByVal FileCharset As String = "AUTO" _
) As String
    Dim FileNumber As Long
    Dim ByteBuffer() As Byte
    Dim ByteLength As Long
    Dim TextBuffer As String
    Dim AdodbStream As Object
    Dim ResultByteBuffer() As Byte
    Dim Counter As Long

    ' ファイルをバイナリとしてバイト配列に読み込み
    FileNumber = FreeFile
    Open Filepath For Binary As #FileNumber
    ByteLength = LOF(FileNumber)
    ReDim ByteBuffer(0 To ByteLength - 1)
    Get #FileNumber, 1, ByteBuffer
    Close #FileNumber
    
    If FileCharset = "AUTO" And 3 <= ByteLength Then
        ' BOM 付のものは UTF-8 とみなす
        If ByteBuffer(0) = &HEF And ByteBuffer(1) = &HBB And ByteBuffer(2) = &HBF Then
            FileCharset = "utf-8"
        End If
    End If

    Set AdodbStream = CreateObject("ADODB.Stream")
    With AdodbStream
        .Mode = 3 ' adModeRead(2), adModeReadWrite(3)
        
        .Open
        
        '.Type = 2 ' adTypeBinary(1), adTypeText(2)
        '.Charset = "UTF-8"
        '.LoadFromFile Filepath
        '→ .LoadFromFile だと、開いているファイルは読み込めず
        '  「ファイルを開けませんでした(3002)(adErrOpeningFile)」（.Mode=3の場合）
        '  「このコンテキストで操作は許可されていません(3219)(adErrIllegalOperation)」（.Mode=2の場合）
        '   のようなエラーとなってしまう
        '→一度バイナリで読み込んだものを ADODB.Stream に書き込み、これをテキストとして読み出す
        
        .Type = 1 ' adTypeBinary(1), adTypeText(2)
        .Write ByteBuffer
        
        .Position = 0
        .Type = 2 ' adTypeBinary(1), adTypeText(2)
        .Charset = IIf(FileCharset = "AUTO", "_autodetect_all", FileCharset) ' "Unicode"(default:UTF-16), "Shift-JIS", "UTF-8", "_autodetect_all",
        
        TextBuffer = .ReadText(-1) 'adReadAll(-1) adReadLine(-2)
        .Close
    End With

    TextBuffer = Replace(Replace(TextBuffer, vbCrLf, vbLf), vbCr, vbLf) ' 改行を LF に統一

    Dim Reg As Object: Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Global = True
        .MultiLine = False

        ' UTF-16 BOM(&HFF, &HFE) がついている場合・NULL 文字が含まれている場合は取り除く
        ' ※ファイルの最後に NULL 文字があると、Power Query の Binary.InferContentType 等が誤動作する模様
        .Pattern = "(^\ufeff|\u0000+)"
        TextBuffer = .Replace(TextBuffer, "")

        ' ファイルの最後が改行の場合はこれを取り除く
        ' ※ファイルの最後に改行があると、余分な行ができてしまうため
        .Pattern = "(\n+$)"
        TextBuffer = .Replace(TextBuffer, "")
    End With
 
    If DebugIsValid Then
        Debug.Print Left(TextBuffer, 100)
        Debug.Print Right(TextBuffer, 100)
        Debug.Print UBound(ByteBuffer) & " vs " & LenB(TextBuffer)
    End If
    
    ReadCsvText = TextBuffer
End Function

Public Sub AnalyzeCsvText( _
    CsvText As String, _
    Optional ByRef MaxRowNumber As Long, _
    Optional ByRef MaxColumnNumber As Long, _
    Optional ByRef PersumedHeaderRow As Long, _
    Optional ByRef PersumedColumnAttributes As Variant, _
    Optional ByVal SearchHeaderRowLimit As Long = 30 _
)
    Dim Reg As Object
    Dim TextBuffer As String
    Dim TextLines As Variant
    Dim TempLine As Variant
    Dim TempColumnNumber As Long
    Dim TempFilledNumber As Long
    Dim PresumedHeaderColumnNumber As Long
    Dim ColumnAttribute() As String
    Dim RowNumber As Long
    Dim ColumnNumber As Long
    Dim ColumnMarks As Variant
    Dim ColumnMark As Variant
    Dim StringColumnCounter As Long
    
    TextBuffer = CsvText
    'If Right(TextBuffer, 1) <> vbLf Then TextBuffer = TextBuffer & vbLf
    
    Set Reg = CreateObject("VBScript.RegExp")
    With Reg
        .Global = True
        .MultiLine = False

        ' "～" で括られた一文字以上のデータを S に置換
        .Pattern = "[\x20\f\r\t\v]*""(?:""""|[^""])+""[\x20\f\r\t\v]*([,\n]|$)"
        TextBuffer = .Replace(TextBuffer, "S$1")

        ' 余分な " 及び空白を削除
        .Pattern = "[""\x20\f\r\t\v]+"
        TextBuffer = .Replace(TextBuffer, "")
        
        ' 文字 N を S に置換
        TextBuffer = Replace(TextBuffer, "N", "S")

        ' 数値を N に置換
        .Pattern = "[\-+]?\d+\.?\d*(?:E[\-+]?\d+)?([,\n]|$)"
        TextBuffer = .Replace(TextBuffer, "N$1")

        ' 一文字以上のデータを S に置換
       .Pattern = "(?:[^,\nNS]|[^,\n]{2,})([,\n]|$)"
        TextBuffer = .Replace(TextBuffer, "S$1")
    End With
    
    MaxRowNumber = 0
    MaxColumnNumber = 0
    PersumedHeaderRow = 0
    PresumedHeaderColumnNumber = 0
    
    Reg.Pattern = "[NS]"
    TextLines = Split(TextBuffer, vbLf)
    For Each TempLine In TextLines
        MaxRowNumber = MaxRowNumber + 1
        
        ' 行内のデータ数（カラム数）を取得
        TempColumnNumber = Len(TempLine) - Len(Replace(TempLine, ",", "")) + 1
        If MaxColumnNumber < TempColumnNumber Then
            MaxColumnNumber = TempColumnNumber
        End If

        If MaxRowNumber <= SearchHeaderRowLimit Then
            ' 行内の有効データ数（空データ以外）を取得
            TempFilledNumber = Len(TempLine) - Len(Reg.Replace(TempLine, ""))
            If (PresumedHeaderColumnNumber < MaxColumnNumber) And (TempFilledNumber = MaxColumnNumber) Then
                ' 全てのカラムが有効データ（空データ以外）である最初の行をヘッダ行とみなす
                PresumedHeaderColumnNumber = MaxColumnNumber
                PersumedHeaderRow = MaxRowNumber
            End If
        End If
    Next TempLine
    
    If PresumedHeaderColumnNumber < MaxColumnNumber Then
        ' ヘッダとみなした行のデータ数が最終的な最大データ数に満たない場合にはヘッダ行無しとする
        PersumedHeaderRow = 0
    End If

    If 0 < MaxColumnNumber Then
        ReDim ColumnAttribute(1 To MaxColumnNumber)

        StringColumnCounter = 0

        For RowNumber = PersumedHeaderRow + 1 To MaxRowNumber
            TempLine = TextLines(RowNumber - 1)
            ColumnMarks = Split(TempLine, ",")
            ColumnNumber = 0
            For Each ColumnMark In ColumnMarks
                ColumnNumber = ColumnNumber + 1
                Select Case ColumnMark
                    Case "N"
                        If ColumnAttribute(ColumnNumber) = "" Then
                            ColumnAttribute(ColumnNumber) = "Number"
                        End If

                    Case "S"
                        If ColumnAttribute(ColumnNumber) <> "String" Then
                            ColumnAttribute(ColumnNumber) = "String"
                            StringColumnCounter = StringColumnCounter + 1
                            If MaxColumnNumber <= StringColumnCounter Then
                                GoTo EXIT_COLUMN_ATTRIBUTES_LOOP
                            End If
                        End If
                End Select
            Next ColumnMark
        Next RowNumber
EXIT_COLUMN_ATTRIBUTES_LOOP:
        PersumedColumnAttributes = ColumnAttribute
    End If
End Sub


Public Function CreateTextFile( _
    TextContent As String, _
    Optional Filepath As String = vbNullString, _
    Optional FileCharset As String = "UTF-8", _
    Optional WithoutBOM As Boolean = True _
) As String
    Dim Fso As Object
    Dim Ts As Object
    Dim AdodbStream As Object

    Set Fso = CreateObject("Scripting.FileSystemObject")
    With Fso
        If Filepath = vbNullString Then: Filepath = .BuildPath(.GetSpecialFolder(2), .GetTempName & ".csv")
        Filepath = .GetAbsolutePathName(Filepath)
        FileCharset = LCase(FileCharset)

        If WithoutBOM Then
            Select Case FileCharset
                Case "utf-8", "utf8"
                    WithoutBOM = True
                Case Else
                    WithoutBOM = False
            End Select
        End If

        Select Case FileCharset
            Case "unicode", "utf-16", "utf16"
                Set Ts = .CreateTextFile(Filename:=Filepath, Overwrite:=True, Unicode:=True)
                With Ts
                    .Write TextContent
                    .Close
                End With
            
            Case Else
                Set AdodbStream = CreateObject("ADODB.Stream")
                With AdodbStream
                    .Mode = 3 ' adModeRead(2), adModeReadWrite(3)
                    
                    .Open
                    
                    .Type = 2 ' adTypeBinary(1), adTypeText(2)
                    .Charset = FileCharset ' "Unicode"(default:UTF-16), "Shift-JIS", "UTF-8", "_autodetect_all",
                    
                    .WriteText TextContent, 0 ' adWriteChar(0), adWriteLine(1)

                    If WithoutBOM And 0 < LenB(TextContent) Then
                        ' 覚書：BOM付ファイルを PowerQuery で処理すると（Csv.Document()等）誤動作（頭にゴミが付いた状態になる）
                        Dim ByteBuffer() As Byte

                        .Position = 0
                        .Type = 1 ' adTypeBinary(1), adTypeText(2)
                        .Position = 3
                        ByteBuffer = .Read()
                        .Position = 0
                        .Close

                        Set AdodbStream = CreateObject("ADODB.Stream")
                        With AdodbStream
                            .Mode = 3 ' adModeRead(2), adModeReadWrite(3)
                            .Open
                            .Type = 1 ' adTypeBinary(1), adTypeText(2)
                            .Write ByteBuffer
                            .SaveToFile Filepath, 2 ' adSaveCreateNotExist(1), adSaveCreateOverWrite(2)
                            .Close
                        End With
                    Else
                        .SaveToFile Filepath, 2 ' adSaveCreateNotExist(1), adSaveCreateOverWrite(2)
                        ' 覚書：adTypeTextの際には（BOMがありうるフォーマットならば）BOM付きで保存される
                        .Close
                    End If
                End With
        End Select
    End With

    CreateTextFile = Filepath
End Function


Public Function ImportCsvFromFile( _
    Filepath As String, _
    Optional BaseSheet As Worksheet, _
    Optional FileCharset As String = "UTF-8", _
    Optional MaxColumnNumber As Long = -1, _
    Optional HeaderRow As Long = -1, _
    Optional SearchHeaderRowLimit As Long = 30, _
    Optional ColumnAttributes As Variant, _
    Optional AdjustCellSize As Boolean = True, _
    Optional MaxColumnWidth As Long = 30, _
    Optional TableStyleName As String = "TableStyleMedium7", _
    Optional ByRef ExtraDataRange As Range _
) As Range
    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    Dim TargetRange As Range
    Dim CsvCodePage As Long
    Dim QueryMformula As String
    Dim TargetQuery As WorkbookQuery
    Dim TargetQueryTable As QueryTable
    Dim TargetConnection As WorkbookConnection
    Dim TargetTable As ListObject
    Dim TempSheet As Worksheet
    Dim TempString As String
    Dim TimeMeasurement As ClsTimeMeasurement: Set TimeMeasurement = New ClsTimeMeasurement: TimeMeasurement.TimeMeasurementIsValid = TimeMeasurementIsValid

    If MaxColumnNumber = 0 Then
        Set ImportCsvFromFile = Nothing
        Set ExtraDataRange = Nothing
        Exit Function
    End If

    If BaseSheet Is Nothing Then
        'Set BaseSheet = ActiveWorkbook.ActiveSheet
        Set BaseSheet = ActiveWorkbook.Worksheets(ActiveWorkbook.Worksheets.Count)
        ' TODO: 便宜上、一番右端にあるシートを基準（元となるCSV）としているが、並び変えられていたら誤動作しうる
    End If
    Set TargetBook = BaseSheet.Parent

'   ※覚書：Query は残さないように変更したため、これでは判定できなくなった
'    On Error Resume Next
'    Set TargetQuery = TargetBook.Queries("CSVTable")
'    On Error GoTo 0
'
'    If Not TargetQuery Is Nothing Then
'        Set ImportCsvFromFile = Nothing
'        Set ExtraDataRange = Nothing
'        Exit Function
'    End If

    For Each TempSheet In TargetBook.Worksheets
        Select Case Left(TempSheet.Name, 1)
            Case "#", "@"
                TempString = Right(TempSheet.Name, Len(TempSheet.Name) - 1)
                If TempString = Left(BaseSheet.Name, Len(TempString)) Then
                    ' 変換済みシートが存在する場合は終了
                    ' TODO: シート名が変更されていたり基準シートの選別が誤っていたら誤動作しうる
                    Set ImportCsvFromFile = Nothing
                    Set ExtraDataRange = Nothing
                    Exit Function
                End If
        End Select
    Next TempSheet
    
    Set TargetSheet = CreateCsvSheet(BaseSheet:=BaseSheet)
    Set TargetRange = TargetSheet.Range("$A$1")

    Set TargetBook = TargetSheet.Parent
    Set TargetRange = TargetRange.Resize(1, 1)
    
    Select Case LCase(FileCharset)
        Case "shift-jis", "shift_jis"
            CsvCodePage = 932
        Case "utf-8", "utf8"
            CsvCodePage = 65001
        Case Else
            CsvCodePage = 1200
    End Select

' ※覚書: Connection:="TEXT;<Filepath>" を使用した方法だと、セル内改行がうまく処理されない模様
'    Set TargetQueryTable = TargetSheet.QueryTables.Add( _
'        Connection:="TEXT;" & Filepath, _
'        Destination:=TargetRange _
'    )
'
'    With TargetQueryTable
'        On Error Resume Next
'        .TextFilePlatform = CsvCodePage
'        ' TODO: 1200 (UTF-16LE) を指定するとエラーになってしまう
'        On Error GoTo 0
'        .TextFileParseType = xlDelimited
'        .TextFileCommaDelimiter = True
'        .RefreshStyle = xlOverwriteCells
'        .AdjustColumnWidth = True
'        .TextFileTextQualifier = xlTextQualifierDoubleQuote
'        .Refresh
'        Set TargetRange = .ResultRange
'        .Delete
'    End With
'
'    MaxColumnNumber = TargetRange.Columns.Count
'
'    TargetSheet.Cells.Delete
'    Set TargetRange = TargetSheet.Range("$A$1")

    TimeMeasurement.GetLapTime Header:="  Pretreatment: "

    If MaxColumnNumber < 0 Then
        ' 最大データ数（カラム数）が指定されていない場合は自動取得（ただし時間がかかる）
        QueryMformula = _
            "let" & _
            "   Source = Binary.InferContentType(" & _
            "       File.Contents(""" & Filepath & """)" & _
            "   )," & _
            "   CsvPotentialDelimiters = Source[Csv.PotentialDelimiters]," & _
            "   FiltersRows = Table.SelectRows(CsvPotentialDelimiters, each ([PotentialDelimiter] = "","") and ([QuoteStyle] = 0))," & _
            "   MaxColumns = FiltersRows{0}[MaxColumns]" & _
            "in MaxColumns"
    
        Set TargetQuery = TargetBook.Queries.Add("MaxColumns", QueryMformula)
    
        Set TargetTable = TargetSheet.ListObjects.Add( _
            SourceType:=xlSrcExternal, _
            Source:= _
                "OLEDB;" & _
                "Provider=Microsoft.Mashup.OleDb.1;" & _
                "Data Source=$Workbook$;" & _
                "Location=MaxColumns;" & _
                "Extended Properties=""""", _
            xlListObjectHasHeaders:=xlNo, _
            Destination:=TargetRange _
        )
    
        Set TargetQueryTable = TargetTable.QueryTable
    
        With TargetQueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [MaxColumns]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = False
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = False
            .AdjustColumnWidth = False
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh
            .Delete
        End With

        For Each TargetConnection In TargetBook.Connections
            TargetConnection.Delete
        Next TargetConnection
    
        TargetQuery.Delete

        MaxColumnNumber = CLng(TargetRange.Offset(1, 0).Value)
    End If
    
    If DebugIsValid Then Debug.Print "MaxColumnNumber=" & CStr(MaxColumnNumber)

    TimeMeasurement.GetLapTime Header:="  Get MaxColumnNumber: "

    TargetSheet.Cells.Delete
    Set TargetRange = TargetSheet.Range("$A$1")

    ' 覚書： 先頭行にヘッダがあるとは限らないため、Table.PromoteHeaders(Source, [PromoteAllScalars=true]) によるヘッダ昇格は行わない
    QueryMformula = _
        "let" & _
        "   CSVTable = Csv.Document(" & _
        "       File.Contents(""" & Filepath & """)," & _
        "       [" & _
        "           Delimiter="",""," & _
        "           Encoding=" & CStr(CsvCodePage) & "," & _
        "           Columns=" & CStr(MaxColumnNumber) & "," & _
        "           QuoteStyle=QuoteStyle.Csv" & _
        "       ]" & _
        "   )" & _
        "in CSVTable"

    Set TargetQuery = TargetBook.Queries.Add("CSVTable", QueryMformula)

    Set TargetTable = TargetSheet.ListObjects.Add( _
        SourceType:=xlSrcExternal, _
        Source:= _
            "OLEDB;" & _
            "Provider=Microsoft.Mashup.OleDb.1;" & _
            "Data Source=$Workbook$;" & _
            "Location=CSVTable;" & _
            "Extended Properties=""""", _
        Destination:=TargetRange, _
        xlListObjectHasHeaders:=xlNo, _
        TableStyleName:=TableStyleName _
    )

    TargetTable.Name = "CSVTable0"

    Set TargetQueryTable = TargetTable.QueryTable
    
    With TargetQueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [CSVTable]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = False
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = False
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .ListObject.DisplayName = "CSVTable0"
        .Refresh

        Set TargetRange = .ResultRange

        .Delete
    End With

    For Each TargetConnection In TargetBook.Connections
        TargetConnection.Delete
    Next TargetConnection

    TargetQuery.Delete

    Set ImportCsvFromFile = TargetRange
    Set ExtraDataRange = Nothing

    TimeMeasurement.GetLapTime Header:="  Convert CSV to Table: "

    If TargetRange.Rows.Count <= 1 Then
        Exit Function
    End If

    Dim SourceSheet As Worksheet
    Dim SourceRange As Range
    Dim SourceValues As Variant
    Dim IndexX As Long
    Dim IndexY As Long
    Dim EmptyCellExists As Boolean

    Set SourceSheet = TargetSheet
    Set SourceRange = TargetRange.Offset(1, 0).Resize(TargetRange.Rows.Count - 1, TargetRange.Columns.Count)
    
    If HeaderRow < 0 Then
        ' ヘッダ行が指定されていない場合には自動検索
        If SourceRange.Rows.Count < SearchHeaderRowLimit Then
            SearchHeaderRowLimit = SourceRange.Rows.Count
        End If
    
        If SearchHeaderRowLimit < 1 Or SourceRange.Cells.Count < 2 Then
            Exit Function
        End If
        
        SourceValues = SourceRange.Resize(SearchHeaderRowLimit, SourceRange.Columns.Count)
        HeaderRow = 0
    
        For IndexY = LBound(SourceValues, 1) To UBound(SourceValues, 1)
            EmptyCellExists = False
            For IndexX = LBound(SourceValues, 2) To UBound(SourceValues, 2)
                If DebugIsValid Then Debug.Print IndexY, IndexX, SourceValues(IndexY, IndexX)
    
                If SourceValues(IndexY, IndexX) = "" Then
                    EmptyCellExists = True
                    Exit For
                End If
            Next IndexX
            If Not EmptyCellExists Then
                HeaderRow = 1 + IndexY - LBound(SourceValues, 1)
                Exit For
            End If
        Next IndexY
    End If

    TimeMeasurement.GetLapTime Header:="  Search Header: "

    If HeaderRow < 1 Then
        Exit Function
    End If

    Set SourceRange = SourceRange.Offset(HeaderRow - 1, 0).Resize(SourceRange.Rows.Count + 1 - HeaderRow, SourceRange.Columns.Count)

    Dim DataSheet As Worksheet
    Dim DataRange As Range
    Dim DataTable As ListObject
    Dim TempRange As Range

    Set DataSheet = CreateCsvSheet(BaseSheet:=BaseSheet, SheetNamePrefix:="#")
    
    SourceRange.Copy Destination:=DataSheet.Range("$A$1")
    Set DataRange = DataSheet.Range("$A$1").Resize(SourceRange.Rows.Count, SourceRange.Columns.Count)
    DataRange.ClearFormats

    Set DataTable = DataSheet.ListObjects.Add( _
        SourceType:=xlSrcRange, _
        Source:=DataRange, _
        xlListObjectHasHeaders:=xlYes, _
        TableStyleName:=TableStyleName _
    )
    DataTable.Name = "CSVTable"

    Set DataRange = DataTable.Range
    Set ImportCsvFromFile = DataRange

    TimeMeasurement.GetLapTime Header:="  Split Table: "

    If Not IsEmpty(ColumnAttributes) Then
        Dim ColumnAttribute As Variant
        Dim ColumnNumber As Long

        ColumnNumber = 0
        On Error Resume Next
        For Each ColumnAttribute In ColumnAttributes
            ColumnNumber = ColumnNumber + 1
            Select Case ColumnAttribute
                Case "Number"
                    DataRange.Columns(ColumnNumber).TextToColumns _
                        Destination:=DataRange.Columns(ColumnNumber), _
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
                Case "String"
                    'DataRange.Columns(ColumnNumber).NumberFormatLocal = "@"
            End Select
        Next ColumnAttribute
        On Error GoTo 0
    End If
    
    If AdjustCellSize Then
        With DataRange.EntireColumn
            .AutoFit
        End With

        With DataRange.Parent.Cells
            .WrapText = True
        End With
        
        If 0 < MaxColumnWidth Then
            For Each TempRange In DataRange.Columns
                If MaxColumnWidth < TempRange.ColumnWidth Then
                    TempRange.ColumnWidth = MaxColumnWidth
                End If
            Next TempRange
        End If

        With DataRange.EntireRow
            .AutoFit
        End With
    
    End If
    
    If HeaderRow = 1 Then
        SourceSheet.Delete
        Set ExtraDataRange = Nothing
    Else
        SourceRange.Delete
        Set ExtraDataRange = TargetTable.Range

        If AdjustCellSize Then
            With ExtraDataRange.EntireColumn
                .AutoFit
            End With
        
            With ExtraDataRange.Parent.Cells
                .WrapText = True
            End With
    
            If 0 < MaxColumnWidth Then
                For Each TempRange In ExtraDataRange.Columns
                    If MaxColumnWidth < TempRange.ColumnWidth Then
                        TempRange.ColumnWidth = MaxColumnWidth
                    End If
                Next TempRange
            End If
            
            With ExtraDataRange.EntireRow
                .AutoFit
            End With
        End If
    End If

    TimeMeasurement.GetLapTime Header:="  Adjust Cell Size: "
End Function


Public Function CreateCsvSheet( _
    Optional ByRef BaseSheet As Worksheet, _
    Optional SheetNamePrefix As String = "@" _
) As Worksheet
    If BaseSheet Is Nothing Then Set BaseSheet = ActiveWorkbook.ActiveSheet

    Dim TargetBook As Workbook
    Dim TargetSheet As Worksheet
    Dim TargetSheetName As String
    
    Set TargetBook = BaseSheet.Parent
    Set TargetSheet = TargetBook.Sheets.Add(Before:=BaseSheet)

    TargetSheetName = SheetNamePrefix & BaseSheet.Name
    On Error Resume Next
    TargetSheet.Name = TargetSheetName
    If TargetSheet.Name <> TargetSheetName Then
        TargetSheetName = Left(TargetSheetName, Len(TargetSheetName) - Len(SheetNamePrefix))
        TargetSheet.Name = TargetSheetName
    End If
    On Error GoTo 0
    
    Set CreateCsvSheet = TargetSheet
End Function


Public Sub ConvertCsv( _
    Optional TargetBook As Workbook, _
    Optional IsManual As Boolean = False _
)
    Dim TimeMeasurement As ClsTimeMeasurement: Set TimeMeasurement = New ClsTimeMeasurement: TimeMeasurement.TimeMeasurementIsValid = TimeMeasurementIsValid

    If TargetBook Is Nothing Then Set TargetBook = ActiveWorkbook

    Dim Filepath As String

    Filepath = vbNullString
    On Error Resume Next
    Filepath = TargetBook.FullName
    On Error GoTo 0

    If Not IsCsvFile(Filepath) Then
        If IsManual Then
            MsgBox IIf( _
                Application.International(xlCountryCode) = 81, _
                "CSVファイルではありません", _
                "It's not a CSV file !" _
            )
        End If
        Exit Sub
    End If
    If DebugIsValid Then Debug.Print "WorkbookOpen : " & Filepath

'    Dim SavedCalculation As Long
'    Dim SavedEnableEvents As Boolean
'    Dim SavedScreenUpdating As Boolean
'    Dim SavedDisplayAlerts As Boolean

    With Application
'       TODO：不具合等で Application.Calculation に不正な値が入れられた場合にはオーバーフローエラーになったりする
'           →値を保存して使うのはやめておく
'        SavedCalculation = .Calculation
'        SavedEnableEvents = .EnableEvents
'        SavedScreenUpdating = .ScreenUpdating
'        SavedDisplayAlerts = .DisplayAlerts

        If DebugIsValid Then
            ' デバッグ時はデフォルトのまま
        Else
            .Calculation = xlCalculationManual
            .EnableEvents = False
            .ScreenUpdating = False
            .DisplayAlerts = False
        End If

        .StatusBar = "Start: Converting CSV file to table"
    End With

    If DebugIsValid Then
        ' デバッグ時はエラーを発生させる
    Else
        On Error GoTo EXIT_SUB
    End If

    Dim CsvText As String
    Dim ConvertedCSVFilepath As String
    Dim DataRange As Range
    Dim ExtraDataRange As Range
    
    Application.StatusBar = "Reading CSV file ..."
    CsvText = ReadCsvText(Filepath)

    TimeMeasurement.GetLapTime Header:="ReadCsvText(): "

    If LenB(CsvText) < 1 Then
        GoTo EXIT_SUB
    End If

    Dim MaxRowNumber As Long
    Dim MaxColumnNumber As Long
    Dim PersumedHeaderRow As Long
    Dim PersumedColumnAttributes As Variant

    Application.StatusBar = "Analyzing CSV file ..."
    AnalyzeCsvText _
        CsvText:=CsvText, _
        MaxRowNumber:=MaxRowNumber, _
        MaxColumnNumber:=MaxColumnNumber, _
        PersumedHeaderRow:=PersumedHeaderRow, _
        PersumedColumnAttributes:=PersumedColumnAttributes
    
    If DebugIsValid Then Debug.Print "AnalyzeCsvText():", MaxRowNumber, MaxColumnNumber, PersumedHeaderRow

    TimeMeasurement.GetLapTime Header:="AnalyzeCsvText(): "

    Application.StatusBar = "Converting CSV file to UTF-8 ..."
    ConvertedCSVFilepath = CreateTextFile(CsvText)

    TimeMeasurement.GetLapTime Header:="CreateTextFile(): "

    If DebugIsValid Then Debug.Print "ConvertedCSVFilepath: " & ConvertedCSVFilepath

    Application.StatusBar = "Converting CSV file to tables ..."
    Set DataRange = ImportCsvFromFile( _
        ConvertedCSVFilepath, _
        MaxColumnNumber:=MaxColumnNumber, _
        HeaderRow:=PersumedHeaderRow, _
        ColumnAttributes:=PersumedColumnAttributes, _
        ExtraDataRange:=ExtraDataRange _
    )

    TimeMeasurement.GetLapTime Header:="ImportCsvFromFile(): "

    If DataRange Is Nothing Then
        Debug.Print "DataRange Is Nothing: It failed to convert or has already converted !"
        If IsManual Then
            MsgBox IIf( _
                Application.International(xlCountryCode) = 81, _
                "既に変換済みのようです", _
                "It failed to convert or has already converted !" _
            )
        End If
    Else
        If DebugIsValid Then
            Debug.Print "DataRange: " & CStr(DataRange.Rows.Count) & " x " & CStr(DataRange.Columns.Count)
            If Not ExtraDataRange Is Nothing Then
                Debug.Print "ExtraDataRange: " & CStr(ExtraDataRange.Rows.Count) & " x " & CStr(ExtraDataRange.Columns.Count)
            End If
        End If
    End If

    If DebugIsValid Then
        ' デバッグ用として変換した CSV ファイルを残しておく
    Else
        Kill ConvertedCSVFilepath
    End If

EXIT_SUB:
    On Error GoTo 0

    With Application
'        .Calculation = SavedCalculation
'        .EnableEvents = SavedEnableEvents
'        .ScreenUpdating = SavedScreenUpdating
'        .DisplayAlerts = SavedDisplayAlerts

        .Calculation = xlCalculationAutomatic
        .EnableEvents = True
        .ScreenUpdating = True
        .DisplayAlerts = True

        .StatusBar = False
    End With

    TimeMeasurement.GetTotalTime Header:="[Total] "
End Sub


Private Sub TargetExcelApplication_WorkbookOpen( _
    ByVal TargetBook As Workbook _
)
    ConvertCsv TargetBook
End Sub

