Function FetchDataFromPDFs(pdfPaths As Variant) As Scripting.Dictionary
    Dim wb As Workbook, ws As Worksheet, currWB As Workbook
    Dim dict As Scripting.Dictionary            ' Add from references to use a dict (Optional)

    Set currWB = ThisWorkbook
    ' Create a new workbook to temporarily store extracted table's data
    Set wb = Workbooks.Add
    Set ws = wb.Worksheets("Sheet1")
    ws.Visible = False

    Set dict = New Scripting.Dictionary

    Dim pdfPath As Variant, sampleName As String, fileName As String

    For Each pdfPath In pdfPaths
        fileName = Right(pdfPath, Len(pdfPath) - InStrRev(pdfPath, "\"))
        sampleName = Split(fileName)(0)
        Set dict(sampleName) = New Scripting.Dictionary
        Set dict(sampleName) = LoadDataTables(pdfPath)  ' This function Fetches data from pdf  
    Next pdfPath

    wb.Close False
    currWB.Activate

    Set FetchDataFromPDFs = dict
End Function


Function LoadDataTables(filePath As Variant) As Scripting.Dictionary
    Dim Name As String, TableIds As Variant, TableId As String, fileName As String, idx As Long
    Dim dict As Scripting.Dictionary

    Debug.Print filePath

    Set dict = New Scripting.Dictionary             ' Store data from pdf into this dict
    Name = "Table Query"
    TableIds = GetPDFTablesIdList(filePath) ' This function fetches id's of all the tables present in the pdf.

    For idx = 0 To UBound(TableIds) - LBound(TableIds) - 1  ' Iterate over each table.
        TableId = TableIds(idx)
        fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))

        ' M script code (Power Query)
        ActiveWorkbook.Queries.Add Name:=Name, Formula:= _
            "let" & _
            "    Source = Pdf.Tables(File.Contents(""" & filePath & """), [Implementation=""1.3""])," & _
            "    " & TableId & " = Source{[Id=""" & TableId & """]}[Data]," & _
            "    #""Promoted Headers"" = Table.PromoteHeaders(" & TableId & ", [PromoteAllScalars=true])," & _
            "    #""Changed Type"" = Table.TransformColumnTypes(" & _
                    "#""Promoted Headers"", List.Transform(Table.ColumnNames(#""Promoted Headers""), each {_, type text})" & _
                ")" & _
            "in" & _
            "    #""Changed Type"""

        With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
            "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & Name & """;Extended Properties=""""" _
            , Destination:=Range("$A$1")).QueryTable
            .CommandType = xlCmdSql
            .CommandText = Array("SELECT * FROM [" & Name & "]")
            .RowNumbers = False
            .FillAdjacentFormulas = False
            .PreserveFormatting = True
            .RefreshOnFileOpen = False
            .BackgroundQuery = True
            .RefreshStyle = xlInsertDeleteCells
            .SavePassword = False
            .SaveData = True
            .AdjustColumnWidth = True
            .RefreshPeriod = 0
            .PreserveColumnInfo = True
            .Refresh BackgroundQuery:=False
        End With

        ''' Processing...
        Dim tbl As ListObject

        Set tbl = ActiveSheet.ListObjects(1)            ' This varibale contains the nth table data
        Dim i As Long
        For i = 0 To tbl.ListRows.count
            Dim j As Long
            For j = 1 To tbl.ListColumns.count
                ' Cell value of ith row and jth col.
                cellValue = tbl.DataBodyRange(i, j).Value
                dict(CStr(i) & " " & CStr(j)) = cellValue
            Next j
        Next i

        ' Remove current query
        Dim Qus As WorkbookQuery
        For Each Qus In ActiveWorkbook.Queries
          Qus.Delete
        Next
        ActiveSheet.Cells.Clear
    Next idx

    Set LoadDataTables = dict
End Function


Function GetPDFTablesIdList(filePath As Variant) As Variant
    Dim fileName As String, Name As String

    Name = "Dummy Name"
    fileName = Right(filePath, Len(filePath) - InStrRev(filePath, "\"))

    ActiveWorkbook.Queries.Add Name:=Name, Formula:= _
        "let" & _
        "    Source = Pdf.Tables(File.Contents(""" & filePath & """), [Implementation=""1.3""])," & _
        "    #""Filtered Rows"" = Table.SelectRows(Source, each ([Kind] = ""Table""))," & _
        "    #""Removed Columns"" = Table.RemoveColumns(#""Filtered Rows"",{""Kind"", ""Data""})" & _
        "in" & _
        "    #""Removed Columns"""

    With ActiveSheet.ListObjects.Add(SourceType:=0, Source:= _
        "OLEDB;Provider=Microsoft.Mashup.OleDb.1;Data Source=$Workbook$;Location=""" & Name & """;Extended Properties=""""" _
        , Destination:=Range("$A$1")).QueryTable
        .CommandType = xlCmdSql
        .CommandText = Array("SELECT * FROM [" & Name & "]")
        .RowNumbers = False
        .FillAdjacentFormulas = False
        .PreserveFormatting = True
        .RefreshOnFileOpen = False
        .BackgroundQuery = True
        .RefreshStyle = xlInsertDeleteCells
        .SavePassword = False
        .SaveData = True
        .AdjustColumnWidth = True
        .RefreshPeriod = 0
        .PreserveColumnInfo = True
        .Refresh BackgroundQuery:=False
    End With

    Dim tbl As ListObject
    Set tbl = ActiveSheet.ListObjects(1)

    Dim TableIds() As Variant
    ReDim TableIds(tbl.ListRows.count)
    For i = 0 To tbl.ListRows.count
        TableIds(i) = tbl.DataBodyRange(i + 1, 1).Value
    Next i

    ' Remove current query
    Dim Qus As WorkbookQuery
    For Each Qus In ActiveWorkbook.Queries
      Qus.Delete
    Next
    ActiveSheet.Cells.Clear

    GetPDFTablesIdList = TableIds
End Function

