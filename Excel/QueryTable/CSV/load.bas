public sub main() ' {

    dim csv_file_name as string
    csv_file_name = thisWorkbook.path & chr(92) & "data.csv" ' chr(92) is backslash

    call importCSV(csv_file_name := csv_file_name              , _
                 sheet_        := activeSheet                , _
                 range_        := activeSheet.Range("$A$1")  , _
                 name_         :="CSVData" )

    activeWorkbook.saved = true

end sub ' }

private sub importCSV(csv_file_name as string, sheet_ as workSheet, range_ as range, name_ as string) ' {

      With activeSheet.queryTables.Add(                   _
               Connection    := "TEXT;" & csv_file_name , _
               Destination   := range_)

        .name                         = name_

        .fieldNames                   = true
        .rowNumbers                   = false
        .textFilePlatform             = 437
        .textFileStartRow             =   1
        .textFileParseType            = xlDelimited
        .textFileTextQualifier        = xlTextQualifierDoubleQuote
        .textFileConsecutiveDelimiter = false
        .preserveFormatting           = true
        .textFileCommaDelimiter       = true
        .preserveFormatting           = true
        .refreshOnFileOpen            = true
        .saveData                     = false

        .textFilePromptOnRefresh      = false
        .textFileTrailingMinusNumbers = true
'       .textFileTabDelimiter         = false
'       .textFileSemicolonDelimiter   = false
'       .textFileSpaceDelimiter       = false
'       .textFileColumnDataTypes      = Array(1, 1, 1, 1, 1, 1, 1, 1, 1, 1, 1)

'       .refreshStyle                 = xlInsertDeleteCells
'       .adjustColumnWidth            = True
'       .refreshPeriod                = 0

        .refresh backgroundQuery     := false

      end with

end sub
