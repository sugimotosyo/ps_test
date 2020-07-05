#https://qiita.com/S_SenSq/items/6bae890dbeb47385ea02


# Excelのテーブルをオブジェクトとして取得する関数
## $sheet: シートオブジェクト
## $start: 表の左上のセル番号（A1形式での文字列）
## $end:   表の右下のセル番号（A1形式での文字列）
Function Get-TableObject($sheet, $start, $end) {
    # テーブルオブジェクト作成
    $table = [pscustomobject]@{
        start = $sheet.Range($start)
        end   = $sheet.Range($end)
        key   = @()
        data  = @()
    }
    # 古いPowershellだと以下のような書き方をしないとエラーになる
    #$table = New-Object PSCustomObject
    #$table | Add-Member -MemberType NoteProperty -Name start -Value $sheet.Range($start)

    # テーブルをオブジェクト化
    for ($row = $table.start.Row; $row -le $table.end.Row; $row++) {
        # 最初にデータの無い行は無視
        if ($sheet.cells.item($row, $table.start.Column).text -eq "" ) {
            continue
        }
        # 1レコード用オブジェクトを準備
        $record = New-Object PSCustomObject
        $key_ref_number = 0
        for ($col = $table.start.Column; $col -le $table.end.Column; $col++) {
            # 最初にデータの無い列は無視
            if ($sheet.cells.item($table.start.Row, $col).text -eq "" ) {
                continue
            }
            # 1行目の値からキー名作成
            if ($row -eq $table.start.Row) {
                $table.key += $sheet.cells.item($row, $col).text
            }
            # 1レコード作成
            else {
                $key_name = ($table.key[$key_ref_number])
                $val = ($sheet.cells.item($row, $col).text)
                $record | Add-Member -MemberType NoteProperty -Name $key_name -Value $val
                $key_ref_number += 1
            }
        }
        # 1レコード追加
        if ($row -gt $table.start.Row) {
            $table.data += $record
        }
    }
    return $table
}

############################
# 定数
############################
$START_CELL = "A1"
$END_CELL = "G100"
$EXCEL_PATH_BASE = "" # 末尾に\が必要
$EXCEL_FILE_NAME_BASE = "scenario"
$EXCEL_FILE_NAME_EXTENSION = "xlsx"



############################
# 処理
############################

## エクセルファイルを取得
$excelFiles = Get-ChildItem -File $EXCEL_FILE_NAME_BASE*.$EXCEL_FILE_NAME_EXTENSION
$excelFiles.Name[0]


for ($i = 0; $i -lt $excelFiles.Name.Length ; $i++ ) {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible = $true

    $book = $excel.Workbooks.Open($EXCEL_PATH_BASE + $excelFiles.Name[$i])

    $sheet = $book.Sheets("Sheet1")
    # 上記でSheetsメソッドが無いとエラーが出る場合は以下
    # $sheet = $book.Worksheets.Item($input.sheetname)


    $table = Get-TableObject $sheet $START_CELL $END_CELL

    $table.data

    #ブックを閉じる
    $book.Close()
    $excel.Quit()
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($sheet) | Out-Null
    [System.Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) | Out-Null


}


