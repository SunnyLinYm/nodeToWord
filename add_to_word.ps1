param (
    [string]$text,
    [string]$imagePath
)

# 載入 System.Drawing 程式庫以使用顏色功能
Add-Type -AssemblyName System.Drawing

# 建立或取得 Word COM 物件
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # 如果是新增的文件，這裡可以加入這行程式碼
    $word.Documents.Add() | Out-Null
}
# 確保文件處於活動狀態，以便選擇
if ($word.Documents.Count -eq 0) {
    $word.Documents.Add() | Out-Null
}

$activeDocument = $word.ActiveDocument
$selection = $word.Selection

Write-Host "Starting Word automation  $selection ...$text"


# ----------------- 新增的程式碼 -----------------

# 移動選取游標到文件結尾
# 確保新內容插入在最下面
$word.Selection.EndKey([Microsoft.Office.Interop.Word.WdUnits]::wdStory)
# $word.Selection.TypeParagraph() # 插入一個新段落，避免和前面內容黏在一起

# 插入一個 1 行 3 欄的表格
$range = $word.Selection.Range
$table = $word.Selection.Tables.Add($range, 1, 3)

# 設置表格的邊框
$table.Borders.Enable = $true

# 設置每一格的內容
# 第1格：插入文字
$table.Cell(1, 1).Range.Text = $text # 使用傳入的參數

# 第2格：插入圖片
try {
    if (-not (Test-Path $imagePath)) {
        throw "圖片檔案不存在：$imagePath"
    }
    # 確保圖片插入在儲存格內
    $table.Cell(1, 2).Range.InlineShapes.AddPicture($imagePath)
} catch {
    Write-Error "無法插入圖片：$($_.Exception.Message)"
    # 如果圖片插入失敗，可以在這格插入一個提示文字
    $table.Cell(1, 2).Range.Text = "圖片插入失敗"
    exit 1
}

# 第3格：插入文字 "成功"
$table.Cell(1, 3).Range.Text = "成功"

# 調整表格寬度以適應視窗
$table.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)

# ----------------- 結束新增 -----------------

# 可選：保存文件
# $doc.SaveAs("C:\Users\YourUsername\Documents\MyDocument.docx")
# Write-Host "文件已保存"

# 不要立即關閉 Word，讓使用者可以看到結果
# $doc.Close()
# $word.Quit()