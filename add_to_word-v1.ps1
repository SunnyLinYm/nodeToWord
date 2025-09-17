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

# 插入文字
$selection.TypeText($text)
$selection.TypeParagraph() # 換行

Write-Host "TypeParagraph  $selection ...$text"

# 插入圖片
try {
    $selection.InlineShapes.AddPicture($imagePath)
    $selection.TypeParagraph()
} catch {
    Write-Error "無法插入圖片：$_"
}

# 可選：保存文件
# $doc.SaveAs("C:\Users\YourUsername\Documents\MyDocument.docx")
# Write-Host "文件已保存"

# 不要立即關閉 Word，讓使用者可以看到結果
# $doc.Close()
# $word.Quit()