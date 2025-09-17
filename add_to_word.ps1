param (
    [string]$text,
    [string]$imagePath
)

# ���J System.Drawing �{���w�H�ϥ��C��\��
Add-Type -AssemblyName System.Drawing

# �إߩΨ��o Word COM ����
try {
    $word = [Runtime.InteropServices.Marshal]::GetActiveObject('Word.Application')
} catch {
    $word = New-Object -ComObject Word.Application
    $word.Visible = $true
    # �p�G�O�s�W�����A�o�̥i�H�[�J�o��{���X
    $word.Documents.Add() | Out-Null
}
# �T�O���B�󬡰ʪ��A�A�H�K���
if ($word.Documents.Count -eq 0) {
    $word.Documents.Add() | Out-Null
}

$activeDocument = $word.ActiveDocument
$selection = $word.Selection

Write-Host "Starting Word automation  $selection ...$text"


# ----------------- �s�W���{���X -----------------

# ���ʿ����Ш��󵲧�
# �T�O�s���e���J�b�̤U��
$word.Selection.EndKey([Microsoft.Office.Interop.Word.WdUnits]::wdStory)
# $word.Selection.TypeParagraph() # ���J�@�ӷs�q���A�קK�M�e�����e�H�b�@�_

# ���J�@�� 1 �� 3 �檺���
$range = $word.Selection.Range
$table = $word.Selection.Tables.Add($range, 1, 3)

# �]�m��檺���
$table.Borders.Enable = $true

# �]�m�C�@�檺���e
# ��1��G���J��r
$table.Cell(1, 1).Range.Text = $text # �ϥζǤJ���Ѽ�

# ��2��G���J�Ϥ�
try {
    if (-not (Test-Path $imagePath)) {
        throw "�Ϥ��ɮפ��s�b�G$imagePath"
    }
    # �T�O�Ϥ����J�b�x�s�椺
    $table.Cell(1, 2).Range.InlineShapes.AddPicture($imagePath)
} catch {
    Write-Error "�L�k���J�Ϥ��G$($_.Exception.Message)"
    # �p�G�Ϥ����J���ѡA�i�H�b�o�洡�J�@�Ӵ��ܤ�r
    $table.Cell(1, 2).Range.Text = "�Ϥ����J����"
    exit 1
}

# ��3��G���J��r "���\"
$table.Cell(1, 3).Range.Text = "���\"

# �վ���e�ץH�A������
$table.AutoFitBehavior([Microsoft.Office.Interop.Word.WdAutoFitBehavior]::wdAutoFitWindow)

# ----------------- �����s�W -----------------

# �i��G�O�s���
# $doc.SaveAs("C:\Users\YourUsername\Documents\MyDocument.docx")
# Write-Host "���w�O�s"

# ���n�ߧY���� Word�A���ϥΪ̥i�H�ݨ쵲�G
# $doc.Close()
# $word.Quit()