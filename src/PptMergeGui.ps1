$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module "$PSScriptRoot/PptMergeCore.psm1" -Force

# ---- ログ用ヘルパ ----
function Write-Log {
    param([string]$Message)
    $ts = (Get-Date).ToString('HH:mm:ss')
    $script:logBox.AppendText("[$ts] $Message`r`n")
    [System.Windows.Forms.Application]::DoEvents()
}

# ---- フォーム ----
$form = New-Object System.Windows.Forms.Form
$form.Text = 'PPT Merge Tool  v0.1.0'
$form.Size = New-Object System.Drawing.Size(640, 640)
$form.StartPosition = 'CenterScreen'

# 上部ボタン
$btnAddFolder = New-Object System.Windows.Forms.Button
$btnAddFolder.Text = 'フォルダから一括追加'
$btnAddFolder.Location = New-Object System.Drawing.Point(12, 12)
$btnAddFolder.Size = New-Object System.Drawing.Size(160, 30)
$form.Controls.Add($btnAddFolder)

$btnAddFiles = New-Object System.Windows.Forms.Button
$btnAddFiles.Text = 'ファイルを個別追加'
$btnAddFiles.Location = New-Object System.Drawing.Point(180, 12)
$btnAddFiles.Size = New-Object System.Drawing.Size(160, 30)
$form.Controls.Add($btnAddFiles)

# ファイルリスト
$listBox = New-Object System.Windows.Forms.CheckedListBox
$listBox.Location = New-Object System.Drawing.Point(12, 52)
$listBox.Size = New-Object System.Drawing.Size(500, 220)
$listBox.CheckOnClick = $true
$listBox.AllowDrop = $true
$form.Controls.Add($listBox)

# 上下/削除/クリアボタン
$btnUp = New-Object System.Windows.Forms.Button
$btnUp.Text = '↑'
$btnUp.Location = New-Object System.Drawing.Point(522, 52)
$btnUp.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnUp)

$btnDown = New-Object System.Windows.Forms.Button
$btnDown.Text = '↓'
$btnDown.Location = New-Object System.Drawing.Point(522, 86)
$btnDown.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnDown)

$btnRemove = New-Object System.Windows.Forms.Button
$btnRemove.Text = '削除'
$btnRemove.Location = New-Object System.Drawing.Point(522, 130)
$btnRemove.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnRemove)

$btnClear = New-Object System.Windows.Forms.Button
$btnClear.Text = 'クリア'
$btnClear.Location = New-Object System.Drawing.Point(522, 164)
$btnClear.Size = New-Object System.Drawing.Size(90, 30)
$form.Controls.Add($btnClear)

# 出力先
$lblFolder = New-Object System.Windows.Forms.Label
$lblFolder.Text = '出力先フォルダ:'
$lblFolder.Location = New-Object System.Drawing.Point(12, 286)
$lblFolder.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblFolder)

$txtFolder = New-Object System.Windows.Forms.TextBox
$txtFolder.Location = New-Object System.Drawing.Point(120, 284)
$txtFolder.Size = New-Object System.Drawing.Size(390, 22)
$form.Controls.Add($txtFolder)

$btnBrowse = New-Object System.Windows.Forms.Button
$btnBrowse.Text = '参照...'
$btnBrowse.Location = New-Object System.Drawing.Point(522, 282)
$btnBrowse.Size = New-Object System.Drawing.Size(90, 26)
$form.Controls.Add($btnBrowse)

# 出力名
$lblName = New-Object System.Windows.Forms.Label
$lblName.Text = '出力ファイル名:'
$lblName.Location = New-Object System.Drawing.Point(12, 318)
$lblName.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblName)

$txtName = New-Object System.Windows.Forms.TextBox
$txtName.Location = New-Object System.Drawing.Point(120, 316)
$txtName.Size = New-Object System.Drawing.Size(250, 22)
$txtName.Text = 'merged_' + (Get-Date).ToString('yyyyMMdd')
$form.Controls.Add($txtName)

$lblExt = New-Object System.Windows.Forms.Label
$lblExt.Text = '(.pptx / .pdf)'
$lblExt.Location = New-Object System.Drawing.Point(376, 318)
$lblExt.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblExt)

# PDF チェック
$chkPdf = New-Object System.Windows.Forms.CheckBox
$chkPdf.Text = 'マージ後に PDF も作成する'
$chkPdf.Location = New-Object System.Drawing.Point(120, 346)
$chkPdf.Size = New-Object System.Drawing.Size(250, 24)
$chkPdf.Checked = $true
$form.Controls.Add($chkPdf)

# 実行ボタン
$btnRun = New-Object System.Windows.Forms.Button
$btnRun.Text = '▶ マージ実行'
$btnRun.Location = New-Object System.Drawing.Point(120, 378)
$btnRun.Size = New-Object System.Drawing.Size(390, 36)
$form.Controls.Add($btnRun)

# ログ
$lblLog = New-Object System.Windows.Forms.Label
$lblLog.Text = 'ログ:'
$lblLog.Location = New-Object System.Drawing.Point(12, 424)
$lblLog.Size = New-Object System.Drawing.Size(100, 20)
$form.Controls.Add($lblLog)

$logBox = New-Object System.Windows.Forms.TextBox
$logBox.Location = New-Object System.Drawing.Point(12, 446)
$logBox.Size = New-Object System.Drawing.Size(600, 140)
$logBox.Multiline = $true
$logBox.ReadOnly = $true
$logBox.ScrollBars = 'Vertical'
$form.Controls.Add($logBox)
$script:logBox = $logBox

# ---- リスト操作ヘルパ ----
function Add-FileToList {
    param([string]$Path)
    $existing = @($listBox.Items | ForEach-Object { $_.Path })
    if (Test-DuplicatePath -ExistingPaths $existing -NewPath $Path) {
        Write-Log "既にリストにあります: $(Split-Path $Path -Leaf)"
        return
    }
    $item = [PSCustomObject]@{ Path = $Path; Name = (Split-Path $Path -Leaf) }
    $idx = $listBox.Items.Add($item)
    $listBox.SetItemChecked($idx, $true)
}
$listBox.DisplayMember = 'Name'

# ---- イベント: フォルダ一括追加 ----
$btnAddFolder.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $files = Get-PptxFilesInFolder -FolderPath $dlg.SelectedPath
        if (@($files).Count -eq 0) {
            Write-Log "pptx が見つかりません: $($dlg.SelectedPath)"
        } else {
            foreach ($f in $files) { Add-FileToList -Path $f }
            Write-Log "$(@($files).Count) 件を読み込みました: $($dlg.SelectedPath)"
        }
    }
})

# ---- イベント: 個別ファイル追加 ----
$btnAddFiles.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'PowerPoint (*.pptx)|*.pptx'
    $dlg.Multiselect = $true
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        foreach ($f in $dlg.FileNames) { Add-FileToList -Path $f }
        Write-Log "$($dlg.FileNames.Count) 件を追加しました"
    }
})

# ---- イベント: 上移動 ----
$btnUp.Add_Click({
    $i = $listBox.SelectedIndex
    if ($i -gt 0) {
        $item = $listBox.Items[$i]
        $checked = $listBox.GetItemChecked($i)
        $listBox.Items.RemoveAt($i)
        $listBox.Items.Insert($i - 1, $item)
        $listBox.SetItemChecked($i - 1, $checked)
        $listBox.SelectedIndex = $i - 1
    }
})

# ---- イベント: 下移動 ----
$btnDown.Add_Click({
    $i = $listBox.SelectedIndex
    if ($i -ge 0 -and $i -lt $listBox.Items.Count - 1) {
        $item = $listBox.Items[$i]
        $checked = $listBox.GetItemChecked($i)
        $listBox.Items.RemoveAt($i)
        $listBox.Items.Insert($i + 1, $item)
        $listBox.SetItemChecked($i + 1, $checked)
        $listBox.SelectedIndex = $i + 1
    }
})

# ---- イベント: 削除 / クリア ----
$btnRemove.Add_Click({
    $i = $listBox.SelectedIndex
    if ($i -ge 0) { $listBox.Items.RemoveAt($i) }
})
$btnClear.Add_Click({ $listBox.Items.Clear() })

# ---- イベント: 参照 ----
$btnBrowse.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtFolder.Text = $dlg.SelectedPath
    }
})

# ---- ドラッグ&ドロップ ----
$dragEnter = {
    param($s, $e)
    if ($e.Data.GetDataPresent([System.Windows.Forms.DataFormats]::FileDrop)) {
        $e.Effect = [System.Windows.Forms.DragDropEffects]::Copy
    }
}
$dragDrop = {
    param($s, $e)
    $paths = $e.Data.GetData([System.Windows.Forms.DataFormats]::FileDrop)
    foreach ($p in $paths) {
        if (Test-Path -LiteralPath $p -PathType Container) {
            foreach ($f in (Get-PptxFilesInFolder -FolderPath $p)) { Add-FileToList -Path $f }
        } elseif ([System.IO.Path]::GetExtension($p) -ieq '.pptx') {
            Add-FileToList -Path $p
        }
    }
}
$listBox.Add_DragEnter($dragEnter)
$listBox.Add_DragDrop($dragDrop)
$form.AllowDrop = $true
$form.Add_DragEnter($dragEnter)
$form.Add_DragDrop($dragDrop)

# ---- 実行ボタン(Task 7 で実装) ----
$btnRun.Add_Click({
    Write-Log 'マージ実行は未実装です'
})

[void]$form.ShowDialog()
