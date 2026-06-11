$ErrorActionPreference = 'Stop'
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
Import-Module "$PSScriptRoot/PptMergeCore.psm1" -Force

# ---- 設定の保存先 ----
$script:SettingsPath = Join-Path $env:APPDATA 'ppt-merge\settings.json'

function Get-AppSettings {
    if (Test-Path -LiteralPath $script:SettingsPath) {
        try {
            return Get-Content -LiteralPath $script:SettingsPath -Raw -Encoding UTF8 | ConvertFrom-Json
        } catch {
            return $null
        }
    }
    return $null
}

function Save-AppSettings {
    param(
        [string]$OutputFolder,
        [bool]$MakePdf
    )
    $dir = Split-Path $script:SettingsPath -Parent
    if (-not (Test-Path -LiteralPath $dir)) {
        New-Item -ItemType Directory -Path $dir -Force | Out-Null
    }
    $obj = [PSCustomObject]@{
        OutputFolder = $OutputFolder
        MakePdf      = $MakePdf
    }
    $obj | ConvertTo-Json | Set-Content -LiteralPath $script:SettingsPath -Encoding UTF8
}

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
$listBox.CheckOnClick = $false
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

# ---- イベント: ダブルクリックでチェック切替 ----
# CheckOnClick=$false のため、ファイル名クリックでは選択のみ。
# チェックボックスのクリック、またはダブルクリックでオン/オフを切り替える。
$listBox.Add_DoubleClick({
    $i = $listBox.SelectedIndex
    if ($i -ge 0) {
        $listBox.SetItemChecked($i, -not $listBox.GetItemChecked($i))
    }
})

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

# ---- 実行ボタン ----
$btnRun.Add_Click({
    # 対象ファイル(チェック済み・表示順)を収集
    $targets = @()
    for ($i = 0; $i -lt $listBox.Items.Count; $i++) {
        if ($listBox.GetItemChecked($i)) { $targets += $listBox.Items[$i].Path }
    }
    if ($targets.Count -eq 0) {
        [System.Windows.Forms.MessageBox]::Show('マージ対象が選択されていません。', '入力エラー')
        return
    }

    $outFolder = $txtFolder.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($outFolder)) {
        [System.Windows.Forms.MessageBox]::Show('出力先フォルダを入力してください。', '入力エラー')
        return
    }
    if (-not (Test-Path -LiteralPath $outFolder -PathType Container)) {
        # 直接入力されたパスがまだ存在しない場合は作成を確認
        $r = [System.Windows.Forms.MessageBox]::Show(
            "出力先フォルダが存在しません。作成しますか?`r`n$outFolder", '確認', 'YesNo', 'Question')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
        try {
            New-Item -ItemType Directory -Path $outFolder -Force | Out-Null
        } catch {
            [System.Windows.Forms.MessageBox]::Show("フォルダを作成できませんでした:`r`n$($_.Exception.Message)", 'エラー', 'OK', 'Error')
            return
        }
    }
    if (-not (Test-OutputName -Name $txtName.Text.Trim())) {
        [System.Windows.Forms.MessageBox]::Show('出力ファイル名が不正です(禁止文字や空欄)。', '入力エラー')
        return
    }

    $paths = Resolve-OutputPaths -FolderPath $outFolder -Name $txtName.Text.Trim()
    $makePdf = $chkPdf.Checked

    # 上書き確認
    $existsTargets = @()
    if (Test-Path -LiteralPath $paths.PptxPath) { $existsTargets += $paths.PptxPath }
    if ($makePdf -and (Test-Path -LiteralPath $paths.PdfPath)) { $existsTargets += $paths.PdfPath }
    if ($existsTargets.Count -gt 0) {
        $msg = "次のファイルを上書きします:`r`n" + ($existsTargets -join "`r`n") + "`r`n続行しますか?"
        $r = [System.Windows.Forms.MessageBox]::Show($msg, '上書き確認', 'YesNo', 'Warning')
        if ($r -ne [System.Windows.Forms.DialogResult]::Yes) { return }
    }

    # 設定保存(Task 8 で定義する Save-AppSettings を呼ぶ)
    Save-AppSettings -OutputFolder $outFolder -MakePdf $makePdf

    $btnRun.Enabled = $false
    $ppt = $null
    $pres = $null
    try {
        Write-Log 'PowerPoint を起動しています...'
        $ppt = New-Object -ComObject PowerPoint.Application
        $pres = $ppt.Presentations.Add($false)  # WithWindow = $false（非表示）

        $n = $targets.Count
        for ($i = 0; $i -lt $n; $i++) {
            $file = $targets[$i]
            Write-Log ("{0}/{1}: {2} を追加中..." -f ($i + 1), $n, (Split-Path $file -Leaf))
            # InsertFromFile(FileName, Index, [SlideStart], [SlideEnd])
            # Index は「この番号のスライドの後ろに挿入」。現在の枚数を渡すと末尾に追加。
            # SlideStart/SlideEnd を省略すると全スライドを挿入(元デザインを保持)。
            # ※ SlideEnd=0 は不正値となるため、引数自体を渡さないこと。
            [void]$pres.Slides.InsertFromFile($file, $pres.Slides.Count)
            [System.Windows.Forms.Application]::DoEvents()
        }

        Write-Log "保存しています: $($paths.PptxPath)"
        # 24 = ppSaveAsOpenXMLPresentation (.pptx)
        $pres.SaveAs($paths.PptxPath, 24)

        if ($makePdf) {
            Write-Log "PDF を作成しています: $($paths.PdfPath)"
            # 32 = ppSaveAsPDF
            $pres.SaveAs($paths.PdfPath, 32)
        }

        Write-Log '完了しました。'
        $doneMsg = "マージ完了:`r`n$($paths.PptxPath)"
        if ($makePdf) { $doneMsg += "`r`n$($paths.PdfPath)" }
        [System.Windows.Forms.MessageBox]::Show($doneMsg, '完了')
    }
    catch {
        Write-Log "エラー: $($_.Exception.Message)"
        [System.Windows.Forms.MessageBox]::Show("処理を中断しました:`r`n$($_.Exception.Message)", 'エラー', 'OK', 'Error')
    }
    finally {
        if ($pres -ne $null) {
            try { $pres.Close() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($pres)
        }
        if ($ppt -ne $null) {
            try { $ppt.Quit() } catch {}
            [void][System.Runtime.InteropServices.Marshal]::ReleaseComObject($ppt)
        }
        $pres = $null
        $ppt = $null
        [System.GC]::Collect()
        [System.GC]::WaitForPendingFinalizers()
        $btnRun.Enabled = $true
    }
})

# ---- 起動時: 設定復元 ----
$saved = Get-AppSettings
if ($null -ne $saved) {
    if ($saved.OutputFolder -and (Test-Path -LiteralPath $saved.OutputFolder -PathType Container)) {
        $txtFolder.Text = $saved.OutputFolder
    }
    if ($null -ne $saved.MakePdf) {
        $chkPdf.Checked = [bool]$saved.MakePdf
    }
}
# 出力先フォルダが未設定なら、デフォルトとしてデスクトップを入れておく(直接編集可)
if ([string]::IsNullOrWhiteSpace($txtFolder.Text)) {
    $txtFolder.Text = [System.Environment]::GetFolderPath('Desktop')
}

[void]$form.ShowDialog()
