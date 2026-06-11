$ErrorActionPreference = 'Stop'
Import-Module "$PSScriptRoot/../src/PptMergeCore.psm1" -Force

Describe 'Test-OutputName' {
    It '通常の名前は有効' {
        Test-OutputName -Name 'merged_20260611' | Should -Be $true
    }
    It '空文字は無効' {
        Test-OutputName -Name '' | Should -Be $false
    }
    It '空白のみは無効' {
        Test-OutputName -Name '   ' | Should -Be $false
    }
    It '禁止文字を含む名前は無効' {
        Test-OutputName -Name 'bad/name' | Should -Be $false
        Test-OutputName -Name 'bad:name' | Should -Be $false
        Test-OutputName -Name 'bad*name' | Should -Be $false
        Test-OutputName -Name 'bad?name' | Should -Be $false
    }
    It '日本語の名前は有効' {
        Test-OutputName -Name '資料まとめ' | Should -Be $true
    }
}

Describe 'Get-PptxFilesInFolder' {
    BeforeAll {
        $script:tmp = Join-Path $TestDrive 'pptxfolder'
        New-Item -ItemType Directory -Path $script:tmp | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:tmp 'a.pptx') | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:tmp 'b.pptx') | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:tmp 'old.ppt') | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:tmp '~$lock.pptx') | Out-Null
        New-Item -ItemType File -Path (Join-Path $script:tmp 'note.txt') | Out-Null
    }
    It 'pptx のみを返す（ppt/txt は除外）' {
        $result = Get-PptxFilesInFolder -FolderPath $script:tmp
        ($result | ForEach-Object { Split-Path $_ -Leaf }) | Should -Contain 'a.pptx'
        ($result | ForEach-Object { Split-Path $_ -Leaf }) | Should -Contain 'b.pptx'
        ($result | ForEach-Object { Split-Path $_ -Leaf }) | Should -Not -Contain 'old.ppt'
        ($result | ForEach-Object { Split-Path $_ -Leaf }) | Should -Not -Contain 'note.txt'
    }
    It '一時ロックファイル(~$)を除外する' {
        $result = Get-PptxFilesInFolder -FolderPath $script:tmp
        ($result | ForEach-Object { Split-Path $_ -Leaf }) | Should -Not -Contain '~$lock.pptx'
    }
    It 'フルパスを返す' {
        $result = Get-PptxFilesInFolder -FolderPath $script:tmp
        $result[0] | Should -Match '^[A-Za-z]:\\'
    }
    It '存在しないフォルダは空配列を返す' {
        $result = @(Get-PptxFilesInFolder -FolderPath (Join-Path $TestDrive 'nope'))
        $result.Count | Should -Be 0
    }
}

Describe 'Test-DuplicatePath' {
    It '既存リストに同じパスがあれば $true' {
        $existing = @('C:\a\x.pptx', 'C:\a\y.pptx')
        Test-DuplicatePath -ExistingPaths $existing -NewPath 'C:\a\x.pptx' | Should -Be $true
    }
    It '大文字小文字を区別せず一致とみなす' {
        $existing = @('C:\a\X.pptx')
        Test-DuplicatePath -ExistingPaths $existing -NewPath 'c:\a\x.pptx' | Should -Be $true
    }
    It '存在しなければ $false' {
        $existing = @('C:\a\x.pptx')
        Test-DuplicatePath -ExistingPaths $existing -NewPath 'C:\a\z.pptx' | Should -Be $false
    }
    It '空リストなら $false' {
        Test-DuplicatePath -ExistingPaths @() -NewPath 'C:\a\z.pptx' | Should -Be $false
    }
}