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