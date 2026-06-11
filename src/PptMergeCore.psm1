function Test-OutputName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyString()]
        [string]$Name
    )
    if ([string]::IsNullOrWhiteSpace($Name)) { return $false }
    $invalid = '\/:*?"<>|'
    foreach ($ch in $invalid.ToCharArray()) {
        if ($Name.Contains($ch)) { return $false }
    }
    return $true
}

function Get-PptxFilesInFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$FolderPath
    )
    if (-not (Test-Path -LiteralPath $FolderPath -PathType Container)) {
        return @()
    }
    Get-ChildItem -LiteralPath $FolderPath -Filter '*.pptx' -File |
        Where-Object { -not $_.Name.StartsWith('~$') } |
        ForEach-Object { $_.FullName }
}

function Test-DuplicatePath {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [string[]]$ExistingPaths,
        [Parameter(Mandatory)]
        [string]$NewPath
    )
    foreach ($p in $ExistingPaths) {
        if ([string]::Equals($p, $NewPath, [StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }
    return $false
}

Export-ModuleMember -Function Test-OutputName, Get-PptxFilesInFolder, Test-DuplicatePath
