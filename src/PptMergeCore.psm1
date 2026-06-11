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

Export-ModuleMember -Function Test-OutputName
