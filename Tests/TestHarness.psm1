function Invoke-TestHarness
{
    [CmdletBinding()]
    param
    (
        [Parameter(Mandatory = $false)]
        [System.String]
        $TestResultsFile,

        [Parameter(Mandatory = $false)]
        [System.String]
        $DscTestsPath,

        [Parameter(Mandatory = $false)]
        [Switch]
        $IgnoreCodeCoverage
    )

    Write-Verbose -Message 'Starting All Tests'

    return $true
}