function Invoke-TestHarness
{
    [CmdletBinding()]
    param
    (
        [Parameter()]
        [System.String]
        $TestResultsFile,

        [Parameter()]
        [System.String]
        $TestsPath,

        [Parameter()]
        [Switch]
        $IgnoreCodeCoverage
    )

    Write-Verbose -Message 'Commencing all SharePointDsc tests'

    <#$repoDir = Join-Path -Path $PSScriptRoot -ChildPath '..\' -Resolve

    $testCoverageFiles = @()
    if ($IgnoreCodeCoverage.IsPresent -eq $false)
    {
        Get-ChildItem -Path "$repoDir\PowerShell\Modules\PFE-SharePoint\Cmdlets\**\*.psm1" -Recurse | ForEach-Object {
            $testCoverageFiles += $_.FullName
        }
    }#>

    $testResultSettings = @{}
    if ([String]::IsNullOrEmpty($TestResultsFile) -eq $false) 
    {
        $testResultSettings.Add('OutputFormat', 'NUnitXml' )
        $testResultSettings.Add('OutputFile', $TestResultsFile)
    }

    <#$mainModule = $repoDir + "PowerShell\Modules\PFE-SharePoint\PFE-SharePoint.psd1"
    Import-Module $mainModule
    $testsToRun = @()

    # Run Unit Tests
    $versionsPath = Join-Path -Path $repoDir -ChildPath "\Tests\Unit\Stubs\SharePoint\"
    $versionsToTest = (Get-ChildItem -Path $versionsPath).Name
    # Import the first stub found so that there is a base module loaded before the tests start
    $firstVersion = $versionsToTest | Select-Object -First 1
    $firstStub = Join-Path -Path $repoDir `
                           -ChildPath "\Tests\Unit\Stubs\SharePoint\$firstVersion\Microsoft.SharePoint.PowerShell.psm1"
    Import-Module $firstStub -WarningAction SilentlyContinue

    $versionsToTest | ForEach-Object -Process {
        $stubPath = Join-Path -Path $repoDir `
                              -ChildPath "\Tests\Unit\Stubs\SharePoint\$_\Microsoft.SharePoint.PowerShell.psm1"
        $testsToRun += @(@{
            'Path' = (Join-Path -Path $repoDir -ChildPath "\Tests\Unit")
            'Parameters' = @{
                'SharePointStubsModule' = $stubPath
            }
        })
    }#>

    # Integration Tests (not run in appveyor due to time/reboots needed to install SharePoint)
    #$integrationTestsPath = Join-Path -Path $repoDir -ChildPath 'Tests\Integration'
    #$testsToRun += @( (Get-ChildItem -Path $integrationTestsPath -Filter '*.Tests.ps1').FullName )

    # DSC Common Tests
    <#if ($PSBoundParameters.ContainsKey('TestsPath') -eq $true)
    {
        $getChildItemParameters = @{
            Path = $TestsPath
            Recurse = $true
            Filter = '*.Tests.ps1'
        }

        # Get all tests '*.Tests.ps1'.
        $commonTestFiles = Get-ChildItem @getChildItemParameters

        $testsToRun += @( $commonTestFiles.FullName )
    }#>

    if ($IgnoreCodeCoverage.IsPresent -eq $false)
    {
        $testResultSettings.Add('CodeCoverage', $testCoverageFiles)
    }

    $results = Invoke-Pester -Script $testsToRun -PassThru @testResultSettings

    return $results
}
