﻿<#
    MIT License

    Copyright (c) Microsoft Corporation.

    Permission is hereby granted, free of charge, to any person obtaining a copy
    of this software and associated documentation files (the "Software"), to deal
    in the Software without restriction, including without limitation the rights
    to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
    copies of the Software, and to permit persons to whom the Software is
    furnished to do so, subject to the following conditions:

    The above copyright notice and this permission notice shall be included in all
    copies or substantial portions of the Software.

    THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
    IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
    FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
    AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
    LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
    OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
    SOFTWARE
#>

# Version 22.11.14.1812

<#
.NOTES
	Name: HealthChecker.ps1
	Requires: Exchange Management Shell and administrator rights on the target Exchange
	server as well as the local machine.
    Major Release History:
        4/20/2021  - Initial Public Release on CSS-Exchange
        11/10/2020 - Initial Public Release of version 3.
        1/18/2017 - Initial Public Release of version 2.
        3/30/2015 - Initial Public Release.

.SYNOPSIS
	Checks the target Exchange server for various configuration recommendations from the Exchange product group.
.DESCRIPTION
	This script checks the Exchange server for various configuration recommendations outlined in the
	"Exchange 2013 Performance Recommendations" section on Microsoft Docs, found here:

	https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help

	Informational items are reported in Grey.  Settings found to match the recommendations are
	reported in Green.  Warnings are reported in yellow.  Settings that can cause performance
	problems are reported in red.  Please note that most of these recommendations only apply to Exchange
	2013/2016.  The script will run against Exchange 2010/2007 but the output is more limited.
.PARAMETER Server
	This optional parameter allows the target Exchange server to be specified.  If it is not the
	local server is assumed.
.PARAMETER OutputFilePath
	This optional parameter allows an output directory to be specified.  If it is not the local
	directory is assumed.  This parameter must not end in a \.  To specify the folder "logs" on
	the root of the E: drive you would use "-OutputFilePath E:\logs", not "-OutputFilePath E:\logs\".
.PARAMETER MailboxReport
	This optional parameter gives a report of the number of active and passive databases and
	mailboxes on the server.
.PARAMETER LoadBalancingReport
    This optional parameter will check the connection count of the Default Web Site for every server
    running Exchange 2013/2016 with the Client Access role in the org.  It then breaks down servers by percentage to
    give you an idea of how well the load is being balanced.
.PARAMETER CasServerList
    Used with -LoadBalancingReport.  A comma separated list of CAS servers to operate against.  Without
    this switch the report will use all 2013/2016 Client Access servers in the organization.
.PARAMETER SiteName
	Used with -LoadBalancingReport.  Specifies a site to pull CAS servers from instead of querying every server
    in the organization.
.PARAMETER XMLDirectoryPath
    Used in combination with BuildHtmlServersReport switch for the location of the HealthChecker XML files for servers
    which you want to be included in the report. Default location is the current directory.
.PARAMETER BuildHtmlServersReport
    Switch to enable the script to build the HTML report for all the servers XML results in the XMLDirectoryPath location.
.PARAMETER HtmlReportFile
    Name of the HTML output file from the BuildHtmlServersReport. Default is ExchangeAllServersReport.html
.PARAMETER DCCoreRatio
    Gathers the Exchange to DC/GC Core ratio and displays the results in the current site that the script is running in.
.PARAMETER AnalyzeDataOnly
    Switch to analyze the existing HealthChecker XML files. The results are displayed on the screen and an HTML report is generated.
.PARAMETER SkipVersionCheck
    No version check is performed when this switch is used.
.PARAMETER SaveDebugLog
    The debug log is kept even if the script is executed successfully.
.PARAMETER ScriptUpdateOnly
    Switch to check for the latest version of the script and perform an auto update. No elevated permissions or EMS are required.
.PARAMETER Verbose
	This optional parameter enables verbose logging.
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME
	Run against a single remote Exchange server
.EXAMPLE
	.\HealthChecker.ps1 -Server SERVERNAME -MailboxReport -Verbose
	Run against a single remote Exchange server with verbose logging and mailbox report enabled.
.EXAMPLE
    Get-ExchangeServer | ?{$_.AdminDisplayVersion -Match "^Version 15"} | %{.\HealthChecker.ps1 -Server $_.Name}
    Run against all Exchange 2013/2016 servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport
    Run a load balancing report comparing all Exchange 2013/2016 CAS servers in the Organization.
.EXAMPLE
    .\HealthChecker.ps1 -LoadBalancingReport -CasServerList CAS01,CAS02,CAS03
    Run a load balancing report comparing servers named CAS01, CAS02, and CAS03.
.LINK
    https://docs.microsoft.com/en-us/exchange/exchange-2013-sizing-and-configuration-recommendations-exchange-2013-help
    https://docs.microsoft.com/en-us/exchange/exchange-2013-virtualization-exchange-2013-help#requirements-for-hardware-virtualization
    https://docs.microsoft.com/en-us/exchange/plan-and-deploy/virtualization?view=exchserver-2019#requirements-for-hardware-virtualization
#>
[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are being used')]
[CmdletBinding(DefaultParameterSetName = "HealthChecker", SupportsShouldProcess)]
param(
    [Parameter(Mandatory = $false, ParameterSetName = "HealthChecker")]
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [string]$Server = ($env:COMPUTERNAME),
    [Parameter(Mandatory = $false)]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$OutputFilePath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "MailboxReport")]
    [switch]$MailboxReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [switch]$LoadBalancingReport,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [array]$CasServerList = $null,
    [Parameter(Mandatory = $false, ParameterSetName = "LoadBalancingReport")]
    [string]$SiteName = ([string]::Empty),
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [ValidateScript( { -not $_.ToString().EndsWith('\') })][string]$XMLDirectoryPath = ".",
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [switch]$BuildHtmlServersReport,
    [Parameter(Mandatory = $false, ParameterSetName = "HTMLReport")]
    [string]$HtmlReportFile = "ExchangeAllServersReport.html",
    [Parameter(Mandatory = $false, ParameterSetName = "DCCoreReport")]
    [switch]$DCCoreRatio,
    [Parameter(Mandatory = $false, ParameterSetName = "AnalyzeDataOnly")]
    [switch]$AnalyzeDataOnly,
    [Parameter(Mandatory = $false)][switch]$SkipVersionCheck,
    [Parameter(Mandatory = $false)][switch]$SaveDebugLog,
    [Parameter(Mandatory = $false, ParameterSetName = "ScriptUpdateOnly")]
    [switch]$ScriptUpdateOnly
)

$BuildVersion = "22.11.14.1812"

$Script:VerboseEnabled = $false
#this is to set the verbose information to a different color
if ($PSBoundParameters["Verbose"]) {
    #Write verbose output in cyan since we already use yellow for warnings
    $Script:VerboseEnabled = $true
    $VerboseForeground = $Host.PrivateData.VerboseForegroundColor
    $Host.PrivateData.VerboseForegroundColor = "Cyan"
}



function Add-AnalyzedResultInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [HealthChecker.AnalyzedInformation]$AnalyzedInformation,
        [object]$Details,
        [string]$Name,
        [string]$TestingName,
        [object]$OutColumns,
        [scriptblock[]]$OutColumnsColorTests,
        [string]$HtmlName,
        [object]$DisplayGroupingKey,
        [int]$DisplayCustomTabNumber = -1,
        [object]$DisplayTestingValue,
        [string]$DisplayWriteType = "Grey",
        [bool]$AddDisplayResultsLineInfo = $true,
        [bool]$AddHtmlDetailRow = $true,
        [string]$HtmlDetailsCustomValue = "",
        [bool]$AddHtmlOverviewValues = $false,
        [bool]$AddHtmlActionRow = $false
        #[string]$ActionSettingClass = "",
        #[string]$ActionSettingValue,
        #[string]$ActionRecommendedDetailsClass = "",
        #[string]$ActionRecommendedDetailsValue,
        #[string]$ActionMoreInformationClass = "",
        #[string]$ActionMoreInformationValue,
    )
    begin {
        Write-Verbose "Calling $($MyInvocation.MyCommand): $name"
        function GetOutColumnsColorObject {
            param(
                [object[]]$OutColumns,
                [scriptblock[]]$OutColumnsColorTests,
                [string]$DefaultDisplayColor = ""
            )

            $returnValue = New-Object System.Collections.Generic.List[object]

            foreach ($obj in $OutColumns) {
                $objectValue = New-Object PSCustomObject
                foreach ($property in $obj.PSObject.Properties.Name) {
                    $displayColor = $DefaultDisplayColor
                    foreach ($func in $OutColumnsColorTests) {
                        $result = $func.Invoke($obj, $property)
                        if (-not [string]::IsNullOrEmpty($result)) {
                            $displayColor = $result[0]
                            break
                        }
                    }

                    $objectValue | Add-Member -MemberType NoteProperty -Name $property -Value ([PSCustomObject]@{
                            Value        = $obj.$property
                            DisplayColor = $displayColor
                        })
                }
                $returnValue.Add($objectValue)
            }
            return $returnValue
        }
    }
    process {
        Write-Verbose "Calling $($MyInvocation.MyCommand): $name"

        if ($AddDisplayResultsLineInfo) {
            if (!($AnalyzedInformation.DisplayResults.ContainsKey($DisplayGroupingKey))) {
                Write-Verbose "Adding Display Grouping Key: $($DisplayGroupingKey.Name)"
                [System.Collections.Generic.List[HealthChecker.DisplayResultsLineInfo]]$list = New-Object System.Collections.Generic.List[HealthChecker.DisplayResultsLineInfo]
                $AnalyzedInformation.DisplayResults.Add($DisplayGroupingKey, $list)
            }

            $lineInfo = New-Object HealthChecker.DisplayResultsLineInfo

            if ($null -ne $OutColumns) {
                $lineInfo.OutColumns = $OutColumns
                $lineInfo.WriteType = "OutColumns"
                $lineInfo.TestingValue = (GetOutColumnsColorObject -OutColumns $OutColumns.DisplayObject -OutColumnsColorTests $OutColumnsColorTests -DefaultDisplayColor "Grey")
                $lineInfo.TestingName = $TestingName
            } else {

                $lineInfo.DisplayValue = $Details
                $lineInfo.Name = $Name

                if ($DisplayCustomTabNumber -ne -1) {
                    $lineInfo.TabNumber = $DisplayCustomTabNumber
                } else {
                    $lineInfo.TabNumber = $DisplayGroupingKey.DefaultTabNumber
                }

                if ($null -ne $DisplayTestingValue) {
                    $lineInfo.TestingValue = $DisplayTestingValue
                } else {
                    $lineInfo.TestingValue = $Details
                }

                if (-not ([string]::IsNullOrEmpty($TestingName))) {
                    $lineInfo.TestingName = $TestingName
                } else {
                    $lineInfo.TestingName = $Name
                }

                $lineInfo.WriteType = $DisplayWriteType
            }

            $AnalyzedInformation.DisplayResults[$DisplayGroupingKey].Add($lineInfo)
        }

        if ($AddHtmlDetailRow) {
            if (!($analyzedResults.HtmlServerValues.ContainsKey("ServerDetails"))) {
                [System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]]$list = New-Object System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]
                $AnalyzedInformation.HtmlServerValues.Add("ServerDetails", $list)
            }

            $detailRow = New-Object HealthChecker.HtmlServerInformationRow

            if ($displayWriteType -ne "Grey") {
                $detailRow.Class = $displayWriteType
            }

            if ([string]::IsNullOrEmpty($HtmlName)) {
                $detailRow.Name = $Name
            } else {
                $detailRow.Name = $HtmlName
            }

            if ($null -ne $OutColumns) {
                $detailRow.TableValue = (GetOutColumnsColorObject -OutColumns $OutColumns.DisplayObject -OutColumnsColorTests $OutColumnsColorTests)
            } elseif ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
                $detailRow.DetailValue = $Details
            } else {
                $detailRow.DetailValue = $HtmlDetailsCustomValue
            }

            $AnalyzedInformation.HtmlServerValues["ServerDetails"].Add($detailRow)
        }

        if ($AddHtmlOverviewValues) {
            if (!($analyzedResults.HtmlServerValues.ContainsKey("OverviewValues"))) {
                [System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]]$list = New-Object System.Collections.Generic.List[HealthChecker.HtmlServerInformationRow]
                $AnalyzedInformation.HtmlServerValues.Add("OverviewValues", $list)
            }

            $overviewValue = New-Object HealthChecker.HtmlServerInformationRow

            if ($displayWriteType -ne "Grey") {
                $overviewValue.Class = $displayWriteType
            }

            if ([string]::IsNullOrEmpty($HtmlName)) {
                $overviewValue.Name = $Name
            } else {
                $overviewValue.Name = $HtmlName
            }

            if ([string]::IsNullOrEmpty($HtmlDetailsCustomValue)) {
                $overviewValue.DetailValue = $Details
            } else {
                $overviewValue.DetailValue = $HtmlDetailsCustomValue
            }

            $AnalyzedInformation.HtmlServerValues["OverviewValues"].Add($overviewValue)
        }

        if ($AddHtmlActionRow) {
            #TODO
        }
    }
}

function Get-DisplayResultsGroupingKey {
    param(
        [string]$Name,
        [bool]$DisplayGroupName = $true,
        [int]$DisplayOrder,
        [int]$DefaultTabNumber = 1
    )
    $obj = New-Object HealthChecker.DisplayResultsGroupingKey
    $obj.Name = $Name
    $obj.DisplayGroupName = $DisplayGroupName
    $obj.DisplayOrder = $DisplayOrder
    $obj.DefaultTabNumber = $DefaultTabNumber
    return $obj
}



function Invoke-CatchActionError {
    [CmdletBinding()]
    param(
        [scriptblock]$CatchActionFunction
    )

    if ($null -ne $CatchActionFunction) {
        & $CatchActionFunction
    }
}

# Common method used to handle Invoke-Command within a script.
# Avoids using Invoke-Command when running locally on a server.
function Invoke-ScriptBlockHandler {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,

        [Parameter(Mandatory = $true)]
        [scriptblock]
        $ScriptBlock,

        [string]
        $ScriptBlockDescription,

        [object]
        $ArgumentList,

        [bool]
        $IncludeNoProxyServerOption,

        [scriptblock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $returnValue = $null
    }
    process {

        if (-not([string]::IsNullOrEmpty($ScriptBlockDescription))) {
            Write-Verbose "Description: $ScriptBlockDescription"
        }

        try {

            if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {

                $params = @{
                    ComputerName = $ComputerName
                    ScriptBlock  = $ScriptBlock
                    ErrorAction  = "Stop"
                }

                if ($IncludeNoProxyServerOption) {
                    Write-Verbose "Including SessionOption"
                    $params.Add("SessionOption", (New-PSSessionOption -ProxyAccessType NoProxyServer))
                }

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Invoke-Command with argument list"
                    $params.Add("ArgumentList", $ArgumentList)
                } else {
                    Write-Verbose "Running Invoke-Command without argument list"
                }

                $returnValue = Invoke-Command @params
            } else {

                if ($null -ne $ArgumentList) {
                    Write-Verbose "Running Script Block Locally with argument list"

                    # if an object array type expect the result to be multiple parameters
                    if ($ArgumentList.GetType().Name -eq "Object[]") {
                        $returnValue = & $ScriptBlock @ArgumentList
                    } else {
                        $returnValue = & $ScriptBlock $ArgumentList
                    }
                } else {
                    Write-Verbose "Running Script Block Locally without argument list"
                    $returnValue = & $ScriptBlock
                }
            }
        } catch {
            Write-Verbose "Failed to run $($MyInvocation.MyCommand)"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $returnValue
    }
}

function Get-VisualCRedistributableInstalledVersion {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $softwareList = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        $installedSoftware = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
            -ScriptBlock { Get-ItemProperty HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\* } `
            -ScriptBlockDescription "Querying for software" `
            -CatchActionFunction $CatchActionFunction

        foreach ($software in $installedSoftware) {

            if ($software.PSObject.Properties.Name -contains "DisplayName" -and $software.DisplayName -like "Microsoft Visual C++ *") {
                Write-Verbose "Microsoft Visual C++ Found: $($software.DisplayName)"
                $softwareList.Add([PSCustomObject]@{
                        DisplayName       = $software.DisplayName
                        DisplayVersion    = $software.DisplayVersion
                        InstallDate       = $software.InstallDate
                        VersionIdentifier = $software.Version
                    })
            }
        }
    }
    end {
        Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
        return $softwareList
    }
}

function Get-VisualCRedistributableInfo {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year
    )

    if ($Year -eq 2012) {
        return [PSCustomObject]@{
            VersionNumber = 184610406
            DownloadUrl   = "https://www.microsoft.com/en-us/download/details.aspx?id=30679"
            DisplayName   = "Microsoft Visual C++ 2012*"
        }
    } else {
        return [PSCustomObject]@{
            VersionNumber = 201367256
            DownloadUrl   = "https://support.microsoft.com/en-us/topic/update-for-visual-c-2013-redistributable-package-d8ccd6a5-4e26-c290-517b-8da6cfdf4f10"
            DisplayName   = "Microsoft Visual C++ 2013*"
        }
    }
}

function Test-VisualCRedistributableInstalled {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object { $_.DisplayName -like $desired.DisplayName }))
}

function Test-VisualCRedistributableUpToDate {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true, Position = 0)]
        [ValidateSet(2012, 2013)]
        [int]
        $Year,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]
        $Installed
    )

    $desired = Get-VisualCRedistributableInfo $Year

    return ($null -ne ($Installed | Where-Object {
                $_.DisplayName -like $desired.DisplayName -and $_.VersionIdentifier -eq $desired.VersionNumber
            }))
}




function WriteErrorInformationBase {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0],
        [ValidateSet("Write-Host", "Write-Verbose")]
        [string]$Cmdlet
    )

    if ($null -ne $CurrentError.OriginInfo) {
        & $Cmdlet "Error Origin Info: $($CurrentError.OriginInfo.ToString())"
    }

    & $Cmdlet "$($CurrentError.CategoryInfo.Activity) : $($CurrentError.ToString())"

    if ($null -ne $CurrentError.Exception -and
        $null -ne $CurrentError.Exception.StackTrace) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception.StackTrace)"
    } elseif ($null -ne $CurrentError.Exception) {
        & $Cmdlet "Inner Exception: $($CurrentError.Exception)"
    }

    if ($null -ne $CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage) {
        & $Cmdlet "Position Message: $($CurrentError.Exception.SerializedRemoteInvocationInfo.PositionMessage)"
    }

    if ($null -ne $CurrentError.ScriptStackTrace) {
        & $Cmdlet "Script Stack: $($CurrentError.ScriptStackTrace)"
    }
}

function Write-VerboseErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Verbose"
}

function Write-HostErrorInformation {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    WriteErrorInformationBase $CurrentError "Write-Host"
}

function Invoke-CatchActions {
    [CmdletBinding()]
    param(
        [object]$CurrentError = $Error[0]
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $script:ErrorsExcluded += $CurrentError
    Write-Verbose "Error Excluded Count: $($Script:ErrorsExcluded.Count)"
    Write-Verbose "Error Count: $($Error.Count)"
    Write-VerboseErrorInformation $CurrentError
}

function Get-UnhandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -eq $handledError) {
                    return [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index++
                    }
                }
            }
}

function Get-HandledErrors {
    [CmdletBinding()]
    param ()
    $index = 0
    return $Error |
        ForEach-Object {
            $currentError = $_
            $handledError = $Script:ErrorsExcluded |
                Where-Object { $_.Equals($currentError) }

                if ($null -ne $handledError) {
                    return [PSCustomObject]@{
                        ErrorInformation = $currentError
                        Index            = $index++
                    }
                }
            }
}

function Test-UnhandledErrorsOccurred {
    return $Error.Count -ne $Script:ErrorsExcluded.Count
}

function Invoke-ErrorCatchActionLoopFromIndex {
    [CmdletBinding()]
    param(
        [int]$StartIndex
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Verbose "Start Index: $StartIndex Error Count: $($Error.Count)"

    if ($StartIndex -ne $Error.Count) {
        # Write the errors out in reverse in the order that they came in.
        $index = $Error.Count - $StartIndex - 1
        do {
            Invoke-CatchActions $Error[$index]
            $index--
        } while ($index -ge 0)
    }
}

function Invoke-ErrorMonitoring {
    # Always clear out the errors
    # setup variable to monitor errors that occurred
    $Error.Clear()
    $Script:ErrorsExcluded = @()
}

function Invoke-WriteDebugErrorsThatOccurred {

    function WriteErrorInformation {
        [CmdletBinding()]
        param(
            [object]$CurrentError
        )
        Write-VerboseErrorInformation $CurrentError
        Write-Verbose "-----------------------------------`r`n`r`n"
    }

    if ($Error.Count -gt 0) {
        Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

        Get-UnhandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }

        Write-Verbose "`r`n`r`nErrors that were handled"
        Get-HandledErrors | ForEach-Object {
            Write-Verbose "Error Index: $($_.Index)"
            WriteErrorInformation $_.ErrorInformation
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

function Invoke-AnalyzerKnownBuildIssues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [string]$CurrentBuild,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Extract for Pester Testing - Start
    function GetVersionFromString {
        param(
            [object]$VersionString
        )
        try {
            return New-Object System.Version $VersionString -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to convert '$VersionString' in $($MyInvocation.MyCommand)"
            Invoke-CatchActions
        }
    }

    function GetKnownIssueInformation {
        param(
            [string]$Name,
            [string]$Url
        )

        return [PSCustomObject]@{
            Name = $Name
            Url  = $Url
        }
    }

    function GetKnownIssueBuildInformation {
        param(
            [string]$BuildNumber,
            [string]$FixBuildNumber,
            [bool]$BuildBound = $true
        )

        return [PSCustomObject]@{
            BuildNumber    = $BuildNumber
            FixBuildNumber = $FixBuildNumber
            BuildBound     = $BuildBound
        }
    }

    function TestOnKnownBuildIssue {
        [CmdletBinding()]
        [OutputType("System.Boolean")]
        param(
            [object]$IssueBuildInformation,
            [version]$CurrentBuild
        )
        $knownIssue = GetVersionFromString $IssueBuildInformation.BuildNumber
        Write-Verbose "Testing Known Issue Build $knownIssue"

        if ($null -eq $knownIssue -or
            $CurrentBuild.Minor -ne $knownIssue.Minor) { return $false }

        $fixValueNull = [string]::IsNullOrEmpty($IssueBuildInformation.FixBuildNumber)
        if ($fixValueNull) {
            $resolvedBuild = GetVersionFromString "0.0.0.0"
        } else {
            $resolvedBuild = GetVersionFromString $IssueBuildInformation.FixBuildNumber
        }

        Write-Verbose "Testing against possible resolved build number $resolvedBuild"
        $buildBound = $IssueBuildInformation.BuildBound
        $withinBuildBoundRange = $CurrentBuild.Build -eq $knownIssue.Build
        $fixNeeded = $fixValueNull -or $CurrentBuild -lt $resolvedBuild
        Write-Verbose "BuildBound: $buildBound | WithinBuildBoundRage: $withinBuildBoundRange | FixNeeded: $fixNeeded"
        if ($CurrentBuild -ge $knownIssue) {
            if ($buildBound) {
                return $withinBuildBoundRange -and $fixNeeded
            } else {
                return $fixNeeded
            }
        }

        return $false
    }

    # Extract for Pester Testing - End

    function TestForKnownBuildIssues {
        param(
            [version]$CurrentVersion,
            [object[]]$KnownBuildIssuesToFixes,
            [object]$InformationUrl
        )
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Testing CurrentVersion $CurrentVersion"

        if ($null -eq $Script:CachedKnownIssues) {
            $Script:CachedKnownIssues = @()
        }

        foreach ($issue in $KnownBuildIssuesToFixes) {

            if ((TestOnKnownBuildIssue $issue $CurrentVersion) -and
                    (-not($Script:CachedKnownIssues.Contains($InformationUrl)))) {
                Write-Verbose "Known issue Match detected"
                if (-not ($Script:DisplayKnownIssueHeader)) {
                    $Script:DisplayKnownIssueHeader = $true

                    $params = $baseParams + @{
                        Name             = "Known Issue Detected"
                        Details          = "True"
                        DisplayWriteType = "Yellow"
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Details                = "This build has a known issue(s) which may or may not have been addressed. See the below link(s) for more information.`r`n"
                        DisplayWriteType       = "Yellow"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                $params = $baseParams + @{
                    Details                = "$($InformationUrl.Name):`r`n`t`t`t$($InformationUrl.Url)"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params

                if (-not ($Script:CachedKnownIssues.Contains($InformationUrl))) {
                    $Script:CachedKnownIssues += $InformationUrl
                    Write-Verbose "Added known issue to cache"
                }
            }
        }
    }

    try {
        $currentVersion = New-Object System.Version $CurrentBuild -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to set the current build to a version type object. $CurrentBuild"
        Invoke-CatchActions
    }

    try {
        Write-Verbose "Working on November 2021 Security Updates - OWA redirection"
        TestForKnownBuildIssues -CurrentVersion $currentVersion `
            -KnownBuildIssuesToFixes @(
            (GetKnownIssueBuildInformation "15.2.986.14" "15.2.986.15"),
            (GetKnownIssueBuildInformation "15.2.922.19" "15.2.922.20"),
            (GetKnownIssueBuildInformation "15.1.2375.17" "15.1.2375.18"),
            (GetKnownIssueBuildInformation "15.1.2308.20" "15.1.2308.21"),
            (GetKnownIssueBuildInformation "15.0.1497.26" "15.0.1497.28")
        ) `
            -InformationUrl (GetKnownIssueInformation `
                "OWA redirection doesn't work after installing November 2021 security updates for Exchange Server 2019, 2016, or 2013" `
                "https://support.microsoft.com/help/5008997")

        Write-Verbose "Working on March 2022 Security Updates - MSExchangeServiceHost service may crash"
        TestForKnownBuildIssues -CurrentVersion $currentVersion `
            -KnownBuildIssuesToFixes @(
            (GetKnownIssueBuildInformation "15.2.1118.7" "15.2.1118.9"),
            (GetKnownIssueBuildInformation "15.2.986.22" "15.2.986.26"),
            (GetKnownIssueBuildInformation "15.2.922.27" $null),
            (GetKnownIssueBuildInformation "15.1.2507.6" "15.1.2507.9"),
            (GetKnownIssueBuildInformation "15.1.2375.24" "15.1.2375.28"),
            (GetKnownIssueBuildInformation "15.1.2308.27" $null),
            (GetKnownIssueBuildInformation "15.0.1497.33" "15.0.1497.36")
        ) `
            -InformationUrl (GetKnownIssueInformation `
                "Exchange Service Host service fails after installing March 2022 security update (KB5013118)" `
                "https://support.microsoft.com/kb/5013118")
    } catch {
        Write-Verbose "Failed to run TestForKnownBuildIssues"
        Invoke-CatchActions
    }
}
function Invoke-AnalyzerExchangeInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $keyExchangeInformation = Get-DisplayResultsGroupingKey -Name "Exchange Information"  -DisplayOrder $Order
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keyExchangeInformation
    }

    $params = $baseParams + @{
        Name                  = "Name"
        Details               = $HealthServerObject.ServerName
        AddHtmlOverviewValues = $true
        HtmlName              = "Server Name"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Generation Time"
        Details               = $HealthServerObject.GenerationTime
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Version"
        Details               = $exchangeInformation.BuildInformation.FriendlyName
        AddHtmlOverviewValues = $true
        HtmlName              = "Exchange Version"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "Build Number"
        Details = $exchangeInformation.BuildInformation.ExchangeSetup.FileVersion
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.BuildInformation.SupportedBuild -eq $false) {
        $daysOld = ($date - ([System.Convert]::ToDateTime([DateTime]$exchangeInformation.BuildInformation.ReleaseDate,
                    [System.Globalization.DateTimeFormatInfo]::InvariantInfo))).Days

        $params = $baseParams + @{
            Name                   = "Error"
            Details                = "Out of date Cumulative Update. Please upgrade to one of the two most recently released Cumulative Updates. Currently running on a build that is $daysOld days old."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
            TestingName            = "Out of Date"
            DisplayTestingValue    = $true
            HtmlName               = "Out of date"
        }
        Add-AnalyzedResultInformation @params
    }

    $extendedSupportDate = [System.Convert]::ToDateTime([DateTime]$exchangeInformation.BuildInformation.ExtendedSupportDate,
        [System.Globalization.DateTimeFormatInfo]::InvariantInfo)
    if ($extendedSupportDate -le ([DateTime]::Now.AddYears(1))) {
        $displayWriteType = "Yellow"

        if ($extendedSupportDate -le ([DateTime]::Now.AddDays(178))) {
            $displayWriteType = "Red"
        }

        $displayValue = "$($exchangeInformation.BuildInformation.ExtendedSupportDate) - Please note of the End Of Life date and plan to migrate soon."

        if ($extendedSupportDate -le ([DateTime]::Now)) {
            $displayValue = "$($exchangeInformation.BuildInformation.ExtendedSupportDate) - Error: You are past the End Of Life of Exchange."
        }

        $params = $baseParams + @{
            Name                   = "End Of Life"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if (-not ([string]::IsNullOrEmpty($exchangeInformation.BuildInformation.LocalBuildNumber))) {
        $local = $exchangeInformation.BuildInformation.LocalBuildNumber
        $remote = $exchangeInformation.BuildInformation.BuildNumber

        if ($local.Substring(0, $local.LastIndexOf(".")) -ne $remote.Substring(0, $remote.LastIndexOf("."))) {
            $params = $baseParams + @{
                Name                   = "Warning"
                Details                = "Running commands from a different version box can cause issues. Local Tools Server Version: $local"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($null -ne $exchangeInformation.BuildInformation.KBsInstalled) {
        Add-AnalyzedResultInformation -Name "Exchange IU or Security Hotfix Detected" @baseParams

        foreach ($kb in $exchangeInformation.BuildInformation.KBsInstalled) {
            $params = $baseParams + @{
                Details                = $kb
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    }

    $params = @{
        AnalyzeResults     = $AnalyzeResults
        DisplayGroupingKey = $keyExchangeInformation
        CurrentBuild       = $exchangeInformation.BuildInformation.ExchangeSetup.FileVersion
    }
    Invoke-AnalyzerKnownBuildIssues @params

    $params = $baseParams + @{
        Name                  = "Server Role"
        Details               = $exchangeInformation.BuildInformation.ServerRole
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.BuildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox) {
        $dagName = [System.Convert]::ToString($exchangeInformation.GetMailboxServer.DatabaseAvailabilityGroup)
        if ([System.String]::IsNullOrWhiteSpace($dagName)) {
            $dagName = "Standalone Server"
        }
        $params = $baseParams + @{
            Name    = "DAG Name"
            Details = $dagName
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "AD Site"
        Details = ([System.Convert]::ToString(($exchangeInformation.GetExchangeServer.Site)).Split("/")[-1])
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "MAPI/HTTP Enabled"
        Details = $exchangeInformation.MapiHttpEnabled
    }
    Add-AnalyzedResultInformation @params

    if (($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) -and
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Hub) -and
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::None)) {

        Write-Verbose "Working on MRS Proxy Settings"
        $mrsProxyDetails = $exchangeInformation.GetWebServicesVirtualDirectory.MRSProxyEnabled
        if ($exchangeInformation.GetWebServicesVirtualDirectory.MRSProxyEnabled) {
            $mrsProxyDetails = "$mrsProxyDetails`n`r`t`tKeep MRS Proxy disabled if you do not plan to move mailboxes cross-forest or remote"
            $mrsProxyWriteType = "Yellow"
        } else {
            $mrsProxyWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name             = "MRS Proxy Enabled"
            Details          = $mrsProxyDetails
            DisplayWriteType = $mrsProxyWriteType
        }
        Add-AnalyzedResultInformation @params
    }

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Mailbox) {

        if ($null -ne $exchangeInformation.ApplicationPools -and
            $exchangeInformation.ApplicationPools.Count -gt 0) {
            $mapiFEAppPool = $exchangeInformation.ApplicationPools["MSExchangeMapiFrontEndAppPool"]
            [bool]$enabled = $mapiFEAppPool.GCServerEnabled
            [bool]$unknown = $mapiFEAppPool.GCUnknown
            $warning = [string]::Empty
            $displayWriteType = "Green"
            $displayValue = "Server"

            if ($hardwareInformation.TotalMemory -ge 21474836480 -and
                $enabled -eq $false) {
                $displayWriteType = "Red"
                $displayValue = "Workstation --- Error"
                $warning = "To Fix this issue go into the file MSExchangeMapiFrontEndAppPool_CLRConfig.config in the Exchange Bin directory and change the GCServer to true and recycle the MAPI Front End App Pool"
            } elseif ($unknown) {
                $displayValue = "Unknown --- Warning"
                $displayWriteType = "Yellow"
            } elseif (!($enabled)) {
                $displayWriteType = "Yellow"
                $displayValue = "Workstation --- Warning"
                $warning = "You could be seeing some GC issues within the Mapi Front End App Pool. However, you don't have enough memory installed on the system to recommend switching the GC mode by default without consulting a support professional."
            }

            $params = $baseParams + @{
                Name                   = "MAPI Front End App Pool GC Mode"
                Details                = $displayValue
                DisplayCustomTabNumber = 2
                DisplayWriteType       = $displayWriteType
            }
            Add-AnalyzedResultInformation @params
        } else {
            $warning = "Unable to determine MAPI Front End App Pool GC Mode status. This may be a temporary issue. You should try to re-run the script"
        }

        if ($warning -ne [string]::Empty) {
            $params = $baseParams + @{
                Details                = $warning
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    $internetProxy = $exchangeInformation.GetExchangeServer.InternetWebProxy

    $params = $baseParams + @{
        Name    = "Internet Web Proxy"
        Details = $( if ([string]::IsNullOrEmpty($internetProxy)) { "Not Set" } else { $internetProxy } )
    }
    Add-AnalyzedResultInformation @params

    if (-not ([string]::IsNullOrWhiteSpace($exchangeInformation.GetWebServicesVirtualDirectory.InternalNLBBypassUrl))) {
        $params = $baseParams + @{
            Name             = "EWS Internal Bypass URL Set"
            Details          = "$($exchangeInformation.GetWebServicesVirtualDirectory.InternalNLBBypassUrl) - Can cause issues after KB 5001779"
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    Write-Verbose "Working on results from Test-ServiceHealth"
    $servicesNotRunning = $exchangeInformation.ExchangeServicesNotRunning

    if ($null -ne $servicesNotRunning) {
        Add-AnalyzedResultInformation -Name "Services Not Running" @baseParams

        foreach ($stoppedService in $servicesNotRunning) {
            $params = $baseParams + @{
                Details                = $stoppedService
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Working on Exchange Dependent Services"
    if ($null -ne $exchangeInformation.DependentServices) {

        if ($exchangeInformation.DependentServices.Critical.Count -gt 0) {
            Write-Verbose "Critical Services found to be not running."
            Add-AnalyzedResultInformation -Name "Critical Services Not Running" @baseParams

            foreach ($service in $exchangeInformation.DependentServices.Critical) {
                $params = $baseParams + @{
                    Details                = "$($service.Name) - Status: $($service.Status) - StartType: $($service.StartType)"
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Red"
                    TestingName            = "Critical $($service.Name)"
                }
                Add-AnalyzedResultInformation @params
            }
        }
        if ($exchangeInformation.DependentServices.Common.Count -gt 0) {
            Write-Verbose "Common Services found to be not running."
            Add-AnalyzedResultInformation -Name "Common Services Not Running" @baseParams

            foreach ($service in $exchangeInformation.DependentServices.Common) {
                $params = $baseParams + @{
                    Details                = "$($service.Name) - Status: $($service.Status) - StartType: $($service.StartType)"
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Yellow"
                    TestingName            = "Common $($service.Name)"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
        $null -ne $exchangeInformation.ExtendedProtectionConfig) {
        $params = $baseParams + @{
            Name    = "Extended Protection Enabled (Any Vdir)"
            Details = $exchangeInformation.ExtendedProtectionConfig.ExtendedProtectionConfigured
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $exchangeInformation.SettingOverrides) {

        $overridesDetected = $null -ne $exchangeInformation.SettingOverrides.SettingOverrides
        $params = $baseParams + @{
            Name    = "Setting Overrides Detected"
            Details = $overridesDetected
        }
        Add-AnalyzedResultInformation @params

        if ($overridesDetected) {
            $params = $baseParams + @{
                OutColumns = ([PSCustomObject]@{
                        DisplayObject = $exchangeInformation.SettingOverrides.SimpleSettingOverrides
                        IndentSpaces  = 12
                    })
                HtmlName   = "Setting Overrides"
            }
            Add-AnalyzedResultInformation @params
        }
    }

    Write-Verbose "Working on Exchange Server Maintenance"
    $serverMaintenance = $exchangeInformation.ServerMaintenance
    $getMailboxServer = $exchangeInformation.GetMailboxServer

    if (($serverMaintenance.InactiveComponents).Count -eq 0 -and
        ($null -eq $serverMaintenance.GetClusterNode -or
        $serverMaintenance.GetClusterNode.State -eq "Up") -and
        ($null -eq $getMailboxServer -or
            ($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -eq $false -and
        $getMailboxServer.DatabaseCopyAutoActivationPolicy.ToString() -eq "Unrestricted"))) {
        $params = $baseParams + @{
            Name             = "Exchange Server Maintenance"
            Details          = "Server is not in Maintenance Mode"
            DisplayWriteType = "Green"
        }
        Add-AnalyzedResultInformation @params
    } else {
        Add-AnalyzedResultInformation -Details "Exchange Server Maintenance" @baseParams

        if (($serverMaintenance.InactiveComponents).Count -ne 0) {
            foreach ($inactiveComponent in $serverMaintenance.InactiveComponents) {
                $params = $baseParams + @{
                    Name                   = "Component"
                    Details                = $inactiveComponent
                    DisplayCustomTabNumber = 2
                    DisplayWriteType       = "Red"
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details                = "For more information: https://aka.ms/HC-ServerComponentState"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Yellow"
            }
            Add-AnalyzedResultInformation @params
        }

        if ($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow -or
            $getMailboxServer.DatabaseCopyAutoActivationPolicy -eq "Blocked") {
            $displayValue = "`r`n`t`tDatabaseCopyActivationDisabledAndMoveNow: $($getMailboxServer.DatabaseCopyActivationDisabledAndMoveNow) --- should be 'false'"
            $displayValue += "`r`n`t`tDatabaseCopyAutoActivationPolicy: $($getMailboxServer.DatabaseCopyAutoActivationPolicy) --- should be 'unrestricted'"

            $params = $baseParams + @{
                Name                   = "Database Copy Maintenance"
                Details                = $displayValue
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $serverMaintenance.GetClusterNode -and
            $serverMaintenance.GetClusterNode.State -ne "Up") {
            $params = $baseParams + @{
                Name                   = "Cluster Node"
                Details                = "'$($serverMaintenance.GetClusterNode.State)' --- should be 'Up'"
                DisplayCustomTabNumber = 2
                DisplayWriteType       = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    }
}

function Invoke-AnalyzerHybridInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = Get-DisplayResultsGroupingKey -Name "Hybrid Information"  -DisplayOrder $Order
    }
    $exchangeInformation = $HealthServerObject.ExchangeInformation

    if ($exchangeInformation.BuildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
        $null -ne $exchangeInformation.GetHybridConfiguration) {

        $params = $baseParams + @{
            Name    = "Organization Hybrid Enabled"
            Details = "True"
        }
        Add-AnalyzedResultInformation @params

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.OnPremisesSmartHost))) {
            $onPremSmartHostDomain = ($exchangeInformation.GetHybridConfiguration.OnPremisesSmartHost).ToString()
            $onPremSmartHostWriteType = "Grey"
        } else {
            $onPremSmartHostDomain = "No on-premises smart host domain configured for hybrid use"
            $onPremSmartHostWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "On-Premises Smart Host Domain"
            Details          = $onPremSmartHostDomain
            DisplayWriteType = $onPremSmartHostWriteType
        }
        Add-AnalyzedResultInformation @params

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.Domains))) {
            $domainsConfiguredForHybrid = $exchangeInformation.GetHybridConfiguration.Domains
            $domainsConfiguredForHybridWriteType = "Grey"
        } else {
            $domainsConfiguredForHybridWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "Domain(s) configured for Hybrid use"
            DisplayWriteType = $domainsConfiguredForHybridWriteType
        }
        Add-AnalyzedResultInformation @params

        if ($domainsConfiguredForHybrid.Count -ge 1) {
            foreach ($domain in $domainsConfiguredForHybrid) {
                $params = $baseParams + @{
                    Details                = $domain
                    DisplayWriteType       = $domainsConfiguredForHybridWriteType
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            $params = $baseParams + @{
                Details                = "No domain configured for Hybrid use"
                DisplayWriteType       = $domainsConfiguredForHybridWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.EdgeTransportServers))) {
            Add-AnalyzedResultInformation -Name "Edge Transport Server(s)" @baseParams

            foreach ($edgeServer in $exchangeInformation.GetHybridConfiguration.EdgeTransportServers) {
                $params = $baseParams + @{
                    Details                = $edgeServer
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.ReceivingTransportServers)) -or
            (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.SendingTransportServers)))) {
                $params = $baseParams + @{
                    Details                = "When configuring the EdgeTransportServers parameter, you must configure the ReceivingTransportServers and SendingTransportServers parameter values to null"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            Add-AnalyzedResultInformation -Name "Receiving Transport Server(s)" @baseParams

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.ReceivingTransportServers))) {
                foreach ($receivingTransportSrv in $exchangeInformation.GetHybridConfiguration.ReceivingTransportServers) {
                    $params = $baseParams + @{
                        Details                = $receivingTransportSrv
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            } else {
                $params = $baseParams + @{
                    Details                = "No Receiving Transport Server configured for Hybrid use"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            Add-AnalyzedResultInformation -Name "Sending Transport Server(s)" @baseParams

            if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.SendingTransportServers))) {
                foreach ($sendingTransportSrv in $exchangeInformation.GetHybridConfiguration.SendingTransportServers) {
                    $params = $baseParams + @{
                        Details                = $sendingTransportSrv
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            } else {
                $params = $baseParams + @{
                    Details                = "No Sending Transport Server configured for Hybrid use"
                    DisplayWriteType       = "Yellow"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        }

        if ($exchangeInformation.GetHybridConfiguration.ServiceInstance -eq 1) {
            $params = $baseParams + @{
                Name    = "Service Instance"
                Details = "Office 365 operated by 21Vianet"
            }
            Add-AnalyzedResultInformation @params
        } elseif ($exchangeInformation.GetHybridConfiguration.ServiceInstance -ne 0) {
            $params = $baseParams + @{
                Name             = "Service Instance"
                Details          = $exchangeInformation.GetHybridConfiguration.ServiceInstance
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details          = "You are using an invalid value. Please set this value to 0 (null) or re-run HCW"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.TlsCertificateName))) {
            $params = $baseParams + @{
                Name    = "TLS Certificate Name"
                Details = ($exchangeInformation.GetHybridConfiguration.TlsCertificateName).ToString()
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name             = "TLS Certificate Name"
                Details          = "No valid certificate found"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }

        Add-AnalyzedResultInformation -Name "Feature(s) enabled for Hybrid use" @baseParams

        if (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.Features))) {
            foreach ($feature in $exchangeInformation.GetHybridConfiguration.Features) {
                $params = $baseParams + @{
                    Details                = $feature
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            $params = $baseParams + @{
                Details                = "No feature(s) enabled for Hybrid use"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $exchangeInformation.ExchangeConnectors) {
            foreach ($connector in $exchangeInformation.ExchangeConnectors) {
                $cloudConnectorWriteType = "Yellow"
                if (($connector.TransportRole -ne "HubTransport") -and
                    ($connector.CloudEnabled -eq $true)) {

                    $params = $baseParams + @{
                        Details          = "`r"
                        AddHtmlDetailRow = $false
                    }
                    Add-AnalyzedResultInformation @params

                    if (($connector.CertificateDetails.CertificateMatchDetected) -and
                        ($connector.CertificateDetails.GoodTlsCertificateSyntax)) {
                        $cloudConnectorWriteType = "Green"
                    }

                    $params = $baseParams + @{
                        Name    = "Connector Name"
                        Details = $connector.Name
                    }
                    Add-AnalyzedResultInformation @params

                    $cloudConnectorEnabledWriteType = "Gray"
                    if ($connector.Enabled -eq $false) {
                        $cloudConnectorEnabledWriteType = "Yellow"
                    }

                    $params = $baseParams + @{
                        Name             = "Connector Enabled"
                        Details          = $connector.Enabled
                        DisplayWriteType = $cloudConnectorEnabledWriteType
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Name    = "Cloud Mail Enabled"
                        Details = $connector.CloudEnabled
                    }
                    Add-AnalyzedResultInformation @params

                    $params = $baseParams + @{
                        Name    = "Connector Type"
                        Details = $connector.ConnectorType
                    }
                    Add-AnalyzedResultInformation @params

                    if (($connector.ConnectorType -eq "Send") -and
                        ($null -ne $connector.TlsAuthLevel)) {
                        # Check if send connector is configured to relay mails to the internet via M365
                        switch ($connector) {
                            { ($_.SmartHosts -like "*.mail.protection.outlook.com") } {
                                $smartHostsPointToExo = $true
                            }
                            { ([System.Management.Automation.WildcardPattern]::ContainsWildcardCharacters($_.AddressSpaces)) } {
                                $addressSpacesContainsWildcard = $true
                            }
                        }

                        if (($smartHostsPointToExo -eq $false) -or
                            ($addressSpacesContainsWildcard -eq $false)) {

                            $tlsAuthLevelWriteType = "Gray"
                            if ($connector.TlsAuthLevel -eq "DomainValidation") {
                                # DomainValidation: In addition to channel encryption and certificate validation,
                                # the Send connector also verifies that the FQDN of the target certificate matches
                                # the domain specified in the TlsDomain parameter. If no domain is specified in the TlsDomain parameter,
                                # the FQDN on the certificate is compared with the recipient's domain.
                                $tlsAuthLevelWriteType = "Green"
                                if ($null -eq $connector.TlsDomain) {
                                    $tlsAuthLevelWriteType = "Yellow"
                                    $tlsAuthLevelAdditionalInfo = "'TlsDomain' is empty which means that the FQDN of the certificate is compared with the recipient's domain.`r`n`t`tMore information: https://aka.ms/HC-HybridConnector"
                                }
                            }

                            $params = $baseParams + @{
                                Name             = "TlsAuthLevel"
                                Details          = $connector.TlsAuthLevel
                                DisplayWriteType = $tlsAuthLevelWriteType
                            }
                            Add-AnalyzedResultInformation @params

                            if ($null -ne $tlsAuthLevelAdditionalInfo) {
                                $params = $baseParams + @{
                                    Details                = $tlsAuthLevelAdditionalInfo
                                    DisplayWriteType       = $tlsAuthLevelWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }
                        }
                    }

                    if (($smartHostsPointToExo) -and
                        ($addressSpacesContainsWildcard)) {
                        # Seems like this send connector is configured to relay mails to the internet via M365 - skipping some checks
                        # https://docs.microsoft.com/exchange/mail-flow-best-practices/use-connectors-to-configure-mail-flow/set-up-connectors-to-route-mail#2-set-up-your-email-server-to-relay-mail-to-the-internet-via-microsoft-365-or-office-365
                        $params = $baseParams + @{
                            Name    = "Relay Internet Mails via M365"
                            Details = $true
                        }
                        Add-AnalyzedResultInformation @params

                        switch ($connector.TlsAuthLevel) {
                            "EncryptionOnly" {
                                $tlsAuthLevelM365RelayWriteType = "Yellow";
                                break
                            }
                            "CertificateValidation" {
                                $tlsAuthLevelM365RelayWriteType = "Green";
                                break
                            }
                            "DomainValidation" {
                                if ($null -eq $connector.TlsDomain) {
                                    $tlsAuthLevelM365RelayWriteType = "Red"
                                } else {
                                    $tlsAuthLevelM365RelayWriteType = "Green"
                                };
                                break
                            }
                            default { $tlsAuthLevelM365RelayWriteType = "Red" }
                        }

                        $params = $baseParams + @{
                            Name             = "TlsAuthLevel"
                            Details          = $connector.TlsAuthLevel
                            DisplayWriteType = $tlsAuthLevelM365RelayWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($tlsAuthLevelM365RelayWriteType -ne "Green") {
                            $params = $baseParams + @{
                                Details                = "'TlsAuthLevel' should be set to 'CertificateValidation'. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $tlsAuthLevelM365RelayWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        }

                        $requireTlsWriteType = "Red"
                        if ($connector.RequireTLS) {
                            $requireTlsWriteType = "Green"
                        }

                        $params = $baseParams + @{
                            Name             = "RequireTls Enabled"
                            Details          = $connector.RequireTLS
                            DisplayWriteType = $requireTlsWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($requireTlsWriteType -eq "Red") {
                            $params = $baseParams + @{
                                Details                = "'RequireTLS' must be set to 'true' to ensure a working mail flow. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $requireTlsWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        }
                    } else {
                        $cloudConnectorTlsCertificateName = "Not set"
                        if ($null -ne $connector.CertificateDetails.TlsCertificateName) {
                            $cloudConnectorTlsCertificateName = $connector.CertificateDetails.TlsCertificateName
                        }

                        $params = $baseParams + @{
                            Name             = "TlsCertificateName"
                            Details          = $cloudConnectorTlsCertificateName
                            DisplayWriteType = $cloudConnectorWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        $params = $baseParams + @{
                            Name             = "Certificate Found On Server"
                            Details          = $connector.CertificateDetails.CertificateMatchDetected
                            DisplayWriteType = $cloudConnectorWriteType
                        }
                        Add-AnalyzedResultInformation @params

                        if ($connector.CertificateDetails.TlsCertificateNameStatus -eq "TlsCertificateNameEmpty") {
                            $params = $baseParams + @{
                                Details                = "There is no 'TlsCertificateName' configured for this cloud mail enabled connector.`r`n`t`tThis will cause mail flow issues in hybrid scenarios. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $cloudConnectorWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        } elseif ($connector.CertificateDetails.CertificateMatchDetected -eq $false) {
                            $params = $baseParams + @{
                                Details                = "The configured 'TlsCertificateName' was not found on the server.`r`n`t`tThis may cause mail flow issues. More information: https://aka.ms/HC-HybridConnector"
                                DisplayWriteType       = $cloudConnectorWriteType
                                DisplayCustomTabNumber = 2
                            }
                            Add-AnalyzedResultInformation @params
                        } else {
                            Add-AnalyzedResultInformation -Name "Certificate Thumbprint(s)" @baseParams

                            foreach ($thumbprint in $($connector.CertificateDetails.CertificateLifetimeInfo).keys) {
                                $params = $baseParams + @{
                                    Details                = $thumbprint
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }

                            Add-AnalyzedResultInformation -Name "Lifetime In Days" @baseParams

                            foreach ($thumbprint in $($connector.CertificateDetails.CertificateLifetimeInfo).keys) {
                                switch ($($connector.CertificateDetails.CertificateLifetimeInfo)[$thumbprint]) {
                                    { ($_ -ge 60) } { $certificateLifetimeWriteType = "Green"; break }
                                    { ($_ -ge 30) } { $certificateLifetimeWriteType = "Yellow"; break }
                                    default { $certificateLifetimeWriteType = "Red" }
                                }

                                $params = $baseParams + @{
                                    Details                = ($connector.CertificateDetails.CertificateLifetimeInfo)[$thumbprint]
                                    DisplayWriteType       = $certificateLifetimeWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }

                            $connectorCertificateMatchesHybridCertificate = $false
                            $connectorCertificateMatchesHybridCertificateWritingType = "Yellow"
                            if (($connector.CertificateDetails.TlsCertificateSet) -and
                                (-not([System.String]::IsNullOrEmpty($exchangeInformation.GetHybridConfiguration.TlsCertificateName))) -and
                                ($connector.CertificateDetails.TlsCertificateName -eq $exchangeInformation.GetHybridConfiguration.TlsCertificateName)) {
                                $connectorCertificateMatchesHybridCertificate = $true
                                $connectorCertificateMatchesHybridCertificateWritingType = "Green"
                            }

                            $params = $baseParams + @{
                                Name             = "Certificate Matches Hybrid Certificate"
                                Details          = $connectorCertificateMatchesHybridCertificate
                                DisplayWriteType = $connectorCertificateMatchesHybridCertificateWritingType
                            }
                            Add-AnalyzedResultInformation @params

                            if (($connector.CertificateDetails.TlsCertificateNameStatus -eq "TlsCertificateNameSyntaxInvalid") -or
                                (($connector.CertificateDetails.GoodTlsCertificateSyntax -eq $false) -and
                                    ($null -ne $connector.CertificateDetails.TlsCertificateName))) {
                                $params = $baseParams + @{
                                    Name             = "TlsCertificateName Syntax Invalid"
                                    Details          = "True"
                                    DisplayWriteType = $cloudConnectorWriteType
                                }
                                Add-AnalyzedResultInformation @params

                                $params = $baseParams + @{
                                    Details                = "The correct syntax is: '<I>X.500Issuer<S>X.500Subject'"
                                    DisplayWriteType       = $cloudConnectorWriteType
                                    DisplayCustomTabNumber = 2
                                }
                                Add-AnalyzedResultInformation @params
                            }
                        }
                    }
                }
            }
        }
    }
}

function Invoke-AnalyzerOsInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Operating System Information"  -DisplayOrder $Order)
    }

    $params = $baseParams + @{
        Name                  = "Version"
        Details               = $osInformation.BuildInformation.FriendlyName
        AddHtmlOverviewValues = $true
        HtmlName              = "OS Version"
    }
    Add-AnalyzedResultInformation @params

    $upTime = "{0} day(s) {1} hour(s) {2} minute(s) {3} second(s)" -f $osInformation.ServerBootUp.Days,
    $osInformation.ServerBootUp.Hours,
    $osInformation.ServerBootUp.Minutes,
    $osInformation.ServerBootUp.Seconds

    $params = $baseParams + @{
        Name                = "System Up Time"
        Details             = $upTime
        DisplayTestingValue = $osInformation.ServerBootUp
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Time Zone"
        Details               = $osInformation.TimeZone.CurrentTimeZone
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $writeValue = $false
    $warning = @("Windows can not properly detect any DST rule changes in your time zone. Set 'Adjust for daylight saving time automatically to on'")

    if ($osInformation.TimeZone.DstIssueDetected) {
        $writeType = "Red"
    } elseif ($osInformation.TimeZone.DynamicDaylightTimeDisabled -ne 0) {
        $writeType = "Yellow"
    } else {
        $warning = [string]::Empty
        $writeValue = $true
        $writeType = "Grey"
    }

    $params = $baseParams + @{
        Name             = "Dynamic Daylight Time Enabled"
        Details          = $writeValue
        DisplayWriteType = $writeType
    }
    Add-AnalyzedResultInformation @params

    if ($warning -ne [string]::Empty) {
        $params = $baseParams + @{
            Details                = $warning
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if ([string]::IsNullOrEmpty($osInformation.TimeZone.TimeZoneKeyName)) {
        $params = $baseParams + @{
            Name             = "Time Zone Key Name"
            Details          = "Empty --- Warning Need to switch your current time zone to a different value, then switch it back to have this value populated again."
            DisplayWriteType = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($exchangeInformation.NETFramework.OnRecommendedVersion) {
        $params = $baseParams + @{
            Name                  = ".NET Framework"
            Details               = $osInformation.NETFramework.FriendlyName
            DisplayWriteType      = "Green"
            AddHtmlOverviewValues = $true
        }
        Add-AnalyzedResultInformation @params
    } else {
        $displayFriendly = Get-NETFrameworkVersion -NetVersionKey $exchangeInformation.NETFramework.MaxSupportedVersion
        $displayValue = "{0} - Warning Recommended .NET Version is {1}" -f $osInformation.NETFramework.FriendlyName, $displayFriendly.FriendlyName
        $testValue = [PSCustomObject]@{
            CurrentValue        = $osInformation.NETFramework.FriendlyName
            MaxSupportedVersion = $exchangeInformation.NETFramework.MaxSupportedVersion
        }
        $params = $baseParams + @{
            Name                   = ".NET Framework"
            Details                = $displayValue
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $testValue
            HtmlDetailsCustomValue = $osInformation.NETFramework.FriendlyName
            AddHtmlOverviewValues  = $true
        }
        Add-AnalyzedResultInformation @params
    }

    $displayValue = [string]::Empty
    $displayWriteType = "Yellow"
    $totalPhysicalMemory = [Math]::Round($hardwareInformation.TotalMemory / 1MB)
    $instanceCount = 0
    Write-Verbose "Evaluating PageFile Information"
    Write-Verbose "Total Memory: $totalPhysicalMemory"

    foreach ($pageFile in $osInformation.PageFile) {

        $pageFileDisplayTemplate = "{0} Size: {1}MB"
        $pageFileAdditionalDisplayValue = $null

        Write-Verbose "Working on PageFile: $($pageFile.Name)"
        Write-Verbose "Max PageFile Size: $($pageFile.MaximumSize)"
        $pageFileObj = [PSCustomObject]@{
            Name                = $pageFile.Name
            TotalPhysicalMemory = $totalPhysicalMemory
            MaxPageSize         = $pageFile.MaximumSize
            MultiPageFile       = (($osInformation.PageFile).Count -gt 1)
            RecommendedPageFile = 0
        }

        if ($pageFileObj.MaxPageSize -eq 0) {
            Write-Verbose "Unconfigured PageFile detected"
            if ([System.String]::IsNullOrEmpty($pageFileObj.Name)) {
                Write-Verbose "System-wide automatically managed PageFile detected"
                $displayValue = ($pageFileDisplayTemplate -f "System is set to automatically manage the PageFile", $pageFileObj.MaxPageSize)
            } else {
                Write-Verbose "Specific system-managed PageFile detected"
                $displayValue = ($pageFileDisplayTemplate -f $pageFileObj.Name, $pageFileObj.MaxPageSize)
            }
            $displayWriteType = "Red"
        } else {
            Write-Verbose "Configured PageFile detected"
            $displayValue = ($pageFileDisplayTemplate -f $pageFileObj.Name, $pageFileObj.MaxPageSize)
        }

        if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
            $recommendedPageFile = [Math]::Round($totalPhysicalMemory / 4)
            $pageFileObj.RecommendedPageFile = $recommendedPageFile
            Write-Verbose "System is running Exchange 2019. Recommended PageFile Size: $recommendedPageFile"

            $recommendedPageFileWording2019 = "On Exchange 2019, the recommended PageFile size is 25% ({0}MB) of the total system memory ({1}MB)."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = ("Error: $recommendedPageFileWording2019" -f $recommendedPageFile, $totalPhysicalMemory)
            } elseif ($recommendedPageFile -ne $pageFileObj.MaxPageSize) {
                $pageFileAdditionalDisplayValue = ("Warning: $recommendedPageFileWording2019" -f $recommendedPageFile, $totalPhysicalMemory)
            } else {
                $displayWriteType = "Grey"
            }
        } elseif ($totalPhysicalMemory -ge 32768) {
            Write-Verbose "System is not running Exchange 2019 and has more than 32GB memory. Recommended PageFile Size: 32778MB"

            $recommendedPageFileWording32GBPlus = "PageFile should be capped at 32778MB for 32GB plus 10MB."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = "Error: $recommendedPageFileWording32GBPlus"
            } elseif ($pageFileObj.MaxPageSize -eq 32778) {
                $displayWriteType = "Grey"
            } else {
                $pageFileAdditionalDisplayValue = "Warning: $recommendedPageFileWording32GBPlus"
            }
        } else {
            $recommendedPageFile = $totalPhysicalMemory + 10
            Write-Verbose "System is not running Exchange 2019 and has less than 32GB of memory. Recommended PageFile Size: $recommendedPageFile"

            $recommendedPageFileWordingBelow32GB = "PageFile is not set to total system memory plus 10MB which should be {0}MB."
            if ($pageFileObj.MaxPageSize -eq 0) {
                $pageFileAdditionalDisplayValue = ("Error: $recommendedPageFileWordingBelow32GB" -f $recommendedPageFile)
            } elseif ($recommendedPageFile -ne $pageFileObj.MaxPageSize) {
                $pageFileAdditionalDisplayValue = ("Warning: $recommendedPageFileWordingBelow32GB" -f $recommendedPageFile)
            } else {
                $displayWriteType = "Grey"
            }
        }

        $params = $baseParams + @{
            Name                = "PageFile"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            TestingName         = "PageFile Size $instanceCount"
            DisplayTestingValue = $pageFileObj
        }
        Add-AnalyzedResultInformation @params

        if ($null -ne $pageFileAdditionalDisplayValue) {
            $params = $baseParams + @{
                Details                = $pageFileAdditionalDisplayValue
                DisplayWriteType       = $displayWriteType
                TestingName            = "PageFile Additional Information"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "More information: https://aka.ms/HC-PageFile"
                DisplayWriteType       = $displayWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $instanceCount++
    }

    if ($null -ne $osInformation.PageFile -and
        $osInformation.PageFile.Count -gt 1) {
        $params = $baseParams + @{
            Details                = "`r`n`t`tError: Multiple PageFiles detected. This has been known to cause performance issues, please address this."
            DisplayWriteType       = "Red"
            TestingName            = "Multiple PageFile Detected"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($osInformation.PowerPlan.HighPerformanceSet) {
        $params = $baseParams + @{
            Name             = "Power Plan"
            Details          = $osInformation.PowerPlan.PowerPlanSetting
            DisplayWriteType = "Green"
        }
        Add-AnalyzedResultInformation @params
    } else {
        $params = $baseParams + @{
            Name             = "Power Plan"
            Details          = "$($osInformation.PowerPlan.PowerPlanSetting) --- Error"
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    $displayWriteType = "Grey"
    $displayValue = $osInformation.NetworkInformation.HttpProxy.ProxyAddress

    if (($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne "None") -and
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge)) {
        $displayValue = "$($osInformation.NetworkInformation.HttpProxy.ProxyAddress) --- Warning this can cause client connectivity issues."
        $displayWriteType = "Yellow"
    }

    $params = $baseParams + @{
        Name                = "Http Proxy Setting"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $osInformation.NetworkInformation.HttpProxy
    }
    Add-AnalyzedResultInformation @params

    if ($displayWriteType -eq "Yellow") {
        $params = $baseParams + @{
            Name             = "Http Proxy By Pass List"
            Details          = "$($osInformation.NetworkInformation.HttpProxy.ByPassList)"
            DisplayWriteType = "Yellow"
        }
        Add-AnalyzedResultInformation @params
    }

    if (($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne "None") -and
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) -and
        ($osInformation.NetworkInformation.HttpProxy.ProxyAddress -ne $exchangeInformation.GetExchangeServer.InternetWebProxy.Authority)) {
        $params = $baseParams + @{
            Details                = "Error: Exchange Internet Web Proxy doesn't match OS Web Proxy."
            DisplayWriteType       = "Red"
            TestingName            = "Proxy Doesn't Match"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $displayWriteType2012 = $displayWriteType2013 = "Red"
    $displayValue2012 = $displayValue2013 = $defaultValue = "Error --- Unknown"

    if ($null -ne $osInformation.VcRedistributable) {

        if (Test-VisualCRedistributableUpToDate -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2012 = "Green"
            $displayValue2012 = "$((Get-VisualCRedistributableInfo 2012).VersionNumber) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2012 -Installed $osInformation.VcRedistributable) {
            $displayValue2012 = "Redistributable is outdated"
            $displayWriteType2012 = "Yellow"
        }

        if (Test-VisualCRedistributableUpToDate -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayWriteType2013 = "Green"
            $displayValue2013 = "$((Get-VisualCRedistributableInfo 2013).VersionNumber) Version is current"
        } elseif (Test-VisualCRedistributableInstalled -Year 2013 -Installed $osInformation.VcRedistributable) {
            $displayValue2013 = "Redistributable is outdated"
            $displayWriteType2013 = "Yellow"
        }
    }

    $params = $baseParams + @{
        Name             = "Visual C++ 2012"
        Details          = $displayValue2012
        DisplayWriteType = $displayWriteType2012
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
        $params = $baseParams + @{
            Name             = "Visual C++ 2013"
            Details          = $displayValue2013
            DisplayWriteType = $displayWriteType2013
        }
        Add-AnalyzedResultInformation @params
    }

    if (($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
            ($displayWriteType2012 -eq "Yellow" -or
            $displayWriteType2013 -eq "Yellow")) -or
        $displayWriteType2012 -eq "Yellow") {

        $params = $baseParams + @{
            Details                = "Note: For more information about the latest C++ Redistributable please visit: https://aka.ms/HC-LatestVC`r`n`t`tThis is not a requirement to upgrade, only a notification to bring to your attention."
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($defaultValue -eq $displayValue2012 -or
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge -and
        $displayValue2013 -eq $defaultValue)) {

        $params = $baseParams + @{
            Details                = "ERROR: Unable to find one of the Visual C++ Redistributable Packages. This can cause a wide range of issues on the server.`r`n`t`tPlease install the missing package as soon as possible. Latest C++ Redistributable please visit: https://aka.ms/HC-LatestVC"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    $displayValue = "False"
    $writeType = "Grey"

    if ($osInformation.ServerPendingReboot.PendingReboot) {
        $displayValue = "True --- Warning a reboot is pending and can cause issues on the server."
        $writeType = "Yellow"
    }

    $params = $baseParams + @{
        Name                = "Server Pending Reboot"
        Details             = $displayValue
        DisplayWriteType    = $writeType
        DisplayTestingValue = $osInformation.ServerPendingReboot.PendingReboot
    }
    Add-AnalyzedResultInformation @params

    if ($osInformation.ServerPendingReboot.PendingReboot -and
        $osInformation.ServerPendingReboot.PendingRebootLocations.Count -gt 0) {

        foreach ($line in $osInformation.ServerPendingReboot.PendingRebootLocations) {
            $params = $baseParams + @{
                Details                = $line
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
                TestingName            = $line
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-RebootPending"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Reboot More Information"
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerHardwareInformation {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Processor/Hardware Information"  -DisplayOrder $Order)
    }

    $params = $baseParams + @{
        Name                  = "Type"
        Details               = $hardwareInformation.ServerType
        AddHtmlOverviewValues = $true
        HtmlName              = "Hardware Type"
    }
    Add-AnalyzedResultInformation @params

    if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
        $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
        $params = $baseParams + @{
            Name    = "Manufacturer"
            Details = $hardwareInformation.Manufacturer
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name    = "Model"
            Details = $hardwareInformation.Model
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "Processor"
        Details = $hardwareInformation.Processor.Name
    }
    Add-AnalyzedResultInformation @params

    $numberOfProcessors = $hardwareInformation.Processor.NumberOfProcessors
    $displayWriteType = "Green"
    $displayValue = $numberOfProcessors

    if ($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::Physical) {
        $displayWriteType = "Grey"
    } elseif ($numberOfProcessors -gt 2) {
        $displayWriteType = "Red"
        $displayValue = "$numberOfProcessors - Error: Recommended to only have 2 Processors"
    }

    $params = $baseParams + @{
        Name                = "Number of Processors"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $numberOfProcessors
    }
    Add-AnalyzedResultInformation @params

    $physicalValue = $hardwareInformation.Processor.NumberOfPhysicalCores
    $logicalValue = $hardwareInformation.Processor.NumberOfLogicalCores
    $physicalValueDisplay = $physicalValue
    $logicalValueDisplay = $logicalValue
    $displayWriteTypeLogic = $displayWriteTypePhysical = "Green"

    if (($logicalValue -gt 24 -and
            $exchangeInformation.BuildInformation.MajorVersion -lt [HealthChecker.ExchangeMajorVersion]::Exchange2019) -or
        $logicalValue -gt 48) {
        $displayWriteTypeLogic = "Red"

        if (($physicalValue -gt 24 -and
                $exchangeInformation.BuildInformation.MajorVersion -lt [HealthChecker.ExchangeMajorVersion]::Exchange2019) -or
            $physicalValue -gt 48) {
            $physicalValueDisplay = "$physicalValue - Error"
            $displayWriteTypePhysical = "Red"
        }

        $logicalValueDisplay = "$logicalValue - Error"
    }

    $params = $baseParams + @{
        Name             = "Number of Physical Cores"
        Details          = $physicalValueDisplay
        DisplayWriteType = $displayWriteTypePhysical
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                  = "Number of Logical Cores"
        Details               = $logicalValueDisplay
        DisplayWriteType      = $displayWriteTypeLogic
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params

    $displayValue = "Disabled"
    $displayWriteType = "Green"
    $displayTestingValue = $false
    $additionalDisplayValue = [string]::Empty
    $additionalWriteType = "Red"

    if ($logicalValue -gt $physicalValue) {

        if ($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::HyperV) {
            $displayValue = "Enabled --- Error: Having Hyper-Threading enabled goes against best practices and can cause performance issues. Please disable as soon as possible."
            $displayTestingValue = $true
            $displayWriteType = "Red"
        } else {
            $displayValue = "Enabled --- Not Applicable"
            $displayTestingValue = $true
            $displayWriteType = "Grey"
        }

        if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
            $additionalDisplayValue = "Error: For high-performance computing (HPC) application, like Exchange, Amazon recommends that you have Hyper-Threading Technology disabled in their service. More information: https://aka.ms/HC-EC2HyperThreading"
        }

        if ($hardwareInformation.Processor.Name.StartsWith("AMD")) {
            $additionalDisplayValue = "This script may incorrectly report that Hyper-Threading is enabled on certain AMD processors. Check with the manufacturer to see if your model supports SMT."
            $additionalWriteType = "Yellow"
        }
    }

    $params = $baseParams + @{
        Name                = "Hyper-Threading"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $displayTestingValue
    }
    Add-AnalyzedResultInformation @params

    if (!([string]::IsNullOrEmpty($additionalDisplayValue))) {
        $params = $baseParams + @{
            Details                = $additionalDisplayValue
            DisplayWriteType       = $additionalWriteType
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    #NUMA BIOS CHECK - AKA check to see if we can properly see all of our cores on the box
    $displayWriteType = "Yellow"
    $testingValue = "Unknown"
    $displayValue = [string]::Empty

    if ($hardwareInformation.Model.Contains("ProLiant")) {
        $name = "NUMA Group Size Optimization"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If this is set to Clustered, this can cause multiple types of issues on the server"
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Clustered `r`n`t`tError: This setting should be set to Flat. By having this set to Clustered, we will see multiple different types of issues."
            $testingValue = "Clustered"
            $displayWriteType = "Red"
        } else {
            $displayValue = "Flat"
            $testingValue = "Flat"
            $displayWriteType = "Green"
        }
    } else {
        $name = "All Processor Cores Visible"

        if ($hardwareInformation.Processor.EnvironmentProcessorCount -eq -1) {
            $displayValue = "Unknown `r`n`t`tWarning: If we aren't able to see all processor cores from Exchange, we could see performance related issues."
        } elseif ($hardwareInformation.Processor.EnvironmentProcessorCount -ne $logicalValue) {
            $displayValue = "Failed `r`n`t`tError: Not all Processor Cores are visible to Exchange and this will cause a performance impact"
            $displayWriteType = "Red"
            $testingValue = "Failed"
        } else {
            $displayWriteType = "Green"
            $displayValue = "Passed"
            $testingValue = "Passed"
        }
    }

    $params = $baseParams + @{
        Name                = $name
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $testingValue
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name    = "Max Processor Speed"
        Details = $hardwareInformation.Processor.MaxMegacyclesPerCore
    }
    Add-AnalyzedResultInformation @params

    if ($hardwareInformation.Processor.ProcessorIsThrottled) {
        $params = $baseParams + @{
            Name                = "Current Processor Speed"
            Details             = "$($hardwareInformation.Processor.CurrentMegacyclesPerCore) --- Error: Processor appears to be throttled."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $hardwareInformation.Processor.CurrentMegacyclesPerCore
        }
        Add-AnalyzedResultInformation @params

        $displayValue = "Error: Power Plan is NOT set to `"High Performance`". This change doesn't require a reboot and takes affect right away. Re-run script after doing so"

        if ($osInformation.PowerPlan.HighPerformanceSet) {
            $displayValue = "Error: Power Plan is set to `"High Performance`", so it is likely that we are throttling in the BIOS of the computer settings."
        }

        $params = $baseParams + @{
            Details             = $displayValue
            DisplayWriteType    = "Red"
            TestingName         = "HighPerformanceSet"
            DisplayTestingValue = $osInformation.PowerPlan.HighPerformanceSet
            AddHtmlDetailRow    = $false
        }
        Add-AnalyzedResultInformation @params
    }

    $totalPhysicalMemory = [System.Math]::Round($hardwareInformation.TotalMemory / 1024 / 1024 / 1024)
    $displayWriteType = "Yellow"
    $displayDetails = [string]::Empty

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {

        if ($totalPhysicalMemory -gt 256) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 256 GB of Memory" -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 64 -and
            $exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::Edge) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 64GB of RAM installed on the machine." -f $totalPhysicalMemory
        } elseif ($totalPhysicalMemory -lt 128) {
            $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to have a minimum of 128GB of RAM installed on the machine." -f $totalPhysicalMemory
        } else {
            $displayDetails = "{0} GB" -f $totalPhysicalMemory
            $displayWriteType = "Grey"
        }
    } elseif ($totalPhysicalMemory -gt 192 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 192 GB of Memory." -f $totalPhysicalMemory
    } elseif ($totalPhysicalMemory -gt 96 -and
        $exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $displayDetails = "{0} GB `r`n`t`tWarning: We recommend for the best performance to be scaled at or below 96GB of Memory." -f $totalPhysicalMemory
    } else {
        $displayDetails = "{0} GB" -f $totalPhysicalMemory
        $displayWriteType = "Grey"
    }

    $params = $baseParams + @{
        Name                  = "Physical Memory"
        Details               = $displayDetails
        DisplayWriteType      = $displayWriteType
        DisplayTestingValue   = $totalPhysicalMemory
        AddHtmlOverviewValues = $true
    }
    Add-AnalyzedResultInformation @params
}

function Invoke-AnalyzerNicSettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "NIC Settings Per Active Adapter"  -DisplayOrder $Order -DefaultTabNumber 2)
    }
    $osInformation = $HealthServerObject.OSInformation
    $hardwareInformation = $HealthServerObject.HardwareInformation

    foreach ($adapter in $osInformation.NetworkInformation.NetworkAdapters) {

        if ($adapter.Description -eq "Remote NDIS Compatible Device") {
            Write-Verbose "Remote NDSI Compatible Device found. Ignoring NIC."
            continue
        }

        $params = $baseParams + @{
            Name                   = "Interface Description"
            Details                = "$($adapter.Description) [$($adapter.Name)]"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        if ($osInformation.BuildInformation.MajorVersion -ge [HealthChecker.OSServerVersion]::Windows2012R2) {
            Write-Verbose "On Windows 2012 R2 or new. Can provide more details on the NICs"

            $driverDate = $adapter.DriverDate
            $detailsValue = $driverDate

            if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
                $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {

                if ($null -eq $driverDate -or
                    $driverDate -eq [DateTime]::MaxValue) {
                    $detailsValue = "Unknown"
                } elseif ((New-TimeSpan -Start $date -End $driverDate).Days -lt [int]-365) {
                    $params = $baseParams + @{
                        Details          = "Warning: NIC driver is over 1 year old. Verify you are at the latest version."
                        DisplayWriteType = "Yellow"
                        AddHtmlDetailRow = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
            }

            $params = $baseParams + @{
                Name    = "Driver Date"
                Details = $detailsValue
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Driver Version"
                Details = $adapter.DriverVersion
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "MTU Size"
                Details = $adapter.MTUSize
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Max Processors"
                Details = $adapter.NetAdapterRss.MaxProcessors
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Max Processor Number"
                Details = $adapter.NetAdapterRss.MaxProcessorNumber
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Name    = "Number of Receive Queues"
                Details = $adapter.NetAdapterRss.NumberOfReceiveQueues
            }
            Add-AnalyzedResultInformation @params

            $writeType = "Yellow"
            $testingValue = $null

            if ($adapter.RssEnabledValue -eq 0) {
                $detailsValue = "False --- Warning: Enabling RSS is recommended."
                $testingValue = $false
            } elseif ($adapter.RssEnabledValue -eq 1) {
                $detailsValue = "True"
                $testingValue = $true
                $writeType = "Green"
            } else {
                $detailsValue = "No RSS Feature Detected."
            }

            $params = $baseParams + @{
                Name                = "RSS Enabled"
                Details             = $detailsValue
                DisplayWriteType    = $writeType
                DisplayTestingValue = $testingValue
            }
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "On Windows 2012 or older and can't get advanced NIC settings"
        }

        $linkSpeed = $adapter.LinkSpeed
        $displayValue = "{0} --- This may not be accurate due to virtualized hardware" -f $linkSpeed

        if ($hardwareInformation.ServerType -eq [HealthChecker.ServerType]::Physical -or
            $hardwareInformation.ServerType -eq [HealthChecker.ServerType]::AmazonEC2) {
            $displayValue = $linkSpeed
        }

        $params = $baseParams + @{
            Name                = "Link Speed"
            Details             = $displayValue
            DisplayTestingValue = $linkSpeed
        }
        Add-AnalyzedResultInformation @params

        $displayValue = "{0}" -f $adapter.IPv6Enabled
        $displayWriteType = "Grey"
        $testingValue = $adapter.IPv6Enabled

        if ($osInformation.NetworkInformation.IPv6DisabledComponents -ne 255 -and
            $adapter.IPv6Enabled -eq $false) {
            $displayValue = "{0} --- Warning" -f $adapter.IPv6Enabled
            $displayWriteType = "Yellow"
            $testingValue = $false
        }

        $params = $baseParams + @{
            Name                = "IPv6 Enabled"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            DisplayTestingValue = $testingValue
        }
        Add-AnalyzedResultInformation @params

        Add-AnalyzedResultInformation -Name "IPv4 Address" @baseParams

        foreach ($address in $adapter.IPv4Addresses) {
            $displayValue = "{0}\{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $params = $baseParams + @{
                Name                   = "Address"
                Details                = $displayValue
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        Add-AnalyzedResultInformation -Name "IPv6 Address" @baseParams

        foreach ($address in $adapter.IPv6Addresses) {
            $displayValue = "{0}\{1}" -f $address.Address, $address.Subnet

            if ($address.DefaultGateway -ne [string]::Empty) {
                $displayValue += " Gateway: {0}" -f $address.DefaultGateway
            }

            $params = $baseParams + @{
                Name                   = "Address"
                Details                = $displayValue
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Name    = "DNS Server"
            Details = $adapter.DnsServer
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name    = "Registered In DNS"
            Details = $adapter.RegisteredInDns
        }
        Add-AnalyzedResultInformation @params

        #Assuming that all versions of Hyper-V doesn't allow sleepy NICs
        if (($hardwareInformation.ServerType -ne [HealthChecker.ServerType]::HyperV) -and ($adapter.PnPCapabilities -ne "MultiplexorNoPnP")) {
            $displayWriteType = "Grey"
            $displayValue = $adapter.SleepyNicDisabled

            if (!$adapter.SleepyNicDisabled) {
                $displayWriteType = "Yellow"
                $displayValue = "False --- Warning: It's recommended to disable NIC power saving options`r`n`t`t`tMore Information: https://aka.ms/HC-NICPowerManagement"
            }

            $params = $baseParams + @{
                Name                = "Sleepy NIC Disabled"
                Details             = $displayValue
                DisplayWriteType    = $displayWriteType
                DisplayTestingValue = $adapter.SleepyNicDisabled
            }
            Add-AnalyzedResultInformation @params
        }

        $adapterDescription = $adapter.Description
        $cookedValue = 0
        $foundCounter = $false

        if ($null -eq $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            Write-Verbose "PacketsReceivedDiscarded is null"
            continue
        }

        foreach ($prdInstance in $osInformation.NetworkInformation.PacketsReceivedDiscarded) {
            $instancePath = $prdInstance.Path
            $startIndex = $instancePath.IndexOf("(") + 1
            $charLength = $instancePath.Substring($startIndex, ($instancePath.IndexOf(")") - $startIndex)).Length
            $instanceName = $instancePath.Substring($startIndex, $charLength)
            $possibleInstanceName = $adapterDescription.Replace("#", "_")

            if ($instanceName -eq $adapterDescription -or
                $instanceName -eq $possibleInstanceName) {
                $cookedValue = $prdInstance.CookedValue
                $foundCounter = $true
                break
            }
        }

        $displayWriteType = "Yellow"
        $displayValue = $cookedValue
        $baseDisplayValue = "{0} --- {1}: This value should be at 0."
        $knownIssue = $false

        if ($foundCounter) {

            if ($cookedValue -eq 0) {
                $displayWriteType = "Green"
            } elseif ($cookedValue -lt 1000) {
                $displayValue = $baseDisplayValue -f $cookedValue, "Warning"
            } else {
                $displayWriteType = "Red"
                $displayValue = [string]::Concat(($baseDisplayValue -f $cookedValue, "Error"), "We are also seeing this value being rather high so this can cause a performance impacted on a system.")
            }

            if ($adapterDescription -like "*vmxnet3*" -and
                $cookedValue -gt 0) {
                $knownIssue = $true
            }
        } else {
            $displayValue = "Couldn't find value for the counter."
            $cookedValue = $null
            $displayWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name                = "Packets Received Discarded"
            Details             = $displayValue
            DisplayWriteType    = $displayWriteType
            DisplayTestingValue = $cookedValue
        }
        Add-AnalyzedResultInformation @params

        if ($knownIssue) {
            $params = $baseParams + @{
                Details                = "Known Issue with vmxnet3: 'Large packet loss at the guest operating system level on the VMXNET3 vNIC in ESXi (2039495)' - https://aka.ms/HC-VMwareLostPackets"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 3
                AddHtmlDetailRow       = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }

    if ($osInformation.NetworkInformation.NetworkAdapters.Count -gt 1) {
        $params = $baseParams + @{
            Details          = "Multiple active network adapters detected. Exchange 2013 or greater may not need separate adapters for MAPI and replication traffic.  For details please refer to https://aka.ms/HC-PlanHA#network-requirements"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    if ($osInformation.NetworkInformation.IPv6DisabledOnNICs) {
        $displayWriteType = "Grey"
        $displayValue = "True"
        $testingValue = $true

        if ($osInformation.NetworkInformation.IPv6DisabledComponents -eq -1) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not correctly disabled via DisabledComponents registry value. It is currently set to '-1'. `r`n`t`tThis setting cause a system startup delay of 5 seconds. For details please refer to: `r`n`t`thttps://aka.ms/HC-ConfigureIPv6"
        } elseif ($osInformation.NetworkInformation.IPv6DisabledComponents -ne 255) {
            $displayWriteType = "Red"
            $testingValue = $false
            $displayValue = "False `r`n`t`tError: IPv6 is disabled on some NIC level settings but not fully disabled. DisabledComponents registry value currently set to '{0}'. For details please refer to the following articles: `r`n`t`thttps://aka.ms/HC-DisableIPv6`r`n`t`thttps://aka.ms/HC-ConfigureIPv6" -f $osInformation.NetworkInformation.IPv6DisabledComponents
        }

        $params = $baseParams + @{
            Name                   = "Disable IPv6 Correctly"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params
    }

    $noDNSRegistered = ($osInformation.NetworkInformation.NetworkAdapters | Where-Object { $_.RegisteredInDns -eq $true }).Count -eq 0

    if ($noDNSRegistered) {
        $params = $baseParams + @{
            Name                   = "No NIC Registered In DNS"
            Details                = "Error: This will cause server to crash and odd mail flow issues. Exchange Depends on the primary NIC to have the setting Registered In DNS set."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerFrequentConfigurationIssues {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation
    $tcpKeepAlive = $osInformation.NetworkInformation.TCPKeepAlive

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Frequent Configuration Issues"  -DisplayOrder $Order)
    }

    if ($tcpKeepAlive -eq 0) {
        $displayValue = "Not Set `r`n`t`tError: Without this value the KeepAliveTime defaults to two hours, which can cause connectivity and performance issues between network devices such as firewalls and load balancers depending on their configuration. `r`n`t`tMore details: https://aka.ms/HC-TcpIpSettingsCheck"
        $displayWriteType = "Red"
    } elseif ($tcpKeepAlive -lt 900000 -or
        $tcpKeepAlive -gt 1800000) {
        $displayValue = "$tcpKeepAlive `r`n`t`tWarning: Not configured optimally, recommended value between 15 to 30 minutes (900000 and 1800000 decimal). `r`n`t`tMore details: https://aka.ms/HC-TcpIpSettingsCheck"
        $displayWriteType = "Yellow"
    } else {
        $displayValue = $tcpKeepAlive
        $displayWriteType = "Green"
    }

    $params = $baseParams + @{
        Name                = "TCP/IP Settings"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $tcpKeepAlive
        HtmlName            = "TCPKeepAlive"
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                = "RPC Min Connection Timeout"
        Details             = "$($osInformation.NetworkInformation.RpcMinConnectionTimeout) `r`n`t`tMore Information: https://aka.ms/HC-RPCSetting"
        DisplayTestingValue = $osInformation.NetworkInformation.RpcMinConnectionTimeout
        HtmlName            = "RPC Minimum Connection Timeout"
    }
    Add-AnalyzedResultInformation @params

    if ($exchangeInformation.RegistryValues.DisableGranularReplication -ne 0) {
        $params = $baseParams + @{
            Name                = "DisableGranularReplication"
            Details             = "$($exchangeInformation.RegistryValues.DisableGranularReplication) - Error this can cause work load management issues."
            DisplayWriteType    = "Red"
            DisplayTestingValue = $true
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name     = "FIPS Algorithm Policy Enabled"
        Details  = $exchangeInformation.RegistryValues.FipsAlgorithmPolicyEnabled
        HtmlName = "FipsAlgorithmPolicy-Enabled"
    }
    Add-AnalyzedResultInformation @params

    $displayValue = $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    $displayWriteType = "Green"

    if ($exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage -ne 0) {
        $displayWriteType = "Red"
        $displayValue = "{0} `r`n`t`tError: This can cause an impact to the server's search performance. This should only be used a temporary fix if no other options are available vs a long term solution." -f $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
    }

    $params = $baseParams + @{
        Name                = "CTS Processor Affinity Percentage"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $exchangeInformation.RegistryValues.CtsProcessorAffinityPercentage
        HtmlName            = "CtsProcessorAffinityPercentage"
    }
    Add-AnalyzedResultInformation @params

    $displayValue = $exchangeInformation.RegistryValues.DisableAsyncNotification
    $displayWriteType = "Grey"

    if ($displayValue -ne 0) {
        $displayWriteType = "Yellow"
        $displayValue = "$($exchangeInformation.RegistryValues.DisableAsyncNotification) Warning: This value should be set back to 0 after you no longer need it for the workaround described in http://support.microsoft.com/kb/5013118"
    }

    $params = $baseParams + @{
        Name                = "Disable Async Notification"
        Details             = $displayValue
        DisplayWriteType    = $displayWriteType
        DisplayTestingValue = $true
    }
    Add-AnalyzedResultInformation @params

    $displayValue = $osInformation.CredentialGuardEnabled
    $displayWriteType = "Grey"

    if ($osInformation.CredentialGuardEnabled) {
        $displayValue = "{0} `r`n`t`tError: Credential Guard is not supported on an Exchange Server. This can cause a performance hit on the server." -f $osInformation.CredentialGuardEnabled
        $displayWriteType = "Red"
    }

    $params = $baseParams + @{
        Name                = "Credential Guard Enabled"
        Details             = $displayValue
        DisplayTestingValue = $osInformation.CredentialGuardEnabled
        DisplayWriteType    = $displayWriteType
    }
    Add-AnalyzedResultInformation @params

    if ($null -ne $exchangeInformation.ApplicationConfigFileStatus -and
        $exchangeInformation.ApplicationConfigFileStatus.Count -ge 1) {

        foreach ($configKey in $exchangeInformation.ApplicationConfigFileStatus.Keys) {
            $configStatus = $exchangeInformation.ApplicationConfigFileStatus[$configKey]
            $writeType = "Green"
            $writeValue = $configStatus.Present

            if (!$configStatus.Present) {
                $writeType = "Red"
                $writeValue = "{0} --- Error" -f $writeValue
            }

            $params = $baseParams + @{
                Name             = "$configKey Present"
                Details          = $writeValue
                DisplayWriteType = $writeType
            }
            Add-AnalyzedResultInformation @params
        }
    }

    $displayWriteType = "Grey"
    $displayValue = "Not Set"
    $additionalDisplayValue = [string]::Empty

    if ($null -ne $exchangeInformation.WildCardAcceptedDomain) {

        if ($exchangeInformation.WildCardAcceptedDomain -eq "Unknown") {
            $displayValue = "Unknown - Unable to run Get-AcceptedDomain"
            $displayWriteType = "Yellow"
        } else {
            $displayWriteType = "Red"
            $domain = $exchangeInformation.WildCardAcceptedDomain
            $displayValue = "Error --- Accepted Domain `"$($domain.Id)`" is set to a Wild Card (*) Domain Name with a domain type of $($domain.DomainType.ToString()). This is not recommended as this is an open relay for the entire environment.`r`n`t`tMore Information: https://aka.ms/HC-OpenRelayDomain"

            if ($domain.DomainType.ToString() -eq "InternalRelay" -and
                (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016 -and
                    $exchangeInformation.BuildInformation.CU -ge [HealthChecker.ExchangeCULevel]::CU22) -or
                ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019 -and
                $exchangeInformation.BuildInformation.CU -ge [HealthChecker.ExchangeCULevel]::CU11))) {

                $additionalDisplayValue = "`r`n`t`tERROR: You have an open relay set as Internal Replay Type and on a CU that is known to cause issues with transport services crashing. Follow the above article for more information."
            } elseif ($domain.DomainType.ToString() -eq "InternalRelay") {
                $additionalDisplayValue = "`r`n`t`tWARNING: You have an open relay set as Internal Relay Type. You are not on a CU yet that is having issue, recommended to change this prior to upgrading. Follow the above article for more information."
            }
        }
    }

    $params = $baseParams + @{
        Name             = "Open Relay Wild Card Domain"
        Details          = $displayValue
        DisplayWriteType = $displayWriteType
    }
    Add-AnalyzedResultInformation @params

    if ($additionalDisplayValue -ne [string]::Empty) {
        $params = $baseParams + @{
            Details          = $additionalDisplayValue
            DisplayWriteType = "Red"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($null -ne $exchangeInformation.IISSettings.IISConfigurationSettings) {
        $iisConfigurationSettings = $exchangeInformation.IISSettings.IISConfigurationSettings |
            Where-Object {
                if ($exchangeInformation.BuildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2016 -or
                    $exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::MultiRole) {
                    return $_
                } elseif ($exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::Mailbox -and
                    $_.Location -like "*ClientAccess*") {
                    return $_
                } elseif ($exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::ClientAccess -and
                    $_.Location -like "*FrontEnd\HttpProxy*") {
                    return $_
                }
            } |
            Where-Object {
                # these are locations that don't by default have configuration files.
                $_.Location -notlike "*\ClientAccess\web.config" -and $_.Location -notlike "*\ClientAccess\exchweb\EWS\bin\web.config" -and
                $_.Location -notlike "*\ClientAccess\Autodiscover\bin\web.config" -and $_.Location -notlike "*\ClientAccess\Autodiscover\help\web.config"
            }

        $missingConfigFile = $iisConfigurationSettings | Where-Object { $_.Exist -eq $false }
        $defaultVariableDetected = $iisConfigurationSettings | Where-Object { $null -ne ($_.Content | Select-String "%ExchangeInstallDir%") }
        $binSearchFoldersNotFound = $iisConfigurationSettings |
            Where-Object { $_.Location -like "*\ClientAccess\ecp\web.config" -and $_.Exist -eq $true } |
            Where-Object {
                $binSearchFolders = $_.Content | Select-String "BinSearchFolders" | Select-Object -ExpandProperty Line
                $startIndex = $binSearchFolders.IndexOf("value=`"") + 7
                $paths = $binSearchFolders.Substring($startIndex, $binSearchFolders.LastIndexOf("`"") - $startIndex).Split(";").Trim().ToLower()
                $paths | ForEach-Object { Write-Verbose "BinSearchFolder: $($_)" }
                $installPath = $exchangeInformation.RegistryValues.MisInstallPath
                foreach ($binTestPath in  @("bin", "bin\CmdletExtensionAgents", "ClientAccess\Owa\bin")) {
                    $testPath = [System.IO.Path]::Combine($installPath, $binTestPath).ToLower()
                    Write-Verbose "Testing path: $testPath"
                    if (-not ($paths.Contains($testPath))) {
                        return $_
                    }
                }
            }

        if ($null -ne $missingConfigFile) {
            $params = $baseParams + @{
                Name                = "Missing Configuration File"
                DisplayWriteType    = "Red"
                DisplayTestingValue = $true
            }
            Add-AnalyzedResultInformation @params

            foreach ($file in $missingConfigFile) {
                $params = $baseParams + @{
                    Details                = "Missing: $($file.Location)"
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details                = "More Information: https://aka.ms/HC-MissingConfig"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $defaultVariableDetected) {
            $params = $baseParams + @{
                Name                = "Default Variable Detected"
                DisplayWriteType    = "Red"
                DisplayTestingValue = $true
            }
            Add-AnalyzedResultInformation @params

            foreach ($file in $defaultVariableDetected) {
                $params = $baseParams + @{
                    Details                = "$($file.Location)"
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details                = "More Information: https://aka.ms/HC-DefaultVariableDetected"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $binSearchFoldersNotFound) {
            $params = $baseParams + @{
                Name                = "Bin Search Folder Not Found"
                DisplayWriteType    = "Red"
                DisplayTestingValue = $true
            }
            Add-AnalyzedResultInformation @params

            foreach ($file in $binSearchFoldersNotFound) {
                $params = $baseParams + @{
                    Details                = "$($file.Location)"
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            $params = $baseParams + @{
                Details                = "More Information: https://aka.ms/HC-BinSearchFolder"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    }
}

function Invoke-AnalyzerWebAppPools {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "Exchange Web App Pools"  -DisplayOrder $Order)
    }

    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
        Write-Verbose "Working on Exchange Web App GC Mode"

        $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]
        foreach ($webAppKey in $exchangeInformation.ApplicationPools.Keys) {

            $appPool = $exchangeInformation.ApplicationPools[$webAppKey]
            $appRestarts = $appPool.AppSettings.add.recycling.periodicRestart
            $appRestartSet = ($appRestarts.PrivateMemory -ne "0" -or
                $appRestarts.Memory -ne "0" -or
                $appRestarts.Requests -ne "0" -or
                $null -ne $appRestarts.Schedule -or
                ($appRestarts.Time -ne "00:00:00" -and
                    ($webAppKey -ne "MSExchangeOWAAppPool" -and
                $webAppKey -ne "MSExchangeECPAppPool")))

            $outputObjectDisplayValue.Add(([PSCustomObject]@{
                        AppPoolName         = $webAppKey
                        State               = $appPool.AppSettings.state
                        GCServerEnabled     = $appPool.GCServerEnabled
                        RestartConditionSet = $appRestartSet
                    })
            )
        }

        $sbStarted = { param($o, $p) if ($p -eq "State") { if ($o."$p" -eq "Started") { "Green" } else { "Red" } } }
        $sbRestart = { param($o, $p) if ($p -eq "RestartConditionSet") { if ($o."$p") { "Red" } else { "Green" } } }
        $params = $baseParams + @{
            OutColumns       = ([PSCustomObject]@{
                    DisplayObject      = $outputObjectDisplayValue
                    ColorizerFunctions = @($sbStarted, $sbRestart)
                    IndentSpaces       = 8
                })
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params

        $periodicStartAppPools = $outputObjectDisplayValue | Where-Object { $_.RestartConditionSet -eq $true }

        if ($null -ne $periodicStartAppPools) {

            $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

            foreach ($appPool in $periodicStartAppPools) {
                $periodicRestart = $exchangeInformation.ApplicationPools[$appPool.AppPoolName].AppSettings.add.recycling.periodicRestart
                $schedule = $periodicRestart.Schedule

                if ([string]::IsNullOrEmpty($schedule)) {
                    $schedule = "null"
                }

                $outputObjectDisplayValue.Add(([PSCustomObject]@{
                            AppPoolName   = $appPool.AppPoolName
                            PrivateMemory = $periodicRestart.PrivateMemory
                            Memory        = $periodicRestart.Memory
                            Requests      = $periodicRestart.Requests
                            Schedule      = $schedule
                            Time          = $periodicRestart.Time
                        }))
            }

            $sbColorizer = {
                param($o, $p)
                switch ($p) {
                    { $_ -in "PrivateMemory", "Memory", "Requests" } {
                        if ($o."$p" -eq "0") { "Green" } else { "Red" }
                    }
                    "Time" {
                        if ($o."$p" -eq "00:00:00") { "Green" } else { "Red" }
                    }
                    "Schedule" {
                        if ($o."$p" -eq "null") { "Green" } else { "Red" }
                    }
                }
            }

            $params = $baseParams + @{
                OutColumns       = ([PSCustomObject]@{
                        DisplayObject      = $outputObjectDisplayValue
                        ColorizerFunctions = @($sbColorizer)
                        IndentSpaces       = 8
                    })
                AddHtmlDetailRow = $false
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details          = "Error: The above app pools currently have the periodic restarts set. This restart will cause disruption to end users."
                DisplayWriteType = "Red"
                AddHtmlDetailRow = $false
            }
            Add-AnalyzedResultInformation @params
        }
    }
}


function Invoke-AnalyzerSecurityExchangeCertificates {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    foreach ($certificate in $exchangeInformation.ExchangeCertificates) {

        if ($certificate.LifetimeInDays -ge 60) {
            $displayColor = "Green"
        } elseif ($certificate.LifetimeInDays -ge 30) {
            $displayColor = "Yellow"
        } else {
            $displayColor = "Red"
        }

        $params = $baseParams + @{
            Name                   = "Certificate"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "FriendlyName"
            Details                = $certificate.FriendlyName
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Thumbprint"
            Details                = $certificate.Thumbprint
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Lifetime in days"
            Details                = $certificate.LifetimeInDays
            DisplayCustomTabNumber = 2
            DisplayWriteType       = $displayColor
        }
        Add-AnalyzedResultInformation @params

        $displayValue = $false
        $displayWriteType = "Grey"
        if ($certificate.LifetimeInDays -lt 0) {
            $displayValue = $true
            $displayWriteType = "Red"
        }

        $params = $baseParams + @{
            Name                   = "Certificate has expired"
            Details                = $displayValue
            DisplayWriteType       = $displayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $certStatusWriteType = [string]::Empty

        if ($null -ne $certificate.Status) {
            switch ($certificate.Status) {
                ("Unknown") { $certStatusWriteType = "Yellow" }
                ("Valid") { $certStatusWriteType = "Grey" }
                ("Revoked") { $certStatusWriteType = "Red" }
                ("DateInvalid") { $certStatusWriteType = "Red" }
                ("Untrusted") { $certStatusWriteType = "Yellow" }
                ("Invalid") { $certStatusWriteType = "Red" }
                ("RevocationCheckFailure") { $certStatusWriteType = "Yellow" }
                ("PendingRequest") { $certStatusWriteType = "Yellow" }
                default { $certStatusWriteType = "Yellow" }
            }

            $params = $baseParams + @{
                Name                   = "Certificate status"
                Details                = $certificate.Status
                DisplayCustomTabNumber = 2
                DisplayWriteType       = $certStatusWriteType
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Certificate status"
                Details                = "Unknown"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($certificate.PublicKeySize -lt 2048) {
            $params = $baseParams + @{
                Name                   = "Key size"
                Details                = $certificate.PublicKeySize
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "It's recommended to use a key size of at least 2048 bit"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Key size"
                Details                = $certificate.PublicKeySize
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($certificate.SignatureHashAlgorithmSecure -eq 1) {
            $shaDisplayWriteType = "Yellow"
        } else {
            $shaDisplayWriteType = "Grey"
        }

        $params = $baseParams + @{
            Name                   = "Signature Algorithm"
            Details                = $certificate.SignatureAlgorithm
            DisplayWriteType       = $shaDisplayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Signature Hash Algorithm"
            Details                = $certificate.SignatureHashAlgorithm
            DisplayWriteType       = $shaDisplayWriteType
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        if ($shaDisplayWriteType -eq "Yellow") {
            $params = $baseParams + @{
                Details                = "It's recommended to use a hash algorithm from the SHA-2 family `r`n`t`tMore information: https://aka.ms/HC-SSLBP"
                DisplayWriteType       = $shaDisplayWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($null -ne $certificate.Services) {
            $params = $baseParams + @{
                Name                   = "Bound to services"
                Details                = $certificate.Services
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
            $params = $baseParams + @{
                Name                   = "Current Auth Certificate"
                Details                = $certificate.IsCurrentAuthConfigCertificate
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }

        $params = $baseParams + @{
            Name                   = "SAN Certificate"
            Details                = $certificate.IsSanCertificate
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                   = "Namespaces"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        foreach ($namespace in $certificate.Namespaces) {
            $params = $baseParams + @{
                Details                = $namespace
                DisplayCustomTabNumber = 3
            }
            Add-AnalyzedResultInformation @params
        }

        if ($certificate.IsCurrentAuthConfigCertificate -eq $true) {
            $currentAuthCertificate = $certificate
        }
    }

    if ($null -ne $currentAuthCertificate) {
        if ($currentAuthCertificate.LifetimeInDays -gt 0) {
            $params = $baseParams + @{
                Name                   = "Valid Auth Certificate Found On Server"
                Details                = $true
                DisplayWriteType       = "Green"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params
        } else {
            $params = $baseParams + @{
                Name                   = "Valid Auth Certificate Found On Server"
                Details                = $false
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 1
            }
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "Auth Certificate has expired `r`n`t`tMore Information: https://aka.ms/HC-OAuthExpired"
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    } elseif ($exchangeInformation.BuildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::Edge) {
        $params = $baseParams + @{
            Name                   = "Valid Auth Certificate Found On Server"
            Details                = $false
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "We can't check for Auth Certificates on Edge Transport Servers"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    } else {
        $params = $baseParams + @{
            Name                   = "Valid Auth Certificate Found On Server"
            Details                = $false
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 1
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "No valid Auth Certificate found. This may cause several problems. `r`n`t`tMore Information: https://aka.ms/HC-FindOAuthHybrid"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerSecurityAMSIConfigState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    $osInformation = $HealthServerObject.OSInformation
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # AMSI integration is only available on Windows Server 2016 or higher and only on
    # Exchange Server 2016 CU21+ or Exchange Server 2019 CU10+.
    # AMSI is also not available on Edge Transport Servers (no http component available).
    if ((($osInformation.BuildInformation.MajorVersion -ge [HealthChecker.OSServerVersion]::Windows2016) -and
        (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU21)) -or
        (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU10))) -and
        ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge)) {

        $amsiInformation = $HealthServerObject.ExchangeInformation.AMSIConfiguration
        $amsiWriteType = "Yellow"
        $amsiConfigurationWarning = "`r`n`t`tThis may pose a security risk to your servers`r`n`t`tMore Information: https://aka.ms/HC-AMSIExchange"

        if (($amsiInformation.Count -eq 1) -and
            ($amsiInformation.QuerySuccessful -eq $true )) {
            $amsiState = $amsiInformation.Enabled
            if ($amsiInformation.Enabled -eq $true) {
                $amsiWriteType = "Green"
            } elseif ($amsiInformation.Enabled -eq $false) {
                switch ($amsiInformation.OrgWideSetting) {
                    ($true) { $additionalAMSIDisplayValue = "Setting applies to all servers of the organization" }
                    ($false) {
                        $additionalAMSIDisplayValue = "Setting applies to the following server(s) of the organization:"
                        foreach ($server in $amsiInformation.Server) {
                            $additionalAMSIDisplayValue += "`r`n`t`t{0}" -f $server
                        }
                    }
                }
                $additionalAMSIDisplayValue += $amsiConfigurationWarning
            } else {
                $additionalAMSIDisplayValue = "Exchange AMSI integration state is unknown"
            }
        } elseif ($amsiInformation.Count -gt 1) {
            $amsiState = "Multiple overrides detected"
            $additionalAMSIDisplayValue = "Exchange AMSI integration state is unknown"
            $i = 0
            foreach ($amsi in $amsiInformation) {
                $i++
                $additionalAMSIDisplayValue += "`r`n`t`tOverride `#{0}" -f $i
                $additionalAMSIDisplayValue += "`r`n`t`t`tName: {0}" -f $amsi.Name
                $additionalAMSIDisplayValue += "`r`n`t`t`tEnabled: {0}" -f $amsi.Enabled
                if ($amsi.OrgWideSetting) {
                    $additionalAMSIDisplayValue += "`r`n`t`t`tSetting applies to all servers of the organization"
                } else {
                    $additionalAMSIDisplayValue += "`r`n`t`t`tSetting applies to the following server(s) of the organization:"
                    foreach ($server in $amsi.Server) {
                        $additionalAMSIDisplayValue += "`r`n`t`t`t{0}" -f $server
                    }
                }
            }
            $additionalAMSIDisplayValue += $amsiConfigurationWarning
        } else {
            $additionalAMSIDisplayValue = "Unable to query Exchange AMSI integration state"
        }

        $params = $baseParams + @{
            Name             = "AMSI Enabled"
            Details          = $amsiState
            DisplayWriteType = $amsiWriteType
        }
        Add-AnalyzedResultInformation @params

        if ($null -ne $additionalAMSIDisplayValue) {
            $params = $baseParams + @{
                Details                = $additionalAMSIDisplayValue
                DisplayWriteType       = $amsiWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "AMSI integration is not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU"
    }
}

function Invoke-AnalyzerSecurityMitigationService {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    $mitigationService = $exchangeInformation.ExchangeEmergencyMitigationService
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }
    #Description: Check for Exchange Emergency Mitigation Service (EEMS)
    #Introduced in: Exchange 2016 CU22, Exchange 2019 CU11
    if (((($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU22)) -or
            (($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
                ($exchangeCU -ge [HealthChecker.ExchangeCULevel]::CU11))) -and
        $exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {

        if (-not([String]::IsNullOrEmpty($mitigationService.MitigationServiceOrgState))) {
            if (($mitigationService.MitigationServiceOrgState) -and
                ($mitigationService.MitigationServiceSrvState)) {
                $eemsWriteType = "Green"
                $eemsOverallState = "Enabled"
            } elseif (($mitigationService.MitigationServiceOrgState -eq $false) -and
                ($mitigationService.MitigationServiceSrvState)) {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled on org level"
            } elseif (($mitigationService.MitigationServiceSrvState -eq $false) -and
                ($mitigationService.MitigationServiceOrgState)) {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled on server level"
            } else {
                $eemsWriteType = "Yellow"
                $eemsOverallState = "Disabled"
            }

            $params = $baseParams + @{
                Name             = "Exchange Emergency Mitigation Service"
                Details          = $eemsOverallState
                DisplayWriteType = $eemsWriteType
            }
            Add-AnalyzedResultInformation @params

            if ($eemsWriteType -ne "Green") {
                $params = $baseParams + @{
                    Details                = "More Information: https://aka.ms/HC-EEMS"
                    DisplayWriteType       = $eemsWriteType
                    DisplayCustomTabNumber = 2
                    AddHtmlDetailRow       = $false
                }
                Add-AnalyzedResultInformation @params
            }

            $eemsWinSrvWriteType = "Yellow"
            if (-not([String]::IsNullOrEmpty($mitigationService.MitigationWinServiceState))) {
                if ($mitigationService.MitigationWinServiceState -eq "Running") {
                    $eemsWinSrvWriteType = "Grey"
                }
                $details = $mitigationService.MitigationWinServiceState
            } else {
                $details = "Unknown"
            }

            $params = $baseParams + @{
                Name                   = "Windows service"
                Details                = $details
                DisplayWriteType       = $eemsWinSrvWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            if ($mitigationService.MitigationServiceEndpoint -eq 200) {
                $eemsPatternServiceWriteType = "Grey"
                $eemsPatternServiceStatus = ("{0} - Reachable" -f $mitigationService.MitigationServiceEndpoint)
            } else {
                $eemsPatternServiceWriteType = "Yellow"
                $eemsPatternServiceStatus = "Unreachable`r`n`t`tMore information: https://aka.ms/HelpConnectivityEEMS"
            }
            $params = $baseParams + @{
                Name                   = "Pattern service"
                Details                = $eemsPatternServiceStatus
                DisplayWriteType       = $eemsPatternServiceWriteType
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params

            if (-not([String]::IsNullOrEmpty($mitigationService.MitigationsApplied))) {
                foreach ($mitigationApplied in $mitigationService.MitigationsApplied) {
                    $params = $baseParams + @{
                        Name                   = "Mitigation applied"
                        Details                = $mitigationApplied
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }

                $params = $baseParams + @{
                    Details                = "Run: 'Get-Mitigations.ps1' from: '$exscripts' to learn more."
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }

            if (-not([String]::IsNullOrEmpty($mitigationService.MitigationsBlocked))) {
                foreach ($mitigationBlocked in $mitigationService.MitigationsBlocked) {
                    $params = $baseParams + @{
                        Name                   = "Mitigation blocked"
                        Details                = $mitigationBlocked
                        DisplayWriteType       = "Yellow"
                        DisplayCustomTabNumber = 2
                    }
                    Add-AnalyzedResultInformation @params
                }
            }

            if (-not([String]::IsNullOrEmpty($mitigationService.DataCollectionEnabled))) {
                $params = $baseParams + @{
                    Name                   = "Telemetry enabled"
                    Details                = $mitigationService.DataCollectionEnabled
                    DisplayCustomTabNumber = 2
                }
                Add-AnalyzedResultInformation @params
            }
        } else {
            Write-Verbose "Unable to validate Exchange Emergency Mitigation Service state"
            $params = $baseParams + @{
                Name             = "Exchange Emergency Mitigation Service"
                Details          = "Failed to query config"
                DisplayWriteType = "Red"
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "Exchange Emergency Mitigation Service feature not available because we are on: $($exchangeInformation.BuildInformation.MajorVersion) $exchangeCU or on Edge Transport Server"
    }
}
function Invoke-AnalyzerSecuritySettings {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $osInformation = $HealthServerObject.OSInformation
    $keySecuritySettings = (Get-DisplayResultsGroupingKey -Name "Security Settings"  -DisplayOrder $Order)
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keySecuritySettings
    }

    ##############
    # TLS Settings
    ##############
    Write-Verbose "Working on TLS Settings"

    function NewDisplayObject {
        param (
            [string]$RegistryKey,
            [string]$Location,
            [object]$Value
        )
        return [PSCustomObject]@{
            RegistryKey = $RegistryKey
            Location    = $Location
            Value       = if ($null -eq $Value) { "NULL" } else { $Value }
        }
    }

    $tlsVersions = @("1.0", "1.1", "1.2", "1.3")
    $currentNetVersion = $osInformation.TLSSettings.Registry.NET["NETv4"]

    $tlsSettings = $osInformation.TLSSettings.Registry.TLS
    $misconfiguredClientServerSettings = ($tlsSettings.Values | Where-Object { $_.TLSMisconfigured -eq $true }).Count -ne 0
    $displayLinkToDocsPage = ($tlsSettings.Values | Where-Object { $_.TLSConfiguration -ne "Enabled" -and $_.TLSConfiguration -ne "Disabled" }).Count -ne 0
    $lowerTlsVersionDisabled = ($tlsSettings.Values | Where-Object { $_.TLSVersionDisabled -eq $true -and ($_.TLSVersion -ne "1.2" -and $_.TLSVersion -ne "1.3") }).Count -ne 0
    $tls13NotDisabled = ($tlsSettings.Values | Where-Object { $_.TLSConfiguration -ne "Disabled" -and $_.TLSVersion -eq "1.3" }).Count -gt 0

    $sbValue = {
        param ($o, $p)
        if ($p -eq "Value") {
            if ($o.$p -eq "NULL" -and -not $o.Location.Contains("1.3")) {
                "Red"
            } elseif ($o.$p -ne "NULL" -and
                $o.$p -ne 1 -and
                $o.$p -ne 0) {
                "Red"
            }
        }
    }

    foreach ($tlsKey in $tlsVersions) {
        $currentTlsVersion = $osInformation.TLSSettings.Registry.TLS[$tlsKey]
        $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]
        $outputObjectDisplayValue.Add((NewDisplayObject "Enabled" -Location $currentTlsVersion.ServerRegistryPath -Value $currentTlsVersion.ServerEnabledValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "DisabledByDefault" -Location $currentTlsVersion.ServerRegistryPath -Value $currentTlsVersion.ServerDisabledByDefaultValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "Enabled" -Location $currentTlsVersion.ClientRegistryPath -Value $currentTlsVersion.ClientEnabledValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "DisabledByDefault" -Location $currentTlsVersion.ClientRegistryPath -Value $currentTlsVersion.ClientDisabledByDefaultValue))
        $displayWriteType = "Green"

        # Any TLS version is Misconfigured or Half Disabled is Red
        # Only TLS 1.2 being Disabled is Red
        # Currently TLS 1.3 being Enabled is Red
        # TLS 1.0 or 1.1 being Enabled is Yellow as we recommend to disable this weak protocol versions
        if (($currentTlsVersion.TLSConfiguration -eq "Misconfigured" -or
                $currentTlsVersion.TLSConfiguration -eq "Half Disabled") -or
                ($tlsKey -eq "1.2" -and $currentTlsVersion.TLSConfiguration -eq "Disabled") -or
                ($tlsKey -eq "1.3" -and $currentTlsVersion.TLSConfiguration -eq "Enabled")) {
            $displayWriteType = "Red"
        } elseif ($currentTlsVersion.TLSConfiguration -eq "Enabled" -and
            ($tlsKey -eq "1.1" -or $tlsKey -eq "1.0")) {
            $displayWriteType = "Yellow"
        }

        $params = $baseParams + @{
            Name             = "TLS $tlsKey"
            Details          = $currentTlsVersion.TLSConfiguration
            DisplayWriteType = $displayWriteType
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            OutColumns           = ([PSCustomObject]@{
                    DisplayObject      = $outputObjectDisplayValue
                    ColorizerFunctions = @($sbValue)
                    IndentSpaces       = 8
                })
            OutColumnsColorTests = @($sbValue)
            HtmlName             = "TLS Settings $tlsKey"
            TestingName          = "TLS Settings Group $tlsKey"
        }
        Add-AnalyzedResultInformation @params
    }

    $netVersions = @("NETv4", "NETv2")
    $outputObjectDisplayValue = New-Object System.Collections.Generic.List[object]

    $sbValue = {
        param ($o, $p)
        if ($p -eq "Value") {
            if ($o.$p -eq "NULL" -and $o.Location -like "*v4.0.30319") {
                "Red"
            }
        }
    }

    foreach ($netVersion in $netVersions) {
        $currentNetVersion = $osInformation.TLSSettings.Registry.NET[$netVersion]
        $outputObjectDisplayValue.Add((NewDisplayObject "SystemDefaultTlsVersions" -Location $currentNetVersion.MicrosoftRegistryLocation -Value $currentNetVersion.SystemDefaultTlsVersionsValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SchUseStrongCrypto" -Location $currentNetVersion.MicrosoftRegistryLocation -Value $currentNetVersion.SchUseStrongCryptoValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SystemDefaultTlsVersions" -Location $currentNetVersion.WowRegistryLocation -Value $currentNetVersion.WowSystemDefaultTlsVersionsValue))
        $outputObjectDisplayValue.Add((NewDisplayObject "SchUseStrongCrypto" -Location $currentNetVersion.WowRegistryLocation -Value $currentNetVersion.WowSchUseStrongCryptoValue))
    }

    $params = $baseParams + @{
        OutColumns  = ([PSCustomObject]@{
                DisplayObject      = $outputObjectDisplayValue
                ColorizerFunctions = @($sbValue)
                IndentSpaces       = 8
            })
        HtmlName    = "TLS NET Settings"
        TestingName = "NET TLS Settings Group"
    }
    Add-AnalyzedResultInformation @params

    $testValues = @("ServerEnabledValue", "ClientEnabledValue", "ServerDisabledByDefaultValue", "ClientDisabledByDefaultValue")

    foreach ($testValue in $testValues) {
        # If value not set to a 0 or a 1.
        $results = $tlsSettings.Values | Where-Object { $null -ne $_."$testValue" -and $_."$testValue" -ne 0 -and $_."$testValue" -ne 1 }

        if ($null -ne $results) {
            $displayLinkToDocsPage = $true
            foreach ($result in $results) {
                $params = $baseParams + @{
                    Name             = "$($result.TLSVersion) $testValue"
                    Details          = "$($result."$testValue") --- Error: Must be a value of 1 or 0."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }

        # if value not defined, we should call that out.
        $results = $tlsSettings.Values | Where-Object { $null -eq $_."$testValue" -and $_.TLSVersion -ne "1.3" }

        if ($null -ne $results) {
            $displayLinkToDocsPage = $true
            foreach ($result in $results) {
                $params = $baseParams + @{
                    Name             = "$($result.TLSVersion) $testValue"
                    Details          = "NULL --- Error: Value should be defined in registry for consistent results."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    # Check for NULL values on NETv4 registry settings
    $testValues = @("SystemDefaultTlsVersionsValue", "SchUseStrongCryptoValue", "WowSystemDefaultTlsVersionsValue", "WowSchUseStrongCryptoValue")

    foreach ($testValue in $testValues) {
        $results = $osInformation.TLSSettings.Registry.NET["NETv4"] | Where-Object { $null -eq $_."$testValue" }
        if ($null -ne $results) {
            $displayLinkToDocsPage = $true
            foreach ($result in $results) {
                $params = $baseParams + @{
                    Name             = "$($result.NetVersion) $testValue"
                    Details          = "NULL --- Error: Value should be defined in registry for consistent results."
                    DisplayWriteType = "Red"
                }
                Add-AnalyzedResultInformation @params
            }
        }
    }

    if ($lowerTlsVersionDisabled -and
        ($osInformation.TLSSettings.Registry.NET["NETv4"].SystemDefaultTlsVersions -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].WowSystemDefaultTlsVersions -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].SchUseStrongCrypto -eq $false -or
        $osInformation.TLSSettings.Registry.NET["NETv4"].WowSchUseStrongCrypto -eq $false)) {
        $params = $baseParams + @{
            Details                = "Error: SystemDefaultTlsVersions or SchUseStrongCrypto is not set to the recommended value. Please visit on how to properly enable TLS 1.2 https://aka.ms/HC-TLSGuide"
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($misconfiguredClientServerSettings) {
        $params = $baseParams + @{
            Details                = "Error: Mismatch in TLS version for client and server. Exchange can be both client and a server. This can cause issues within Exchange for communication."
            DisplayWriteType       = "Red"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "For More Information on how to properly set TLS follow this guide: https://aka.ms/HC-TLSGuide"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Detected TLS Mismatch Display More Info"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($tls13NotDisabled) {
        $displayLinkToDocsPage = $true
        $params = $baseParams + @{
            Details                = "Error: TLS 1.3 is not disabled and not supported currently on Exchange and is known to cause issues within the cluster."
            DisplayWriteType       = "Red"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "TLS 1.3 not disabled"
        }
        Add-AnalyzedResultInformation @params
    }

    if ($lowerTlsVersionDisabled -eq $false) {
        $displayLinkToDocsPage = $true
        $params = $baseParams + @{
            Name = "TLS hardening recommendations"
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Details                = "Microsoft recommends customers proactively address weak TLS usage by removing TLS 1.0/1.1 dependencies in their environments and disabling TLS 1.0/1.1 at the operating system level where possible."
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
        }
        Add-AnalyzedResultInformation @params
    }

    if ($displayLinkToDocsPage) {
        $params = $baseParams + @{
            Details                = "More Information: https://aka.ms/HC-TLSConfigDocs"
            DisplayWriteType       = "Yellow"
            DisplayTestingValue    = $true
            DisplayCustomTabNumber = 2
            TestingName            = "Display Link to Docs Page"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "SecurityProtocol"
        Details = $osInformation.TLSSettings.SecurityProtocol
    }
    Add-AnalyzedResultInformation @params

    if ($null -ne $osInformation.TLSSettings.TlsCipherSuite) {
        $outputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'

        foreach ($tlsCipher in $osInformation.TLSSettings.TlsCipherSuite) {
            $outputObjectDisplayValue.Add(([PSCustomObject]@{
                        TlsCipherSuiteName = $tlsCipher.Name
                        CipherSuite        = $tlsCipher.CipherSuite
                        Cipher             = $tlsCipher.Cipher
                        Certificate        = $tlsCipher.Certificate
                    })
            )
        }

        $params = $baseParams + @{
            OutColumns  = ([PSCustomObject]@{
                    DisplayObject = $outputObjectDisplayValue
                    IndentSpaces  = 8
                })
            HtmlName    = "TLS Cipher Suite"
            TestingName = "TLS Cipher Suite Group"
        }
        Add-AnalyzedResultInformation @params
    }

    $params = $baseParams + @{
        Name    = "LmCompatibilityLevel Settings"
        Details = $osInformation.LmCompatibility.RegistryValue
    }
    Add-AnalyzedResultInformation @params

    $params = $baseParams + @{
        Name                   = "Description"
        Details                = $osInformation.LmCompatibility.Description
        DisplayCustomTabNumber = 2
        AddHtmlDetailRow       = $false
    }
    Add-AnalyzedResultInformation @params

    $additionalDisplayValue = [string]::Empty
    $smb1Settings = $osInformation.Smb1ServerSettings

    if ($osInformation.BuildInformation.MajorVersion -gt [HealthChecker.OSServerVersion]::Windows2012) {
        $displayValue = "False"
        $writeType = "Green"

        if (-not ($smb1Settings.SuccessfulGetInstall)) {
            $displayValue = "Failed to get install status"
            $writeType = "Yellow"
        } elseif ($smb1Settings.Installed) {
            $displayValue = "True"
            $writeType = "Red"
            $additionalDisplayValue = "SMB1 should be uninstalled"
        }

        $params = $baseParams + @{
            Name             = "SMB1 Installed"
            Details          = $displayValue
            DisplayWriteType = $writeType
        }
        Add-AnalyzedResultInformation @params
    }

    $writeType = "Green"
    $displayValue = "True"

    if (-not ($smb1Settings.SuccessfulGetBlocked)) {
        $displayValue = "Failed to get block status"
        $writeType = "Yellow"
    } elseif (-not($smb1Settings.IsBlocked)) {
        $displayValue = "False"
        $writeType = "Red"
        $additionalDisplayValue += " SMB1 should be blocked"
    }

    $params = $baseParams + @{
        Name             = "SMB1 Blocked"
        Details          = $displayValue
        DisplayWriteType = $writeType
    }
    Add-AnalyzedResultInformation @params

    if ($additionalDisplayValue -ne [string]::Empty) {
        $additionalDisplayValue += "`r`n`t`tMore Information: https://aka.ms/HC-SMB1"

        $params = $baseParams + @{
            Details                = $additionalDisplayValue.Trim()
            DisplayWriteType       = "Yellow"
            DisplayCustomTabNumber = 2
            AddHtmlDetailRow       = $false
        }
        Add-AnalyzedResultInformation @params
    }

    Invoke-AnalyzerSecurityExchangeCertificates -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecurityAMSIConfigState -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings
    Invoke-AnalyzerSecurityMitigationService -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecuritySettings

    if ($null -ne $HealthServerObject.ExchangeInformation.BuildInformation.FIPFSUpdateIssue) {
        $fipfsInfoObject = $HealthServerObject.ExchangeInformation.BuildInformation.FIPFSUpdateIssue
        $highestVersion = $fipfsInfoObject.HighesVersionNumberDetected
        $fipfsIssueBaseParams = @{
            Name             = "FIP-FS Update Issue Detected"
            Details          = $true
            DisplayWriteType = "Red"
        }
        $moreInformation = "More Information: https://aka.ms/HC-FIPFSUpdateIssue"

        if ($fipfsInfoObject.ServerRoleAffected -eq $false) {
            # Server role is not affected by the FIP-FS issue so we don't need to check for the other conditions.
            Write-Verbose "The Exchange server runs a role which is not affected by the FIP-FS issue"
        } elseif (($fipfsInfoObject.FIPFSFixedBuild -eq $false) -and
            ($fipfsInfoObject.BadVersionNumberDirDetected)) {
            # Exchange doesn't run a build which is resitent against the problematic pattern
            # and a folder with the problematic version number was detected on the computer.
            $params = $baseParams + $fipfsIssueBaseParams
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = $moreInformation
                DisplayWriteType       = "Red"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } elseif (($fipfsInfoObject.FIPFSFixedBuild) -and
            ($fipfsInfoObject.BadVersionNumberDirDetected)) {
            # Exchange runs a build that can handle the problematic pattern. However, we found
            # a high-version folder which should be removed (recommendation).
            $fipfsIssueBaseParams.DisplayWriteType = "Yellow"
            $params = $baseParams + $fipfsIssueBaseParams
            Add-AnalyzedResultInformation @params

            $params = $baseParams + @{
                Details                = "Detected problematic FIP-FS version $highestVersion directory`r`n`t`tAlthough it should not cause any problems, we recommend performing a FIP-FS reset`r`n`t`t$moreInformation"
                DisplayWriteType       = "Yellow"
                DisplayCustomTabNumber = 2
            }
            Add-AnalyzedResultInformation @params
        } elseif ($null -eq $fipfsInfoObject.HighesVersionNumberDetected) {
            # No scan engine was found on the Exchange server. This will cause multiple issues on transport.
            $fipfsIssueBaseParams.Details = "Error: Failed to find the scan engines on server, this can cause issues with transport rules as well as the malware agent."
            $params = $baseParams + $fipfsIssueBaseParams
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "Server runs a FIP-FS fixed build: $($fipfsInfoObject.FIPFSFixedBuild) - Highest version number: $highestVersion"
        }
    } else {
        $fipfsIssueBaseParams.Details = "Warning: Unable to check if the system is vulnerable to the FIP-FS bad pattern issue. Please re-run. $moreInformation"
        $fipfsIssueBaseParams.DisplayWriteType = "Yellow"
        $params = $baseParams + $fipfsIssueBaseParams
        Add-AnalyzedResultInformation @params
    }
}



function Invoke-AnalyzerSecurityCve-2020-0796 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    #Description: Check for CVE-2020-0796 SMBv3 vulnerability
    #Affected OS versions: Windows 10 build 1903 and 1909
    #Fix: KB4551762
    #Workaround: Disable SMBv3 compression

    if ($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
        Write-Verbose "Testing CVE: CVE-2020-0796"
        $buildNumber = $SecurityObject.OsInformation.BuildInformation.VersionBuild.Split(".")[2]

        if (($buildNumber -eq 18362 -or
                $buildNumber -eq 18363) -and
            ($SecurityObject.OsInformation.RegistryValues.CurrentVersionUbr -lt 720)) {
            Write-Verbose "Build vulnerable to CVE-2020-0796. Checking if workaround is in place."
            $writeType = "Red"
            $writeValue = "System Vulnerable"

            if ($SecurityObject.OsInformation.RegistryValues.LanManServerDisabledCompression -eq 1) {
                Write-Verbose "Workaround to disable affected SMBv3 compression is in place."
                $writeType = "Yellow"
                $writeValue = "Workaround is in place"
            } else {
                Write-Verbose "Workaround to disable affected SMBv3 compression is NOT in place."
            }

            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "CVE-2020-0796"
                Details             = "$writeValue`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796 for more information."
                DisplayWriteType    = $writeType
                DisplayTestingValue = "CVE-2020-0796"
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2020-0796. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/CVE-2020-0796"
        }
    } else {
        Write-Verbose "Operating System NOT vulnerable to CVE-2020-0796."
    }
}

function Invoke-AnalyzerSecurityCve-2020-1147 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    #Description: Check for CVE-2020-1147
    #Affected OS versions: Every OS supporting .NET Core 2.1 and 3.1 and .NET Framework 2.0 SP2 or above
    #Fix: https://portal.msrc.microsoft.com/en-US/security-guidance/advisory/CVE-2020-1147
    #Workaround: N/A
    $dllFileBuildPartToCheckAgainst = 3630

    if ($SecurityObject.OsInformation.NETFramework.NetMajorVersion -eq [HealthChecker.NetMajorVersion]::Net4d8) {
        $dllFileBuildPartToCheckAgainst = 4190
    }

    $systemDataDll = $SecurityObject.OsInformation.NETFramework.FileInformation["System.Data.dll"]
    $systemConfigurationDll = $SecurityObject.OsInformation.NETFramework.FileInformation["System.Configuration.dll"]
    Write-Verbose "System.Data.dll FileBuildPart: $($systemDataDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemDataDll.LastWriteTimeUtc)"
    Write-Verbose "System.Configuration.dll FileBuildPart: $($systemConfigurationDll.VersionInfo.FileBuildPart) | LastWriteTimeUtc: $($systemConfigurationDll.LastWriteTimeUtc)"

    if ($systemDataDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemConfigurationDll.VersionInfo.FileBuildPart -ge $dllFileBuildPartToCheckAgainst -and
        $systemDataDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo)) -and
        $systemConfigurationDll.LastWriteTimeUtc -ge ([System.Convert]::ToDateTime("06/05/2020", [System.Globalization.DateTimeFormatInfo]::InvariantInfo))) {
        Write-Verbose ("System NOT vulnerable to {0}. Information URL: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0}" -f "CVE-2020-1147")
    } else {
        $params = @{
            AnalyzedInformation = $AnalyzeResults
            DisplayGroupingKey  = $DisplayGroupingKey
            Name                = "Security Vulnerability"
            Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/en-us/security-guidance/advisory/{0} for more information." -f "CVE-2020-1147")
            DisplayWriteType    = "Red"
            DisplayTestingValue = "CVE-2020-1147"
            AddHtmlDetailRow    = $false
        }
        Add-AnalyzedResultInformation @params
    }
}

function Invoke-AnalyzerSecurityCve-2021-1730 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: Check for CVE-2021-1730 vulnerability
    #Fix available for: Exchange 2016 CU18+, Exchange 2019 CU7+
    #Fix: Configure Download Domains feature
    #Workaround: N/A

    if (((($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
                ($SecurityObject.CU -ge [HealthChecker.ExchangeCULevel]::CU18)) -or
            (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
                ($SecurityObject.CU -ge [HealthChecker.ExchangeCULevel]::CU7))) -and
        $SecurityObject.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {

        $downloadDomainsEnabled = $SecurityObject.ExchangeInformation.EnableDownloadDomains
        $owaVDirObject = $SecurityObject.ExchangeInformation.GetOwaVirtualDirectory
        $displayWriteType = "Green"

        if (-not ($downloadDomainsEnabled)) {
            $downloadDomainsOrgDisplayValue = "Download Domains are not configured. You should configure them to be protected against CVE-2021-1730.`r`n`t`tConfiguration instructions: https://aka.ms/HC-DownloadDomains"
            $displayWriteType = "Red"
        } else {
            if (-not ([String]::IsNullOrEmpty($OwaVDirObject.ExternalDownloadHostName))) {
                if (($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                            ($OwaVDirObject.ExternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadExternalDisplayValue = "Set to the same as Internal Or External URL as OWA."
                    $displayWriteType = "Red"
                } else {
                    $downloadExternalDisplayValue = "Set Correctly."
                }
            } else {
                $downloadExternalDisplayValue = "Not Configured"
                $displayWriteType = "Red"
            }

            if (-not ([string]::IsNullOrEmpty($owaVDirObject.InternalDownloadHostName))) {
                if (($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.ExternalUrl.Host) -or
                            ($OwaVDirObject.InternalDownloadHostName -eq $OwaVDirObject.InternalUrl.Host)) {
                    $downloadInternalDisplayValue = "Set to the same as Internal Or External URL as OWA."
                    $displayWriteType = "Red"
                } else {
                    $downloadInternalDisplayValue = "Set Correctly."
                }
            } else {
                $displayWriteType = "Red"
                $downloadInternalDisplayValue = "Not Configured"
            }

            $downloadDomainsOrgDisplayValue = "Download Domains are configured.`r`n`t`tExternalDownloadHostName: $downloadExternalDisplayValue`r`n`t`tInternalDownloadHostName: $downloadInternalDisplayValue`r`n`t`tConfiguration instructions: https://aka.ms/HC-DownloadDomains"
        }

        #Only display security vulnerability if present
        if ($displayWriteType -eq "Red") {
            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "Security Vulnerability"
                Details             = $downloadDomainsOrgDisplayValue
                DisplayWriteType    = "Red"
                TestingName         = "CVE-2021-1730"
                DisplayTestingValue = ([PSCustomObject]@{
                        DownloadDomainsEnabled   = $downloadDomainsEnabled
                        ExternalDownloadHostName = $downloadExternalDisplayValue
                        InternalDownloadHostName = $downloadInternalDisplayValue
                    })
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "Download Domains feature not available because we are on: $($SecurityObject.MajorVersion) $($SecurityObject.CU) or on Edge Transport Server"
    }
}

function Invoke-AnalyzerSecurityCve-2021-34470 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: Check for CVE-2021-34470 rights elevation vulnerability
    #Affected Exchange versions: 2013, 2016, 2019
    #Fix:
    ##Exchange 2013 CU23 + July 2021 SU + /PrepareSchema,
    ##Exchange 2016 CU20 + July 2021 SU + /PrepareSchema or CU21,
    ##Exchange 2019 CU9 + July 2021 SU + /PrepareSchema or CU10
    #Workaround: N/A

    if (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) -or
        (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($SecurityObject.CU -lt [HealthChecker.ExchangeCULevel]::CU21)) -or
        (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($SecurityObject.CU -lt [HealthChecker.ExchangeCULevel]::CU10))) {
        Write-Verbose "Testing CVE: CVE-2021-34470"

        $displayWriteTypeColor = $null
        if ($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $SecurityObject.BuildRevision -SecurityFixedBuilds "1497.23" -CVENames "CVE-2021-34470"
        }

        if ($null -eq $SecurityObject.ExchangeInformation.msExchStorageGroup) {
            Write-Verbose "Unable to query classSchema: 'ms-Exch-Storage-Group' information"
            $details = "CVE-2021-34470`r`n`t`tWarning: Unable to query classSchema: 'ms-Exch-Storage-Group' to perform testing."
            $displayWriteTypeColor = "Yellow"
        } elseif ($SecurityObject.ExchangeInformation.msExchStorageGroup.Properties.posssuperiors -eq "computer") {
            Write-Verbose "Attribute: 'possSuperiors' with value: 'computer' detected in classSchema: 'ms-Exch-Storage-Group'"
            $details = "CVE-2021-34470`r`n`t`tPrepareSchema required: https://aka.ms/HC-July21SU"
            $displayWriteTypeColor = "Red"
        } else {
            Write-Verbose "System NOT vulnerable to CVE-2021-34470"
        }

        if ($null -ne $displayWriteTypeColor) {
            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "Security Vulnerability"
                Details             = $details
                DisplayWriteType    = $displayWriteTypeColor
                DisplayTestingValue = "CVE-2021-34470"
                AddHtmlDetailRow    = $false
            }
            Add-AnalyzedResultInformation @params
        }
    } else {
        Write-Verbose "System NOT vulnerable to CVE-2021-34470"
    }
}

function Invoke-AnalyzerSecurityCve-2022-21978 {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $displayWriteTypeColor = $null

    # Description: Check for CVE-2022-21978 vulnerability
    # Affected Exchange versions: 2013, 2016, 2019
    # Fix:
    # Exchange 2013 CU23 + May 2022 SU + /PrepareDomain or /PrepareAllDomains,
    # Exchange 2016 CU22/CU23 + May 2022 SU + /PrepareDomain or /PrepareAllDomains,
    # Exchange 2019 CU11/CU12 + May 2022 SU + /PrepareDomain or /PrepareAllDomains
    # Workaround: N/A

    if ((($SecurityObject.MajorVersion -le [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($SecurityObject.CU -le [HealthChecker.ExchangeCULevel]::CU23)) -or
        (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($SecurityObject.CU -le [HealthChecker.ExchangeCULevel]::CU12)) -and
        ($SecurityObject.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge)) {
        Write-Verbose "Testing CVE: CVE-2022-21978"

        if ($null -ne $SecurityObject.ExchangeInformation.ExchangeAdPermissions) {
            Write-Verbose "Exchange AD permission information found - performing vulnerability testing"
            foreach ($entry in $SecurityObject.ExchangeInformation.ExchangeAdPermissions) {
                if ($entry.CheckPass -eq $false) {
                    $details = "CVE-2022-21978`r`n`t`tInstall the May 2022 SU and run /PrepareDomain or /PrepareAllDomains - See: https://aka.ms/HC-May22SU"
                    $displayWriteTypeColor = "Red"
                }
            }

            if ($displayWriteTypeColor -ne "Red") {
                Write-Verbose "System NOT vulnerable to CVE-2022-21978"
            }
        } else {
            Write-Verbose "Unable to perform CVE-2022-21978 vulnerability testing"
            $details = "CVE-2022-21978`r`n`t`tUnable to perform vulnerability testing. If Exchange admins do not have domain permissions this might be expected, please re-run with domain or enterprise admin account. - See: https://aka.ms/HC-May22SU"
            $displayWriteTypeColor = "Yellow"
        }

        if ($null -ne $displayWriteTypeColor) {
            $params = @{
                AnalyzedInformation = $AnalyzeResults
                DisplayGroupingKey  = $DisplayGroupingKey
                Name                = "Security Vulnerability"
                Details             = $details
                DisplayWriteType    = $displayWriteTypeColor
                DisplayTestingValue = "CVE-2022-21978"
            }
            Add-AnalyzedResultInformation @params
        }
    }
}

function Invoke-AnalyzerSecurityCve-MarchSuSpecial {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    #Description: March 2021 Exchange vulnerabilities Security Update (SU) check for outdated version (CUs)
    #Affected Exchange versions: Exchange 2013, Exchange 2016, Exchange 2016 (we only provide this special SU for these versions)
    #Fix: Update to a supported CU and apply KB5000871
    if (($SecurityObject.ExchangeInformation.BuildInformation.March2021SUInstalled) -and
        ($SecurityObject.ExchangeInformation.BuildInformation.SupportedBuild -eq $false)) {
        if (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) -and
            ($SecurityObject.CU -lt [HealthChecker.ExchangeCULevel]::CU23)) {
            switch ($SecurityObject.CU) {
                ([HealthChecker.ExchangeCULevel]::CU21) { $KBCveComb = @{KB4340731 = "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" } }
                ([HealthChecker.ExchangeCULevel]::CU22) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018" } }
            }
        } elseif (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) -and
            ($SecurityObject.CU -lt [HealthChecker.ExchangeCULevel]::CU18)) {
            switch ($SecurityObject.CU) {
                ([HealthChecker.ExchangeCULevel]::CU8) { $KBCveComb = @{KB4073392 = "CVE-2018-0924", "CVE-2018-0940", "CVE-2018-0941"; KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159" } }
                ([HealthChecker.ExchangeCULevel]::CU9) { $KBCveComb = @{KB4092041 = "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159"; KB4340731 = "CVE-2018-8374", "CVE-2018-8302" } }
                ([HealthChecker.ExchangeCULevel]::CU10) { $KBCveComb = @{KB4340731 = "CVE-2018-8374", "CVE-2018-8302"; KB4459266 = "CVE-2018-8265", "CVE-2018-8448"; KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588" } }
                ([HealthChecker.ExchangeCULevel]::CU11) { $KBCveComb = @{KB4468741 = "CVE-2018-8604"; KB4471389 = "CVE-2019-0586", "CVE-2019-0588"; KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018" } }
                ([HealthChecker.ExchangeCULevel]::CU12) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2018-0858"; KB4503027 = "ADV190018"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" } }
                ([HealthChecker.ExchangeCULevel]::CU13) { $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" } }
                ([HealthChecker.ExchangeCULevel]::CU14) { $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU15) { $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU16) { $KBCveComb = @{KB4577352 = "CVE-2020-16875" } }
                ([HealthChecker.ExchangeCULevel]::CU17) { $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" } }
            }
        } elseif (($SecurityObject.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) -and
            ($SecurityObject.CU -lt [HealthChecker.ExchangeCULevel]::CU7)) {
            switch ($SecurityObject.CU) {
                ([HealthChecker.ExchangeCULevel]::RTM) { $KBCveComb = @{KB4471389 = "CVE-2019-0586", "CVE-2019-0588"; KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018" } }
                ([HealthChecker.ExchangeCULevel]::CU1) { $KBCveComb = @{KB4487563 = "CVE-2019-0817", "CVE-2019-0858"; KB4503027 = "ADV190018"; KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266" } }
                ([HealthChecker.ExchangeCULevel]::CU2) { $KBCveComb = @{KB4509409 = "CVE-2019-1084", "CVE-2019-1137"; KB4515832 = "CVE-2019-1233", "CVE-2019-1266"; KB4523171 = "CVE-2019-1373" } }
                ([HealthChecker.ExchangeCULevel]::CU3) { $KBCveComb = @{KB4523171 = "CVE-2019-1373"; KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU4) { $KBCveComb = @{KB4536987 = "CVE-2020-0688", "CVE-2020-0692"; KB4540123 = "CVE-2020-0903" } }
                ([HealthChecker.ExchangeCULevel]::CU5) { $KBCveComb = @{KB4577352 = "CVE-2020-16875" } }
                ([HealthChecker.ExchangeCULevel]::CU6) { $KBCveComb = @{KB4577352 = "CVE-2020-16875"; KB4581424 = "CVE-2020-16969"; KB4588741 = "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"; KB4593465 = "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143" } }
            }
        } else {
            Write-Verbose "No need to call 'Show-March2021SUOutdatedCUWarning'"
        }
        if ($null -ne $KBCveComb) {
            foreach ($kbName in $KBCveComb.Keys) {
                foreach ($cveName in $KBCveComb[$kbName]) {
                    $params = @{
                        AnalyzedInformation = $AnalyzeResults
                        DisplayGroupingKey  = $DisplayGroupingKey
                        Name                = "March 2021 Exchange Security Update for unsupported CU detected"
                        Details             = "`r`n`t`tPlease make sure $kbName is installed to be fully protected against: $cveName"
                        DisplayWriteType    = "Yellow"
                        DisplayTestingValue = $cveName
                        AddHtmlDetailRow    = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
            }
        }
    }
}

function Invoke-AnalyzerSecurityExtendedProtectionConfigState {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $extendedProtection = $SecurityObject.ExchangeInformation.ExtendedProtectionConfig

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Supported server roles are: Mailbox and ClientAccess
    if (($SecurityObject.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013) -and
            ($SecurityObject.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge)) {

        if ($null -ne $extendedProtection) {
            Write-Verbose "Exchange extended protection information found - performing vulnerability testing"

            # Description: Check for CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134 vulnerability
            # Affected Exchange versions: 2013, 2016, 2019
            # Fix: Install Aug 2022 SU & enable extended protection
            # Extended protection is available with IIS 7.5 or higher
            Write-Verbose "Testing CVE: CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134"
            if (($extendedProtection.ExtendedProtectionConfiguration.ProperlySecuredConfiguration.Contains($false)) -or
                ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                Write-Verbose "At least one vDir is not configured properly and so, the system may be at risk"
                if (($extendedProtection.ExtendedProtectionConfiguration.SupportedExtendedProtection.Contains($false)) -and
                    ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                    # This combination means that EP is configured for at least one vDir, but the Exchange build doesn't support it.
                    # Such a combination can break several things like mailbox access, EMS... .
                    # Recommended action: Disable EP, upgrade to a supported build (Aug 2022 SU+) and enable afterwards.
                    $epDetails = "Extended Protection is configured, but not supported on this Exchange Server build"
                } elseif ((-not($extendedProtection.ExtendedProtectionConfiguration.SupportedExtendedProtection.Contains($false))) -and
                    ($extendedProtection.SupportedVersionForExtendedProtection -eq $false)) {
                    # This combination means that EP is not configured and the Exchange build doesn't support it.
                    # Recommended action: Upgrade to a supported build (Aug 2022 SU+) and enable EP afterwards.
                    $epDetails = "Your Exchange server is at risk. Install the latest SU and enable Extended Protection"
                } else {
                    # This means that EP is supported but not configured for at least one vDir.
                    # Recommended action: Enable EP for each vDir on the system by using the script provided by us.
                    $epDetails += "Extended Protection isn't configured as expected"
                }

                $epCveParams = $baseParams + @{
                    Name             = "Security Vulnerability"
                    Details          = "CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134"
                    DisplayWriteType = "Red"
                }
                $epBasicParams = $baseParams + @{
                    DisplayWriteType       = "Red"
                    DisplayCustomTabNumber = 2
                    Details                = "$epDetails"
                }
                Add-AnalyzedResultInformation @epCveParams
                Add-AnalyzedResultInformation @epBasicParams

                $epFrontEndOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'
                $epBackEndOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'
                $mitigationOutputObjectDisplayValue = New-Object 'System.Collections.Generic.List[object]'

                foreach ($entry in $extendedProtection.ExtendedProtectionConfiguration) {
                    $vDirArray = $entry.VirtualDirectoryName.Split("/", 2)
                    $ssl = $entry.Configuration.SslSettings

                    $listToAdd = $epFrontEndOutputObjectDisplayValue
                    if ($vDirArray[0] -eq "Exchange Back End") {
                        $listToAdd = $epBackEndOutputObjectDisplayValue
                    }

                    $listToAdd.Add(([PSCustomObject]@{
                                $vDirArray[0]     = $vDirArray[1]
                                Value             = $entry.ExtendedProtection
                                SupportedValue    = if ($entry.MitigationSupported -and $entry.MitigationEnabled) { "None" } else { $entry.ExpectedExtendedConfiguration }
                                ConfigSupported   = $entry.ProperlySecuredConfiguration
                                RequireSSL        = "$($ssl.RequireSSL) $(if($ssl.Ssl128Bit) { "(128-bit)" })".Trim()
                                ClientCertificate = $ssl.ClientCertificate
                                IPFilterEnabled   = $entry.MitigationEnabled
                            })
                    )

                    if ($entry.MitigationEnabled) {
                        $mitigationOutputObjectDisplayValue.Add([PSCustomObject]@{
                                VirtualDirectory = $entry.VirtualDirectoryName
                                Details          = $entry.Configuration.MitigationSettings.Restrictions
                            })
                    }
                }

                $epConfig = {
                    param ($o, $p)
                    if ($p -eq "ConfigSupported") {
                        if ($o.$p -ne $true) {
                            "Red"
                        } else {
                            "Green"
                        }
                    } elseif ($p -eq "IPFilterEnabled") {
                        if ($o.$p -eq $true) {
                            "Green"
                        }
                    }
                }

                $epFrontEndParams = $baseParams + @{
                    Name                = "Security Vulnerability"
                    OutColumns          = ([PSCustomObject]@{
                            DisplayObject      = $epFrontEndOutputObjectDisplayValue
                            ColorizerFunctions = @($epConfig)
                            IndentSpaces       = 8
                        })
                    DisplayTestingValue = "CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134"
                }

                $epBackEndParams = $baseParams + @{
                    Name                = "Security Vulnerability"
                    OutColumns          = ([PSCustomObject]@{
                            DisplayObject      = $epBackEndOutputObjectDisplayValue
                            ColorizerFunctions = @($epConfig)
                            IndentSpaces       = 8
                        })
                    DisplayTestingValue = "CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134"
                }

                Add-AnalyzedResultInformation @epFrontEndParams
                Add-AnalyzedResultInformation @epBackEndParams
                if ($mitigationOutputObjectDisplayValue.Count -ge 1) {
                    foreach ($mitigation in $mitigationOutputObjectDisplayValue) {
                        $epMitigationvDir = $baseParams + @{
                            Details          = "$($mitigation.Details.Count) IPs in filter list on vDir: '$($mitigation.VirtualDirectory)'"
                            DisplayWriteType = "Yellow"
                        }
                        Add-AnalyzedResultInformation @epMitigationvDir
                        $mitigationOutputObjectDisplayValue.Details.GetEnumerator() | ForEach-Object {
                            Write-Verbose "IP Address: $($_.key) is allowed to connect? $($_.value)"
                        }
                    }
                }

                $moreInformationParams = $baseParams + @{
                    DisplayWriteType = "Red"
                    Details          = "For more information about Extended Protection and how to configure, please read this article:`n`thttps://aka.ms/HC-ExchangeEPDoc"
                }
                Add-AnalyzedResultInformation @moreInformationParams
            } else {
                Write-Verbose "System NOT vulnerable to CVE-2022-24516, CVE-2022-21979, CVE-2022-21980, CVE-2022-24477, CVE-2022-30134"
            }
        } else {
            Write-Verbose "No Extended Protection configuration found - check will be skipped"
        }
    }
}


function Invoke-AnalyzerSecurityIISModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$SecurityObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exchangeInformation = $SecurityObject.ExchangeInformation
    $moduleInformation = $exchangeInformation.IISSettings.IISModulesInformation

    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $DisplayGroupingKey
    }

    # Description: Check for modules which are loaded by IIS and not signed by Microsoft or not signed at all
    if ($exchangeInformation.BuildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
        if ($null -ne $moduleInformation) {
            $iisModulesOutputList = New-Object 'System.Collections.Generic.List[object]'
            $modulesWriteType = "Grey"

            foreach ($m in $moduleInformation.ModuleList) {
                if ($m.Signed -eq $false) {
                    $modulesWriteType = "Red"

                    $iisModulesOutputList.Add([PSCustomObject]@{
                            Module = $m.Name
                            Path   = $m.Path
                            Signer = "N/A"
                            Status = "Not signed"
                        })
                } elseif (($m.SignatureDetails.IsMicrosoftSigned -eq $false) -or
                    ($m.SignatureDetails.SignatureStatus -ne 0)) {
                    if ($modulesWriteType -ne "Red") {
                        $modulesWriteType = "Yellow"
                    }

                    $iisModulesOutputList.Add([PSCustomObject]@{
                            Module = $m.Name
                            Path   = $m.Path
                            Signer = $m.SignatureDetails.Signer
                            Status = $m.SignatureDetails.SignatureStatus
                        })
                }
            }
            $params = $baseParams + @{
                Name             = "IIS module anomalies detected"
                Details          = ($iisModulesOutputList.Count -ge 1)
                DisplayWriteType = $modulesWriteType
            }
            Add-AnalyzedResultInformation @params

            if ($iisModulesOutputList.Count -ge 1) {
                if ($moduleInformation.AllModulesSigned -eq $false) {
                    $params = $baseParams + @{
                        Details                = "Modules that are loaded by IIS but NOT SIGNED - possibly a security risk"
                        DisplayCustomTabNumber = 2
                        DisplayWriteType       = "Red"
                    }
                    Add-AnalyzedResultInformation @params
                }

                if (($moduleInformation.AllSignedModulesSignedByMSFT -eq $false) -or
                    ($moduleInformation.AllSignaturesValid -eq $false)) {
                    $params = $baseParams + @{
                        Details                = "Modules that are loaded but NOT SIGNED BY Microsoft OR that have a problem with their signature"
                        DisplayCustomTabNumber = 2
                        DisplayWriteType       = "Yellow"
                    }
                    Add-AnalyzedResultInformation @params
                }

                $iisModulesConfig = {
                    param ($o, $p)
                    if ($p -eq "Signer") {
                        if ($o.$p -eq "N/A") {
                            "Red"
                        } else {
                            "Yellow"
                        }
                    } elseif ($p -eq "Status") {
                        if ($o.$p -eq "Not signed") {
                            "Red"
                        } elseif ($o.$p -ne 0) {
                            "Yellow"
                        }
                    }
                }

                $iisModulesParams = $baseParams + @{
                    Name       = "IIS Modules"
                    OutColumns = ([PSCustomObject]@{
                            DisplayObject      = $iisModulesOutputList
                            ColorizerFunctions = @($iisModulesConfig)
                            IndentSpaces       = 8
                        })
                }
                Add-AnalyzedResultInformation @iisModulesParams
            }
        } else {
            Write-Verbose "No modules were returned by previous call"
        }
    } else {
        Write-Verbose "IIS is not available on Edge Transport Server - check will be skipped"
    }
}
function Invoke-AnalyzerSecurityCveCheck {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [object]$DisplayGroupingKey
    )

    function TestVulnerabilitiesByBuildNumbersForDisplay {
        param(
            [Parameter(Mandatory = $true)][string]$ExchangeBuildRevision,
            [Parameter(Mandatory = $true)][array]$SecurityFixedBuilds,
            [Parameter(Mandatory = $true)][array]$CVENames
        )
        [int]$fileBuildPart = ($split = $ExchangeBuildRevision.Split("."))[0]
        [int]$filePrivatePart = $split[1]
        $Script:breakpointHit = $false

        foreach ($securityFixedBuild in $SecurityFixedBuilds) {
            [int]$securityFixedBuildPart = ($split = $securityFixedBuild.Split("."))[0]
            [int]$securityFixedPrivatePart = $split[1]

            if ($fileBuildPart -eq $securityFixedBuildPart) {
                $Script:breakpointHit = $true
            }

            if (($fileBuildPart -lt $securityFixedBuildPart) -or
                    ($fileBuildPart -eq $securityFixedBuildPart -and
                $filePrivatePart -lt $securityFixedPrivatePart)) {
                foreach ($cveName in $CVENames) {
                    $params = @{
                        AnalyzedInformation = $AnalyzeResults
                        DisplayGroupingKey  = $DisplayGroupingKey
                        Name                = "Security Vulnerability"
                        Details             = ("{0}`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/{0} for more information." -f $cveName)
                        DisplayWriteType    = "Red"
                        DisplayTestingValue = $cveName
                        AddHtmlDetailRow    = $false
                    }
                    Add-AnalyzedResultInformation @params
                }
                break
            }

            if ($Script:breakpointHit) {
                break
            }
        }
    }

    $exchangeInformation = $HealthServerObject.ExchangeInformation
    $osInformation = $HealthServerObject.OSInformation

    [string]$buildRevision = ("{0}.{1}" -f $exchangeInformation.BuildInformation.ExchangeSetup.FileBuildPart, `
            $exchangeInformation.BuildInformation.ExchangeSetup.FilePrivatePart)
    $exchangeCU = $exchangeInformation.BuildInformation.CU
    Write-Verbose "Exchange Build Revision: $buildRevision"
    Write-Verbose "Exchange CU: $exchangeCU"

    if ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU19) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1347.5", "1365.3" `
                -CVENames "CVE-2018-0924", "CVE-2018-0940"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU20) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1365.7", "1367.6" `
                -CVENames "CVE-2018-8151", "CVE-2018-8154", "CVE-2018-8159"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU21) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1367.9", "1395.7" `
                -CVENames "CVE-2018-8302"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1395.8" `
                -CVENames "CVE-2018-8265", "CVE-2018-8448"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1395.10" `
                -CVENames "CVE-2019-0586", "CVE-2019-0588"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU22) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1473.3" `
                -CVENames "CVE-2019-0686", "CVE-2019-0724"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1473.4" `
                -CVENames "CVE-2019-0817", "CVE-2019-0858"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1473.5" `
                -CVENames "ADV190018"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU23) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.3" `
                -CVENames "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.4" `
                -CVENames "CVE-2019-1373"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.6" `
                -CVENames "CVE-2020-0688", "CVE-2020-0692"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.7" `
                -CVENames "CVE-2020-16969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.8" `
                -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.10" `
                -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17142", "CVE-2020-17143"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1395.12", "1473.6", "1497.12" `
                -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.12" `
                -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.15" `
                -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.18" `
                -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.23" `
                -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.24" `
                -CVENames "CVE-2021-26427"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.26" `
                -CVENames "CVE-2021-42305", "CVE-2021-41349"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.28" `
                -CVENames "CVE-2022-21855", "CVE-2022-21846", "CVE-2022-21969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.33" `
                -CVENames "CVE-2022-23277"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1497.44" `
                -CVENames "CVE-2022-41040", "CVE-2022-41082", "CVE-2022-41079", "CVE-2022-41078", "CVE-2022-41080"
        }
    } elseif ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU8) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1261.39", "1415.4" `
                -CVENames "CVE-2018-0924", "CVE-2018-0940", "CVE-2018-0941"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU9) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1415.7", "1466.8" `
                -CVENames "CVE-2018-8151", "CVE-2018-8152", "CVE-2018-8153", "CVE-2018-8154", "CVE-2018-8159"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU10) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1466.9", "1531.6" `
                -CVENames "CVE-2018-8374", "CVE-2018-8302"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1531.8" `
                -CVENames "CVE-2018-8265", "CVE-2018-8448"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU11) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1531.8", "1591.11" `
                -CVENames "CVE-2018-8604"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1531.10", "1591.13" `
                -CVENames "CVE-2019-0586", "CVE-2019-0588"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU12) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1591.16", "1713.6" `
                -CVENames "CVE-2019-0817", "CVE-2018-0858"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1591.17", "1713.7" `
                -CVENames "ADV190018"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1713.5" `
                -CVENames "CVE-2019-0686", "CVE-2019-0724"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU13) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1713.8", "1779.4" `
                -CVENames "CVE-2019-1084", "CVE-2019-1136", "CVE-2019-1137"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1713.9", "1779.5" `
                -CVENames "CVE-2019-1233", "CVE-2019-1266"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU14) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1779.7", "1847.5" `
                -CVENames "CVE-2019-1373"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU15) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1847.7", "1913.7" `
                -CVENames "CVE-2020-0688", "CVE-2020-0692"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1847.10", "1913.10" `
                -CVENames "CVE-2020-0903"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU17) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1979.6", "2044.6" `
                -CVENames "CVE-2020-16875"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU18) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2106.2" `
                -CVENames "CVE-2021-1730"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2044.7", "2106.3" `
                -CVENames "CVE-2020-16969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2044.8", "2106.4" `
                -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2044.12", "2106.6" `
                -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU19) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2106.8", "2176.4" `
                -CVENames "CVE-2021-24085"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "1415.8", "1466.13", "1531.12", "1591.18", "1713.10", "1779.8", "1847.12", "1913.12", "1979.8", "2044.13", "2106.13", "2176.9" `
                -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2106.13", "2176.9" `
                -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU20) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2176.12", "2242.8" `
                -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2176.14", "2242.10" `
                -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU21) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2242.12", "2308.14" `
                -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU22) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2308.15", "2375.7" `
                -CVENames "CVE-2021-26427"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2308.15", "2375.12" `
                -CVENames "CVE-2021-41350", "CVE-2021-41348", "CVE-2021-34453"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2308.20", "2375.17" `
                -CVENames "CVE-2021-42305", "CVE-2021-41349", "CVE-2021-42321"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2308.21", "2375.18" `
                -CVENames "CVE-2022-21855", "CVE-2022-21846", "CVE-2022-21969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2308.27", "2375.24" `
                -CVENames "CVE-2022-23277", "CVE-2022-24463"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU23) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2375.31", "2507.12" `
                -CVENames "CVE-2022-34692"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "2375.37", "2507.16" `
                -CVENames "CVE-2022-41040", "CVE-2022-41082", "CVE-2022-41079", "CVE-2022-41078", "CVE-2022-41080", "CVE-2022-41123"
        }
    } elseif ($exchangeInformation.BuildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU1) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "221.14" `
                -CVENames "CVE-2019-0586", "CVE-2019-0588"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "221.16", "330.7" `
                -CVENames "CVE-2019-0817", "CVE-2019-0858"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "221.17", "330.8" `
                -CVENames "ADV190018"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "330.6" `
                -CVENames "CVE-2019-0686", "CVE-2019-0724"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU2) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "330.9", "397.5" `
                -CVENames "CVE-2019-1084", "CVE-2019-1137"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "397.6", "330.10" `
                -CVENames "CVE-2019-1233", "CVE-2019-1266"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU3) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "397.9", "464.7" `
                -CVENames "CVE-2019-1373"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU4) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "464.11", "529.8" `
                -CVENames "CVE-2020-0688", "CVE-2020-0692"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "464.14", "529.11" `
                -CVENames "CVE-2020-0903"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU6) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "595.6", "659.6" `
                -CVENames "CVE-2020-16875"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU7) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "721.2" `
                -CVENames "CVE-2021-1730"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "659.7", "721.3" `
                -CVENames "CVE-2020-16969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "659.8", "721.4" `
                -CVENames "CVE-2020-17083", "CVE-2020-17084", "CVE-2020-17085"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "659.11", "721.6" `
                -CVENames "CVE-2020-17117", "CVE-2020-17132", "CVE-2020-17141", "CVE-2020-17142", "CVE-2020-17143"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU8) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "721.8", "792.5" `
                -CVENames "CVE-2021-24085"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "221.18", "330.11", "397.11", "464.15", "529.13", "595.8", "659.12", "721.13", "792.10" `
                -CVENames "CVE-2021-26855", "CVE-2021-26857", "CVE-2021-26858", "CVE-2021-27065"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "721.13", "792.10" `
                -CVENames "CVE-2021-26412", "CVE-2021-27078", "CVE-2021-26854"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU9) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "792.13", "858.10" `
                -CVENames "CVE-2021-28480", "CVE-2021-28481", "CVE-2021-28482", "CVE-2021-28483"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "792.15", "858.12" `
                -CVENames "CVE-2021-31195", "CVE-2021-31198", "CVE-2021-31207", "CVE-2021-31209"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU10) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "858.15", "922.13" `
                -CVENames "CVE-2021-31206", "CVE-2021-31196", "CVE-2021-33768"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU11) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "922.14", "986.5" `
                -CVENames "CVE-2021-26427"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "922.14", "986.9" `
                -CVENames "CVE-2021-41350", "CVE-2021-41348", "CVE-2021-34453"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "922.19", "986.14" `
                -CVENames "CVE-2021-42305", "CVE-2021-41349", "CVE-2021-42321"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "922.20", "986.15" `
                -CVENames "CVE-2022-21855", "CVE-2022-21846", "CVE-2022-21969"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "922.27", "986.22" `
                -CVENames "CVE-2022-23277", "CVE-2022-24463"
        }

        if ($exchangeCU -le [HealthChecker.ExchangeCULevel]::CU12) {
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "986.29", "1118.12" `
                -CVENames "CVE-2022-34692"
            TestVulnerabilitiesByBuildNumbersForDisplay -ExchangeBuildRevision $buildRevision `
                -SecurityFixedBuilds "986.36", "1118.20" `
                -CVENames "CVE-2022-41040", "CVE-2022-41082", "CVE-2022-41079", "CVE-2022-41078", "CVE-2022-41080", "CVE-2022-41123"
        }
    } else {
        Write-Verbose "Unknown Version of Exchange"
    }

    $securityObject = [PSCustomObject]@{
        MajorVersion        = $exchangeInformation.BuildInformation.MajorVersion
        ServerRole          = $exchangeInformation.BuildInformation.ServerRole
        CU                  = $exchangeCU
        BuildRevision       = $buildRevision
        ExchangeInformation = $exchangeInformation
        OsInformation       = $osInformation
    }

    Invoke-AnalyzerSecurityIISModules -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2020-0796 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2020-1147 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2021-1730 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2021-34470 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-2022-21978 -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    Invoke-AnalyzerSecurityCve-MarchSuSpecial -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
    # Make sure that these stay as the last one to keep the output more readable
    Invoke-AnalyzerSecurityExtendedProtectionConfigState -AnalyzeResults $AnalyzeResults -SecurityObject $securityObject -DisplayGroupingKey $DisplayGroupingKey
}
function Invoke-AnalyzerSecurityVulnerability {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ref]$AnalyzeResults,

        [Parameter(Mandatory = $true)]
        [object]$HealthServerObject,

        [Parameter(Mandatory = $true)]
        [int]$Order
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $keySecurityVulnerability = Get-DisplayResultsGroupingKey -Name "Security Vulnerability"  -DisplayOrder $Order
    $baseParams = @{
        AnalyzedInformation = $AnalyzeResults
        DisplayGroupingKey  = $keySecurityVulnerability
    }

    Invoke-AnalyzerSecurityCveCheck -AnalyzeResults $AnalyzeResults -HealthServerObject $HealthServerObject -DisplayGroupingKey $keySecurityVulnerability

    $allSecurityVulnerabilities = $AnalyzeResults.Value.DisplayResults[$keySecurityVulnerability]
    $securityVulnerabilities = $allSecurityVulnerabilities | Where-Object { $_.Name -ne "IIS module anomalies detected" }
    $iisModule = $allSecurityVulnerabilities | Where-Object { $_.Name -eq "IIS module anomalies detected" }

    if ($null -eq $securityVulnerabilities -and
        ($null -ne $iisModule -or $iisModule.DisplayValue -eq $false)) {
        $params = $baseParams + @{
            Details          = "All known security issues in this version of the script passed."
            DisplayWriteType = "Green"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                      = "Security Vulnerabilities"
            Details                   = "None"
            AddDisplayResultsLineInfo = $false
            AddHtmlOverviewValues     = $true
        }
        Add-AnalyzedResultInformation @params
    } elseif ($null -ne $securityVulnerabilities -or
        ($null -ne $iisModule -and $iisModule.DisplayValue -eq $true)) {

        $details = $securityVulnerabilities.DisplayValue |
            ForEach-Object {
                return $_ + "<br>"
            }

        # If details are null, but iisModule is showing a vulnerability,
        # then just provide see IIS Module section
        if ($null -eq $details) { $details = "See IIS module anomalies detected section above" }

        $params = $baseParams + @{
            Name                      = "Security Vulnerabilities"
            Details                   = $details
            DisplayWriteType          = "Red"
            AddDisplayResultsLineInfo = $false
        }
        Add-AnalyzedResultInformation @params

        $params = $baseParams + @{
            Name                      = "Vulnerability Detected"
            Details                   = $true
            AddDisplayResultsLineInfo = $false
            DisplayWriteType          = "Red"
            AddHtmlOverviewValues     = $true
            AddHtmlDetailRow          = $false
        }
        Add-AnalyzedResultInformation @params
    }
}
function Invoke-AnalyzerEngine {
    param(
        [HealthChecker.HealthCheckerExchangeServer]$HealthServerObject
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $analyzedResults = New-Object HealthChecker.AnalyzedInformation
    $analyzedResults.HealthCheckerExchangeServer = $HealthServerObject

    #Display Grouping Keys
    $order = 1
    $baseParams = @{
        AnalyzedInformation = $analyzedResults
        DisplayGroupingKey  = (Get-DisplayResultsGroupingKey -Name "BeginningInfo" -DisplayGroupName $false -DisplayOrder 0 -DefaultTabNumber 0)
    }

    if (!$Script:DisplayedScriptVersionAlready) {
        $params = $baseParams + @{
            Name             = "Exchange Health Checker Version"
            Details          = $BuildVersion
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    $VirtualizationWarning = @"
Virtual Machine detected.  Certain settings about the host hardware cannot be detected from the virtual machine.  Verify on the VM Host that:

    - There is no more than a 1:1 Physical Core to Virtual CPU ratio (no oversubscribing)
    - If Hyper-Threading is enabled do NOT count Hyper-Threaded cores as physical cores
    - Do not oversubscribe memory or use dynamic memory allocation

Although Exchange technically supports up to a 2:1 physical core to vCPU ratio, a 1:1 ratio is strongly recommended for performance reasons.  Certain third party Hyper-Visors such as VMWare have their own guidance.

VMWare recommends a 1:1 ratio.  Their guidance can be found at https://aka.ms/HC-VMwareBP2019.
Related specifically to VMWare, if you notice you are experiencing packet loss on your VMXNET3 adapter, you may want to review the following article from VMWare:  https://aka.ms/HC-VMwareLostPackets.

For further details, please review the virtualization recommendations on Microsoft Docs here: https://aka.ms/HC-Virtualization.

"@

    if ($HealthServerObject.HardwareInformation.ServerType -eq [HealthChecker.ServerType]::VMWare -or
        $HealthServerObject.HardwareInformation.ServerType -eq [HealthChecker.ServerType]::HyperV) {
        $params = $baseParams + @{
            Details          = $VirtualizationWarning
            DisplayWriteType = "Yellow"
            AddHtmlDetailRow = $false
        }
        Add-AnalyzedResultInformation @params
    }

    Invoke-AnalyzerExchangeInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerHybridInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerOsInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerHardwareInformation -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerNicSettings -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerFrequentConfigurationIssues -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerSecuritySettings -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerSecurityVulnerability -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Invoke-AnalyzerWebAppPools -AnalyzeResults ([ref]$analyzedResults) -HealthServerObject $HealthServerObject -Order ($order++)
    Write-Debug("End of Analyzer Engine")
    return $analyzedResults
}




function Get-ExtendedProtectionConfiguration {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $false)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        [Parameter(Mandatory = $false)]
        [System.Version]$ExSetupVersion,

        [Parameter(Mandatory = $false)]
        [bool]$IsMailboxServer = $true,

        [Parameter(Mandatory = $false)]
        [bool]$IsClientAccessServer = $true,

        [Parameter(Mandatory = $false)]
        [bool]$ExcludeEWS = $false,

        [Parameter(Mandatory = $false)]
        [ValidateSet("Exchange Back End/EWS")]
        [string[]]$SiteVDirLocations,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        function NewVirtualDirMatchingEntry {
            param(
                [Parameter(Mandatory = $true)]
                [string]$VirtualDirectory,
                [Parameter(Mandatory = $true)]
                [ValidateSet("Default Web Site", "Exchange Back End")]
                [string[]]$WebSite,
                [Parameter(Mandatory = $true)]
                [ValidateSet("None", "Allow", "Require")]
                [string[]]$ExtendedProtection,
                # Need to define this twice once for Default Web Site and Exchange Back End for the default values
                [Parameter(Mandatory = $false)]
                [string[]]$SslFlags = @("Ssl,Ssl128", "Ssl,Ssl128")
            )

            if ($WebSite.Count -ne $ExtendedProtection.Count) {
                throw "Argument count mismatch on $VirtualDirectory"
            }

            for ($i = 0; $i -lt $WebSite.Count; $i++) {
                # special conditions for Exchange 2013
                # powershell is on front and back so skip over those
                if ($IsExchange2013 -and $virtualDirectory -ne "Powershell") {
                    # No API virtual directory
                    if ($virtualDirectory -eq "API") { return }
                    if ($IsClientAccessServer -eq $false -and $WebSite[$i] -eq "Default Web Site") { continue }
                    if ($IsMailboxServer -eq $false -and $WebSite[$i] -eq "Exchange Back End") { continue }
                }
                # Set EWS Vdir to None for known issues
                if ($ExcludeEWS -and $virtualDirectory -eq "EWS") { $ExtendedProtection[$i] = "None" }

                if ($null -ne $SiteVDirLocations -and
                    $SiteVDirLocations.Count -gt 0) {
                    foreach ($SiteVDirLocation in $SiteVDirLocations) {
                        if ($SiteVDirLocation -eq "$($WebSite[$i])/$virtualDirectory") {
                            Write-Verbose "Set Extended Protection to None because of restriction override '$($WebSite[$i])\$virtualDirectory'"
                            $ExtendedProtection[$i] = "None"
                            break;
                        }
                    }
                }

                [PSCustomObject]@{
                    VirtualDirectory   = $virtualDirectory
                    WebSite            = $WebSite[$i]
                    ExtendedProtection = $ExtendedProtection[$i]
                    SslFlags           = $SslFlags[$i]
                }
            }
        }

        # Intended for inside of Invoke-Command.
        function GetApplicationHostConfig {
            $appHostConfig = New-Object -TypeName Xml
            try {
                $appHostConfigPath = "$($env:WINDIR)\System32\inetsrv\config\applicationHost.config"
                $appHostConfig.Load($appHostConfigPath)
            } catch {
                Write-Verbose "Failed to loaded application host config file. $_"
                $appHostConfig = $null
            }
            return $appHostConfig
        }

        function GetExtendedProtectionConfiguration {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [System.Xml.XmlNode]$Xml,
                [Parameter(Mandatory = $true)]
                [string]$Path
            )
            process {
                try {
                    $nodePath = [string]::Empty
                    $extendedProtection = "None"
                    $ipRestictionsHashTable = @{}
                    $pathIndex = [array]::IndexOf(($Xml.configuration.location.path).ToLower(), $Path.ToLower())
                    $rootIndex = [array]::IndexOf(($Xml.configuration.location.path).ToLower(), ($Path.Split("/")[0]).ToLower())

                    if ($pathIndex -ne -1) {
                        $configNode = $Xml.configuration.location[$pathIndex]
                        $nodePath = $configNode.Path
                        $ep = $configNode.'system.webServer'.security.authentication.windowsAuthentication.extendedProtection.tokenChecking
                        $ipRestrictions = $configNode.'system.webServer'.security.ipSecurity

                        if (-not ([string]::IsNullOrEmpty($ep))) {
                            Write-Verbose "Found tokenChecking: $ep"
                            $extendedProtection = $ep
                        } else {
                            Write-Verbose "Failed to find tokenChecking. Using default value of None."
                        }

                        [string]$sslSettings = $configNode.'system.webServer'.security.access.sslFlags

                        if ([string]::IsNullOrEmpty($sslSettings)) {
                            Write-Verbose "Failed to find SSL settings for the path. Falling back to the root."

                            if ($rootIndex -ne -1) {
                                Write-Verbose "Found root path."
                                $rootConfigNode = $Xml.configuration.location[$rootIndex]
                                [string]$sslSettings = $rootConfigNode.'system.webServer'.security.access.sslFlags
                            }
                        }

                        if (-not([string]::IsNullOrEmpty($ipRestrictions))) {
                            Write-Verbose "IP-filtered restrictions detected"
                            foreach ($restriction in $ipRestrictions.add) {
                                $ipRestictionsHashTable.Add($restriction.ipAddress, $restriction.allowed)
                            }
                        }

                        Write-Verbose "SSLSettings: $sslSettings"

                        if ($null -ne $sslSettings) {
                            [array]$sslFlags = ($sslSettings.Split(",").ToLower()).Trim()
                        } else {
                            $sslFlags = $null
                        }

                        # SSL flags: https://docs.microsoft.com/iis/configuration/system.webserver/security/access#attributes
                        $requireSsl = $false
                        $ssl128Bit = $false
                        $clientCertificate = "Unknown"

                        if ($null -eq $sslFlags) {
                            Write-Verbose "Failed to find SSLFlags"
                        } elseif ($sslFlags.Contains("none")) {
                            $clientCertificate = "Ignore"
                        } else {
                            if ($sslFlags.Contains("ssl")) { $requireSsl = $true }
                            if ($sslFlags.Contains("ssl128")) { $ssl128Bit = $true }
                            if ($sslFlags.Contains("sslnegotiatecert")) {
                                $clientCertificate = "Accept"
                            } elseif ($sslFlags.Contains("sslrequirecert")) {
                                $clientCertificate = "Require"
                            } else {
                                $clientCertificate = "Ignore"
                            }
                        }
                    }
                } catch {
                    Write-Verbose "Ran into some error trying to parse the application host config for $Path."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } end {
                return [PSCustomObject]@{
                    ExtendedProtection = $extendedProtection
                    ValidPath          = ($pathIndex -ne -1)
                    NodePath           = $nodePath
                    SslSettings        = [PSCustomObject]@{
                        RequireSsl        = $requireSsl
                        Ssl128Bit         = $ssl128Bit
                        ClientCertificate = $clientCertificate
                        Value             = $sslSettings
                    }
                    MitigationSettings = [PScustomObject]@{
                        AllowUnlisted = $ipRestrictions.allowUnlisted
                        Restrictions  = $ipRestictionsHashTable
                    }
                }
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $computerResult = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlock { return $env:COMPUTERNAME }
        $serverConnected = $null -ne $computerResult

        if ($null -eq $computerResult) {
            Write-Verbose "Failed to connect to server $ComputerName"
            return
        }

        if ($null -eq $ExSetupVersion) {
            [System.Version]$ExSetupVersion = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlock {
                (Get-Command Exsetup.exe |
                    ForEach-Object { $_.FileVersionInfo } |
                    Select-Object -First 1).FileVersion
            }

            if ($null -eq $ExSetupVersion) {
                throw "Failed to determine Exchange build number"
            }
        } else {
            # Hopefully the caller knows what they are doing, best be from the correct server!!
            Write-Verbose "Caller passed the ExSetupVersion information"
        }

        if ($null -eq $ApplicationHostConfig) {
            Write-Verbose "Trying to load the application host config from $ComputerName"
            $params = @{
                ComputerName        = $ComputerName
                ScriptBlock         = ${Function:GetApplicationHostConfig}
                CatchActionFunction = $CatchActionFunction
            }

            $ApplicationHostConfig = Invoke-ScriptBlockHandler @params

            if ($null -eq $ApplicationHostConfig) {
                throw "Failed to load application host config from $ComputerName"
            }
        } else {
            # Hopefully the caller knows what they are doing, best be from the correct server!!
            Write-Verbose "Caller passed the application host config."
        }

        $default = "Default Web Site"
        $backend = "Exchange Back End"
        $Script:IsExchange2013 = $ExSetupVersion.Major -eq 15 -and $ExSetupVersion.Minor -eq 0
        try {
            $VirtualDirectoryMatchEntries = @(
                (NewVirtualDirMatchingEntry "API" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "Autodiscover" -WebSite $default, $backend -ExtendedProtection "None", "None")
                (NewVirtualDirMatchingEntry "ECP" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "EWS" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                (NewVirtualDirMatchingEntry "Microsoft-Server-ActiveSync" -WebSite $default, $backend -ExtendedProtection "Allow", "Require")
                (NewVirtualDirMatchingEntry "OAB" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "Powershell" -WebSite $default, $backend -ExtendedProtection "Require", "Require" -SslFlags "SslNegotiateCert", "Ssl,Ssl128,SslNegotiateCert")
                (NewVirtualDirMatchingEntry "OWA" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "RPC" -WebSite $default, $backend -ExtendedProtection "Require", "Require")
                (NewVirtualDirMatchingEntry "MAPI" -WebSite $default -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "PushNotifications" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "RPCWithCert" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "MAPI/emsmdb" -WebSite $backend -ExtendedProtection "Require")
                (NewVirtualDirMatchingEntry "MAPI/nspi" -WebSite $backend -ExtendedProtection "Require")
            )
        } catch {
            # Don't handle with Catch Error as this is a bug in the script.
            throw "Failed to create NewVirtualDirMatchingEntry. Inner Exception $_"
        }

        # Is Supported build of Exchange to have the configuration set.
        # Edge Server is not accounted for. It is the caller's job to not try to collect this info on Edge.
        $supportedVersion = $false
        $extendedProtectionList = New-Object 'System.Collections.Generic.List[object]'

        if ($ExSetupVersion.Major -eq 15) {
            if ($ExSetupVersion.Minor -eq 2) {
                $supportedVersion = $ExSetupVersion.Build -gt 1118 -or
                ($ExSetupVersion.Build -eq 1118 -and $ExSetupVersion.Revision -ge 11) -or
                ($ExSetupVersion.Build -eq 986 -and $ExSetupVersion.Revision -ge 28)
            } elseif ($ExSetupVersion.Minor -eq 1) {
                $supportedVersion = $ExSetupVersion.Build -gt 2507 -or
                ($ExSetupVersion.Build -eq 2507 -and $ExSetupVersion.Revision -ge 11) -or
                ($ExSetupVersion.Build -eq 2375 -and $ExSetupVersion.Revision -ge 30)
            } elseif ($ExSetupVersion.Minor -eq 0) {
                $supportedVersion = $ExSetupVersion.Build -gt 1497 -or
                ($ExSetupVersion.Build -eq 1497 -and $ExSetupVersion.Revision -ge 38)
            }
            Write-Verbose "Build $ExSetupVersion is supported: $supportedVersion"
        } else {
            Write-Verbose "Not on Exchange Version 15"
        }

        # Add all vDirs for which the IP filtering mitigation is supported
        $mitigationSupportedvDirs = $MyInvocation.MyCommand.Parameters["SiteVDirLocations"].Attributes |
            Where-Object { $_ -is [System.Management.Automation.ValidateSetAttribute] } |
            ForEach-Object { return $_.ValidValues.ToLower() }
        Write-Verbose "Supported mitigated virtual directories: $([string]::Join(",", $mitigationSupportedvDirs))"
    }
    process {
        try {
            foreach ($matchEntry in $VirtualDirectoryMatchEntries) {
                try {
                    Write-Verbose "Verify extended protection setting for $($matchEntry.VirtualDirectory) on web site $($matchEntry.WebSite)"

                    $extendedConfiguration = GetExtendedProtectionConfiguration -Xml $applicationHostConfig -Path "$($matchEntry.WebSite)/$($matchEntry.VirtualDirectory)"

                    # Extended Protection is a windows security feature which blocks MiTM attacks.
                    # Supported server roles are: Mailbox and ClientAccess
                    # Possible configuration settings are:
                    # <None>: This value specifies that IIS will not perform channel-binding token checking.
                    # <Allow>: This value specifies that channel-binding token checking is enabled, but not required.
                    # <Require>: This value specifies that channel-binding token checking is required.
                    # https://docs.microsoft.com/iis/configuration/system.webserver/security/authentication/windowsauthentication/extendedprotection/

                    if ($extendedConfiguration.ValidPath) {
                        Write-Verbose "Configuration was successfully returned: $($extendedConfiguration.ExtendedProtection)"
                    } else {
                        Write-Verbose "Extended protection setting was not queried because it wasn't found on the system."
                    }

                    $sslFlagsToSet = $extendedConfiguration.SslSettings.Value
                    $currentSetFlags = $sslFlagsToSet.Split(",").Trim()
                    foreach ($sslFlag in $matchEntry.SslFlags.Split(",").Trim()) {
                        if (-not($currentSetFlags.Contains($sslFlag))) {
                            Write-Verbose "Failed to find SSL Flag $sslFlag"
                            # We do not want to include None in the flags as that takes priority over the other options.
                            if ($sslFlagsToSet -eq "None") {
                                $sslFlagsToSet = "$sslFlag"
                            } else {
                                $sslFlagsToSet += ",$sslFlag"
                            }
                            Write-Verbose "Updated SSL Flags Value: $sslFlagsToSet"
                        } else {
                            Write-Verbose "SSL Flag $sslFlag set."
                        }
                    }

                    $expectedExtendedConfiguration = if ($supportedVersion) { $matchEntry.ExtendedProtection } else { "None" }
                    $virtualDirectoryName = "$($matchEntry.WebSite)/$($matchEntry.VirtualDirectory)"

                    # Properly Secured Configuration is only a concern if Required is the Expected value
                    # If the Expected value is None or Allow, you can have it configured however you would like and from a security standpoint, it shouldn't be a concern.
                    # For a mitigation scenario, like EWS BE, Required is the Expected value. Therefore, on those directories, we need to verify that IP filtering is set if not set to Require.
                    if ($expectedExtendedConfiguration -eq "Require") {
                        $properlySecuredConfiguration = $expectedExtendedConfiguration -eq $extendedConfiguration.ExtendedProtection

                        if ($properlySecuredConfiguration -eq $false) {
                            # Only care about virtual directories that we allow mitigation for
                            $properlySecuredConfiguration = $mitigationSupportedvDirs.Contains($virtualDirectoryName.ToLower()) -and
                            $extendedConfiguration.MitigationSettings.AllowUnlisted -eq "false"
                        }
                    } else {
                        $properlySecuredConfiguration = $true
                    }

                    $extendedProtectionList.Add([PSCustomObject]@{
                            VirtualDirectoryName          = $virtualDirectoryName
                            Configuration                 = $extendedConfiguration
                            ExtendedProtection            = $extendedConfiguration.ExtendedProtection
                            SupportedExtendedProtection   = $expectedExtendedConfiguration -eq $extendedConfiguration.ExtendedProtection
                            ExpectedExtendedConfiguration = $expectedExtendedConfiguration
                            MitigationEnabled             = ($extendedConfiguration.MitigationSettings.AllowUnlisted -eq "false")
                            MitigationSupported           = $mitigationSupportedvDirs.Contains($virtualDirectoryName.ToLower())
                            ProperlySecuredConfiguration  = $properlySecuredConfiguration
                            ExpectedSslFlags              = $matchEntry.SslFlags
                            SslFlagsSetCorrectly          = $sslFlagsToSet.Split(",").Count -eq $currentSetFlags.Count
                            SslFlagsToSet                 = $sslFlagsToSet
                        })
                } catch {
                    Write-Verbose "Failed to get extended protection match entry."
                    Invoke-CatchActionError $CatchActionFunction
                }
            }
        } catch {
            Write-Verbose "Failed to get get extended protection."
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            ComputerName                          = $ComputerName
            ServerConnected                       = $serverConnected
            SupportedVersionForExtendedProtection = $supportedVersion
            ApplicationHostConfig                 = $ApplicationHostConfig
            ExtendedProtectionConfiguration       = $extendedProtectionList
            ExtendedProtectionConfigured          = $null -ne ($extendedProtectionList.ExtendedProtection | Where-Object { $_ -ne "None" })
        }
    }
}

function Get-ExchangeBuildVersionInformation {
    [CmdletBinding()]
    param(
        [object]$AdminDisplayVersion
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed $($AdminDisplayVersion.ToString())"
        $AdminDisplayVersion = $AdminDisplayVersion.ToString()
        $exchangeMajorVersion = [string]::Empty
        [int]$major = 0
        [int]$minor = 0
        [int]$build = 0
        [int]$revision = 0
        $product = $null
        [double]$buildVersion = 0.0
    }
    process {
        $split = $AdminDisplayVersion.Substring(($AdminDisplayVersion.IndexOf(" ")) + 1, 4).Split(".")
        $major = [int]$split[0]
        $minor = [int]$split[1]
        $product = $major + ($minor / 10)

        $buildStart = $AdminDisplayVersion.LastIndexOf(" ") + 1
        $split = $AdminDisplayVersion.Substring($buildStart, ($AdminDisplayVersion.LastIndexOf(")") - $buildStart)).Split(".")
        $build = [int]$split[0]
        $revision = [int]$split[1]
        $revisionDecimal = if ($revision -lt 10) { $revision / 10 } else { $revision / 100 }
        $buildVersion = $build + $revisionDecimal

        Write-Verbose "Determining Major Version based off of $product"

        switch ([string]$product) {
            "14.3" { $exchangeMajorVersion = "Exchange2010" }
            "15" { $exchangeMajorVersion = "Exchange2013" }
            "15.1" { $exchangeMajorVersion = "Exchange2016" }
            "15.2" { $exchangeMajorVersion = "Exchange2019" }
            default { $exchangeMajorVersion = "Unknown" }
        }
    }
    end {
        Write-Verbose "Found Major Version '$exchangeMajorVersion'"
        return [PSCustomObject]@{
            MajorVersion = $exchangeMajorVersion
            Major        = $major
            Minor        = $minor
            Build        = $build
            Revision     = $revision
            Product      = $product
            BuildVersion = $buildVersion
        }
    }
}


function Get-ExchangeSettingOverride {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Server,
        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $updatedTime = [DateTime]::MinValue
        $settingOverrides = $null
        $simpleSettingOverrides = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        try {
            $params = @{
                Process     = "Microsoft.Exchange.Directory.TopologyService"
                Component   = "VariantConfiguration"
                Argument    = "Overrides"
                Server      = $Server
                ErrorAction = "Stop"
            }
            $diagnosticInfo = Get-ExchangeDiagnosticInfo @params
            Write-Verbose "Successfully got the Exchange Diagnostic Information"
            $xml = [xml]$diagnosticInfo.Result
            $overrides = $xml.Diagnostics.Components.VariantConfiguration.Overrides
            $updatedTime = $overrides.Updated
            $settingOverrides = $overrides.SettingOverride

            foreach ($override in $settingOverrides) {
                Write-Verbose "Working on $($override.Name)"
                $simpleSettingOverrides.Add([PSCustomObject]@{
                        Name          = $override.Name
                        ComponentName = $override.ComponentName
                        SectionName   = $override.SectionName
                        Status        = $override.Status
                        Parameters    = $override.Parameters.Parameter
                    })
            }
        } catch {
            Write-Verbose "Failed to get the Exchange setting override"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            Server                 = $Server
            LastUpdated            = $updatedTime
            SettingOverrides       = $settingOverrides
            SimpleSettingOverrides = $simpleSettingOverrides
        }
    }
}


function Get-AppPool {
    [CmdletBinding()]
    param ()

    begin {
        function Get-IndentLevel ($line) {
            if ($line.StartsWith(" ")) {
                ($line | Select-String "^ +").Matches[0].Length
            } else {
                0
            }
        }

        function Convert-FromAppPoolText {
            [CmdletBinding()]
            param (
                [Parameter(Mandatory = $true)]
                [string[]]
                $Text,

                [Parameter(Mandatory = $false)]
                [int]
                $Line = 0,

                [Parameter(Mandatory = $false)]
                [int]
                $MinimumIndentLevel = 2
            )

            if ($Line -ge $Text.Count) {
                return $null
            }

            $startingIndentLevel = Get-IndentLevel $Text[$Line]
            if ($startingIndentLevel -lt $MinimumIndentLevel) {
                return $null
            }

            $hash = @{}

            while ($Line -lt $Text.Count) {
                $indentLevel = Get-IndentLevel $Text[$Line]
                if ($indentLevel -gt $startingIndentLevel) {
                    # Skip until we get to the next thing at this level
                } elseif ($indentLevel -eq $startingIndentLevel) {
                    # We have a property at this level. Add it to the object.
                    if ($Text[$Line] -match "\[(\S+)\]") {
                        $name = $Matches[1]
                        $value = Convert-FromAppPoolText -Text $Text -Line ($Line + 1) -MinimumIndentLevel $startingIndentLevel
                        $hash[$name] = $value
                    } elseif ($Text[$Line] -match "\s+(\S+):`"(.*)`"") {
                        $name = $Matches[1]
                        $value = $Matches[2].Trim("`"")
                        $hash[$name] = $value
                    }
                } else {
                    # IndentLevel is less than what we started with, so return
                    [PSCustomObject]$hash
                    return
                }

                ++$Line
            }

            [PSCustomObject]$hash
        }

        $appPoolCmd = "$env:windir\System32\inetsrv\appcmd.exe"
    }

    process {
        $appPoolNames = & $appPoolCmd list apppool |
            Select-String "APPPOOL `"(\S+)`" " |
            ForEach-Object { $_.Matches.Groups[1].Value }

        foreach ($appPoolName in $appPoolNames) {
            $appPoolText = & $appPoolCmd list apppool $appPoolName /text:*
            Convert-FromAppPoolText -Text $appPoolText -Line 1
        }
    }
}
function Get-ExchangeAppPoolsInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $appPool = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-AppPool} `
        -ScriptBlockDescription "Getting App Pool information" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $exchangeAppPoolsInfo = @{}

    $appPool |
        Where-Object { $_.add.name -like "MSExchange*" } |
        ForEach-Object {
            $configContent = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock {
                param(
                    $FilePath
                )
                if (Test-Path $FilePath) {
                    return (Get-Content $FilePath)
                }
                return [string]::Empty
            } `
                -ScriptBlockDescription "Getting Content file for $($_.add.name)" `
                -ArgumentList $_.add.CLRConfigFile `
                -CatchActionFunction ${Function:Invoke-CatchActions}

            $gcUnknown = $true
            $gcServerEnabled = $false

            if (-not ([string]::IsNullOrEmpty($configContent))) {
                $gcSetting = ([xml]$configContent).Configuration.Runtime.gcServer.Enabled
                $gcUnknown = $gcSetting -ne "true" -and $gcSetting -ne "false"
                $gcServerEnabled = $gcSetting -eq "true"
            }
            $exchangeAppPoolsInfo.Add($_.add.Name, [PSCustomObject]@{
                    ConfigContent   = $configContent
                    AppSettings     = $_
                    GCUnknown       = $gcUnknown
                    GCServerEnabled = $gcServerEnabled
                })
        }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exchangeAppPoolsInfo
}



function Get-ExchangeIISConfigSettings {
    [CmdletBinding()]
    param(
        [string]$MachineName,
        [string[]]$FilePath,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        function GetExchangeIISConfigSettings {
            param(
                [string[]]$FilePath
            )

            $results = New-Object 'System.Collections.Generic.List[object]'
            $sharedConfigure = @()
            $ca = "ClientAccess\"
            $hp = "HttpProxy\"
            $sharedWebConfig = "SharedWebConfig.config"

            foreach ($location in $FilePath) {

                Write-Verbose "Working on location: $location"
                $exist = Test-Path $location
                $content = $null
                $sharedLocation = $null

                if ($exist) {
                    Write-Verbose "File exists. Getting content"
                    $content = Get-Content $location
                    $linkedConfiguration = ($content | Select-String linkedConfiguration).Line

                    if ($null -ne $linkedConfiguration) {
                        Write-Verbose "Found linkedConfiguration"
                        $clientAccessSharedIndex = $location.IndexOf($ca)
                        $httpProxySharedIndex = $location.IndexOf($hp)

                        if ($clientAccessSharedIndex -ne -1) {
                            $sharedLocation = [System.IO.Path]::Combine($location.Substring(0, $clientAccessSharedIndex + $ca.Length), $sharedWebConfig)
                        } elseif ($httpProxySharedIndex -ne -1) {
                            $sharedLocation = [System.IO.Path]::Combine($location.Substring(0, $httpProxySharedIndex + $hp.Length), $sharedWebConfig)
                        }
                    }

                    if ($null -ne $sharedLocation -and
                        (-not ($sharedConfigure.Contains($sharedLocation)))) {
                        Write-Verbose "Adding Shared Location of: $sharedLocation"
                        $sharedConfigure += $sharedLocation
                        $results.Add([PSCustomObject]@{
                                Location = $sharedLocation
                                Content  = if (Test-Path $sharedLocation) { Get-Content $sharedLocation } else { $null }
                                Exist    = $(Test-Path $sharedLocation)
                            })
                    }
                }

                $results.Add([PSCustomObject]@{
                        Location = $location
                        Content  = $content
                        Exist    = $exist
                    })
            }
            return $results
        }
    } process {
        $params = @{
            ComputerName        = $MachineName
            ScriptBlock         = ${Function:GetExchangeIISConfigSettings}
            ArgumentList        = $FilePath
            CatchActionFunction = $CatchActionFunction
        }
        return Invoke-ScriptBlockHandler @params
    }
}

function Get-IISWebApplication {
    $webApplications = Get-WebApplication
    $returnList = New-Object 'System.Collections.Generic.List[object]'

    foreach ($webApplication in $webApplications) {
        $returnList.Add([PSCustomObject]@{
                Path                       = $webApplication.Path
                ApplicationPool            = $webApplication.applicationPool
                EnabledProtocols           = $webApplication.enabledProtocols
                ServiceAutoStartEnabled    = $webApplication.serviceAutoStartEnabled
                ServiceAutoStartProvider   = $webApplication.serviceAutoStartProvider
                PreloadEnabled             = $webApplication.preloadEnabled
                PreviouslyEnabledProtocols = $webApplication.previouslyEnabledProtocols
                ServiceAutoStartMode       = $webApplication.serviceAutoStartMode
                VirtualDirectoryDefaults   = $webApplication.virtualDirectoryDefaults
                Collection                 = $webApplication.Collection
                Location                   = $webApplication.Location
                ItemXPath                  = $webApplication.ItemXPath
                PhysicalPath               = $webApplication.PhysicalPath.Replace("%windir%", $env:windir).Replace("%SystemDrive%", $env:SystemDrive)
            })
    }

    return $returnList
}

function Get-IISWebSite {
    param(
        [array]$WebSitesToProcess
    )

    $returnList = New-Object 'System.Collections.Generic.List[object]'
    $webSites = New-Object 'System.Collections.Generic.List[object]'

    if ($null -eq $WebSitesToProcess) {
        $webSites.AddRange((Get-WebSite))
    } else {
        foreach ($iisWebSite in $WebSitesToProcess) {
            $webSites.Add((Get-WebSite -Name $($iisWebSite)))
        }
    }

    $bindings = Get-WebBinding

    foreach ($site in $webSites) {
        $siteBindings = $bindings |
            Where-Object { $_.ItemXPath -like "*@name='$($site.name)' and @id='$($site.id)'*" }
        $returnList.Add([PSCustomObject]@{
                Name                       = $site.Name
                Id                         = $site.Id
                State                      = $site.State
                Bindings                   = $siteBindings
                Limits                     = $site.Limits
                LogFile                    = $site.logFile
                TraceFailedRequestsLogging = $site.traceFailedRequestsLogging
                Hsts                       = $site.hsts
                ApplicationDefaults        = $site.applicationDefaults
                VirtualDirectoryDefaults   = $site.virtualDirectoryDefaults
                Collection                 = $site.collection
                ApplicationPool            = $site.applicationPool
                EnabledProtocols           = $site.enabledProtocols
                PhysicalPath               = $site.physicalPath.Replace("%windir%", $env:windir).Replace("%SystemDrive%", $env:SystemDrive)
            }
        )
    }
    return $returnList
}




function Get-ExchangeContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
    $exchangeContainerPath = ("CN=Microsoft Exchange,CN=Services," + $rootDSE.configurationNamingContext)
    $exchangeContainer = [ADSI]("LDAP://" + $exchangeContainerPath)
    Write-Verbose "Exchange Container Path: $($exchangeContainer.path)"
    return $exchangeContainer
}

function Get-OrganizationContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param ()

    $exchangeContainer = Get-ExchangeContainer
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($exchangeContainer, "(objectClass=msExchOrganizationContainer)", @("distinguishedName"))
    return $searcher.FindOne().GetDirectoryEntry()
}

function Get-ExchangeProtocolContainer {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.DirectoryEntry])]
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    $ComputerName = $ComputerName.Split(".")[0]

    $organizationContainer = Get-OrganizationContainer
    $protocolContainerPath = ("CN=Protocols,CN=" + $ComputerName + ",CN=Servers,CN=Exchange Administrative Group (FYDIBOHF23SPDLT),CN=Administrative Groups," + $organizationContainer.distinguishedName)
    $protocolContainer = [ADSI]("LDAP://" + $protocolContainerPath)
    Write-Verbose "Protocol Container Path: $($protocolContainer.Path)"
    return $protocolContainer
}

function Get-ExchangeWebSitesFromAd {
    [CmdletBinding()]
    [OutputType([System.Object])]
    param (
        [string]$ComputerName = $env:COMPUTERNAME
    )

    begin {
        function GetExchangeWebSiteFromCn {
            param (
                [string]$Site
            )

            if ($null -ne $Site) {
                $index = $Site.IndexOf("(") + 1
                if ($index -ne 0) {
                    return ($Site.Substring($index, ($Site.LastIndexOf(")") - $index)))
                }
            }
        }

        $processedExchangeWebSites = New-Object 'System.Collections.Generic.List[array]'
    }
    process {
        $protocolContainer = Get-ExchangeProtocolContainer -ComputerName $ComputerName
        if ($null -ne $protocolContainer) {
            $httpProtocol = $protocolContainer.Children | Where-Object {
                ($_.name -eq "HTTP")
            }

            foreach ($cn in $httpProtocol.Children.cn) {
                $processedExchangeWebSites.Add((GetExchangeWebSiteFromCn $cn))
            }
        }
    }
    end {
        return ($processedExchangeWebSites | Select-Object -Unique)
    }
}


function Get-ApplicationHostConfig {
    [CmdletBinding()]
    [OutputType([System.Xml.XmlNode])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,
        [scriptblock]$CatchActionFunction
    )
    function LoadApplicationHostConfig {
        param()
        $appHostConfig = New-Object -TypeName Xml
        $appHostConfigPath = "$($env:WINDIR)\System32\inetsrv\config\applicationHost.config"
        $appHostConfig.Load($appHostConfigPath)
        return $appHostConfig
    }

    $params = @{
        ComputerName           = $ComputerName
        ScriptBlockDescription = "Getting applicationHost.config"
        ScriptBlock            = ${Function:LoadApplicationHostConfig}
        CatchActionFunction    = $CatchActionFunction
    }

    return Invoke-ScriptBlockHandler @params
}


function Get-IISModules {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [System.Xml.XmlNode]$ApplicationHostConfig,

        [Parameter(Mandatory = $false)]
        [bool]$SkipLegacyOSModulesCheck = $false,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $modulesToCheckList = New-Object 'System.Collections.Generic.List[object]'

        # Add all modules here which should be skipped on legacy OS (pre-Windows Server 2016)
        $modulesToSkip = @(
            "$env:windir\system32\inetsrv\cachuri.dll",
            "$env:windir\system32\inetsrv\cachfile.dll",
            "$env:windir\system32\inetsrv\cachtokn.dll",
            "$env:windir\system32\inetsrv\cachhttp.dll",
            "$env:windir\system32\inetsrv\compstat.dll",
            "$env:windir\system32\inetsrv\defdoc.dll",
            "$env:windir\system32\inetsrv\dirlist.dll",
            "$env:windir\system32\inetsrv\protsup.dll",
            "$env:windir\system32\inetsrv\redirect.dll",
            "$env:windir\system32\inetsrv\static.dll",
            "$env:windir\system32\inetsrv\authanon.dll",
            "$env:windir\system32\inetsrv\custerr.dll",
            "$env:windir\system32\inetsrv\loghttp.dll",
            "$env:windir\system32\inetsrv\iisetw.dll",
            "$env:windir\system32\inetsrv\iisfreb.dll",
            "$env:windir\system32\inetsrv\iisreqs.dll",
            "$env:windir\system32\inetsrv\isapi.dll",
            "$env:windir\system32\inetsrv\compdyn.dll",
            "$env:windir\system32\inetsrv\authcert.dll",
            "$env:windir\system32\inetsrv\authbas.dll",
            "$env:windir\system32\inetsrv\authsspi.dll",
            "$env:windir\system32\inetsrv\authmd5.dll",
            "$env:windir\system32\inetsrv\modrqflt.dll",
            "$env:windir\system32\inetsrv\filter.dll",
            "$env:windir\system32\rpcproxy\rpcproxy.dll",
            "$env:windir\system32\inetsrv\validcfg.dll",
            "$env:windir\system32\wsmsvc.dll",
            "$env:windir\system32\inetsrv\iprestr.dll",
            "$env:windir\system32\inetsrv\diprestr.dll",
            "$env:windir\system32\inetsrv\iis_ssi.dll",
            "$env:windir\system32\inetsrv\cgi.dll",
            "$env:windir\system32\inetsrv\iisfcgi.dll",
            "$env:windir\system32\inetsrv\iiswsock.dll",
            "$env:windir\system32\inetsrv\warmup.dll")

        function GetModulePath {
            [CmdletBinding()]
            [OutputType([System.String])]
            param(
                [string]$Path
            )

            if (-not([String]::IsNullOrEmpty($Path))) {
                $returnPath = $Path

                if ($Path -match "\%.+\%") {
                    Write-Verbose "Environment variable found in path: $Path"
                    # Assuming that we have the env var always at the beginning of the string and no other vars within the string
                    # Example: %windir%\system32\someexample.dll
                    $preparedPath = ($Path.Split("%", [System.StringSplitOptions]::RemoveEmptyEntries))
                    if ($preparedPath.Count -eq 2) {
                        if ($preparedPath[0] -notmatch "\\.+\\") {
                            $varPath = [System.Environment]::GetEnvironmentVariable($preparedPath[0])
                            $returnPath = [String]::Join("", $varPath, $($preparedPath[1]))
                        }
                    }
                }
            } else {
                $returnPath = $null
            }

            return $returnPath
        }
        function GetIISModulesSignatureStatus {
            [CmdletBinding()]
            param(
                [Parameter(Mandatory = $true)]
                [object[]]$Modules
            )
            process {
                try {
                    $iisModulesList = New-Object 'System.Collections.Generic.List[object]'
                    if ($Modules.Count -ge 1) {
                        Write-Verbose "At least one module is loaded by IIS"
                        foreach ($m in $Modules) {
                            Write-Verbose "Now processing module: $($m.Name)"
                            $isModuleSigned = $false
                            $signatureDetails = [PSCustomObject]@{
                                Signer            = $null
                                SignatureStatus   = -1
                                IsMicrosoftSigned = $null
                            }

                            $moduleFilePath = GetModulePath -Path $m.Image

                            try {
                                Write-Verbose "Querying file signing information"
                                $signature = Get-AuthenticodeSignature -FilePath $moduleFilePath -ErrorAction Stop
                                Write-Verbose "Performing signature status validation. Status: $($signature.Status)"
                                # Signature Status Enum Values:
                                # <0> Valid, <1> UnknownError, <2> NotSigned, <3> HashMismatch,
                                # <4> NotTrusted, <5> NotSupportedFileFormat, <6> Incompatible,
                                # https://docs.microsoft.com/dotnet/api/system.management.automation.signaturestatus
                                if (($null -ne $signature.Status) -and
                                    ($signature.Status -ne 1) -and
                                    ($signature.Status -ne 2) -and
                                    ($signature.Status -ne 5) -and
                                    ($signature.Status -ne 6)) {

                                    $signatureDetails.SignatureStatus = $signature.Status
                                    $isModuleSigned = $true

                                    if ($null -ne $signature.SignerCertificate.Subject) {
                                        Write-Verbose "Signer information found. Subject: $($signature.SignerCertificate.Subject)"
                                        $signatureDetails.Signer = $signature.SignerCertificate.Subject.ToString()
                                        $signatureDetails.IsMicrosoftSigned = $signature.SignerCertificate.Subject -cmatch "O=Microsoft Corporation, L=Redmond, S=Washington"
                                    }
                                }

                                $iisModulesList.Add([PSCustomObject]@{
                                        Name             = $m.Name
                                        Path             = $moduleFilePath
                                        Signed           = $isModuleSigned
                                        SignatureDetails = $signatureDetails
                                    })
                            } catch {
                                Write-Verbose "Unable to validate file signing information"
                                Invoke-CatchActionError $CatchActionFunction
                            }
                        }
                    } else {
                        Write-Verbose "No modules are loaded by IIS"
                    }
                } catch {
                    Write-Verbose "Failed to process global module information. $_"
                    Invoke-CatchActionError $CatchActionFunction
                }
            }
            end {
                return $iisModulesList
            }
        }
    }
    process {
        $ApplicationHostConfig.configuration.'system.webServer'.globalModules.add | ForEach-Object {
            if ($SkipLegacyOSModulesCheck) {
                if ((GetModulePath $_.image) -notin $modulesToSkip) {
                    $modulesToCheckList.Add($_)
                }
            } else {
                $modulesToCheckList.Add($_)
            }
        }

        $modules = GetIISModulesSignatureStatus -Modules $modulesToCheckList

        # Validate if all modules that are loaded are digitally signed
        $allModulesAreSigned = (-not($modules.Signed.Contains($false)))
        Write-Verbose "Are all modules loaded by IIS digitally signed? $allModulesAreSigned"

        # Validate that all modules are signed by Microsoft Corp.
        $allModulesSignedByMSFT = (-not($modules.SignatureDetails.IsMicrosoftSigned.Contains($false)))
        Write-Verbose "Are all modules signed by Microsoft Corporation? $allModulesSignedByMSFT"

        # Validate if all signatures are valid (regardless of whether signed by Microsoft Corp. or not)
        $allSignaturesValid = $null -eq ($modules |
                Where-Object { $_.Signed -and $_.SignatureDetails.SignatureStatus -ne 0 })
    }
    end {
        return [PSCustomObject]@{
            AllSignedModulesSignedByMSFT = $allModulesSignedByMSFT
            AllSignaturesValid           = $allSignaturesValid
            AllModulesSigned             = $allModulesAreSigned
            ModuleList                   = $modules
        }
    }
}

function Get-ExchangeServerIISSettings {
    param(
        [string]$ComputerName,
        [bool]$IsLegacyOS = $false,
        [scriptblock]$CatchActionFunction
    )
    process {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"

        $params = @{
            ComputerName        = $ComputerName
            CatchActionFunction = $CatchActionFunction
        }

        try {
            $exchangeWebSites = Get-ExchangeWebSitesFromAd -ComputerName $ComputerName
            if ($exchangeWebSites.Count -gt 2) {
                Write-Verbose "Multiple OWA/ECP virtual directories detected"
            }
            Write-Verbose "Exchange websites detected: $([string]::Join(", " ,$exchangeWebSites))"
        } catch {
            Write-Verbose "Failed to get the Exchange Web Sites from Ad."
            $exchangeWebSites = $null
            Invoke-CatchActions
        }

        # We need to wrap the array into another array as the -WebSitesToProcess parameter expects an array object
        $webSite = Invoke-ScriptBlockHandler @params -ScriptBlock ${Function:Get-IISWebSite} -ArgumentList (, $exchangeWebSites)
        $webApplication = Invoke-ScriptBlockHandler @params -ScriptBlock ${Function:Get-IISWebApplication}

        $configurationFiles = @($webSite.PhysicalPath)
        $configurationFiles += $webApplication.PhysicalPath | Select-Object -Unique
        $configurationFiles = $configurationFiles | ForEach-Object { [System.IO.Path]::Combine($_, "web.config") }

        $iisConfigParams = @{
            MachineName         = $ComputerName
            FilePath            = $configurationFiles
            CatchActionFunction = $CatchActionFunction
        }
        Write-Verbose "Trying to query the IIS configuration settings"
        $iisConfigurationSettings = Get-ExchangeIISConfigSettings @iisConfigParams

        Write-Verbose "Trying to query the 'applicationHost.config' file"
        $applicationHostConfig = Get-ApplicationHostConfig $ComputerName $CatchActionFunction

        if ($null -ne $applicationHostConfig) {
            Write-Verbose "Trying to query the modules which are loaded by IIS"
            $iisModulesParams = @{
                ApplicationHostConfig    = $applicationHostConfig
                SkipLegacyOSModulesCheck = $IsLegacyOS
                CatchActionFunction      = $CatchActionFunction
            }
            $iisModulesInformation = Get-IISModules @iisModulesParams
        } else {
            Write-Verbose "No 'applicationHost.config' file returned by previous call"
        }
    } end {
        return [PSCustomObject]@{
            applicationHostConfig    = $applicationHostConfig
            IISModulesInformation    = $iisModulesInformation
            IISConfigurationSettings = $iisConfigurationSettings
            IISWebSite               = $webSite
            IISWebApplication        = $webApplication
        }
    }
}



function Get-ExchangeDomainConfigVersion {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [string]
        $Domain
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    if ([System.String]::IsNullOrEmpty($Domain)) {
        Write-Verbose "No domain information passed - using current domain"
        $Domain = ([System.DirectoryServices.ActiveDirectory.Domain]::GetCurrentDomain()).Name
    }

    Write-Verbose "Getting domain information for domain: $Domain"
    $forest = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Forest

    Write-Verbose "Checking if domain is present"
    if ($forest.Domains.Name.Contains($Domain)) {
        Write-Verbose "Domain: $Domain is present in forest: $($forest.Name)"
        $domainObject = $forest.Domains | Where-Object { $_.Name -eq $Domain }
        $domainDN = $domainObject.GetDirectoryEntry().distinguishedName
        $adEntry = [ADSI]("LDAP://CN=Microsoft Exchange System Objects," + $domainDN)
        $sdFinder = New-Object System.DirectoryServices.DirectorySearcher($adEntry)
        try {
            $mesoResult = $sdFinder.FindOne()
        } catch {
            Write-Verbose "No result was returned"
            Invoke-CatchActions
        }

        if ($null -ne $mesoResult) {
            Write-Verbose "MESO (Microsoft Exchange System Objects) container detected"
            [int]$objectVersion = $mesoResult.Properties.objectversion[0]
            $whenChangedInfo = $mesoResult.Properties.whenchanged
        } else {
            Write-Verbose "No MESO (Microsoft Exchange System Objects) container detected"
        }
    } else {
        Write-Verbose "Domain: $Domain is NOT present in forest: $($forest.Name)"
    }

    return [PSCustomObject]@{
        Domain                    = $Domain
        DomainPreparedForExchange = ($mesoResult.Count -gt 0)
        ObjectVersion             = $objectVersion
        WhenChanged               = $whenChangedInfo
    }
}

function Get-ActiveDirectoryAcl {
    [CmdletBinding()]
    [OutputType([System.DirectoryServices.ActiveDirectorySecurity])]
    param (
        [Parameter()]
        [string]
        $DistinguishedName
    )

    $adEntry = [ADSI]("LDAP://$($DistinguishedName)")
    $sdFinder = New-Object System.DirectoryServices.DirectorySearcher($adEntry, "(objectClass=*)", [string[]]("distinguishedName", "ntSecurityDescriptor"), [System.DirectoryServices.SearchScope]::Base)
    $sdResult = $sdFinder.FindOne()
    $ntsdProp = $sdResult.Properties["ntSecurityDescriptor"][0]
    $adSec = New-Object System.DirectoryServices.ActiveDirectorySecurity
    $adSec.SetSecurityDescriptorBinaryForm($ntsdProp)
    return $adSec
}


function Get-ExchangeADSplitPermissionsEnabled {
    [CmdletBinding()]
    [OutputType([bool])]
    param ()

    <#
        The following bullets are AD split permissions indicators:
        - An organizational unit (OU) named Microsoft 'Exchange Protected Groups' is created
        - The 'Exchange Windows Permissions' security group is created/moved in/to the 'Microsoft Exchange Protected Groups' OU
        - The 'Exchange Trusted Subsystem' security group isn't member of the 'Exchange Windows Permissions' security group
        - ACEs that would have been assigned to the 'Exchange Windows Permissions' security group aren't added to the Active Directory domain object
        See: https://learn.microsoft.com/exchange/permissions/split-permissions/split-permissions?view=exchserver-2019#active-directory-split-permissions
    #>

    $isADSplitPermissionsEnabled = $false
    try {
        $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")
        $exchangeTrustedSubsystemDN = ("CN=Exchange Trusted Subsystem,OU=Microsoft Exchange Security Groups," + $rootDSE.rootDomainNamingContext)
        $adSearcher = New-Object DirectoryServices.DirectorySearcher
        $adSearcher.Filter = '(&(objectCategory=group)(cn=Exchange Windows Permissions))'
        $adSearcher.SearchRoot = ("LDAP://OU=Microsoft Exchange Protected Groups," + $rootDSE.rootDomainNamingContext)
        $adSearcherResult = $adSearcher.FindOne()

        if ($null -ne $adSearcherResult) {
            Write-Verbose "'Exchange Windows Permissions' in 'Microsoft Exchange Protected Groups' OU detected"
            # AD split permissions is enabled if 'Exchange Trusted Subsystem' isn't a member of the 'Exchange Windows Permissions' security group
            $isADSplitPermissionsEnabled = (($null -eq $adSearcherResult.Properties.member) -or
            (-not($adSearcherResult.Properties.member).ToLower().Contains($exchangeTrustedSubsystemDN.ToLower())))
        }
    } catch {
        Write-Verbose "OU 'Microsoft Exchange Protected Groups' was not found - AD split permissions not enabled"
        Invoke-CatchActions
    }

    return $isADSplitPermissionsEnabled
}


function Get-ExchangeOtherWellKnownObjects {
    [CmdletBinding()]
    param ()

    $otherWellKnownObjectIds = @{
        "C2F9A9F9D6A1B74A9E068728F8F842EA" = "Organization Management"
        "DB72C41D49580A4DB304FE6981E56297" = "Recipient Management"
        "1A9E39D35ABE5747B979FFC0C6E5EA26" = "View-Only Organization Management"
        "45FA417B3574DC4E929BC4B059699792" = "Public Folder Management"
        "E80CDFB75697934981C898B4DBC5A0C6" = "UM Management"
        "B3DDC6BE2A3BE84B97EB2DCE9477E389" = "Help Desk"
        "BEA432C94E1D254EAF99B40573360D5B" = "Records Management"
        "C67FDE2E8339674490FBAFDCA3DFDC95" = "Discovery Management"
        "4DB8E7754EB6C1439565612E69A80A4F" = "Server Management"
        "D1281926D1F55B44866D1D6B5BD87A09" = "Delegated Setup"
        "03B709F451F3BF4388E33495369B6771" = "Hygiene Management"
        "B30A449BA9B420458C4BB22F33C52766" = "Compliance Management"
        "A7D2016C83F003458132789EEB127B84" = "Exchange Servers"
        "EA876A58DB6DD04C9006939818F800EB" = "Exchange Trusted Subsystem"
        "02522ECF9985984A9232056FC704CC8B" = "Managed Availability Servers"
        "4C17D0117EBE6642AFAEE03BC66D381F" = "Exchange Windows Permissions"
        "9C5B963F67F14A4B936CB8EFB19C4784" = "ExchangeLegacyInterop"
        "776B176BD3CB2A4DA7829EA963693013" = "Security Reader"
        "03D7F0316EF4B3498AC434B6E16F09D9" = "Security Administrator"
        "A2A4102E6F676141A2C4AB50F3C102D5" = "PublicFolderMailboxes"
    }

    $exchangeContainer = Get-ExchangeContainer
    $searcher = New-Object System.DirectoryServices.DirectorySearcher($exchangeContainer, "(objectClass=*)", @("otherWellKnownObjects", "distinguishedName"))
    $result = $searcher.FindOne()
    foreach ($val in $result.Properties["otherWellKnownObjects"]) {
        $matchResults = $val | Select-String "^B:32:([^:]+):(.*)$"
        if ($matchResults.Matches.Groups.Count -ne 3) {
            # Only output the raw value of a corrupted entry
            [PSCustomObject]@{
                WellKnownName     = $null
                WellKnownGuid     = $null
                DistinguishedName = $null
                RawValue          = $val
            }

            continue
        }

        $wkGuid = $matchResults.Matches.Groups[1].Value
        $wkName = $otherWellKnownObjectIds[$wkGuid]

        [PSCustomObject]@{
            WellKnownName     = $wkName
            WellKnownGuid     = $wkGuid
            DistinguishedName = $matchResults.Matches.Groups[2].Value
            RawValue          = $val
        }
    }
}

function Get-ExchangeAdPermissions {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [HealthChecker.ExchangeMajorVersion]
        $ExchangeVersion,
        [Parameter(Mandatory = $true)]
        [HealthChecker.OSServerVersion]
        $OSVersion
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function NewMatchingEntry {
        param(
            [ValidateSet("Domain", "AdminSDHolder")]
            [string]$TargetObject,
            [string]$ObjectTypeGuid,
            [string]$InheritedObjectType
        )

        return [PSCustomObject]@{
            TargetObject        = $TargetObject
            ObjectTypeGuid      = $ObjectTypeGuid
            InheritedObjectType = $InheritedObjectType
        }
    }

    function NewGroupEntry {
        param(
            [string]$Name,
            [object[]]$MatchingEntries
        )

        return [PSCustomObject]@{
            Name     = $Name
            Sid      = $null
            AceEntry = $MatchingEntries
        }
    }

    # Computer Class GUID
    $computerClassGUID = "bf967a86-0de6-11d0-a285-00aa003049e2"

    # userCertificate GUID
    $userCertificateGUID = "bf967a7f-0de6-11d0-a285-00aa003049e2"

    # managedBy GUID
    $managedByGUID = "0296c120-40da-11d1-a9c0-0000f80367c1"

    $writePropertyRight = [System.DirectoryServices.ActiveDirectoryRights]::WriteProperty
    $denyType = [System.Security.AccessControl.AccessControlType]::Deny
    $inheritanceAll = [System.DirectoryServices.ActiveDirectorySecurityInheritance]::All

    $groupLists = @(
        (NewGroupEntry "Exchange Servers" @(
            (NewMatchingEntry -TargetObject "Domain" -ObjectTypeGuid $userCertificateGUID -InheritedObjectType $computerClassGUID)
        )),

        (NewGroupEntry "Exchange Windows Permissions" @(
            (NewMatchingEntry -TargetObject "Domain" -ObjectTypeGuid $managedByGUID -InheritedObjectType $computerClassGUID)
        )))

    $returnedResults = New-Object 'System.Collections.Generic.List[object]'

    try {
        Write-Verbose "Getting the domain information"
        $forest = [System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Forest
        Write-Verbose ("Detected: $($forest.Domains.Count) domain(s)")
        $otherWellKnownObjects = Get-ExchangeOtherWellKnownObjects

        foreach ($group in $groupLists) {
            Write-Verbose "Trying to find: $($group.Name)"
            $wkObject = $otherWellKnownObjects | Where-Object { $_.WellKnownName -eq $group.Name }
            if ($null -ne $wkObject) {
                Write-Verbose "Found DN in otherWellKnownObjects: $($wkObject.DistinguishedName)"
                $entry = [ADSI]("LDAP://$($wkObject.DistinguishedName)")
                $group.Sid = (New-Object System.Security.Principal.SecurityIdentifier($entry.objectSid.Value, 0)).Value
                Write-Verbose "Found Results Set Sid: $($group.Sid)"
            }
        }
    } catch {
        Write-Verbose "Failed collecting domain information"
        Invoke-CatchActions
    }

    foreach ($domain in $forest.Domains) {

        $domainName = $domain.Name
        try {
            $domainDN = $domain.GetDirectoryEntry().distinguishedName
        } catch {
            Write-Verbose "Domain: $domainName - seems to be offline and will be skipped"
            Invoke-CatchActions
            continue
        }
        $adminSdHolderDN = "CN=AdminSDHolder,CN=System,$domainDN"
        $prepareDomainInfo = Get-ExchangeDomainConfigVersion -Domain $domainName

        if ($prepareDomainInfo.DomainPreparedForExchange) {
            Write-Verbose "Working on Domain: $domainName"
            Write-Verbose "MESO object version is: $($prepareDomainInfo.ObjectVersion)"
            Write-Verbose "DomainDN: $domainDN"

            try {
                try {
                    # Check if AD split permissions is enabled and if so, throw to check for objectVersion instead ACE
                    if (Get-ExchangeADSplitPermissionsEnabled) {
                        throw "Active Directory split permissions enabled. Fallback to 'objectVersion (Default)' validation initiated."
                    }

                    # Where() method became available with PowerShell 4.0 (default PS on Server 2012 R2),
                    # throw to initiate objectVersion (Default) testing, as we can't use Where() to check ACE below
                    if ($OSVersion -le [HealthChecker.OSServerVersion]::Windows2012) {
                        throw "Legacy server OS detected, fallback to 'objectVersion (Default)' validation initiated."
                    }
                    $domainAcl = Get-ActiveDirectoryAcl $domainDN.ToString()
                    $adminSdHolderAcl = Get-ActiveDirectoryAcl $adminSdHolderDN

                    if ($null -eq $domainAcl -or
                        $null -eq $domainAcl.Access -or
                        $null -eq $adminSdHolderAcl -or
                        $null -eq $adminSdHolderAcl.Access) {
                        throw "Failed to get required ACL information. Fallback to 'objectVersion (Default)' validation initiated."
                    }
                } catch {
                    Invoke-CatchActions
                    $objectVersionTestingValue = 13243
                    if ($ExchangeVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
                        $objectVersionTestingValue = 13238
                    }

                    $returnedResults.Add([PSCustomObject]@{
                            DomainName = $domainName
                            ObjectDN   = $null
                            ObjectAcl  = $null
                            CheckPass  = ($prepareDomainInfo.ObjectVersion -ge $objectVersionTestingValue)
                        })
                    continue
                }

                foreach ($group in $groupLists) {
                    Write-Verbose "Looking Ace Entries for the group: $($group.Name)"

                    foreach ($entry in $group.AceEntry) {
                        Write-Verbose "Trying to find the entry GUID: $($entry.ObjectTypeGuid)"
                        if ($entry.TargetObject -eq "AdminSDHolder") {
                            Write-Verbose "Looking for AdminSDHolder target object"
                            $objectAcl = $adminSdHolderAcl
                            $objectDN = $adminSdHolderDN
                        } else {
                            Write-Verbose "Looking for Domain target object"
                            $objectAcl = $domainAcl
                            $objectDN = $domainDN
                        }
                        Write-Verbose "ObjectDN: $objectDN"

                        # We need to pass an IdentityReference object to the constructor
                        $groupIdentityRef = New-Object System.Security.Principal.SecurityIdentifier($group.Sid)

                        $ace = New-Object System.DirectoryServices.ActiveDirectoryAccessRule($groupIdentityRef, $writePropertyRight, $denyType, $entry.ObjectTypeGuid, $inheritanceAll, $entry.InheritedObjectType)

                        $checkAce = $objectAcl.Access.Where({
                                    ($_.ActiveDirectoryRights -eq $ace.ActiveDirectoryRights) -and
                                    ($_.InheritanceType -eq $ace.InheritanceType) -and
                                    ($_.ObjectType -eq $ace.ObjectType) -and
                                    ($_.InheritedObjectType -eq $ace.InheritedObjectType) -and
                                    ($_.ObjectFlags -eq $ace.ObjectFlags) -and
                                    ($_.AccessControlType -eq $ace.AccessControlType) -and
                                    ($_.IsInherited -eq $ace.IsInherited) -and
                                    ($_.InheritanceFlags -eq $ace.InheritanceFlags) -and
                                    ($_.PropagationFlags -eq $ace.PropagationFlags) -and
                                    ($_.IdentityReference -eq $ace.IdentityReference.Translate([System.Security.Principal.NTAccount]))
                            })

                        $checkPass = $checkAce.Count -gt 0
                        Write-Verbose "Ace Result Check Passed: $checkPass"

                        $returnedResults.Add([PSCustomObject]@{
                                DomainName = $domainName
                                ObjectDN   = $objectDN
                                ObjectAcl  = $objectAcl
                                CheckPass  = $checkPass
                            })
                    }
                }
            } catch {
                Write-Verbose "Failed while getting ACE information"
                Invoke-CatchActions
            }
        } else {
            Write-Verbose "Domain: $domainName will be skipped because it is not configured to hold Exchange-related objects"
        }
    }
    return $returnedResults
}

function Get-ExchangeAdSchemaClass {
    param(
        [Parameter(Mandatory = $true)][string]$SchemaClassName
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand) to query $SchemaClassName schema class"

    $rootDSE = [ADSI]("LDAP://$([System.DirectoryServices.ActiveDirectory.Domain]::GetComputerDomain().Name)/RootDSE")

    if ([string]::IsNullOrEmpty($rootDSE.schemaNamingContext)) {
        return $null
    }

    $directorySearcher = New-Object System.DirectoryServices.DirectorySearcher
    $directorySearcher.SearchScope = "Subtree"
    $directorySearcher.SearchRoot = [ADSI]("LDAP://" + $rootDSE.schemaNamingContext.ToString())
    $directorySearcher.Filter = "(Name={0})" -f $SchemaClassName

    $findAll = $directorySearcher.FindAll()

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $findAll
}

function Get-ExchangeAMSIConfigurationState {
    [CmdletBinding()]
    param ()

    begin {
        function Get-AMSIStatusFlag {
            [CmdletBinding()]
            [OutputType([bool])]
            param (
                [Parameter(Mandatory = $true)]
                [object]$AMSIParameters
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                switch ($AMSIParameters.Split("=")[1]) {
                    ("False") { return $false }
                    ("True") { return $true }
                    default { return $null }
                }
            } catch {
                Write-Verbose "Ran into an issue when calling Split method. Parameters passed: $AMSIParameters"
                throw
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $amsiState = "Unknown"
        $amsiOrgWideSetting = $true
        $amsiConfigurationQuerySuccessful = $false
    } process {
        try {
            Write-Verbose "Trying to query AMSI configuration state"
            $amsiConfiguration = Get-SettingOverride -ErrorAction Stop | Where-Object { ($_.ComponentName -eq "Cafe") -and ($_.SectionName -eq "HttpRequestFiltering") }
            $amsiConfigurationQuerySuccessful = $true

            if ($null -ne $amsiConfiguration) {
                Write-Verbose "$($amsiConfiguration.Count) override(s) detected for AMSI configuration"
                $amsiMultiConfigObject = @()
                foreach ($amsiConfig in $amsiConfiguration) {
                    try {
                        $amsiState = Get-AMSIStatusFlag -AMSIParameters $amsiConfig.Parameters -ErrorAction Stop
                    } catch {
                        Write-Verbose "Unable to process: $($amsiConfig.Parameters) to determine status flags"
                        $amsiState = "Unknown"
                        Invoke-CatchActions
                    }
                    $amsiOrgWideSetting = ($null -eq $amsiConfig.Server)
                    $amsiConfigTempCustomObject = [PSCustomObject]@{
                        Id              = $amsiConfig.Id
                        Name            = $amsiConfig.Name
                        Reason          = $amsiConfig.Reason
                        Server          = $amsiConfig.Server
                        ModifiedBy      = $amsiConfig.ModifiedBy
                        Enabled         = $amsiState
                        OrgWideSetting  = $amsiOrgWideSetting
                        QuerySuccessful = $amsiConfigurationQuerySuccessful
                    }

                    $amsiMultiConfigObject += $amsiConfigTempCustomObject
                }
            } else {
                Write-Verbose "No setting override found that overrides AMSI configuration"
                $amsiState = $true
            }
        } catch {
            Write-Verbose "Unable to query AMSI configuration state"
            Invoke-CatchActions
        }
    } end {
        if ($amsiMultiConfigObject) {
            return $amsiMultiConfigObject
        }

        return [PSCustomObject]@{
            Enabled         = $amsiState
            QuerySuccessful = $amsiConfigurationQuerySuccessful
        }
    }
}

function Get-ExchangeApplicationConfigurationFileValidation {
    param(
        [string[]]$ConfigFileLocation
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $results = @{}
    $ConfigFileLocation |
        ForEach-Object {
            $obj = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlockDescription "Getting Exchange Application Configuration File Validation" `
                -CatchActionFunction ${Function:Invoke-CatchActions} `
                -ScriptBlock {
                param($Location)
                return [PSCustomObject]@{
                    Present  = ((Test-Path $Location))
                    FileName = ([IO.Path]::GetFileName($Location))
                    FilePath = $Location
                }
            } -ArgumentList $_
            $results.Add($obj.FileName, $obj)
        }
    return $results
}

function Get-ExchangeConnectors {
    [CmdletBinding()]
    [OutputType("System.Object[]")]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,
        [Parameter(Mandatory = $false)]
        [object]
        $CertificateObject
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - Computername: $ComputerName"
        function ExchangeConnectorObjectFactory {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [object]
                $ConnectorObject
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            $exchangeFactoryConnectorReturnObject = [PSCustomObject]@{
                Identity           = $ConnectorObject.Identity
                Name               = $ConnectorObject.Name
                Enabled            = $ConnectorObject.Enabled
                CloudEnabled       = $false
                ConnectorType      = $null
                TransportRole      = $null
                SmartHosts         = $null
                AddressSpaces      = $null
                RequireTLS         = $false
                TlsAuthLevel       = $null
                TlsDomain          = $null
                CertificateDetails = [PSCustomObject]@{
                    CertificateMatchDetected = $false
                    GoodTlsCertificateSyntax = $false
                    TlsCertificateName       = $null
                    TlsCertificateNameStatus = $null
                    TlsCertificateSet        = $false
                    CertificateLifetimeInfo  = $null
                }
            }

            Write-Verbose ("Creating object for Exchange connector: '{0}'" -f $ConnectorObject.Identity)
            if ($null -ne $ConnectorObject.Server) {
                Write-Verbose "Exchange ReceiveConnector detected"
                $exchangeFactoryConnectorReturnObject.ConnectorType =  "Receive"
                $exchangeFactoryConnectorReturnObject.TransportRole = $ConnectorObject.TransportRole
                if (-not([System.String]::IsNullOrEmpty($ConnectorObject.TlsDomainCapabilities))) {
                    $exchangeFactoryConnectorReturnObject.CloudEnabled = $true
                }
            } else {
                Write-Verbose "Exchange SendConnector detected"
                $exchangeFactoryConnectorReturnObject.ConnectorType = "Send"
                $exchangeFactoryConnectorReturnObject.CloudEnabled = $ConnectorObject.CloudServicesMailEnabled
                $exchangeFactoryConnectorReturnObject.TlsDomain = $ConnectorObject.TlsDomain
                if ($null -ne $ConnectorObject.TlsAuthLevel) {
                    $exchangeFactoryConnectorReturnObject.TlsAuthLevel = $ConnectorObject.TlsAuthLevel
                }

                if ($null -ne $ConnectorObject.SmartHosts) {
                    $exchangeFactoryConnectorReturnObject.SmartHosts = $ConnectorObject.SmartHosts
                }

                if ($null -ne $ConnectorObject.AddressSpaces) {
                    $exchangeFactoryConnectorReturnObject.AddressSpaces = $ConnectorObject.AddressSpaces
                }
            }

            if ($null -ne $ConnectorObject.TlsCertificateName) {
                Write-Verbose "TlsCertificateName is configured on this connector"
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateSet = $true
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateName = ($ConnectorObject.TlsCertificateName).ToString()
            } else {
                Write-Verbose "TlsCertificateName is not configured on this connector"
                $exchangeFactoryConnectorReturnObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNameEmpty"
            }

            $exchangeFactoryConnectorReturnObject.RequireTLS = $ConnectorObject.RequireTLS

            return $exchangeFactoryConnectorReturnObject
        }

        function NormalizeTlsCertificateName {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $TlsCertificateName
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                Write-Verbose ("TlsCertificateName that was passed: '{0}'" -f $TlsCertificateName)
                # RegEx to match the recommended value which is "<I>X.500Issuer<S>X.500Subject"
                if ($TlsCertificateName -match "(<i>).*(<s>).*") {
                    $expectedTlsCertificateNameDetected = $true
                    $issuerIndex = $TlsCertificateName.IndexOf("<I>", [System.StringComparison]::OrdinalIgnoreCase)
                    $subjectIndex = $TlsCertificateName.IndexOf("<S>", [System.StringComparison]::OrdinalIgnoreCase)

                    Write-Verbose "TlsCertificateName that matches the expected syntax was passed"
                } else {
                    # Failsafe to detect cases where <I> and <S> are missing in TlsCertificateName
                    $issuerIndex = $TlsCertificateName.IndexOf("CN=", [System.StringComparison]::OrdinalIgnoreCase)
                    $subjectIndex = $TlsCertificateName.LastIndexOf("CN=", [System.StringComparison]::OrdinalIgnoreCase)

                    Write-Verbose "TlsCertificateName with bad syntax was passed"
                }

                # We stop processing if Issuer OR Subject index is -1 (no match found)
                if (($issuerIndex -ne -1) -and
                    ($subjectIndex -ne -1)) {
                    if ($expectedTlsCertificateNameDetected) {
                        $issuer = $TlsCertificateName.Substring(($issuerIndex + 3), ($subjectIndex - 3))
                        $subject = $TlsCertificateName.Substring($subjectIndex + 3)
                    } else {
                        $issuer  = $TlsCertificateName.Substring($issuerIndex, $subjectIndex)
                        $subject = $TlsCertificateName.Substring($subjectIndex)
                    }
                }

                if (($null -ne $issuer) -and
                    ($null -ne $subject)) {
                    return [PSCustomObject]@{
                        Issuer     = $issuer
                        Subject    = $subject
                        GoodSyntax = $expectedTlsCertificateNameDetected
                    }
                }
            } catch {
                Write-Verbose "We hit an exception while parsing the TlsCertificateName string"
                Invoke-CatchActions
            }
        }

        function FindMatchingExchangeCertificate {
            [CmdletBinding()]
            [OutputType("System.Object")]
            param(
                [Parameter(Mandatory = $true)]
                [object]
                $CertificateObject,
                [Parameter(Mandatory = $true)]
                [object]
                $ConnectorCustomObject
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                Write-Verbose ("{0} connector object(s) was/were passed to process" -f $ConnectorCustomObject.Count)
                foreach ($connectorObject in $ConnectorCustomObject) {

                    if ($null -ne $ConnectorObject.CertificateDetails.TlsCertificateName) {
                        $connectorTlsCertificateNormalizedObject = NormalizeTlsCertificateName `
                            -TlsCertificateName $ConnectorObject.CertificateDetails.TlsCertificateName

                        if ($null -eq $connectorTlsCertificateNormalizedObject) {
                            Write-Verbose "Unable to normalize TlsCertificateName - could be caused by an invalid TlsCertificateName configuration"
                            $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNameSyntaxInvalid"
                        } else {
                            if ($connectorTlsCertificateNormalizedObject.GoodSyntax) {
                                $connectorObject.CertificateDetails.GoodTlsCertificateSyntax = $connectorTlsCertificateNormalizedObject.GoodSyntax
                            }

                            $certificateMatches = 0
                            $certificateLifetimeInformation = @{}
                            foreach ($certificate in $CertificateObject) {
                                if (($certificate.Issuer -eq $connectorTlsCertificateNormalizedObject.Issuer) -and
                                    ($certificate.Subject -eq $connectorTlsCertificateNormalizedObject.Subject)) {
                                    Write-Verbose ("Certificate: '{0}' matches Connectors: '{1}' TlsCertificateName: '{2}'" -f $certificate.Thumbprint, $connectorObject.Identity, $connectorObject.CertificateDetails.TlsCertificateName)
                                    $connectorObject.CertificateDetails.CertificateMatchDetected = $true
                                    $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateMatch"
                                    $certificateLifetimeInformation.Add($certificate.Thumbprint, $certificate.LifetimeInDays)

                                    $certificateMatches++
                                }
                            }

                            if ($certificateMatches -eq 0) {
                                Write-Verbose "No matching certificate was found on the server"
                                $connectorObject.CertificateDetails.TlsCertificateNameStatus = "TlsCertificateNotFound"
                            } else {
                                Write-Verbose ("We found: '{0}' matching certificates on the server" -f $certificateMatches)
                                $connectorObject.CertificateDetails.CertificateLifetimeInfo = $certificateLifetimeInformation
                            }
                        }
                    }
                }
            } catch {
                Write-Verbose "Hit an exception while trying to locate the configured certificate on the system"
                Invoke-CatchActions
            }

            return $ConnectorCustomObject
        }
    }
    process {
        Write-Verbose ("Trying to query Exchange connectors for server: '{0}'" -f $ComputerName)
        try {
            $allReceiveConnectors = Get-ReceiveConnector -Server $ComputerName -ErrorAction Stop
            $allSendConnectors = Get-SendConnector -ErrorAction Stop
            $connectorCustomObject = @()

            foreach ($receiveConnector in $allReceiveConnectors) {
                $connectorCustomObject += ExchangeConnectorObjectFactory -ConnectorObject $receiveConnector
            }

            foreach ($sendConnector in $allSendConnectors) {
                $connectorCustomObject += ExchangeConnectorObjectFactory -ConnectorObject $sendConnector
            }

            if (($null -ne $connectorCustomObject) -and
                ($null -ne $CertificateObject)) {
                $connectorReturnObject = FindMatchingExchangeCertificate `
                    -CertificateObject $CertificateObject `
                    -ConnectorCustomObject $connectorCustomObject
            } else {
                Write-Verbose "No connector object which can be processed was returned"
                $connectorReturnObject = $connectorCustomObject
            }
        } catch {
            Write-Verbose "Hit an exception while processing the Exchange Send-/Receive Connectors"
            Invoke-CatchActions
        }
    }
    end {
        return $connectorReturnObject
    }
}

function Get-ExchangeDependentServices {
    [CmdletBinding()]
    param(
        [string]$MachineName
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $criticalWindowServices = @("WinMgmt", "W3Svc", "IISAdmin", "Pla", "MpsSvc",
            "RpcEptMapper", "EventLog").ToLower()
        $criticalExchangeServices = @("MSExchangeADTopology", "MSExchangeDelivery",
            "MSExchangeFastSearch", "MSExchangeFrontEndTransport", "MSExchangeIS",
            "MSExchangeRepl", "MSExchangeRPC", "MSExchangeServiceHost",
            "MSExchangeSubmission", "MSExchangeTransport", "HostControllerService").ToLower()
        $commonExchangeServices = @("MSExchangeAntispamUpdate", "MSComplianceAudit", "MSExchangeCompliance",
            "MSExchangeDagMgmt", "MSExchangeDiagnostics", "MSExchangeEdgeSync",
            "MSExchangeHM", "MSExchangeHMRecovery", "MSExchangeMailboxAssistants",
            "MSExchangeMailboxReplication", "MSExchangeMitigation",
            "MSExchangeThrottling", "MSExchangeTransportLogSearch", "BITS").ToLower()
        $criticalServices = New-Object 'System.Collections.Generic.List[object]'
        $commonServices = New-Object 'System.Collections.Generic.List[object]'
        $getServicesList = New-Object 'System.Collections.Generic.List[object]'
        function TestServiceRunning {
            param(
                [object]$Service
            )
            Write-Verbose "Testing $($Service.Name) - Status: $($Service.Status)"
            if ($Service.Status.ToString() -eq "Running") { return $true }
            return $false
        }

        function NewServiceObject {
            param(
                [object]$Service
            )
            $name = $Service.Name
            $status = "Unknown"
            $startType = "Unknown"
            try {
                $status = $Service.Status.ToString()
            } catch {
                Write-Verbose "Failed to set Status of service '$name'"
                Invoke-CatchActions
            }
            try {
                $startType = $Service.StartType.ToString()
            } catch {
                Write-Verbose "Failed to set Start Type of service '$name'"
                Invoke-CatchActions
            }
            return [PSCustomObject]@{
                Name      = $name
                Status    = $status
                StartType = $startType
            }
        }
    } process {
        try {
            $getServices = Get-Service -ComputerName $MachineName -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to get the services on the server"
            Invoke-CatchActions
            return
        }

        foreach ($service in $getServices) {
            if (($criticalWindowServices.Contains($service.Name.ToLower()) -or
                    $criticalExchangeServices.Contains($service.Name.ToLower())) -and
                (-not (TestServiceRunning $service))) {
                $criticalServices.Add((NewServiceObject $service))
            } elseif ($commonExchangeServices.Contains($service.Name.ToLower()) -and
                (-not (TestServiceRunning $service))) {
                $commonServices.Add((NewServiceObject $service))
            }
            $getServicesList.Add((NewServiceObject $service))
        }
    } end {
        return [PSCustomObject]@{
            Services = $getServicesList
            Critical = $criticalServices
            Common   = $commonServices
        }
    }
}

function Get-ExchangeEmergencyMitigationServiceState {
    [CmdletBinding()]
    [OutputType("System.Object")]
    param(
        [Parameter(Mandatory = $true)]
        [object]
        $RequiredInformation,
        [Parameter(Mandatory = $false)]
        [scriptblock]
        $CatchActionFunction
    )
    begin {
        $computerName = $RequiredInformation.ComputerName
        $emergencyMitigationServiceOrgState = $RequiredInformation.MitigationsEnabled
        $exchangeServerConfiguration = $RequiredInformation.GetExchangeServer
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - Computername: $ComputerName"
    }
    process {
        if ($null -ne $emergencyMitigationServiceOrgState) {
            Write-Verbose "Exchange Emergency Mitigation Service detected"
            try {
                $exchangeEmergencyMitigationWinServiceRating = $null
                $emergencyMitigationWinService = Get-Service -ComputerName $ComputerName -Name MSExchangeMitigation -ErrorAction Stop
                if (($emergencyMitigationWinService.Status.ToString() -eq "Running") -and
                    ($emergencyMitigationWinService.StartType.ToString() -eq "Automatic")) {
                    $exchangeEmergencyMitigationWinServiceRating = "Running"
                } else {
                    $exchangeEmergencyMitigationWinServiceRating = "Investigate"
                }
            } catch {
                Write-Verbose "Failed to query EEMS Windows service data"
                Invoke-CatchActionError $CatchActionFunction
            }

            $eemsEndpoint = Invoke-ScriptBlockHandler -ComputerName $ComputerName -ScriptBlockDescription "Test EEMS pattern service connectivity" `
                -CatchActionFunction $CatchActionFunction `
                -ScriptBlock {
                [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12; `
                    if ($null -ne $args[0]) {
                    Write-Verbose "Proxy Server detected. Going to use: $($args[0])"
                    [System.Net.WebRequest]::DefaultWebProxy = New-Object System.Net.WebProxy($args[0])
                    [System.Net.WebRequest]::DefaultWebProxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
                    [System.Net.WebRequest]::DefaultWebProxy.BypassProxyOnLocal = $true
                }; `
                    Invoke-WebRequest -Method Get -Uri "https://officeclient.microsoft.com/getexchangemitigations" -UseBasicParsing
            } `
                -ArgumentList $exchangeServerConfiguration.InternetWebProxy
        }
    }
    end {
        return [PSCustomObject]@{
            MitigationWinServiceState = $exchangeEmergencyMitigationWinServiceRating
            MitigationServiceOrgState = $emergencyMitigationServiceOrgState
            MitigationServiceSrvState = $exchangeServerConfiguration.MitigationsEnabled
            MitigationServiceEndpoint = $eemsEndpoint.StatusCode
            MitigationsApplied        = $exchangeServerConfiguration.MitigationsApplied
            MitigationsBlocked        = $exchangeServerConfiguration.MitigationsBlocked
            DataCollectionEnabled     = $exchangeServerConfiguration.DataCollectionEnabled
        }
    }
}



function Get-RemoteRegistrySubKey {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Attempting to open the Base Key $RegistryHive on Machine $MachineName"
        $regKey = $null
    }
    process {

        try {
            $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey($RegistryHive, $MachineName)
            Write-Verbose "Attempting to open the Sub Key '$SubKey'"
            $regKey = $reg.OpenSubKey($SubKey)
            Write-Verbose "Opened Sub Key"
        } catch {
            Write-Verbose "Failed to open the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        return $regKey
    }
}

function Get-RemoteRegistryValue {
    [CmdletBinding()]
    param(
        [string]$RegistryHive = "LocalMachine",
        [string]$MachineName,
        [string]$SubKey,
        [string]$GetValue,
        [string]$ValueType,
        [scriptblock]$CatchActionFunction
    )

    <#
    Valid ValueType return values (case-sensitive)
    (https://docs.microsoft.com/en-us/dotnet/api/microsoft.win32.registryvaluekind?view=net-5.0)
    Binary = REG_BINARY
    DWord = REG_DWORD
    ExpandString = REG_EXPAND_SZ
    MultiString = REG_MULTI_SZ
    None = No data type
    QWord = REG_QWORD
    String = REG_SZ
    Unknown = An unsupported registry data type
    #>

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $registryGetValue = $null
    }
    process {

        try {

            $regSubKey = Get-RemoteRegistrySubKey -RegistryHive $RegistryHive `
                -MachineName $MachineName `
                -SubKey $SubKey

            if (-not ([System.String]::IsNullOrWhiteSpace($regSubKey))) {
                Write-Verbose "Attempting to get the value $GetValue"
                $registryGetValue = $regSubKey.GetValue($GetValue)
                Write-Verbose "Finished running GetValue()"

                if ($null -ne $registryGetValue -and
                    (-not ([System.String]::IsNullOrWhiteSpace($ValueType)))) {
                    Write-Verbose "Validating ValueType $ValueType"
                    $registryValueType = $regSubKey.GetValueKind($GetValue)
                    Write-Verbose "Finished running GetValueKind()"

                    if ($ValueType -ne $registryValueType) {
                        Write-Verbose "ValueType: $ValueType is different to the returned ValueType: $registryValueType"
                        $registryGetValue = $null
                    } else {
                        Write-Verbose "ValueType matches: $ValueType"
                    }
                }
            }
        } catch {
            Write-Verbose "Failed to get the value on the registry"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
    end {
        Write-Verbose "Get-RemoteRegistryValue Return Value: '$registryGetValue'"
        return $registryGetValue
    }
}
function Get-ExchangeRegistryValues {
    [CmdletBinding()]
    param(
        [string]$MachineName,
        [scriptblock]$CatchActionFunction
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    $baseParams = @{
        MachineName         = $MachineName
        CatchActionFunction = $CatchActionFunction
    }

    $ctsParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Search\SystemParameters"
        GetValue = "CtsProcessorAffinityPercentage"
    }

    $fipsParams = $baseParams + @{
        SubKey   = "SYSTEM\CurrentControlSet\Control\Lsa\FipsAlgorithmPolicy"
        GetValue = "Enabled"
    }

    $blockReplParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Replay\Parameters"
        GetValue = "DisableGranularReplication"
    }

    $disableAsyncParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15"
        GetValue = "DisableAsyncNotification"
    }

    $installDirectoryParams = $baseParams + @{
        SubKey   = "SOFTWARE\Microsoft\ExchangeServer\v15\Setup"
        GetValue = "MsiInstallPath"
    }

    return [PSCustomObject]@{
        CtsProcessorAffinityPercentage = [int](Get-RemoteRegistryValue @ctsParams)
        FipsAlgorithmPolicyEnabled     = [int](Get-RemoteRegistryValue @fipsParams)
        DisableGranularReplication     = [int](Get-RemoteRegistryValue @blockReplParams)
        DisableAsyncNotification       = [int](Get-RemoteRegistryValue @disableAsyncParams)
        MisInstallPath                 = [string](Get-RemoteRegistryValue @installDirectoryParams)
    }
}

function Get-ExchangeServerCertificates {
    param()

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function NewCertificateExclusionEntry {
        [OutputType("System.Object")]
        param(
            [Parameter(Mandatory = $true)]
            [string]
            $IssuerOrSubjectPattern,
            [Parameter(Mandatory = $true)]
            [bool]
            $IsSelfSigned
        )

        return [PSCustomObject]@{
            IorSPattern  = $IssuerOrSubjectPattern
            IsSelfSigned = $IsSelfSigned
        }
    }

    function ShouldCertificateBeSkipped {
        [OutputType("System.Boolean")]
        param (
            [Parameter(Mandatory = $true)]
            [PSCustomObject]
            $Exclusions,
            [Parameter(Mandatory = $true)]
            [System.Security.Cryptography.X509Certificates.X509Certificate2]
            $Certificate
        )

        $certificateMatch = $Exclusions | Where-Object {
            ((($Certificate.Subject -match $_.IorSPattern) -or
            ($Certificate.Issuer -match $_.IorSPattern)) -and
            ($Certificate.IsSelfSigned -eq $_.IsSelfSigned))
        } | Select-Object -First 1

        if ($null -ne $certificateMatch) {
            return $certificateMatch.IsSelfSigned -eq $Certificate.IsSelfSigned
        }
        return $false
    }

    try {
        Write-Verbose "Build certificate exclusion list"
        # Add the certificates that should be excluded from the Exchange certificate check (we don't return an object for them)
        # Exclude "MS-Organization-P2P-Access [YYYY]" certificate with one day lifetime on Azure hosted machines.
        # See: What are the MS-Organization-P2P-Access certificates present on our Windows 10/11 devices?
        # https://docs.microsoft.com/azure/active-directory/devices/faq
        # Exclude "DC=Windows Azure CRP Certificate Generator" (TenantEncryptionCertificate)
        # The certificates are built by the Azure fabric controller and passed to the Azure VM Agent.
        # If you stop and start the VM every day, the fabric controller might create a new certificate.
        # These certificates can be deleted. The Azure VM Agent re-creates certificates if needed.
        # https://docs.microsoft.com/azure/virtual-machines/extensions/features-windows
        $certificatesToExclude = @(
            NewCertificateExclusionEntry "CN=MS-Organization-P2P-Access \[[12][0-9]{3}\]$" $false
            NewCertificateExclusionEntry "DC=Windows Azure CRP Certificate Generator" $true
        )
        Write-Verbose "Trying to receive certificates from Exchange server: $($Script:Server)"
        $exchangeServerCertificates = Get-ExchangeCertificate -Server $Script:Server -ErrorAction Stop

        if ($null -ne $exchangeServerCertificates) {
            try {
                $authConfig = Get-AuthConfig -ErrorAction Stop
                $authConfigDetected = $true
            } catch {
                $authConfigDetected = $false
                Invoke-CatchActions
            }

            [array]$certObject = @()
            foreach ($cert in $exchangeServerCertificates) {
                try {
                    $certificateLifetime = ([System.Convert]::ToDateTime($cert.NotAfter, [System.Globalization.DateTimeFormatInfo]::InvariantInfo) - (Get-Date)).Days
                    $sanCertificateInfo = $false

                    $excludeCertificate = ShouldCertificateBeSkipped -Exclusions $certificatesToExclude -Certificate $cert

                    if ($excludeCertificate) {
                        Write-Verbose "Excluding certificate $($cert.Subject). Moving to next certificate"
                        continue
                    }

                    $currentErrors = $Error.Count
                    if ($null -ne $cert.DnsNameList -and
                        ($cert.DnsNameList).Count -gt 1) {
                        $sanCertificateInfo = $true
                        $certDnsNameList = $cert.DnsNameList
                    } elseif ($null -eq $cert.DnsNameList) {
                        $certDnsNameList = "None"
                    } else {
                        $certDnsNameList = $cert.DnsNameList
                    }
                    if ($currentErrors -lt $Error.Count) {
                        $i = 0
                        while ($i -lt ($Error.Count - $currentErrors)) {
                            Invoke-CatchActions $Error[$i]
                            $i++
                        }
                    }

                    if ($authConfigDetected) {
                        $isAuthConfigInfo = $false

                        if ($cert.Thumbprint -eq $authConfig.CurrentCertificateThumbprint) {
                            $isAuthConfigInfo = $true
                        }
                    } else {
                        $isAuthConfigInfo = "InvalidAuthConfig"
                    }

                    if ([String]::IsNullOrEmpty($cert.FriendlyName)) {
                        $certFriendlyName = ($certDnsNameList[0]).ToString()
                    } else {
                        $certFriendlyName = $cert.FriendlyName
                    }

                    if ([String]::IsNullOrEmpty($cert.Status)) {
                        $certStatus = "Unknown"
                    } else {
                        $certStatus = ($cert.Status).ToString()
                    }

                    if ([String]::IsNullOrEmpty($cert.SignatureAlgorithm.FriendlyName)) {
                        $certSignatureAlgorithm = "Unknown"
                        $certSignatureHashAlgorithm = "Unknown"
                        $certSignatureHashAlgorithmSecure = 0
                    } else {
                        $certSignatureAlgorithm = $cert.SignatureAlgorithm.FriendlyName
                        <#
                            OID Table
                            https://docs.microsoft.com/en-us/openspecs/windows_protocols/ms-gpnap/a48b02b2-2a10-4eb0-bed4-1807a6d2f5ad
                            SignatureHashAlgorithmSecure = Unknown 0
                            SignatureHashAlgorithmSecure = Insecure/Weak 1
                            SignatureHashAlgorithmSecure = Secure 2
                        #>
                        switch ($cert.SignatureAlgorithm.Value) {
                            "1.2.840.113549.1.1.5" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.4" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.10040.4.3" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.29" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.15" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.3" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.2" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.1.1.3" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.2" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.4" { $certSignatureHashAlgorithm = "md4"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.7.2.3.1" { $certSignatureHashAlgorithm = "md2"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.13" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.3.14.3.2.27" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "2.16.840.1.101.2.1.1.19" { $certSignatureHashAlgorithm = "mosaicSignature"; $certSignatureHashAlgorithmSecure = 0 }
                            "1.3.14.3.2.26" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.113549.2.5" { $certSignatureHashAlgorithm = "md5"; $certSignatureHashAlgorithmSecure = 1 }
                            "2.16.840.1.101.3.4.2.1" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "2.16.840.1.101.3.4.2.2" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "2.16.840.1.101.3.4.2.3" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.11" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.12" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.13" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.113549.1.1.10" { $certSignatureHashAlgorithm = "rsassa-pss"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.1" { $certSignatureHashAlgorithm = "sha1"; $certSignatureHashAlgorithmSecure = 1 }
                            "1.2.840.10045.4.3.2" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3.3" { $certSignatureHashAlgorithm = "sha384"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3.4" { $certSignatureHashAlgorithm = "sha512"; $certSignatureHashAlgorithmSecure = 2 }
                            "1.2.840.10045.4.3" { $certSignatureHashAlgorithm = "sha256"; $certSignatureHashAlgorithmSecure = 2 }
                            default { $certSignatureHashAlgorithm = "Unknown"; $certSignatureHashAlgorithmSecure = 0 }
                        }
                    }

                    $certInformationObj = New-Object PSCustomObject
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Issuer" -Value $cert.Issuer
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Subject" -Value $cert.Subject
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "FriendlyName" -Value $certFriendlyName
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Thumbprint" -Value $cert.Thumbprint
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "PublicKeySize" -Value $cert.PublicKey.Key.KeySize
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureAlgorithm" -Value $certSignatureAlgorithm
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureHashAlgorithm" -Value $certSignatureHashAlgorithm
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "SignatureHashAlgorithmSecure" -Value $certSignatureHashAlgorithmSecure
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "IsSanCertificate" -Value $sanCertificateInfo
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Namespaces" -Value $certDnsNameList
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Services" -Value $cert.Services
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "IsCurrentAuthConfigCertificate" -Value $isAuthConfigInfo
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "LifetimeInDays" -Value $certificateLifetime
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "Status" -Value $certStatus
                    $certInformationObj | Add-Member -MemberType NoteProperty -Name "CertificateObject" -Value $cert

                    $certObject += $certInformationObj
                } catch {
                    Write-Verbose "Unable to process certificate: $($cert.Thumbprint)"
                    Invoke-CatchActions
                }
            }
            Write-Verbose "Processed: $($certObject.Count) certificates"
            return $certObject
        } else {
            Write-Verbose "Failed to find any Exchange certificates"
            return $null
        }
    } catch {
        Write-Verbose "Failed to run Get-ExchangeCertificate. Error: $($Error[0].Exception)."
        Invoke-CatchActions
    }
}

function Get-ExchangeServerMaintenanceState {
    param(
        [Parameter(Mandatory = $false)][array]$ComponentsToSkip
    )
    Write-Verbose "Calling Function: $($MyInvocation.MyCommand)"

    [HealthChecker.ExchangeServerMaintenance]$serverMaintenance = New-Object -TypeName HealthChecker.ExchangeServerMaintenance
    $serverMaintenance.GetServerComponentState = Get-ServerComponentState -Identity $Script:Server -ErrorAction SilentlyContinue

    try {
        $serverMaintenance.GetClusterNode = Get-ClusterNode -Name $Script:Server -ErrorAction Stop
    } catch {
        Write-Verbose "Failed to run Get-ClusterNode"
        Invoke-CatchActions
    }

    Write-Verbose "Running ServerComponentStates checks"

    foreach ($component in $serverMaintenance.GetServerComponentState) {
        if (($null -ne $ComponentsToSkip -and
                $ComponentsToSkip.Count -ne 0) -and
            $ComponentsToSkip -notcontains $component.Component) {
            if ($component.State.ToString() -ne "Active") {
                $latestLocalState = $null
                $latestRemoteState = $null

                if ($null -ne $component.LocalStates -and
                    $component.LocalStates.Count -gt 0) {
                    $latestLocalState = ($component.LocalStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                }

                if ($null -ne $component.RemoteStates -and
                    $component.RemoteStates.Count -gt 0) {
                    $latestRemoteState = ($component.RemoteStates | Sort-Object { $_.TimeStamp } -ErrorAction SilentlyContinue)[-1]
                }

                Write-Verbose "Component: '$($component.Component)' LocalState: '$($latestLocalState.State)' RemoteState: '$($latestRemoteState.State)'"

                if ($latestLocalState.State -eq $latestRemoteState.State) {
                    $serverMaintenance.InactiveComponents += "'{0}' is in Maintenance Mode" -f $component.Component
                } else {
                    if (($null -ne $latestLocalState) -and
                        ($latestLocalState.State -ne "Active")) {
                        $serverMaintenance.InactiveComponents += "'{0}' is in Local Maintenance Mode only" -f $component.Component
                    }

                    if (($null -ne $latestRemoteState) -and
                        ($latestRemoteState.State -ne "Active")) {
                        $serverMaintenance.InactiveComponents += "'{0}' is in Remote Maintenance Mode only" -f $component.Component
                    }
                }
            } else {
                Write-Verbose "Component '$($component.Component)' is Active"
            }
        } else {
            Write-Verbose "Component: $($component.Component) will be skipped"
        }
    }

    return $serverMaintenance
}

function Get-ExchangeUpdates {
    param(
        [Parameter(Mandatory = $true)][HealthChecker.ExchangeMajorVersion]$ExchangeMajorVersion
    )
    Write-Verbose("Calling: $($MyInvocation.MyCommand) Passed: $ExchangeMajorVersion")
    $RegLocation = [string]::Empty

    if ([HealthChecker.ExchangeMajorVersion]::Exchange2013 -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2013"
    } elseif ([HealthChecker.ExchangeMajorVersion]::Exchange2016 -eq $ExchangeMajorVersion) {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2016"
    } else {
        $RegLocation = "SOFTWARE\Microsoft\Updates\Exchange 2019"
    }

    $RegKey = Get-RemoteRegistrySubKey -MachineName $Script:Server `
        -SubKey $RegLocation `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $RegKey) {
        $IU = $RegKey.GetSubKeyNames()
        if ($null -ne $IU) {
            Write-Verbose "Detected fixes installed on the server"
            $fixes = @()
            foreach ($key in $IU) {
                $IUKey = $RegKey.OpenSubKey($key)
                $IUName = $IUKey.GetValue("PackageName")
                Write-Verbose "Found: $IUName"
                $fixes += $IUName
            }
            return $fixes
        } else {
            Write-Verbose "No IUs found in the registry"
        }
    } else {
        Write-Verbose "No RegKey returned"
    }

    Write-Verbose "Exiting: Get-ExchangeUpdates"
    return $null
}

function Get-ExSetupDetails {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $exSetupDetails = [string]::Empty
    function Get-ExSetupDetailsScriptBlock {
        Get-Command ExSetup | ForEach-Object { $_.FileVersionInfo }
    }

    $exSetupDetails = Invoke-ScriptBlockHandler -ComputerName $Script:Server -ScriptBlock ${Function:Get-ExSetupDetailsScriptBlock} -ScriptBlockDescription "Getting ExSetup remotely" -CatchActionFunction ${Function:Invoke-CatchActions}
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $exSetupDetails
}


function Get-FIPFSScanEngineVersionState {
    [CmdletBinding()]
    [OutputType("System.Object")]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $ComputerName,
        [Parameter(Mandatory = $true)]
        [System.Version]
        $ExSetupVersion,
        [Parameter(Mandatory = $true)]
        [HealthChecker.ExchangeServerRole]
        $ServerRole
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        function GetFolderFromExchangeInstallPath {
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $ExchangeSubDir
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"
            try {
                $exSetupPath = (Get-ItemProperty -Path Registry::HKEY_LOCAL_MACHINE\SOFTWARE\Microsoft\ExchangeServer\v15\Setup -ErrorAction Stop).MsiInstallPath
            } catch {
                # since this is a script block, can't call Invoke-CatchActions
                $exSetupPath = $env:ExchangeInstallPath
            }

            $finalPath = Join-Path $exSetupPath $ExchangeSubDir

            if ($ExchangeSubDir -notmatch '\.[a-zA-Z0-9]+$') {

                if (Test-Path $finalPath) {
                    $getDir = Get-ChildItem -Path $finalPath -Attributes Directory
                }

                return ([PSCustomObject]@{
                        Name             = $getDir.Name
                        LastWriteTimeUtc = $getDir.LastWriteTimeUtc
                        Failed           = $null -eq $getDir
                    })
            }
            return $null
        }

        function GetHighestScanEngineVersionNumber {
            param (
                [string]
                $ComputerName
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"

            try {
                $scanEngineVersions = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
                    -ScriptBlock ${Function:GetFolderFromExchangeInstallPath} `
                    -ArgumentList ("FIP-FS\Data\Engines\amd64\Microsoft\Bin") `
                    -CatchActionFunction ${Function:Invoke-CatchActions}

                if ($null -ne $scanEngineVersions) {
                    if ($scanEngineVersions.Failed) {
                        Write-Verbose "Failed to find the scan engine directory"
                    } else {
                        return [Int64]($scanEngineVersions.Name | Measure-Object -Maximum).Maximum
                    }
                } else {
                    Write-Verbose "No FIP-FS scan engine version(s) detected - GetFolderFromExchangeInstallPath returned null"
                }
            } catch {
                Write-Verbose "Error occurred while processing FIP-FS scan engine version(s)"
                Invoke-CatchActions
            }
            return $null
        }

        function IsServerRoleAffected {
            param (
                [HealthChecker.ExchangeServerRole]
                $ServerRole
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"

            # Affected roles are Hub Transport, Mailbox and MultiRole
            if (($ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) -and
                ($ServerRole -ne [HealthChecker.ExchangeServerRole]::None) -and
                ($ServerRole -ne [HealthChecker.ExchangeServerRole]::ClientAccess)) {
                Write-Verbose "Server role is affected by this FIP-FS issue"
                return $true
            } else {
                Write-Verbose "Server role is NOT affected by this FIP-FS issue"
                return $false
            }
        }
        function IsFIPFSFixedBuild {
            param (
                [System.Version]
                $BuildNumber
            )

            Write-Verbose "Calling: $($MyInvocation.MyCommand)"

            $fixedFIPFSBuild = $false

            # Fixed on Exchange side with March 2022 Security update
            if ($BuildNumber.Major -eq 15) {
                if ($BuildNumber.Minor -eq 2) {
                    $fixedFIPFSBuild = ($BuildNumber.Build -gt 986) -or
                        (($BuildNumber.Build -eq 986) -and ($BuildNumber.Revision -ge 22)) -or
                        (($BuildNumber.Build -eq 922) -and ($BuildNumber.Revision -ge 27))
                } elseif ($BuildNumber.Minor -eq 1) {
                    $fixedFIPFSBuild = ($BuildNumber.Build -gt 2375) -or
                        (($BuildNumber.Build -eq 2375) -and ($BuildNumber.Revision -ge 24)) -or
                        (($BuildNumber.Build -eq 2308) -and ($BuildNumber.Revision -ge 27))
                } else {
                    Write-Verbose "Looks like we're on Exchange 2013 which is not affected by this FIP-FS issue"
                    $fixedFIPFSBuild = $true
                }
            } else {
                Write-Verbose "We are not on Exchange version 15"
                $fixedFIPFSBuild = $true
            }

            return $fixedFIPFSBuild
        }
    } process {
        $isAffectedByFIPFSUpdateIssue = $false
        try {

            $serverRoleAffected = IsServerRoleAffected -ServerRole $ServerRole
            if ($serverRoleAffected) {
                $highestScanEngineVersionNumber = GetHighestScanEngineVersionNumber -ComputerName $ComputerName
                $fipfsIssueFixedBuild = IsFIPFSFixedBuild -BuildNumber $ExSetupVersion

                if ($null -eq $highestScanEngineVersionNumber) {
                    Write-Verbose "No scan engine version found on the computer - this can cause issues still with some transport rules"
                } elseif ($highestScanEngineVersionNumber -ge 2201010000) {
                    if ($fipfsIssueFixedBuild) {
                        Write-Verbose "Scan engine: $highestScanEngineVersionNumber detected but Exchange runs a fixed build that doesn't crash"
                    } else {
                        Write-Verbose "Scan engine: $highestScanEngineVersionNumber will cause transport queue or pattern update issues"
                    }
                    $isAffectedByFIPFSUpdateIssue = $true
                } else {
                    Write-Verbose "Scan engine: $highestScanEngineVersionNumber is safe to use"
                }
            }
        } catch {
            Write-Verbose "Failed to check for the FIP-FS update issue"
            Invoke-CatchActions
            return $null
        }
    } end {
        return [PSCustomObject]@{
            FIPFSFixedBuild             = $fipfsIssueFixedBuild
            ServerRoleAffected          = $serverRoleAffected
            HighesVersionNumberDetected = $highestScanEngineVersionNumber
            BadVersionNumberDirDetected = $isAffectedByFIPFSUpdateIssue
        }
    }
}

function Get-ServerRole {
    param(
        [Parameter(Mandatory = $true)][object]$ExchangeServerObj
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $roles = $ExchangeServerObj.ServerRole.ToString()
    Write-Verbose "Roll: $roles"
    #Need to change this to like because of Exchange 2010 with AIO with the hub role.
    if ($roles -like "Mailbox, ClientAccess*") {
        return [HealthChecker.ExchangeServerRole]::MultiRole
    } elseif ($roles -eq "Mailbox") {
        return [HealthChecker.ExchangeServerRole]::Mailbox
    } elseif ($roles -eq "Edge") {
        return [HealthChecker.ExchangeServerRole]::Edge
    } elseif ($roles -like "*ClientAccess*") {
        return [HealthChecker.ExchangeServerRole]::ClientAccess
    } else {
        return [HealthChecker.ExchangeServerRole]::None
    }
}
function Get-ExchangeInformation {
    param(
        [HealthChecker.OSServerVersion]$OSMajorVersion
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand) Passed: OSMajorVersion: $OSMajorVersion"
    [HealthChecker.ExchangeInformation]$exchangeInformation = New-Object -TypeName HealthChecker.ExchangeInformation
    $exchangeInformation.GetExchangeServer = (Get-ExchangeServer -Identity $Script:Server -Status)
    $exchangeInformation.ExchangeCertificates = Get-ExchangeServerCertificates
    $buildInformation = $exchangeInformation.BuildInformation
    $buildVersionInfo = Get-ExchangeBuildVersionInformation -AdminDisplayVersion $exchangeInformation.GetExchangeServer.AdminDisplayVersion
    $buildInformation.MajorVersion = ([HealthChecker.ExchangeMajorVersion]$buildVersionInfo.MajorVersion)
    $buildInformation.BuildNumber = "{0}.{1}.{2}.{3}" -f $buildVersionInfo.Major, $buildVersionInfo.Minor, $buildVersionInfo.Build, $buildVersionInfo.Revision
    $buildInformation.ServerRole = (Get-ServerRole -ExchangeServerObj $exchangeInformation.GetExchangeServer)
    $buildInformation.ExchangeSetup = Get-ExSetupDetails
    $exchangeInformation.DependentServices = (Get-ExchangeDependentServices -MachineName $Script:Server)

    if ($buildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox ) {
        try {
            $exchangeInformation.GetMailboxServer = (Get-MailboxServer -Identity $Script:Server -ErrorAction Stop)
        } catch {
            Write-Verbose "Failed to run Get-MailboxServer"
            Invoke-CatchActions
        }
    }

    if (($buildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2016 -and
            $buildInformation.ServerRole -le [HealthChecker.ExchangeServerRole]::Mailbox) -or
        ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2013 -and
            ($buildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::ClientAccess -or
        $buildInformation.ServerRole -eq [HealthChecker.ExchangeServerRole]::MultiRole))) {
        $exchangeInformation.GetOwaVirtualDirectory = Get-OwaVirtualDirectory -Identity ("{0}\owa (Default Web Site)" -f $Script:Server) -ADPropertiesOnly
        $exchangeInformation.GetWebServicesVirtualDirectory = Get-WebServicesVirtualDirectory -Server $Script:Server
    }

    if ($Script:ExchangeShellComputer.ToolsOnly) {
        $buildInformation.LocalBuildNumber = "{0}.{1}.{2}.{3}" -f $Script:ExchangeShellComputer.Major, $Script:ExchangeShellComputer.Minor, `
            $Script:ExchangeShellComputer.Build, `
            $Script:ExchangeShellComputer.Revision
    }

    #Exchange 2013 or greater
    if ($buildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $netFrameworkExchange = $exchangeInformation.NETFramework
        [System.Version]$adminDisplayVersionFullBuildNumber = $buildInformation.BuildNumber
        Write-Verbose "The AdminDisplayVersion build number is: $adminDisplayVersionFullBuildNumber"
        #Build Numbers: https://docs.microsoft.com/en-us/Exchange/new-features/build-numbers-and-release-dates?view=exchserver-2019
        if ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2019) {
            Write-Verbose "Exchange 2019 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2019 "
            $buildInformation.ExtendedSupportDate = "10/14/2025"

            #Exchange 2019 Information
            if ($adminDisplayVersionFullBuildNumber -lt "15.2.330.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::RTM
                $buildInformation.FriendlyName += "RTM"
                $buildInformation.ReleaseDate = "10/22/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.397.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.464.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "06/18/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.529.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "09/17/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.595.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "12/17/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.659.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "03/17/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.721.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "06/16/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.792.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "09/15/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.858.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "12/15/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.922.7") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "03/16/2021"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.986.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "06/29/2021"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.2.1118.7") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "09/28/2021"
                $buildInformation.SupportedBuild = $true
            } elseif ($adminDisplayVersionFullBuildNumber -ge "15.2.1118.7") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU12
                $buildInformation.FriendlyName += "CU12"
                $buildInformation.ReleaseDate = "04/20/2022"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2019 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        } elseif ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2016) {
            Write-Verbose "Exchange 2016 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2016 "
            $buildInformation.ExtendedSupportDate = "10/14/2025"

            #Exchange 2016 Information
            if ($adminDisplayVersionFullBuildNumber -lt "15.1.466.34") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "03/15/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.544.27") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "06/21/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.669.32") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "09/20/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.845.34") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "12/13/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1034.26") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "03/21/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1261.35") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "06/24/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1415.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "09/16/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1466.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "12/19/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1531.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "03/20/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1591.10") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "06/19/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1713.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "10/16/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1779.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU12
                $buildInformation.FriendlyName += "CU12"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1847.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU13
                $buildInformation.FriendlyName += "CU13"
                $buildInformation.ReleaseDate = "06/18/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1913.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU14
                $buildInformation.FriendlyName += "CU14"
                $buildInformation.ReleaseDate = "09/17/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.1979.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU15
                $buildInformation.FriendlyName += "CU15"
                $buildInformation.ReleaseDate = "12/17/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2044.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU16
                $buildInformation.FriendlyName += "CU16"
                $buildInformation.ReleaseDate = "03/17/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2106.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU17
                $buildInformation.FriendlyName += "CU17"
                $buildInformation.ReleaseDate = "06/16/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2176.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU18
                $buildInformation.FriendlyName += "CU18"
                $buildInformation.ReleaseDate = "09/15/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2242.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU19
                $buildInformation.FriendlyName += "CU19"
                $buildInformation.ReleaseDate = "12/15/2020"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2308.8") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU20
                $buildInformation.FriendlyName += "CU20"
                $buildInformation.ReleaseDate = "03/16/2021"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2375.7") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU21
                $buildInformation.FriendlyName += "CU21"
                $buildInformation.ReleaseDate = "06/29/2021"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.1.2507.6") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU22
                $buildInformation.FriendlyName += "CU22"
                $buildInformation.ReleaseDate = "09/28/2021"
                $buildInformation.SupportedBuild = $true
            } elseif ($adminDisplayVersionFullBuildNumber -ge "15.1.2507.6") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU23
                $buildInformation.FriendlyName += "CU23"
                $buildInformation.ReleaseDate = "04/20/2022"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2016 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU2) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU3) {

                if ($OSMajorVersion -eq [HealthChecker.OSServerVersion]::Windows2016) {
                    $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                    $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                } else {
                    $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                    $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
                }
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU8) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU10) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU11) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU13) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        } else {
            Write-Verbose "Exchange 2013 is detected. Checking build number..."
            $buildInformation.FriendlyName = "Exchange 2013 "
            $buildInformation.ExtendedSupportDate = "04/11/2023"

            #Exchange 2013 Information
            if ($adminDisplayVersionFullBuildNumber -lt "15.0.712.24") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU1
                $buildInformation.FriendlyName += "CU1"
                $buildInformation.ReleaseDate = "04/02/2013"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.775.38") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU2
                $buildInformation.FriendlyName += "CU2"
                $buildInformation.ReleaseDate = "07/09/2013"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.847.32") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU3
                $buildInformation.FriendlyName += "CU3"
                $buildInformation.ReleaseDate = "11/25/2013"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.913.22") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU4
                $buildInformation.FriendlyName += "CU4"
                $buildInformation.ReleaseDate = "02/25/2014"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.995.29") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU5
                $buildInformation.FriendlyName += "CU5"
                $buildInformation.ReleaseDate = "05/27/2014"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1044.25") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU6
                $buildInformation.FriendlyName += "CU6"
                $buildInformation.ReleaseDate = "08/26/2014"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1076.9") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU7
                $buildInformation.FriendlyName += "CU7"
                $buildInformation.ReleaseDate = "12/09/2014"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1104.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU8
                $buildInformation.FriendlyName += "CU8"
                $buildInformation.ReleaseDate = "03/17/2015"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1130.7") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU9
                $buildInformation.FriendlyName += "CU9"
                $buildInformation.ReleaseDate = "06/17/2015"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1156.6") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU10
                $buildInformation.FriendlyName += "CU10"
                $buildInformation.ReleaseDate = "09/15/2015"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1178.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU11
                $buildInformation.FriendlyName += "CU11"
                $buildInformation.ReleaseDate = "12/15/2015"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1210.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU12
                $buildInformation.FriendlyName += "CU12"
                $buildInformation.ReleaseDate = "03/15/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1236.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU13
                $buildInformation.FriendlyName += "CU13"
                $buildInformation.ReleaseDate = "06/21/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1263.5") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU14
                $buildInformation.FriendlyName += "CU14"
                $buildInformation.ReleaseDate = "09/20/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1293.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU15
                $buildInformation.FriendlyName += "CU15"
                $buildInformation.ReleaseDate = "12/13/2016"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1320.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU16
                $buildInformation.FriendlyName += "CU16"
                $buildInformation.ReleaseDate = "03/21/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1347.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU17
                $buildInformation.FriendlyName += "CU17"
                $buildInformation.ReleaseDate = "06/24/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1365.1") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU18
                $buildInformation.FriendlyName += "CU18"
                $buildInformation.ReleaseDate = "09/16/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1367.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU19
                $buildInformation.FriendlyName += "CU19"
                $buildInformation.ReleaseDate = "12/19/2017"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1395.4") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU20
                $buildInformation.FriendlyName += "CU20"
                $buildInformation.ReleaseDate = "03/20/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1473.3") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU21
                $buildInformation.FriendlyName += "CU21"
                $buildInformation.ReleaseDate = "06/19/2018"
            } elseif ($adminDisplayVersionFullBuildNumber -lt "15.0.1497.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU22
                $buildInformation.FriendlyName += "CU22"
                $buildInformation.ReleaseDate = "02/12/2019"
            } elseif ($adminDisplayVersionFullBuildNumber -ge "15.0.1497.2") {
                $buildInformation.CU = [HealthChecker.ExchangeCULevel]::CU23
                $buildInformation.FriendlyName += "CU23"
                $buildInformation.ReleaseDate = "06/18/2019"
                $buildInformation.SupportedBuild = $true
            }

            #Exchange 2013 .NET Information
            if ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU4) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU13) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d2wFix
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d1wFix
            } elseif ($buildInformation.CU -eq [HealthChecker.ExchangeCULevel]::CU15) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d5d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU19) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU21) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d6d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
            } elseif ($buildInformation.CU -lt [HealthChecker.ExchangeCULevel]::CU23) {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d1
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
            } else {
                $netFrameworkExchange.MinSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d7d2
                $netFrameworkExchange.MaxSupportedVersion = [HealthChecker.NetMajorVersion]::Net4d8
            }
        }

        try {
            $organizationConfig = Get-OrganizationConfig -ErrorAction Stop
            $exchangeInformation.GetOrganizationConfig = $organizationConfig
        } catch {
            Write-Yellow "Failed to run Get-OrganizationConfig."
            Invoke-CatchActions
        }

        $mitigationsEnabled = $null
        if ($null -ne $organizationConfig) {
            $mitigationsEnabled = $organizationConfig.MitigationsEnabled
        }

        $exchangeInformation.ExchangeEmergencyMitigationService = Get-ExchangeEmergencyMitigationServiceState `
            -RequiredInformation ([PSCustomObject]@{
                ComputerName       = $Script:Server
                MitigationsEnabled = $mitigationsEnabled
                GetExchangeServer  = $exchangeInformation.GetExchangeServer
            }) `
            -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -ne $organizationConfig) {
            $exchangeInformation.MapiHttpEnabled = $organizationConfig.MapiHttpEnabled
            if ($null -ne $organizationConfig.EnableDownloadDomains) {
                $exchangeInformation.EnableDownloadDomains = $organizationConfig.EnableDownloadDomains
            }
        } else {
            Write-Verbose "MAPI HTTP Enabled and Download Domains Enabled results not accurate"
        }

        try {
            $exchangeInformation.WildCardAcceptedDomain = Get-AcceptedDomain | Where-Object { $_.DomainName.ToString() -eq "*" }
        } catch {
            Write-Verbose "Failed to run Get-AcceptedDomain"
            $exchangeInformation.WildCardAcceptedDomain = "Unknown"
            Invoke-CatchActions
        }

        if (($OSMajorVersion -ge [HealthChecker.OSServerVersion]::Windows2016) -and
            ($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge)) {
            $exchangeInformation.AMSIConfiguration = Get-ExchangeAMSIConfigurationState
        } else {
            Write-Verbose "AMSI Interface is not available on this OS / Exchange server role"
        }

        $exchangeInformation.RegistryValues = Get-ExchangeRegistryValues -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $serverExchangeBinDirectory = [System.Io.Path]::Combine($exchangeInformation.RegistryValues.MisInstallPath, "Bin\")
        Write-Verbose "Found Exchange Bin: $serverExchangeBinDirectory"

        if ($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::Edge) {
            $exchangeInformation.ApplicationPools = Get-ExchangeAppPoolsInformation
            try {
                $exchangeInformation.GetHybridConfiguration = Get-HybridConfiguration -ErrorAction Stop
            } catch {
                Write-Yellow "Failed to run Get-HybridConfiguration"
                Invoke-CatchActions
            }

            Write-Verbose "Query Exchange Connector settings via 'Get-ExchangeConnectors'"
            $exchangeInformation.ExchangeConnectors = Get-ExchangeConnectors `
                -ComputerName $Script:Server `
                -CertificateObject $exchangeInformation.ExchangeCertificates

            $exchangeServerIISParams = @{
                ComputerName        = $Script:Server
                IsLegacyOS          = ($OSMajorVersion -lt [HealthChecker.OSServerVersion]::Windows2016)
                CatchActionFunction = ${Function:Invoke-CatchActions}
            }

            Write-Verbose "Trying to query Exchange Server IIS settings"
            $exchangeInformation.IISSettings = Get-ExchangeServerIISSettings @exchangeServerIISParams

            Write-Verbose "Query Exchange AD permissions for CVE-2022-21978 testing"
            $exchangeInformation.ExchangeAdPermissions = Get-ExchangeAdPermissions -ExchangeVersion $buildInformation.MajorVersion -OSVersion $OSMajorVersion

            Write-Verbose "Query extended protection configuration for multiple CVEs testing"
            $getExtendedProtectionConfigurationParams = @{
                ComputerName        = $Script:Server
                ExSetupVersion      = $buildInformation.ExchangeSetup.FileVersion
                CatchActionFunction = ${Function:Invoke-CatchActions}
            }

            $exchangeInformation.ExtendedProtectionConfig = Get-ExtendedProtectionConfiguration @getExtendedProtectionConfigurationParams
        }

        $exchangeInformation.ApplicationConfigFileStatus = Get-ExchangeApplicationConfigurationFileValidation -ConfigFileLocation ("{0}EdgeTransport.exe.config" -f $serverExchangeBinDirectory)

        $buildInformation.KBsInstalled = Get-ExchangeUpdates -ExchangeMajorVersion $buildInformation.MajorVersion
        if (($null -ne $buildInformation.KBsInstalled) -and ($buildInformation.KBsInstalled -like "*KB5000871*")) {
            Write-Verbose "March 2021 SU: KB5000871 was detected on the system"
            $buildInformation.March2021SUInstalled = $true
        } else {
            Write-Verbose "March 2021 SU: KB5000871 was not detected on the system"
            $buildInformation.March2021SUInstalled = $false
        }

        Write-Verbose "Query schema class information for CVE-2021-34470 testing"
        try {
            $exchangeInformation.msExchStorageGroup = Get-ExchangeAdSchemaClass -SchemaClassName "ms-Exch-Storage-Group"
        } catch {
            Write-Verbose "Failed to run Get-ExchangeAdSchemaClass"
            Invoke-CatchActions
        }

        Write-Verbose "Checking if FIP-FS is affected by the pattern issue"
        $fipfsParams = @{
            ComputerName   = $Script:Server
            ExSetupVersion = $buildInformation.ExchangeSetup.FileVersion
            ServerRole     = $buildInformation.ServerRole
        }

        $buildInformation.FIPFSUpdateIssue = Get-FIPFSScanEngineVersionState @fipfsParams
        $exchangeInformation.ServerMaintenance = Get-ExchangeServerMaintenanceState -ComponentsToSkip "ForwardSyncDaemon", "ProvisioningRps"
        $exchangeInformation.SettingOverrides = Get-ExchangeSettingOverride -Server $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

        if (($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::ClientAccess) -and
            ($buildInformation.ServerRole -ne [HealthChecker.ExchangeServerRole]::None)) {
            try {
                $testServiceHealthResults = Test-ServiceHealth -Server $Script:Server -ErrorAction Stop
                foreach ($notRunningService in $testServiceHealthResults.ServicesNotRunning) {
                    if ($exchangeInformation.ExchangeServicesNotRunning -notcontains $notRunningService) {
                        $exchangeInformation.ExchangeServicesNotRunning += $notRunningService
                    }
                }
            } catch {
                Write-Verbose "Failed to run Test-ServiceHealth"
                Invoke-CatchActions
            }
        }
    } elseif ($buildInformation.MajorVersion -eq [HealthChecker.ExchangeMajorVersion]::Exchange2010) {
        Write-Verbose "Exchange 2010 detected."
        $buildInformation.FriendlyName = "Exchange 2010"
        $buildInformation.BuildNumber = $exchangeInformation.GetExchangeServer.AdminDisplayVersion.ToString()
    }

    Write-Verbose "Exiting: Get-ExchangeInformation"
    return $exchangeInformation
}




function Get-WmiObjectHandler {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingWMICmdlet', '', Justification = 'This is what this function is for')]
    [CmdletBinding()]
    param(
        [string]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [string]
        $Class,

        [string]
        $Filter,

        [string]
        $Namespace,

        [scriptblock]
        $CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | Class: '$Class' | Filter: '$Filter' | Namespace: '$Namespace'"

        $execute = @{
            ComputerName = $ComputerName
            Class        = $Class
        }

        if (-not ([string]::IsNullOrEmpty($Filter))) {
            $execute.Add("Filter", $Filter)
        }

        if (-not ([string]::IsNullOrEmpty($Namespace))) {
            $execute.Add("Namespace", $Namespace)
        }
    }
    process {
        try {
            $wmi = Get-WmiObject @execute -ErrorAction Stop
            return $wmi
        } catch {
            Write-Verbose "Failed to run Get-WmiObject on class '$class'"

            if ($null -ne $CatchActionFunction) {
                & $CatchActionFunction
            }
        }
    }
}
function Get-WmiObjectCriticalHandler {
    [CmdletBinding()]
    param(
        [string]
        $ComputerName = $env:COMPUTERNAME,

        [Parameter(Mandatory = $true)]
        [string]
        $Class,

        [string]
        $Filter,

        [string]
        $Namespace,

        [scriptblock]
        $CatchActionFunction
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $params = @{
        ComputerName        = $ComputerName
        Class               = $Class
        Filter              = $Filter
        Namespace           = $Namespace
        CatchActionFunction = $CatchActionFunction
    }


    $wmi = Get-WmiObjectHandler @params

    if ($null -eq $wmi) {
        # Check for common issues that have been seen. If common issue, the do a Write-Error, custom message, then exit to maintain readability.

        if ($Error[0].Exception.ErrorCode -eq 0x800703FA) {
            Write-Verbose "Registry key marked for deletion."
            Write-Error $Error[0]
            $message = "A registry key is marked for deletion that was attempted to read from for the cmdlet 'Get-WmiObject -Class $Class'.`r`n"
            $message += "`tThis error goes away after some time and/or a reboot of the computer. At that time you should be able to run Health Checker again."
            Write-Warning $message
            exit
        }

        # Grab the English version of hte message and/or the error code. Could get a different error code if service is not disabled.
        if ($Error[0].Exception.Message -like "The service cannot be started, either because it is disabled or because it has no enabled devices associated with it. *" -or
            $Error[0].Exception.ErrorCode -eq 0x80070422) {
            Write-Verbose "winmgmt service is disabled or not working."
            Write-Error $Error[0]
            Write-Warning "The 'winmgmt' service appears to not be working correctly. Please make sure it is set to Automatic and in a running state. This script will fail unless this is working correctly."
            exit
        }

        throw "Failed to get critical information. Stopping the script. InnerException: $($Error[0])"
    }

    return $wmi
}
function Get-ProcessorInformation {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $wmiObject = $null
        $processorName = [string]::Empty
        $maxClockSpeed = 0
        $numberOfLogicalCores = 0
        $numberOfPhysicalCores = 0
        $numberOfProcessors = 0
        $currentClockSpeed = 0
        $processorIsThrottled = $false
        $differentProcessorCoreCountDetected = $false
        $differentProcessorsDetected = $false
        $presentedProcessorCoreCount = 0
        $previousProcessor = $null
    }
    process {
        $wmiObject = @(Get-WmiObjectCriticalHandler -ComputerName $MachineName -Class "Win32_Processor" -CatchActionFunction $CatchActionFunction)
        $processorName = $wmiObject[0].Name
        $maxClockSpeed = $wmiObject[0].MaxClockSpeed
        Write-Verbose "Evaluating processor results"

        foreach ($processor in $wmiObject) {
            $numberOfPhysicalCores += $processor.NumberOfCores
            $numberOfLogicalCores += $processor.NumberOfLogicalProcessors
            $numberOfProcessors++

            if ($processor.CurrentClockSpeed -lt $processor.MaxClockSpeed) {
                Write-Verbose "Processor is being throttled"
                $processorIsThrottled = $true
                $currentClockSpeed = $processor.CurrentClockSpeed
            }

            if ($null -ne $previousProcessor) {

                if ($processor.Name -ne $previousProcessor.Name -or
                    $processor.MaxClockSpeed -ne $previousProcessor.MaxClockSpeed) {
                    Write-Verbose "Different Processors are detected!!! This is an issue."
                    $differentProcessorsDetected = $true
                }

                if ($processor.NumberOfLogicalProcessors -ne $previousProcessor.NumberOfLogicalProcessors) {
                    Write-Verbose "Different Processor core count per processor socket detected. This is an issue."
                    $differentProcessorCoreCountDetected = $true
                }
            }
            $previousProcessor = $processor
        }

        $presentedProcessorCoreCount = Invoke-ScriptBlockHandler -ComputerName $MachineName `
            -ScriptBlock { [System.Environment]::ProcessorCount } `
            -ScriptBlockDescription "Trying to get the System.Environment ProcessorCount" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $presentedProcessorCoreCount) {
            Write-Verbose "Wasn't able to get Presented Processor Core Count on the Server. Setting to -1."
            $presentedProcessorCoreCount = -1
        }
    }
    end {
        Write-Verbose "PresentedProcessorCoreCount: $presentedProcessorCoreCount"
        Write-Verbose "NumberOfPhysicalCores: $numberOfPhysicalCores | NumberOfLogicalCores: $numberOfLogicalCores | NumberOfProcessors: $numberOfProcessors"
        Write-Verbose "ProcessorIsThrottled: $processorIsThrottled | CurrentClockSpeed: $currentClockSpeed"
        Write-Verbose "DifferentProcessorsDetected: $differentProcessorsDetected | DifferentProcessorCoreCountDetected: $differentProcessorCoreCountDetected"
        return [PSCustomObject]@{
            Name                                = $processorName
            MaxMegacyclesPerCore                = $maxClockSpeed
            NumberOfPhysicalCores               = $numberOfPhysicalCores
            NumberOfLogicalCores                = $numberOfLogicalCores
            NumberOfProcessors                  = $numberOfProcessors
            CurrentMegacyclesPerCore            = $currentClockSpeed
            ProcessorIsThrottled                = $processorIsThrottled
            DifferentProcessorsDetected         = $differentProcessorsDetected
            DifferentProcessorCoreCountDetected = $differentProcessorCoreCountDetected
            EnvironmentProcessorCount           = $presentedProcessorCoreCount
            ProcessorClassObject                = $wmiObject
        }
    }
}

function Get-ServerType {
    [CmdletBinding()]
    [OutputType("System.String")]
    param(
        [Parameter(Mandatory = $true)]
        [string]
        $ServerType
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ServerType: $ServerType"
        $returnServerType = [string]::Empty
    }
    process {
        if ($ServerType -like "VMWare*") { $returnServerType = "VMware" }
        elseif ($ServerType -like "*Amazon EC2*") { $returnServerType = "AmazonEC2" }
        elseif ($ServerType -like "*Microsoft Corporation*") { $returnServerType = "HyperV" }
        elseif ($ServerType.Length -gt 0) { $returnServerType = "Physical" }
        else { $returnServerType = "Unknown" }
    }
    end {
        Write-Verbose "Returning: $returnServerType"
        return $returnServerType
    }
}
function Get-HardwareInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.HardwareInformation]$hardware_obj = New-Object HealthChecker.HardwareInformation
    $system = Get-WmiObjectCriticalHandler -ComputerName $Script:Server -Class "Win32_ComputerSystem" -CatchActionFunction ${Function:Invoke-CatchActions}
    $hardware_obj.MemoryInformation = Get-WmiObjectHandler -ComputerName $Script:Server -Class "Win32_PhysicalMemory" -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -eq $hardware_obj.MemoryInformation) {
        Write-Verbose "Using memory from Win32_ComputerSystem class instead. This may cause memory calculation issues."
        $hardware_obj.TotalMemory = $system.TotalPhysicalMemory
    } else {
        foreach ($memory in $hardware_obj.MemoryInformation) {
            $hardware_obj.TotalMemory += $memory.Capacity
        }
    }
    $hardware_obj.Manufacturer = $system.Manufacturer
    $hardware_obj.System = $system
    $hardware_obj.AutoPageFile = $system.AutomaticManagedPagefile
    $hardware_obj.ServerType = (Get-ServerType -ServerType $system.Manufacturer)
    $processorInformation = Get-ProcessorInformation -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

    #Need to do it this way because of Windows 2012R2
    $processor = New-Object HealthChecker.ProcessorInformation
    $processor.Name = $processorInformation.Name
    $processor.NumberOfPhysicalCores = $processorInformation.NumberOfPhysicalCores
    $processor.NumberOfLogicalCores = $processorInformation.NumberOfLogicalCores
    $processor.NumberOfProcessors = $processorInformation.NumberOfProcessors
    $processor.MaxMegacyclesPerCore = $processorInformation.MaxMegacyclesPerCore
    $processor.CurrentMegacyclesPerCore = $processorInformation.CurrentMegacyclesPerCore
    $processor.ProcessorIsThrottled = $processorInformation.ProcessorIsThrottled
    $processor.DifferentProcessorsDetected = $processorInformation.DifferentProcessorsDetected
    $processor.DifferentProcessorCoreCountDetected = $processorInformation.DifferentProcessorCoreCountDetected
    $processor.EnvironmentProcessorCount = $processorInformation.EnvironmentProcessorCount
    $processor.ProcessorClassObject = $processorInformation.ProcessorClassObject

    $hardware_obj.Processor = $processor
    $hardware_obj.Model = $system.Model

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $hardware_obj
}



function Get-ServerRebootPending {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {

        function Get-PendingFileReboot {
            try {
                if ((Get-ItemProperty -Path "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" -Name PendingFileRenameOperations -ErrorAction Stop)) {
                    return $true
                }
                return $false
            } catch {
                throw
            }
        }

        function Get-UpdateExeVolatile {
            try {
                $updateExeVolatileProps = Get-ItemProperty -Path "HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\" -ErrorAction Stop
                if ($null -ne $updateExeVolatileProps -and $null -ne $updateExeVolatileProps.Flags) {
                    return $true
                }
                return $false
            } catch {
                throw
            }
        }

        function Get-PendingCCMReboot {
            try {
                return (Invoke-CimMethod -Namespace 'Root\ccm\clientSDK' -ClassName 'CCM_ClientUtilities' -Name 'DetermineIfRebootPending' -ErrorAction Stop)
            } catch {
                throw
            }
        }

        function Get-PathTestingReboot {
            param(
                [string]$TestingPath
            )

            return (Test-Path $TestingPath)
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $pendingRebootLocations = New-Object 'System.Collections.Generic.List[string]'
    }
    process {
        $pendingFileRenameOperationValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingFileReboot} `
            -ScriptBlockDescription "Get-PendingFileReboot" `
            -CatchActionFunction $CatchActionFunction

        if ($null -eq $pendingFileRenameOperationValue) {
            $pendingFileRenameOperationValue = $false
        }

        $componentBasedServicingPendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Component Based Servicing Reboot Pending" `
            -ArgumentList "HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending" `
            -CatchActionFunction $CatchActionFunction

        $ccmReboot = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PendingCCMReboot} `
            -ScriptBlockDescription "Get-PendingSCCMReboot" `
            -CatchActionFunction $CatchActionFunction

        $autoUpdatePendingRebootValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-PathTestingReboot} `
            -ScriptBlockDescription "Auto Update Pending Reboot" `
            -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired" `
            -CatchActionFunction $CatchActionFunction

        $updateExeVolatileValue = Invoke-ScriptBlockHandler -ComputerName $ServerName -ScriptBlock ${Function:Get-UpdateExeVolatile} `
            -ScriptBlockDescription "UpdateExeVolatile Reboot Pending" `
            -CatchActionFunction $CatchActionFunction

        $ccmRebootPending = $ccmReboot -and ($ccmReboot.RebootPending -or $ccmReboot.IsHardRebootPending)
        $pendingReboot = $ccmRebootPending -or $pendingFileRenameOperationValue -or $componentBasedServicingPendingRebootValue -or $autoUpdatePendingRebootValue -or $updateExeVolatileValue

        if ($ccmRebootPending) {
            Write-Verbose "RebootPending in CCM_ClientUtilities"
            $pendingRebootLocations.Add("CCM_ClientUtilities Showing Reboot Pending")
        }

        if ($pendingFileRenameOperationValue) {
            Write-Verbose "RebootPending at HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations"
            $pendingRebootLocations.Add("HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\PendingFileRenameOperations")
        }

        if ($componentBasedServicingPendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending"
            $pendingRebootLocations.Add("HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending")
        }

        if ($autoUpdatePendingRebootValue) {
            Write-Verbose "RebootPending at HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired"
            $pendingRebootLocations.Add("HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired")
        }

        if ($updateExeVolatileValue) {
            Write-Verbose "RebootPending at HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\Flags"
            $pendingRebootLocations.Add("HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\Flags")
        }
    }
    end {
        return [PSCustomObject]@{
            PendingFileRenameOperations          = $pendingFileRenameOperationValue
            ComponentBasedServicingPendingReboot = $componentBasedServicingPendingRebootValue
            AutoUpdatePendingReboot              = $autoUpdatePendingRebootValue
            UpdateExeVolatileValue               = $updateExeVolatileValue
            CcmRebootPending                     = $ccmRebootPending
            PendingReboot                        = $pendingReboot
            PendingRebootLocations               = $pendingRebootLocations
        }
    }
}


function Get-AllTlsSettingsFromRegistry {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {

        function Get-TLSMemberValue {
            param(
                [Parameter(Mandatory = $true)]
                [string]
                $GetKeyType,

                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter( Mandatory = $false)]
                [bool]
                $NullIsEnabled
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | GetKeyType: $GetKeyType | NullIsEnabled: $NullIsEnabled"
            switch ($GetKeyType) {
                "Enabled" {
                    return ($null -eq $KeyValue -and $NullIsEnabled) -or $KeyValue -eq 1
                }
                "DisabledByDefault" {
                    return $null -ne $KeyValue -and $KeyValue -eq 1
                }
            }
        }

        function Get-NETDefaultTLSValue {
            param(
                [Parameter(Mandatory = $false)]
                [object]
                $KeyValue,

                [Parameter(Mandatory = $true)]
                [string]
                $NetVersion,

                [Parameter(Mandatory = $true)]
                [string]
                $KeyName
            )
            Write-Verbose "KeyValue is null: '$($null -eq $KeyValue)' | KeyValue: '$KeyValue' | NetVersion: '$NetVersion' | KeyName: '$KeyName'"
            return $null -ne $KeyValue -and $KeyValue -eq 1
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - MachineName: '$MachineName'"
        $registryBase = "SYSTEM\CurrentControlSet\Control\SecurityProviders\SCHANNEL\Protocols\TLS {0}\{1}"
        $tlsVersions = @("1.0", "1.1", "1.2", "1.3")
        $enabledKey = "Enabled"
        $disabledKey = "DisabledByDefault"
        $netVersions = @("v2.0.50727", "v4.0.30319")
        $netRegistryBase = "SOFTWARE\{0}\.NETFramework\{1}"
        $allTlsObjects = [PSCustomObject]@{
            "TLS" = @{}
            "NET" = @{}
        }
    }
    process {
        foreach ($tlsVersion in $tlsVersions) {
            $registryServer = $registryBase -f $tlsVersion, "Server"
            $registryClient = $registryBase -f $tlsVersion, "Client"

            # Get the Enabled and DisabledByDefault values
            $serverEnabledValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $registryServer `
                -GetValue $enabledKey `
                -CatchActionFunction $CatchActionFunction
            $serverDisabledByDefaultValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $registryServer `
                -GetValue $disabledKey `
                -CatchActionFunction $CatchActionFunction
            $clientEnabledValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $registryClient `
                -GetValue $enabledKey `
                -CatchActionFunction $CatchActionFunction
            $clientDisabledByDefaultValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $registryClient `
                -GetValue $disabledKey `
                -CatchActionFunction $CatchActionFunction

            $serverEnabled = (Get-TLSMemberValue -GetKeyType $enabledKey -KeyValue $serverEnabledValue -NullIsEnabled ($tlsVersion -ne "1.3"))
            $serverDisabledByDefault = (Get-TLSMemberValue -GetKeyType $disabledKey -KeyValue $serverDisabledByDefaultValue)
            $clientEnabled = (Get-TLSMemberValue -GetKeyType $enabledKey -KeyValue $clientEnabledValue -NullIsEnabled ($tlsVersion -ne "1.3"))
            $clientDisabledByDefault = (Get-TLSMemberValue -GetKeyType $disabledKey -KeyValue $clientDisabledByDefaultValue)
            $disabled = $serverEnabled -eq $false -and ($serverDisabledByDefault -or $null -eq $serverDisabledByDefaultValue) -and
            $clientEnabled -eq $false -and ($clientDisabledByDefault -or $null -eq $clientDisabledByDefaultValue)
            $misconfigured = $serverEnabled -ne $clientEnabled -or $serverDisabledByDefault -ne $clientDisabledByDefault
            # only need to test server settings here, because $misconfigured will be set and will be the official status.
            # want to check for if Server is Disabled and Disabled By Default is not set or the reverse. This would be only part disabled
            # and not what we recommend on the blog post.
            $halfDisabled = ($serverEnabled -eq $false -and $serverDisabledByDefault -eq $false -and $null -ne $serverDisabledByDefaultValue) -or
                ($serverEnabled -and $serverDisabledByDefault)
            $configuration = "Enabled"

            if ($disabled) {
                Write-Verbose "TLS is Disabled"
                $configuration = "Disabled"
            }

            if ($halfDisabled) {
                Write-Verbose "TLS is only half disabled"
                $configuration = "Half Disabled"
            }

            if ($misconfigured) {
                Write-Verbose "TLS is misconfigured"
                $configuration = "Misconfigured"
            }

            $currentTLSObject = [PSCustomObject]@{
                TLSVersion                 = $tlsVersion
                "Server$enabledKey"        = $serverEnabled
                "Server$enabledKey`Value"  = $serverEnabledValue
                "Server$disabledKey"       = $serverDisabledByDefault
                "Server$disabledKey`Value" = $serverDisabledByDefaultValue
                "ServerRegistryPath"       = $registryServer
                "Client$enabledKey"        = $clientEnabled
                "Client$enabledKey`Value"  = $clientEnabledValue
                "Client$disabledKey"       = $clientDisabledByDefault
                "Client$disabledKey`Value" = $clientDisabledByDefaultValue
                "ClientRegistryPath"       = $registryClient
                "TLSVersionDisabled"       = $disabled
                "TLSMisconfigured"         = $misconfigured
                "TLSHalfDisabled"          = $halfDisabled
                "TLSConfiguration"         = $configuration
            }
            $allTlsObjects.TLS.Add($TlsVersion, $currentTLSObject)
        }

        foreach ($netVersion in $netVersions) {

            $msRegistryKey = $netRegistryBase -f "Microsoft", $netVersion
            $wowMsRegistryKey = $netRegistryBase -f "Wow6432Node\Microsoft", $netVersion

            $systemDefaultTlsVersionsValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $msRegistryKey `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $schUseStrongCryptoValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $msRegistryKey `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction
            $wowSystemDefaultTlsVersionsValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $wowMsRegistryKey `
                -GetValue "SystemDefaultTlsVersions" `
                -CatchActionFunction $CatchActionFunction
            $wowSchUseStrongCryptoValue = Get-RemoteRegistryValue `
                -MachineName $MachineName `
                -SubKey $wowMsRegistryKey `
                -GetValue "SchUseStrongCrypto" `
                -CatchActionFunction $CatchActionFunction

            $systemDefaultTlsVersions = (Get-NETDefaultTLSValue -KeyValue $SystemDefaultTlsVersionsValue -NetVersion $netVersion -KeyName "SystemDefaultTlsVersions")
            $wowSystemDefaultTlsVersions = (Get-NETDefaultTLSValue -KeyValue $wowSystemDefaultTlsVersionsValue -NetVersion $netVersion -KeyName "WowSystemDefaultTlsVersions")

            $currentNetTlsDefaultVersionObject = [PSCustomObject]@{
                NetVersion                       = $netVersion
                SystemDefaultTlsVersions         = $systemDefaultTlsVersions
                SystemDefaultTlsVersionsValue    = $systemDefaultTlsVersionsValue
                SchUseStrongCrypto               = (Get-NETDefaultTLSValue -KeyValue $schUseStrongCryptoValue -NetVersion $netVersion -KeyName "SchUseStrongCrypto")
                SchUseStrongCryptoValue          = $schUseStrongCryptoValue
                MicrosoftRegistryLocation        = $msRegistryKey
                WowSystemDefaultTlsVersions      = $wowSystemDefaultTlsVersions
                WowSystemDefaultTlsVersionsValue = $wowSystemDefaultTlsVersionsValue
                WowSchUseStrongCrypto            = (Get-NETDefaultTLSValue -KeyValue $wowSchUseStrongCryptoValue -NetVersion $netVersion -KeyName "WowSchUseStrongCrypto")
                WowSchUseStrongCryptoValue       = $wowSchUseStrongCryptoValue
                WowRegistryLocation              = $wowMsRegistryKey
                SdtvConfiguredCorrectly          = $systemDefaultTlsVersions -eq $wowSystemDefaultTlsVersions
                SdtvEnabled                      = $systemDefaultTlsVersions -and $wowSystemDefaultTlsVersions
            }

            $hashKeyName = "NET{0}" -f ($netVersion.Split(".")[0])
            $allTlsObjects.NET.Add($hashKeyName, $currentNetTlsDefaultVersionObject)
        }
        return $allTlsObjects
    }
}


function Get-TlsCipherSuiteInformation {
    [OutputType("System.Object")]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $tlsCipherReturnObject = New-Object 'System.Collections.Generic.List[object]'
    }
    process {
        # 'Get-TlsCipherSuite' takes account of the cipher suites which are configured by the help of GPO.
        # No need to query the ciphers defined via GPO if this call is successful.
        Write-Verbose "Trying to query TlsCipherSuites via 'Get-TlsCipherSuite'"
        $getTlsCipherSuiteParams = @{
            ComputerName        = $MachineName
            ScriptBlock         = { Get-TlsCipherSuite }
            CatchActionFunction = $CatchActionFunction
        }
        $tlsCipherSuites = Invoke-ScriptBlockHandler @getTlsCipherSuiteParams

        if ($null -eq $tlsCipherSuites) {
            # If we can't get the ciphers via cmdlet, we need to query them via registry call and need to check
            # if ciphers suites are defined via GPO as well. If there are some, these take precedence over what
            # is in the default location.
            Write-Verbose "Failed to query TlsCipherSuites via 'Get-TlsCipherSuite' fallback to registry"

            $policyTlsRegistryParams = @{
                MachineName         = $MachineName
                Subkey              = "SOFTWARE\Policies\Microsoft\Cryptography\Configuration\SSL\00010002"
                GetValue            = "Functions"
                ValueType           = "String"
                CatchActionFunction = $CatchActionFunction
            }

            Write-Verbose "Trying to query cipher suites configured via GPO from registry"
            $policyDefinedCiphers = Get-RemoteRegistryValue @policyTlsRegistryParams

            if ($null -ne $policyDefinedCiphers) {
                Write-Verbose "Ciphers specified via GPO found - these take precedence over what is in the default location"
                $tlsCipherSuites = $policyDefinedCiphers.Split(",")
            } else {
                Write-Verbose "No cipher suites configured via GPO found - going to query the local TLS cipher suites"
                $tlsRegistryParams = @{
                    MachineName         = $MachineName
                    SubKey              = "SYSTEM\CurrentControlSet\Control\Cryptography\Configuration\Local\SSL\00010002"
                    GetValue            = "Functions"
                    ValueType           = "MultiString"
                    CatchActionFunction = $CatchActionFunction
                }

                $tlsCipherSuites = Get-RemoteRegistryValue @tlsRegistryParams
            }
        }

        if ($null -ne $tlsCipherSuites) {
            foreach ($cipher in $tlsCipherSuites) {
                $tlsCipherReturnObject.Add([PSCustomObject]@{
                        Name        = if ($null -eq $cipher.Name) { $cipher } else { $cipher.Name }
                        CipherSuite = if ($null -eq $cipher.CipherSuite) { "N/A" } else { $cipher.CipherSuite }
                        Cipher      = if ($null -eq $cipher.Cipher) { "N/A" } else { $cipher.Cipher }
                        Certificate = if ($null -eq $cipher.Certificate) { "N/A" } else { $cipher.Certificate }
                    })
            }
        }
    }
    end {
        return $tlsCipherReturnObject
    }
}

# Gets all related TLS Settings, from registry or other factors
function Get-AllTlsSettings {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    }
    process {
        return [PSCustomObject]@{
            Registry         = (Get-AllTlsSettingsFromRegistry -MachineName $MachineName -CatchActionFunction $CatchActionFunction)
            SecurityProtocol = (Invoke-ScriptBlockHandler -ComputerName $MachineName -ScriptBlock { ([System.Net.ServicePointManager]::SecurityProtocol).ToString() } -CatchActionFunction $CatchActionFunction)
            TlsCipherSuite   = (Get-TlsCipherSuiteInformation -MachineName $MachineName -CatchActionFunction $CatchActionFunction)
        }
    }
}

function Invoke-CatchActionErrorLoop {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, Position = 0)]
        [int]$CurrentErrors,
        [Parameter(Mandatory = $false, Position = 1)]
        [scriptblock]$CatchActionFunction
    )
    process {
        if ($null -ne $CatchActionFunction -and
            $Error.Count -ne $CurrentErrors) {
            $i = 0
            while ($i -lt ($Error.Count - $currentErrors)) {
                & $CatchActionFunction $Error[$i]
                $i++
            }
        }
    }
}
function Get-AllNicInformation {
    [CmdletBinding()]
    param(
        [string]$ComputerName = $env:COMPUTERNAME,
        [string]$ComputerFQDN,
        [scriptblock]$CatchActionFunction
    )
    begin {

        # Extract for Pester Testing - Start
        function Get-NicPnpCapabilitiesSetting {
            [CmdletBinding()]
            param(
                [ValidateNotNullOrEmpty()]
                [string]$NicAdapterComponentId
            )
            begin {
                $nicAdapterBasicPath = "SYSTEM\CurrentControlSet\Control\Class\{4D36E972-E325-11CE-BFC1-08002bE10318}"
                [int]$i = 0
                Write-Verbose "Probing started to detect NIC adapter registry path"
            }
            process {
                $registrySubKey = Get-RemoteRegistrySubKey -MachineName $ComputerName `
                    -SubKey $nicAdapterBasicPath
                $optionalKeys = $registrySubKey.GetSubKeyNames() | Where-Object { $_ -like "0*" }
                do {
                    $nicAdapterPnPCapabilitiesProbingKey = "$nicAdapterBasicPath\$($optionalKeys[$i])"
                    $netCfgInstanceId = Get-RemoteRegistryValue -MachineName $ComputerName `
                        -SubKey $nicAdapterPnPCapabilitiesProbingKey `
                        -GetValue "NetCfgInstanceId" `
                        -CatchActionFunction $CatchActionFunction

                    if ($netCfgInstanceId -eq $NicAdapterComponentId) {
                        Write-Verbose "Matching ComponentId found - now checking for PnPCapabilitiesValue"
                        $nicAdapterPnPCapabilitiesValue = Get-RemoteRegistryValue -MachineName $ComputerName `
                            -SubKey $nicAdapterPnPCapabilitiesProbingKey `
                            -GetValue "PnPCapabilities" `
                            -CatchActionFunction $CatchActionFunction
                        break
                    } else {
                        Write-Verbose "No matching ComponentId found"
                        $i++
                    }
                } while ($i -lt $optionalKeys.Count)
            }
            end {
                return [PSCustomObject]@{
                    PnPCapabilities   = $nicAdapterPnPCapabilitiesValue
                    SleepyNicDisabled = ($nicAdapterPnPCapabilitiesValue -eq 24 -or $nicAdapterPnPCapabilitiesValue -eq 280)
                }
            }
        }

        # Extract for Pester Testing - End

        function Get-NetworkConfiguration {
            [CmdletBinding()]
            param(
                [string]$ComputerName
            )
            begin {
                $currentErrors = $Error.Count
                $params = @{
                    ErrorAction = "Stop"
                }
            }
            process {
                try {
                    if (($ComputerName).Split(".")[0] -ne $env:COMPUTERNAME) {
                        $cimSession = New-CimSession -ComputerName $ComputerName -ErrorAction Stop
                        $params.Add("CimSession", $cimSession)
                    }
                    $networkIpConfiguration = Get-NetIPConfiguration @params | Where-Object { $_.NetAdapter.MediaConnectionState -eq "Connected" }
                    Invoke-CatchActionErrorLoop -CurrentErrors $currentErrors -CatchActionFunction $CatchActionFunction
                    return $networkIpConfiguration
                } catch {
                    Write-Verbose "Failed to run Get-NetIPConfiguration. Error $($_.Exception)"
                    #just rethrow as caller will handle the catch
                    throw
                }
            }
        }

        function Get-NicInformation {
            [CmdletBinding()]
            param(
                [array]$NetworkConfiguration,
                [bool]$WmiObject
            )
            begin {

                function Get-IpvAddresses {
                    return [PSCustomObject]@{
                        Address        = ([string]::Empty)
                        Subnet         = ([string]::Empty)
                        DefaultGateway = ([string]::Empty)
                    }
                }

                if ($null -eq $NetworkConfiguration) {
                    Write-Verbose "NetworkConfiguration are null in New-NicInformation. Returning a null object."
                    return $null
                }

                $nicObjects = New-Object 'System.Collections.Generic.List[object]'
            }
            process {
                if ($WmiObject) {
                    $networkAdapterConfigurations = Get-WmiObjectHandler -ComputerName $ComputerName `
                        -Class "Win32_NetworkAdapterConfiguration" `
                        -Filter "IPEnabled = True" `
                        -CatchActionFunction $CatchActionFunction
                }

                foreach ($networkConfig in $NetworkConfiguration) {
                    $dnsClient = $null
                    $rssEnabledValue = 2
                    $netAdapterRss = $null
                    $mtuSize = 0
                    $driverDate = [DateTime]::MaxValue
                    $driverVersion = [string]::Empty
                    $description = [string]::Empty
                    $ipv4Address = @()
                    $ipv6Address = @()
                    $ipv6Enabled = $false

                    if (-not ($WmiObject)) {
                        Write-Verbose "Working on NIC: $($networkConfig.InterfaceDescription)"
                        $adapter = $networkConfig.NetAdapter

                        if ($adapter.DriverFileName -ne "NdisImPlatform.sys") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.DeviceID
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        try {
                            $dnsClient = $adapter | Get-DnsClient -ErrorAction Stop
                            $isRegisteredInDns = $dnsClient.RegisterThisConnectionsAddress
                            Write-Verbose "Got DNS Client information"
                        } catch {
                            Write-Verbose "Failed to get the DNS client information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        try {
                            $netAdapterRss = $adapter | Get-NetAdapterRss -ErrorAction Stop
                            Write-Verbose "Got Net Adapter RSS Information"

                            if ($null -ne $netAdapterRss) {
                                [int]$rssEnabledValue = $netAdapterRss.Enabled
                            }
                        } catch {
                            Write-Verbose "Failed to get RSS Information"
                            Invoke-CatchActionError $CatchActionFunction
                        }

                        foreach ($ipAddress in $networkConfig.AllIPAddresses.IPAddress) {
                            if ($ipAddress.Contains(":")) {
                                $ipv6Enabled = $true
                            }
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv4Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv4Address -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv4Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv4Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv4DefaultGateway -and
                                $i -lt $networkConfig.IPv4Address.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv4DefaultGateway[$i].NextHop
                            }
                            $ipv4Address += $newIpvAddress
                        }

                        for ($i = 0; $i -lt $networkConfig.IPv6Address.Count; $i++) {
                            $newIpvAddress = Get-IpvAddresses

                            if ($null -ne $networkConfig.IPv6Address -and
                                $i -lt $networkConfig.IPv6Address.Count) {
                                $newIpvAddress.Address = $networkConfig.IPv6Address[$i].IPAddress
                                $newIpvAddress.Subnet = $networkConfig.IPv6Address[$i].PrefixLength
                            }

                            if ($null -ne $networkConfig.IPv6DefaultGateway -and
                                $i -lt $networkConfig.IPv6DefaultGateway.Count) {
                                $newIpvAddress.DefaultGateway = $networkConfig.IPv6DefaultGateway[$i].NextHop
                            }
                            $ipv6Address += $newIpvAddress
                        }

                        $mtuSize = $adapter.MTUSize
                        $driverDate = $adapter.DriverDate
                        $driverVersion = $adapter.DriverVersionString
                        $description = $adapter.InterfaceDescription
                        $dnsServerToBeUsed = $networkConfig.DNSServer.ServerAddresses
                    } else {
                        Write-Verbose "Working on NIC: $($networkConfig.Description)"
                        $adapter = $networkConfig
                        $description = $adapter.Description

                        if ($adapter.ServiceName -ne "NdisImPlatformMp") {
                            $nicPnpCapabilitiesSetting = Get-NicPnpCapabilitiesSetting -NicAdapterComponentId $adapter.Guid
                        } else {
                            Write-Verbose "Multiplexor adapter detected. Going to skip PnpCapabilities check"
                            $nicPnpCapabilitiesSetting = [PSCustomObject]@{
                                PnPCapabilities = "MultiplexorNoPnP"
                            }
                        }

                        #set the correct $adapterConfiguration to link to the correct $networkConfig that we are on
                        $adapterConfiguration = $networkAdapterConfigurations |
                            Where-Object { $_.SettingID -eq $networkConfig.GUID -or
                                $_.SettingID -eq $networkConfig.InterfaceGuid }

                        if ($null -eq $adapterConfiguration) {
                            Write-Verbose "Failed to find correct adapterConfiguration for this networkConfig."
                            Write-Verbose "GUID: $($networkConfig.GUID) | InterfaceGuid: $($networkConfig.InterfaceGuid)"
                        }

                        $ipv6Enabled = ($adapterConfiguration.IPAddress | Where-Object { $_.Contains(":") }).Count -ge 1

                        if ($null -ne $adapterConfiguration.DefaultIPGateway) {
                            $ipv4Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(".") }
                            $ipv6Gateway = $adapterConfiguration.DefaultIPGateway | Where-Object { $_.Contains(":") }
                        } else {
                            $ipv4Gateway = "No default IPv4 gateway set"
                            $ipv6Gateway = "No default IPv6 gateway set"
                        }

                        for ($i = 0; $i -lt $adapterConfiguration.IPAddress.Count; $i++) {

                            if ($adapterConfiguration.IPAddress[$i].Contains(":")) {
                                $newIpv6Address = Get-IpvAddresses
                                if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                    $newIpv6Address.Address = $adapterConfiguration.IPAddress[$i]
                                    $newIpv6Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                }

                                $newIpv6Address.DefaultGateway = $ipv6Gateway
                                $ipv6Address += $newIpv6Address
                            } else {
                                $newIpv4Address = Get-IpvAddresses
                                if ($i -lt $adapterConfiguration.IPAddress.Count) {
                                    $newIpv4Address.Address = $adapterConfiguration.IPAddress[$i]
                                    $newIpv4Address.Subnet = $adapterConfiguration.IPSubnet[$i]
                                }

                                $newIpv4Address.DefaultGateway = $ipv4Gateway
                                $ipv4Address += $newIpv4Address
                            }
                        }

                        $isRegisteredInDns = $adapterConfiguration.FullDNSRegistrationEnabled
                        $dnsServerToBeUsed = $adapterConfiguration.DNSServerSearchOrder
                    }

                    $nicObjects.Add([PSCustomObject]@{
                            WmiObject         = $WmiObject
                            Name              = $adapter.Name
                            LinkSpeed         = ((($adapter.Speed) / 1000000).ToString() + " Mbps")
                            DriverDate        = $driverDate
                            NetAdapterRss     = $netAdapterRss
                            RssEnabledValue   = $rssEnabledValue
                            IPv6Enabled       = $ipv6Enabled
                            Description       = $description
                            DriverVersion     = $driverVersion
                            MTUSize           = $mtuSize
                            PnPCapabilities   = $nicPnpCapabilitiesSetting.PnpCapabilities
                            SleepyNicDisabled = $nicPnpCapabilitiesSetting.SleepyNicDisabled
                            IPv4Addresses     = $ipv4Address
                            IPv6Addresses     = $ipv6Address
                            RegisteredInDns   = $isRegisteredInDns
                            DnsServer         = $dnsServerToBeUsed
                            DnsClient         = $dnsClient
                        })
                }
            }
            end {
                Write-Verbose "Found $($nicObjects.Count) active adapters on the computer."
                Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
                return $nicObjects
            }
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed - ComputerName: '$ComputerName' | ComputerFQDN: '$ComputerFQDN'"
    }
    process {
        try {
            try {
                $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerName
            } catch {
                Invoke-CatchActionError $CatchActionFunction

                try {
                    if (-not ([string]::IsNullOrEmpty($ComputerFQDN))) {
                        $networkConfiguration = Get-NetworkConfiguration -ComputerName $ComputerFQDN
                    } else {
                        $bypassCatchActions = $true
                        Write-Verbose "No FQDN was passed, going to rethrow error."
                        throw
                    }
                } catch {
                    #Just throw again
                    throw
                }
            }

            if ([String]::IsNullOrEmpty($networkConfiguration)) {
                # Throw if nothing was returned by previous calls.
                # Can be caused when executed on Server 2008 R2 where CIM namespace ROOT/StandardCimv2 is invalid.
                Write-Verbose "No value was returned by 'Get-NetworkConfiguration'. Fallback to WMI."
                throw
            }

            return (Get-NicInformation -NetworkConfiguration $networkConfiguration)
        } catch {
            if (-not $bypassCatchActions) {
                Invoke-CatchActionError $CatchActionFunction
            }

            $wmiNetworkCards = Get-WmiObjectHandler -ComputerName $ComputerName `
                -Class "Win32_NetworkAdapter" `
                -Filter "NetConnectionStatus ='2'" `
                -CatchActionFunction $CatchActionFunction

            return (Get-NicInformation -NetworkConfiguration $wmiNetworkCards -WmiObject $true)
        }
    }
}

function Get-CredentialGuardEnabled {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $registryValue = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Control\LSA" `
        -GetValue "LsaCfgFlags" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -ne $registryValue -and
        $registryValue -ne 0) {
        return $true
    }

    return $false
}

function Get-HttpProxySetting {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function GetWinHttpSettings {
        param(
            [Parameter(Mandatory = $true)][string]$RegistryLocation
        )
        $connections = Get-ItemProperty -Path $RegistryLocation
        $proxyAddress = [string]::Empty
        $byPassList = [string]::Empty

        if (($null -ne $connections) -and
            ($Connections | Get-Member).Name -contains "WinHttpSettings") {
            $onProxy = $true

            foreach ($Byte in $Connections.WinHttpSettings) {
                if ($onProxy -and
                    $Byte -ge 42) {
                    $proxyAddress += [CHAR]$Byte
                } elseif (-not $onProxy -and
                    $Byte -ge 42) {
                    $byPassList += [CHAR]$Byte
                } elseif (-not ([string]::IsNullOrEmpty($proxyAddress)) -and
                    $onProxy -and
                    $Byte -eq 0) {
                    $onProxy = $false
                }
            }
        }

        return [PSCustomObject]@{
            ProxyAddress = $(if ($proxyAddress -eq [string]::Empty) { "None" } else { $proxyAddress })
            ByPassList   = $byPassList
        }
    }

    $httpProxy32 = Invoke-ScriptBlockHandler -ComputerName $Script:Server `
        -ScriptBlock ${Function:GetWinHttpSettings} `
        -ArgumentList "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" `
        -ScriptBlockDescription "Getting 32 Http Proxy Value" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $httpProxy64 = Invoke-ScriptBlockHandler -ComputerName $Script:Server `
        -ScriptBlock ${Function:GetWinHttpSettings} `
        -ArgumentList "HKLM:\SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Internet Settings\Connections" `
        -ScriptBlockDescription "Getting 64 Http Proxy Value" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $httpProxy = [PSCustomObject]@{
        ProxyAddress         = $(if ($httpProxy32.ProxyAddress -ne "None") { $httpProxy32.ProxyAddress } else { $httpProxy64.ProxyAddress })
        ByPassList           = $(if ($httpProxy32.ByPassList -ne [string]::Empty) { $httpProxy32.ByPassList } else { $httpProxy64.ByPassList })
        HttpProxyDifference  = $httpProxy32.ProxyAddress -ne $httpProxy64.ProxyAddress
        HttpByPassDifference = $httpProxy32.ByPassList -ne $httpProxy64.ByPassList
        HttpProxy32          = $httpProxy32
        HttpProxy64          = $httpProxy64
    }

    Write-Verbose "Http Proxy 32: $($httpProxy32.ProxyAddress)"
    Write-Verbose "Http By Pass List 32: $($httpProxy32.ByPassList)"
    Write-Verbose "Http Proxy 64: $($httpProxy64.ProxyAddress)"
    Write-Verbose "Http By Pass List 64: $($httpProxy64.ByPassList)"
    Write-Verbose "Proxy Address: $($httpProxy.ProxyAddress)"
    Write-Verbose "By Pass List: $($httpProxy.ByPassList)"
    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $httpProxy
}

function Get-LmCompatibilityLevelInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.LmCompatibilityLevelInformation]$ServerLmCompatObject = New-Object -TypeName HealthChecker.LmCompatibilityLevelInformation
    $registryValue = Get-RemoteRegistryValue -RegistryHive "LocalMachine" `
        -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Control\Lsa" `
        -GetValue "LmCompatibilityLevel" `
        -ValueType "DWord" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if ($null -eq $registryValue) {
        $registryValue = 3
    }

    $ServerLmCompatObject.RegistryValue = $registryValue
    Write-Verbose "LmCompatibilityLevel Registry Value: $registryValue"

    switch ($ServerLmCompatObject.RegistryValue) {
        0 { $ServerLmCompatObject.Description = "Clients use LM and NTLM authentication, but they never use NTLMv2 session security. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        1 { $ServerLmCompatObject.Description = "Clients use LM and NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        2 { $ServerLmCompatObject.Description = "Clients use only NTLM authentication, and they use NTLMv2 session security if the server supports it. Domain controller accepts LM, NTLM, and NTLMv2 authentication." }
        3 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controllers accept LM, NTLM, and NTLMv2 authentication." }
        4 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM authentication responses, but it accepts NTLM and NTLMv2." }
        5 { $ServerLmCompatObject.Description = "Clients use only NTLMv2 authentication, and they use NTLMv2 session security if the server supports it. Domain controller refuses LM and NTLM authentication responses, but it accepts NTLMv2." }
    }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $ServerLmCompatObject
}

function Get-PageFileInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $pageFiles = @(Get-WmiObjectHandler -ComputerName $Script:Server -Class "Win32_PageFileSetting" -CatchActionFunction ${Function:Invoke-CatchActions})
    $pageFileList = New-Object 'System.Collections.Generic.List[object]'

    if ($null -eq $pageFiles -or
        $pageFiles.Count -eq 0) {
        Write-Verbose "Found No Page File Settings"
        $pageFileList.Add([PSCustomObject]@{
                Name        = [string]::Empty
                InitialSize = 0
                MaximumSize = 0
            })
    } else {
        Write-Verbose "Found $($pageFiles.Count) different page files"
    }

    foreach ($pageFile in $pageFiles) {
        $pageFileList.Add([PSCustomObject]@{
                Name        = $pageFile.Name
                InitialSize = $pageFile.InitialSize
                MaximumSize = $pageFile.MaximumSize
            })
    }

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $pageFileList
}

function Get-ServerOperatingSystemVersion {
    [CmdletBinding()]
    [OutputType("System.String")]
    param(
        [string]$OsCaption
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $osReturnValue = [string]::Empty
    }
    process {
        if ([string]::IsNullOrEmpty($OsCaption)) {
            Write-Verbose "Getting the local machine version build number"
            $OsCaption = (Get-WmiObjectHandler -Class "Win32_OperatingSystem").Caption
        }
        Write-Verbose "OsCaption: '$OsCaption'"

        switch -Wildcard ($OsCaption) {
            "*Server 2008 R2*" { $osReturnValue = "Windows2008R2"; break }
            "*Server 2008*" { $osReturnValue = "Windows2008" }
            "*Server 2012 R2*" { $osReturnValue = "Windows2012R2"; break }
            "*Server 2012*" { $osReturnValue = "Windows2012" }
            "*Server 2016*" { $osReturnValue = "Windows2016" }
            "*Server 2019*" { $osReturnValue = "Windows2019" }
            "*Server 2022*" { $osReturnValue = "Windows2022" }
            "Microsoft Windows Server Standard" { $osReturnValue = "WindowsCore" }
            "Microsoft Windows Server Datacenter" { $osReturnValue = "WindowsCore" }
            default { $osReturnValue = "Unknown" }
        }
    }
    end {
        Write-Verbose "Returned: '$osReturnValue'"
        return [string]$osReturnValue
    }
}

function Get-Smb1ServerSettings {
    [CmdletBinding()]
    param(
        [string]$ServerName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $smbServerConfiguration = $null
        $windowsFeature = $null
    }
    process {
        $smbServerConfiguration = Invoke-ScriptBlockHandler -ComputerName $ServerName `
            -ScriptBlock { Get-SmbServerConfiguration } `
            -CatchActionFunction $CatchActionFunction `
            -ScriptBlockDescription "Get-SmbServerConfiguration"

        try {
            $windowsFeature = Get-WindowsFeature "FS-SMB1" -ComputerName $ServerName -ErrorAction Stop
        } catch {
            Write-Verbose "Failed to Get-WindowsFeature for FS-SMB1"
            Invoke-CatchActionError $CatchActionFunction
        }
    }
    end {
        return [PSCustomObject]@{
            SmbServerConfiguration = $smbServerConfiguration
            WindowsFeature         = $windowsFeature
            SuccessfulGetInstall   = $null -ne $windowsFeature
            SuccessfulGetBlocked   = $null -ne $smbServerConfiguration
            Installed              = $windowsFeature.Installed -eq $true
            IsBlocked              = $smbServerConfiguration.EnableSMB1Protocol -eq $false
        }
    }
}

function Get-TimeZoneInformationRegistrySettings {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $timeZoneInformationSubKey = "SYSTEM\CurrentControlSet\Control\TimeZoneInformation"
        $actionsToTake = @()
        $dstIssueDetected = $false
    }
    process {
        $dynamicDaylightTimeDisabled = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "DynamicDaylightTimeDisabled" -CatchActionFunction $CatchActionFunction
        $timeZoneKeyName = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "TimeZoneKeyName" -CatchActionFunction $CatchActionFunction
        $standardStart = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "StandardStart" -CatchActionFunction $CatchActionFunction
        $daylightStart = Get-RemoteRegistryValue -MachineName $MachineName -SubKey $timeZoneInformationSubKey -GetValue "DaylightStart" -CatchActionFunction $CatchActionFunction

        if ([string]::IsNullOrEmpty($timeZoneKeyName)) {
            Write-Verbose "TimeZoneKeyName is null or empty. Action should be taken to address this."
            $actionsToTake += "TimeZoneKeyName is blank. Need to switch your current time zone to a different value, then switch it back to have this value populated again."
        }

        $standardStartNonZeroValue = ($null -ne ($standardStart | Where-Object { $_ -ne 0 }))
        $daylightStartNonZeroValue = ($null -ne ($daylightStart | Where-Object { $_ -ne 0 }))

        if ($dynamicDaylightTimeDisabled -ne 0 -and
            ($standardStartNonZeroValue -or
            $daylightStartNonZeroValue)) {
            Write-Verbose "Determined that there is a chance the settings set could cause a DST issue."
            $dstIssueDetected = $true
            $actionsToTake += "High Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone. `
            It is possible that you could be running into this issue. Set 'Adjust for daylight saving time automatically to on'"
        } elseif ($dynamicDaylightTimeDisabled -ne 0) {
            Write-Verbose "Daylight savings auto adjustment is disabled."
            $actionsToTake += "Warning: DynamicDaylightTimeDisabled is set, Windows can not properly detect any DST rule changes in your time zone."
        }
    }
    end {
        return [PSCustomObject]@{
            DynamicDaylightTimeDisabled = $dynamicDaylightTimeDisabled
            TimeZoneKeyName             = $timeZoneKeyName
            StandardStart               = $standardStart
            DaylightStart               = $daylightStart
            DstIssueDetected            = $dstIssueDetected
            ActionsToTake               = $actionsToTake
        }
    }
}


# Use this after the counters have been localized.
function Get-CounterSamples {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$MachineName,

        [Parameter(Mandatory = $true)]
        [string[]]$Counter,

        [string]$CustomErrorAction = "Stop"
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    try {
        return (Get-Counter -ComputerName $MachineName -Counter $Counter -ErrorAction $CustomErrorAction).CounterSamples
    } catch {
        Write-Verbose "Failed ot get counter samples"
        Invoke-CatchActions
    }
}

# Use this to localize the counters provided
function Get-LocalizedCounterSamples {
    param(
        [Parameter(Mandatory = $true)]
        [string[]]$MachineName,

        [Parameter(Mandatory = $true)]
        [string[]]$Counter,

        [string]$CustomErrorAction = "Stop"
    )

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $localizedCounters = @()

    foreach ($computer in $MachineName) {

        foreach ($currentCounter in $Counter) {
            $counterObject = Get-CounterFullNameToCounterObject -FullCounterName $currentCounter
            $localizedCounterName = Get-LocalizedPerformanceCounterName -ComputerName $computer -PerformanceCounterName $counterObject.CounterName
            $localizedObjectName = Get-LocalizedPerformanceCounterName -ComputerName $computer -PerformanceCounterName $counterObject.ObjectName
            $localizedFullCounterName = ($counterObject.FullName.Replace($counterObject.CounterName, $localizedCounterName)).Replace($counterObject.ObjectName, $localizedObjectName)

            if (-not ($localizedCounters.Contains($localizedFullCounterName))) {
                $localizedCounters += $localizedFullCounterName
            }
        }
    }

    return (Get-CounterSamples -MachineName $MachineName -Counter $localizedCounters -CustomErrorAction $CustomErrorAction)
}

function Get-LocalizedPerformanceCounterName {
    [CmdletBinding()]
    [OutputType('System.String')]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ComputerName,

        [Parameter(Mandatory = $true)]
        [string]$PerformanceCounterName
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    if ($null -eq $Script:EnglishOnlyOSCache) {
        $Script:EnglishOnlyOSCache = @{}
    }

    if ($null -eq $Script:Counter009Cache) {
        $Script:Counter009Cache = @{}
    }

    if ($null -eq $Script:CounterCurrentLanguageCache) {
        $Script:CounterCurrentLanguageCache = @{}
    }

    if (-not ($Script:EnglishOnlyOSCache.ContainsKey($ComputerName))) {
        $perfLib = Get-RemoteRegistrySubKey -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009" `
            -CatchActionFunction ${Function:Invoke-CatchActions}
        $englishOnlyOS = ($perfLib.GetSubKeyNames() |
                Where-Object { $_ -like "0*" }).Count -eq 1
        $Script:EnglishOnlyOSCache.Add($ComputerName, $englishOnlyOS)
    }

    if ($Script:EnglishOnlyOSCache[$ComputerName]) {
        Write-Verbose "English Only Machine, return same value"
        return $PerformanceCounterName
    }

    if (-not ($Script:Counter009Cache.ContainsKey($ComputerName))) {
        $enUSCounterKeys = Get-RemoteRegistryValue -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\009" `
            -GetValue "Counter" `
            -ValueType "MultiString" `
            -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -eq $enUSCounterKeys) {
            Write-Verbose "No 'en-US' (009) 'Counter' registry value found."
            return $null
        } else {
            $Script:Counter009Cache.Add($ComputerName, $enUSCounterKeys)
        }
    }

    if (-not ($Script:CounterCurrentLanguageCache.ContainsKey($ComputerName))) {
        $currentCounterKeys = Get-RemoteRegistryValue -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion\Perflib\CurrentLanguage" `
            -GetValue "Counter" `
            -ValueType "MultiString" `
            -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($null -eq $currentCounterKeys) {
            Write-Verbose "No 'localized' (CurrentLanguage) 'Counter' registry value found"
            return $null
        } else {
            $Script:CounterCurrentLanguageCache.Add($ComputerName, $currentCounterKeys)
        }
    }

    $counterName = $PerformanceCounterName.ToLower()
    Write-Verbose "Trying to query ID index for Performance Counter: $counterName"
    $enUSCounterKeys = $Script:Counter009Cache[$ComputerName]
    $currentCounterKeys = $Script:CounterCurrentLanguageCache[$ComputerName]
    $counterIdIndex = ($enUSCounterKeys.ToLower().IndexOf("$counterName") - 1)

    if ($counterIdIndex -ge 0) {
        Write-Verbose "Counter ID Index: $counterIdIndex"
        Write-Verbose "Verify Value: $($enUSCounterKeys[$counterIdIndex + 1])"
        $counterId = $enUSCounterKeys[$counterIdIndex]
        Write-Verbose "Counter ID: $counterId"
        $localizedCounterNameIndex = ($currentCounterKeys.IndexOf("$counterId") + 1)

        if ($localizedCounterNameIndex -gt 0) {
            $localCounterName = $currentCounterKeys[$localizedCounterNameIndex]
            Write-Verbose "Found Localized Counter Index: $localizedCounterNameIndex"
            Write-Verbose "Localized Counter Name: $localCounterName"
            return $localCounterName
        } else {
            Write-Verbose "Failed to find Localized Counter Index"
        }
    } else {
        Write-Verbose "Failed to find the counter ID."
    }
}

function Get-CounterFullNameToCounterObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FullCounterName
    )

    # Supported Scenarios
    # \\adt-e2k13aio1\logicaldisk(harddiskvolume1)\avg. disk sec/read
    # \\adt-e2k13aio1\\logicaldisk(harddiskvolume1)\avg. disk sec/read
    # \logicaldisk(harddiskvolume1)\avg. disk sec/read
    if (-not ($FullCounterName.StartsWith("\"))) {
        throw "Full Counter Name Should start with '\'"
    } elseif ($FullCounterName.StartsWith("\\")) {
        $endOfServerIndex = $FullCounterName.IndexOf("\", 2)
        $serverName = $FullCounterName.Substring(2, $endOfServerIndex - 2)
    } else {
        $endOfServerIndex = 0
    }
    $startOfCounterIndex = $FullCounterName.LastIndexOf("\") + 1
    $endOfCounterObjectIndex = $FullCounterName.IndexOf("(")

    if ($endOfCounterObjectIndex -eq -1) {
        $endOfCounterObjectIndex = $startOfCounterIndex - 1
    } else {
        $instanceName = $FullCounterName.Substring($endOfCounterObjectIndex + 1, ($FullCounterName.IndexOf(")") - $endOfCounterObjectIndex - 1))
    }

    $doubleSlash = 0
    if (($FullCounterName.IndexOf("\\", 2) -ne -1)) {
        $doubleSlash = 1
    }

    return [PSCustomObject]@{
        FullName     = $FullCounterName
        ServerName   = $serverName
        ObjectName   = ($FullCounterName.Substring($endOfServerIndex + 1 + $doubleSlash, $endOfCounterObjectIndex - $endOfServerIndex - 1 - $doubleSlash))
        InstanceName = $instanceName
        CounterName  = $FullCounterName.Substring($startOfCounterIndex)
    }
}
function Get-OperatingSystemInformation {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.OperatingSystemInformation]$osInformation = New-Object HealthChecker.OperatingSystemInformation
    $win32_OperatingSystem = Get-WmiObjectCriticalHandler -ComputerName $Script:Server -Class Win32_OperatingSystem -CatchActionFunction ${Function:Invoke-CatchActions}
    $win32_PowerPlan = Get-WmiObjectHandler -ComputerName $Script:Server -Class Win32_PowerPlan -Namespace 'root\cimv2\power' -Filter "isActive='true'" -CatchActionFunction ${Function:Invoke-CatchActions}
    $currentDateTime = Get-Date
    $lastBootUpTime = [Management.ManagementDateTimeConverter]::ToDateTime($win32_OperatingSystem.lastbootuptime)
    $osInformation.BuildInformation.VersionBuild = $win32_OperatingSystem.Version
    $osInformation.BuildInformation.MajorVersion = (Get-ServerOperatingSystemVersion -OsCaption $win32_OperatingSystem.Caption)
    $osInformation.BuildInformation.FriendlyName = $win32_OperatingSystem.Caption
    $osInformation.BuildInformation.OperatingSystem = $win32_OperatingSystem
    $osInformation.ServerBootUp.Days = ($currentDateTime - $lastBootUpTime).Days
    $osInformation.ServerBootUp.Hours = ($currentDateTime - $lastBootUpTime).Hours
    $osInformation.ServerBootUp.Minutes = ($currentDateTime - $lastBootUpTime).Minutes
    $osInformation.ServerBootUp.Seconds = ($currentDateTime - $lastBootUpTime).Seconds

    if ($null -ne $win32_PowerPlan) {

        if ($win32_PowerPlan.InstanceID -eq "Microsoft:PowerPlan\{8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c}") {
            Write-Verbose "High Performance Power Plan is set to true"
            $osInformation.PowerPlan.HighPerformanceSet = $true
        } else { Write-Verbose "High Performance Power Plan is NOT set to true" }
        $osInformation.PowerPlan.PowerPlanSetting = $win32_PowerPlan.ElementName
    } else {
        Write-Verbose "Power Plan Information could not be read"
        $osInformation.PowerPlan.PowerPlanSetting = "N/A"
    }
    $osInformation.PowerPlan.PowerPlan = $win32_PowerPlan
    $osInformation.PageFile = Get-PageFileInformation
    $osInformation.NetworkInformation.NetworkAdapters = (Get-AllNicInformation -ComputerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions} -ComputerFQDN $Script:ServerFQDN)
    foreach ($adapter in $osInformation.NetworkInformation.NetworkAdapters) {

        if (!$adapter.IPv6Enabled) {
            $osInformation.NetworkInformation.IPv6DisabledOnNICs = $true
            break
        }
    }

    $osInformation.NetworkInformation.IPv6DisabledComponents = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters" `
        -GetValue "DisabledComponents" `
        -ValueType "DWord" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.TCPKeepAlive = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\Tcpip\Parameters" `
        -GetValue "KeepAliveTime" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.RpcMinConnectionTimeout = Get-RemoteRegistryValue -MachineName $Script:Server `
        -SubKey "Software\Policies\Microsoft\Windows NT\RPC\" `
        -GetValue "MinimumConnectionTimeout" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.NetworkInformation.HttpProxy = Get-HttpProxySetting
    $osInformation.InstalledUpdates.HotFixes = (Get-HotFix -ComputerName $Script:Server -ErrorAction SilentlyContinue) #old school check still valid and faster and a failsafe
    $osInformation.LmCompatibility = Get-LmCompatibilityLevelInformation
    $counterSamples = (Get-LocalizedCounterSamples -MachineName $Script:Server -Counter "\Network Interface(*)\Packets Received Discarded")

    if ($null -ne $counterSamples) {
        $osInformation.NetworkInformation.PacketsReceivedDiscarded = $counterSamples
    }

    $osInformation.ServerPendingReboot = (Get-ServerRebootPending -ServerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions})
    $timeZoneInformation = Get-TimeZoneInformationRegistrySettings -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.TimeZone.DynamicDaylightTimeDisabled = $timeZoneInformation.DynamicDaylightTimeDisabled
    $osInformation.TimeZone.TimeZoneKeyName = $timeZoneInformation.TimeZoneKeyName
    $osInformation.TimeZone.StandardStart = $timeZoneInformation.StandardStart
    $osInformation.TimeZone.DaylightStart = $timeZoneInformation.DaylightStart
    $osInformation.TimeZone.DstIssueDetected = $timeZoneInformation.DstIssueDetected
    $osInformation.TimeZone.ActionsToTake = $timeZoneInformation.ActionsToTake
    $osInformation.TimeZone.CurrentTimeZone = Invoke-ScriptBlockHandler -ComputerName $Script:Server `
        -ScriptBlock { ([System.TimeZone]::CurrentTimeZone).StandardName } `
        -ScriptBlockDescription "Getting Current Time Zone" `
        -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.TLSSettings = Get-AllTlsSettings -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.VcRedistributable = Get-VisualCRedistributableInstalledVersion -ComputerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
    $osInformation.CredentialGuardEnabled = Get-CredentialGuardEnabled
    $osInformation.RegistryValues.CurrentVersionUbr = Get-RemoteRegistryValue `
        -MachineName $Script:Server `
        -SubKey "SOFTWARE\Microsoft\Windows NT\CurrentVersion" `
        -GetValue "UBR" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $osInformation.RegistryValues.LanManServerDisabledCompression = Get-RemoteRegistryValue `
        -MachineName $Script:Server `
        -SubKey "SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters" `
        -GetValue "DisableCompression" `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    $osInformation.Smb1ServerSettings = Get-Smb1ServerSettings -ServerName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}

    Write-Verbose "Exiting: $($MyInvocation.MyCommand)"
    return $osInformation
}

function Get-DotNetDllFileVersions {
    [CmdletBinding()]
    [OutputType("System.Collections.Hashtable")]
    param(
        [string]$ComputerName,
        [array]$FileNames,
        [scriptblock]$CatchActionFunction
    )

    begin {
        function Invoke-ScriptBlockGetItem {
            param(
                [string]$FilePath
            )
            $getItem = Get-Item $FilePath

            $returnObject = ([PSCustomObject]@{
                    GetItem          = $getItem
                    LastWriteTimeUtc = $getItem.LastWriteTimeUtc
                    VersionInfo      = ([PSCustomObject]@{
                            FileMajorPart   = $getItem.VersionInfo.FileMajorPart
                            FileMinorPart   = $getItem.VersionInfo.FileMinorPart
                            FileBuildPart   = $getItem.VersionInfo.FileBuildPart
                            FilePrivatePart = $getItem.VersionInfo.FilePrivatePart
                        })
                })

            return $returnObject
        }

        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $dotNetInstallPath = [string]::Empty
        $files = @{}
    }
    process {
        $dotNetInstallPath = Get-RemoteRegistryValue -MachineName $ComputerName `
            -SubKey "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
            -GetValue "InstallPath" `
            -CatchActionFunction $CatchActionFunction

        if ([string]::IsNullOrEmpty($dotNetInstallPath)) {
            Write-Verbose "Failed to determine .NET install path"
            return
        }

        foreach ($fileName in $FileNames) {
            Write-Verbose "Querying for .NET DLL File $fileName"
            $getItem = Invoke-ScriptBlockHandler -ComputerName $ComputerName `
                -ScriptBlock ${Function:Invoke-ScriptBlockGetItem} `
                -ArgumentList ("{0}\{1}" -f $dotNetInstallPath, $filename) `
                -CatchActionFunction $CatchActionFunction
            $files.Add($fileName, $getItem)
        }
    }
    end {
        return $files
    }
}


function Get-NETFrameworkVersion {
    [CmdletBinding()]
    param(
        [string]$MachineName = $env:COMPUTERNAME,
        [int]$NetVersionKey = -1,
        [scriptblock]$CatchActionFunction
    )
    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        $friendlyName = [string]::Empty
        $minValue = -1
    }
    process {

        if ($NetVersionKey -eq -1) {
            [int]$NetVersionKey = Get-RemoteRegistryValue -MachineName $MachineName `
                -SubKey "SOFTWARE\Microsoft\NET Framework Setup\NDP\v4\Full" `
                -GetValue "Release" `
                -CatchActionFunction $CatchActionFunction
        }

        #Using Minimum Version as per https://docs.microsoft.com/en-us/dotnet/framework/migration-guide/how-to-determine-which-versions-are-installed?redirectedfrom=MSDN#minimum-version
        if ($NetVersionKey -lt 378389) {
            $friendlyName = "Unknown"
            $minValue = -1
        } elseif ($NetVersionKey -lt 378675) {
            $friendlyName = "4.5"
            $minValue = 378389
        } elseif ($NetVersionKey -lt 379893) {
            $friendlyName = "4.5.1"
            $minValue = 378675
        } elseif ($NetVersionKey -lt 393295) {
            $friendlyName = "4.5.2"
            $minValue = 379893
        } elseif ($NetVersionKey -lt 394254) {
            $friendlyName = "4.6"
            $minValue = 393295
        } elseif ($NetVersionKey -lt 394802) {
            $friendlyName = "4.6.1"
            $minValue = 394254
        } elseif ($NetVersionKey -lt 460798) {
            $friendlyName = "4.6.2"
            $minValue = 394802
        } elseif ($NetVersionKey -lt 461308) {
            $friendlyName = "4.7"
            $minValue = 460798
        } elseif ($NetVersionKey -lt 461808) {
            $friendlyName = "4.7.1"
            $minValue = 461308
        } elseif ($NetVersionKey -lt 528040) {
            $friendlyName = "4.7.2"
            $minValue = 461808
        } elseif ($NetVersionKey -ge 528040) {
            $friendlyName = "4.8"
            $minValue = 528040
        }
    }
    end {
        Write-Verbose "FriendlyName: $friendlyName | RegistryValue: $netVersionKey | MinimumValue: $minValue"
        return [PSCustomObject]@{
            FriendlyName  = $friendlyName
            RegistryValue = $NetVersionKey
            MinimumValue  = $minValue
        }
    }
}
function Get-HealthCheckerExchangeServer {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    [HealthChecker.HealthCheckerExchangeServer]$HealthExSvrObj = New-Object -TypeName HealthChecker.HealthCheckerExchangeServer
    $HealthExSvrObj.ServerName = $Script:Server
    $HealthExSvrObj.HardwareInformation = Get-HardwareInformation
    $HealthExSvrObj.OSInformation = Get-OperatingSystemInformation
    $HealthExSvrObj.ExchangeInformation = Get-ExchangeInformation -OSMajorVersion $HealthExSvrObj.OSInformation.BuildInformation.MajorVersion

    if ($HealthExSvrObj.ExchangeInformation.BuildInformation.MajorVersion -ge [HealthChecker.ExchangeMajorVersion]::Exchange2013) {
        $netFrameworkVersion = Get-NETFrameworkVersion -MachineName $Script:Server -CatchActionFunction ${Function:Invoke-CatchActions}
        $HealthExSvrObj.OSInformation.NETFramework.FriendlyName = $netFrameworkVersion.FriendlyName
        $HealthExSvrObj.OSInformation.NETFramework.RegistryValue = $netFrameworkVersion.RegistryValue
        $HealthExSvrObj.OSInformation.NETFramework.NetMajorVersion = $netFrameworkVersion.MinimumValue
        $HealthExSvrObj.OSInformation.NETFramework.FileInformation = Get-DotNetDllFileVersions -ComputerName $Script:Server -FileNames @("System.Data.dll", "System.Configuration.dll") -CatchActionFunction ${Function:Invoke-CatchActions}

        if ($netFrameworkVersion.MinimumValue -eq $HealthExSvrObj.ExchangeInformation.NETFramework.MaxSupportedVersion) {
            $HealthExSvrObj.ExchangeInformation.NETFramework.OnRecommendedVersion = $true
        }
    }
    $HealthExSvrObj.HealthCheckerVersion = $BuildVersion
    $HealthExSvrObj.GenerationTime = [datetime]::Now
    Write-Verbose "Finished building health Exchange Server Object for server: $Script:Server"
    return $HealthExSvrObj
}

function Get-ErrorsThatOccurred {

    function WriteErrorInformation {
        [CmdletBinding()]
        param(
            [object]$CurrentError
        )
        Write-VerboseErrorInformation $CurrentError
        Write-Verbose "-----------------------------------`r`n`r`n"
    }

    if ($Error.Count -gt 0) {
        Write-Grey(" "); Write-Grey(" ")
        function Write-Errors {
            Write-Verbose "`r`n`r`nErrors that occurred that wasn't handled"

            Get-UnhandledErrors | ForEach-Object {
                Write-Verbose "Error Index: $($_.Index)"
                WriteErrorInformation $_.ErrorInformation
            }

            Write-Verbose "`r`n`r`nErrors that were handled"
            Get-HandledErrors | ForEach-Object {
                Write-Verbose "Error Index: $($_.Index)"
                WriteErrorInformation $_.ErrorInformation
            }
        }

        if ((Test-UnhandledErrorsOccurred)) {
            Write-Red("There appears to have been some errors in the script. To assist with debugging of the script, please send the HealthChecker-Debug_*.txt, HealthChecker-Errors.json, and .xml file to ExToolsFeedback@microsoft.com.")
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
            #Need to convert Error to Json because running into odd issues with trying to export $Error out in my lab. Got StackOverflowException for one of the errors i always see there.
            try {
                $Error |
                    ConvertTo-Json |
                    Out-File ("$OutputFilePath\HealthChecker-Errors.json")
            } catch {
                Write-Red("Failed to export the HealthChecker-Errors.json")
                Invoke-CatchActions
            }
        } elseif ($Script:VerboseEnabled -or
            $SaveDebugLog) {
            Write-Verbose "All errors that occurred were in try catch blocks and was handled correctly."
            $Script:Logger.PreventLogCleanup = $true
            Write-Errors
        }
    } else {
        Write-Verbose "No errors occurred in the script."
    }
}

function Get-HealthCheckFilesItemsFromLocation {
    $items = Get-ChildItem $XMLDirectoryPath | Where-Object { $_.Name -like "HealthChecker-*-*.xml" }

    if ($null -eq $items) {
        Write-Host("Doesn't appear to be any Health Check XML files here....stopping the script")
        exit
    }
    return $items
}

function Get-OnlyRecentUniqueServersXMLs {
    param(
        [Parameter(Mandatory = $true)][array]$FileItems
    )
    $aObject = @()

    foreach ($item in $FileItems) {
        $obj = New-Object PSCustomObject
        [string]$itemName = $item.Name
        $ServerName = $itemName.Substring(($itemName.IndexOf("-") + 1), ($itemName.LastIndexOf("-") - $itemName.IndexOf("-") - 1))
        $obj | Add-Member -MemberType NoteProperty -Name ServerName -Value $ServerName
        $obj | Add-Member -MemberType NoteProperty -Name FileName -Value $itemName
        $obj | Add-Member -MemberType NoteProperty -Name FileObject -Value $item
        $aObject += $obj
    }

    $grouped = $aObject | Group-Object ServerName
    $FilePathList = @()

    foreach ($gServer in $grouped) {

        if ($gServer.Count -gt 1) {
            #going to only use the most current file for this server providing that they are using the newest updated version of Health Check we only need to sort by name
            $groupData = $gServer.Group #because of win2008
            $FilePathList += ($groupData | Sort-Object FileName -Descending | Select-Object -First 1).FileObject.VersionInfo.FileName
        } else {
            $FilePathList += ($gServer.Group).FileObject.VersionInfo.FileName
        }
    }
    return $FilePathList
}

function Import-MyData {
    param(
        [Parameter(Mandatory = $true)][array]$FilePaths
    )
    [System.Collections.Generic.List[System.Object]]$myData = New-Object -TypeName System.Collections.Generic.List[System.Object]

    foreach ($filePath in $FilePaths) {
        $importData = Import-Clixml -Path $filePath
        $myData.Add($importData)
    }
    return $myData
}



# Confirm that either Remote Shell or EMS is loaded from an Edge Server, Exchange Server, or a Tools box.
# It does this by also initializing the session and running Get-EventLogLevel. (Server Management RBAC right)
# All script that require Confirm-ExchangeShell should be at least using Server Management RBAC right for the user running the script.
function Confirm-ExchangeShell {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $false)]
        [bool]$LoadExchangeShell = $true,

        [Parameter(Mandatory = $false)]
        [scriptblock]$CatchActionFunction
    )

    begin {
        Write-Verbose "Calling: $($MyInvocation.MyCommand)"
        Write-Verbose "Passed: LoadExchangeShell: $LoadExchangeShell"
        $currentErrors = $Error.Count
        $passed = $false
        $edgeTransportKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\EdgeTransportRole'
        $setupKey = 'HKLM:\SOFTWARE\Microsoft\ExchangeServer\v15\Setup'
        $remoteShell = (!(Test-Path $setupKey))
        $toolsServer = (Test-Path $setupKey) -and (!(Test-Path $edgeTransportKey)) -and `
        ($null -eq (Get-ItemProperty -Path $setupKey -Name "Services" -ErrorAction SilentlyContinue))
        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
    }
    process {
        try {
            $currentErrors = $Error.Count
            $isEMS = Get-EventLogLevel -ErrorAction Stop |
                Select-Object -First 1 |
                ForEach-Object { if ($_.GetType().Name -eq "EventCategoryObject") { return $true } return $false }
            Write-Verbose "Exchange PowerShell Module already loaded."
            $passed = $true
            Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
        } catch {
            Write-Verbose "Failed to run Get-EventLogLevel"
            Invoke-CatchActionError $CatchActionFunction
            if (-not ($LoadExchangeShell)) { return }

            #Test 32 bit process, as we can't see the registry if that is the case.
            if (-not ([System.Environment]::Is64BitProcess)) {
                Write-Warning "Open a 64 bit PowerShell process to continue"
                return
            }

            if (Test-Path "$setupKey") {
                $currentErrors = $Error.Count
                Write-Verbose "We are on Exchange 2013 or newer"

                try {
                    if (Test-Path $edgeTransportKey) {
                        Write-Verbose "We are on Exchange Edge Transport Server"
                        [xml]$PSSnapIns = Get-Content -Path "$env:ExchangeInstallPath\Bin\exshell.psc1" -ErrorAction Stop

                        foreach ($PSSnapIn in $PSSnapIns.PSConsoleFile.PSSnapIns.PSSnapIn) {
                            Write-Verbose ("Trying to add PSSnapIn: {0}" -f $PSSnapIn.Name)
                            Add-PSSnapin -Name $PSSnapIn.Name -ErrorAction Stop
                        }

                        Import-Module $env:ExchangeInstallPath\bin\Exchange.ps1 -ErrorAction Stop
                    } else {
                        Import-Module $env:ExchangeInstallPath\bin\RemoteExchange.ps1 -ErrorAction Stop
                        Connect-ExchangeServer -Auto -ClientApplication:ManagementShell
                    }

                    Write-Verbose "Imported Module. Trying Get-EventLogLevel Again"
                    try {
                        $isEMS = Get-EventLogLevel -ErrorAction Stop |
                            Select-Object -First 1 |
                            ForEach-Object { if ($_.GetType().Name -eq "EventCategoryObject") { return $true } return $false }
                        $passed = $true
                        Write-Verbose "Successfully loaded Exchange Management Shell"
                        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction
                    } catch {
                        Write-Verbose "Failed to run Get-EventLogLevel again"
                        Invoke-CatchActionError $CatchActionFunction
                    }
                } catch {
                    Write-Warning "Failed to Load Exchange PowerShell Module..."
                    Invoke-CatchActionError $CatchActionFunction
                }
            } else {
                Write-Verbose "Not on an Exchange or Tools server"
            }
        }
    }
    end {

        $currentErrors = $Error.Count
        $returnObject = [PSCustomObject]@{
            ShellLoaded = $passed
            Major       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMajor" -ErrorAction SilentlyContinue).MsiProductMajor)
            Minor       = ((Get-ItemProperty -Path $setupKey -Name "MsiProductMinor" -ErrorAction SilentlyContinue).MsiProductMinor)
            Build       = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMajor" -ErrorAction SilentlyContinue).MsiBuildMajor)
            Revision    = ((Get-ItemProperty -Path $setupKey -Name "MsiBuildMinor" -ErrorAction SilentlyContinue).MsiBuildMinor)
            EdgeServer  = $passed -and (Test-Path $setupKey) -and (Test-Path $edgeTransportKey)
            ToolsOnly   = $passed -and $toolsServer
            RemoteShell = $passed -and $remoteShell
            EMS         = $isEMS
        }

        Invoke-CatchActionErrorLoop $currentErrors $CatchActionFunction

        return $returnObject
    }
}
function Invoke-ScriptLogFileLocation {
    param(
        [Parameter(Mandatory = $true)][string]$FileName,
        [Parameter(Mandatory = $false)][bool]$IncludeServerName = $false
    )
    $endName = "-{0}.txt" -f $dateTimeStringFormat

    if ($IncludeServerName) {
        $endName = "-{0}{1}" -f $Script:Server, $endName
    }

    $Script:OutputFullPath = "{0}\{1}{2}" -f $OutputFilePath, $FileName, $endName
    $Script:OutXmlFullPath = $Script:OutputFullPath.Replace(".txt", ".xml")

    if ($AnalyzeDataOnly -or
        $BuildHtmlServersReport -or
        $ScriptUpdateOnly) {
        return
    }

    $Script:ExchangeShellComputer = Confirm-ExchangeShell `
        -CatchActionFunction ${Function:Invoke-CatchActions}

    if (!($Script:ExchangeShellComputer.ShellLoaded)) {
        Write-Yellow("Failed to load Exchange Shell... stopping script")
        $Script:Logger.PreventLogCleanup = $true
        exit
    }

    if ($Script:ExchangeShellComputer.ToolsOnly -and
        $env:COMPUTERNAME -eq $Script:Server -and
        !($LoadBalancingReport)) {
        Write-Yellow("Can't run Exchange Health Checker Against a Tools Server. Use the -Server Parameter and provide the server you want to run the script against.")
        $Script:Logger.PreventLogCleanup = $true
        exit
    }

    Write-Verbose("Script Executing on Server $env:COMPUTERNAME")
    Write-Verbose("ToolsOnly: $($Script:ExchangeShellComputer.ToolsOnly) | RemoteShell $($Script:ExchangeShellComputer.RemoteShell)")
}

function Test-RequiresServerFqdn {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $tempServerName = ($Script:Server).Split(".")

    if ($tempServerName[0] -eq $env:COMPUTERNAME) {
        Write-Verbose "Executed against the local machine. No need to pass '-ComputerName' parameter."
        return
    } else {
        try {
            $Script:ServerFQDN = (Get-ExchangeServer $Script:Server -ErrorAction Stop).FQDN
        } catch {
            Invoke-CatchActions
            Write-Verbose "Unable to query Fqdn via 'Get-ExchangeServer'"
        }
    }

    try {
        Invoke-Command -ComputerName $Script:Server -ScriptBlock { Get-Date | Out-Null } -ErrorAction Stop
        Write-Verbose "Connected successfully using: $($Script:Server)."
    } catch {
        Invoke-CatchActions
        if ($tempServerName.Count -gt 1) {
            $Script:Server = $tempServerName[0]
        } else {
            $Script:Server = $Script:ServerFQDN
        }

        try {
            Invoke-Command -ComputerName $Script:Server -ScriptBlock { Get-Date | Out-Null } -ErrorAction Stop
            Write-Verbose "Fallback to: $($Script:Server) Connection was successfully established."
        } catch {
            Write-Red("Failed to run against: {0}. Please try to run the script locally on: {0} for results. " -f $Script:Server)
            exit
        }
    }
}

$healthCheckerCustomClass = @"
using System;
using System.Collections;
    namespace HealthChecker
    {
        public class HealthCheckerExchangeServer
        {
            public string ServerName;        //String of the server that we are working with
            public HardwareInformation HardwareInformation;  // Hardware Object Information
            public OperatingSystemInformation  OSInformation; // OS Version Object Information
            public ExchangeInformation ExchangeInformation; //Detailed Exchange Information
            public string HealthCheckerVersion; //To determine the version of the script on the object.
            public DateTime GenerationTime; //Time stamp of running the script
        }

        // ExchangeInformation
        public class ExchangeInformation
        {
            public ExchangeBuildInformation BuildInformation = new ExchangeBuildInformation();   //Exchange build information
            public object GetExchangeServer;      //Stores the Get-ExchangeServer Object
            public object GetMailboxServer;       //Stores the Get-MailboxServer Object
            public object GetOwaVirtualDirectory; //Stores the Get-OwaVirtualDirectory Object
            public object GetWebServicesVirtualDirectory; //stores the Get-WebServicesVirtualDirectory object
            public object GetOrganizationConfig; //Stores the result from Get-OrganizationConfig
            public object ExchangeAdPermissions; //Stores the Exchange AD permissions for vulnerability testing
            public object ExtendedProtectionConfig; //Stores the extended protection configuration
            public object msExchStorageGroup;   //Stores the properties of the 'ms-Exch-Storage-Group' Schema class
            public object GetHybridConfiguration; //Stores the Get-HybridConfiguration Object
            public object ExchangeConnectors; //Stores the Get-ExchangeConnectors Object
            public bool EnableDownloadDomains = new bool(); //True if Download Domains are enabled on org level
            public object WildCardAcceptedDomain; // for issues with * accepted domain.
            public System.Array AMSIConfiguration; //Stores the Setting Override for AMSI Interface
            public ExchangeNetFrameworkInformation NETFramework = new ExchangeNetFrameworkInformation();
            public bool MapiHttpEnabled; //Stored from organization config
            public System.Array ExchangeServicesNotRunning; //Contains the Exchange services not running by Test-ServiceHealth
            public Hashtable ApplicationPools = new Hashtable();
            public object RegistryValues; //stores all Exchange Registry values
            public ExchangeServerMaintenance ServerMaintenance;
            public System.Array ExchangeCertificates;           //stores all the Exchange certificates on the servers.
            public object ExchangeEmergencyMitigationService;   //stores the Exchange Emergency Mitigation Service (EEMS) object
            public Hashtable ApplicationConfigFileStatus = new Hashtable();
            public object DependentServices; // store the results for the dependent services of Exchange.
            public object IISSettings;  //Stores the IISConfigurationSettings, applicationHostConfig and IISModulesInformation
            public object SettingOverrides; //Stores the information regarding the Exchange Setting Overrides on the server.
        }

        public class ExchangeBuildInformation
        {
            public ExchangeServerRole ServerRole; //Roles that are currently set and installed.
            public ExchangeMajorVersion MajorVersion; //Exchange Version (Exchange 2010/2013/2019)
            public ExchangeCULevel CU;             // Exchange CU Level
            public string FriendlyName;     //Exchange Friendly Name is provided
            public string BuildNumber;      //Exchange Build Number
            public string LocalBuildNumber; //Local Build Number. Is only populated if from a Tools Machine
            public string ReleaseDate;      // Exchange release date for which the CU they are currently on
            public string ExtendedSupportDate; // End of Life Support Date.
            public bool SupportedBuild;     //Determines if we are within the correct build of Exchange.
            public object ExchangeSetup;    //Stores the Get-Command ExSetup object
            public System.Array KBsInstalled;  //Stored object IU or Security KB fixes
            public bool March2021SUInstalled;    //True if March 2021 SU is installed
            public object FIPFSUpdateIssue; //Stores FIP-FS update issue information
        }

        public class ExchangeNetFrameworkInformation
        {
            public NetMajorVersion MinSupportedVersion; //Min Supported .NET Framework version
            public NetMajorVersion MaxSupportedVersion; //Max (Recommended) Supported .NET version.
            public bool OnRecommendedVersion; //RecommendedNetVersion Info includes all the factors. Windows Version & CU.
            public string DisplayWording; //Display if we are in Support or not
        }

        public class ExchangeServerMaintenance
        {
            public System.Array InactiveComponents;
            public object GetServerComponentState;
            public object GetClusterNode;
            public object GetMailboxServer; //TODO: Remove this
        }

        //enum for CU levels of Exchange
        public enum ExchangeCULevel
        {
            Unknown,
            Preview,
            RTM,
            CU1,
            CU2,
            CU3,
            CU4,
            CU5,
            CU6,
            CU7,
            CU8,
            CU9,
            CU10,
            CU11,
            CU12,
            CU13,
            CU14,
            CU15,
            CU16,
            CU17,
            CU18,
            CU19,
            CU20,
            CU21,
            CU22,
            CU23
        }

        //enum for the server roles that the computer is
        public enum ExchangeServerRole
        {
            MultiRole,
            Mailbox,
            ClientAccess,
            Hub,
            Edge,
            None
        }

        //enum for the Exchange version
        public enum ExchangeMajorVersion
        {
            Unknown,
            Exchange2010,
            Exchange2013,
            Exchange2016,
            Exchange2019
        }
        // End ExchangeInformation

        // OperatingSystemInformation
        public class OperatingSystemInformation
        {
            public OSBuildInformation BuildInformation = new OSBuildInformation(); // contains build information
            public NetworkInformation NetworkInformation = new NetworkInformation(); //stores network information and settings
            public PowerPlanInformation PowerPlan = new PowerPlanInformation(); //stores the power plan information
            public object PageFile;             //stores the page file information
            public LmCompatibilityLevelInformation LmCompatibility; // stores Lm Compatibility Level Information
            public object ServerPendingReboot; // determine if server is pending a reboot.
            public TimeZoneInformation TimeZone = new TimeZoneInformation();    //stores time zone information
            public object TLSSettings;            // stores the TLS settings on the server.
            public InstalledUpdatesInformation InstalledUpdates = new InstalledUpdatesInformation();  //store the install update
            public ServerBootUpInformation ServerBootUp = new ServerBootUpInformation();   // stores the server boot up time information
            public System.Array VcRedistributable;            //stores the Visual C++ Redistributable
            public OSNetFrameworkInformation NETFramework = new OSNetFrameworkInformation();          //stores OS Net Framework
            public bool CredentialGuardEnabled;
            public OSRegistryValues RegistryValues = new OSRegistryValues();
            public object Smb1ServerSettings;
        }

        public class OSBuildInformation
        {
            public OSServerVersion MajorVersion; //OS Major Version
            public string VersionBuild;           //hold the build number
            public string FriendlyName;           //string holder of the Windows Server friendly name
            public object OperatingSystem;        // holds Win32_OperatingSystem
        }

        public class NetworkInformation
        {
            public double TCPKeepAlive;           // value used for the TCP/IP keep alive value in the registry
            public double RpcMinConnectionTimeout;  //holds the value for the RPC minimum connection timeout.
            public object HttpProxy;                // holds the setting for HttpProxy if one is set.
            public object PacketsReceivedDiscarded;   //hold all the packets received discarded on the server.
            public double IPv6DisabledComponents;    //value stored in the registry HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\Tcpip6\Parameters\DisabledComponents
            public bool IPv6DisabledOnNICs;          //value that determines if we have IPv6 disabled on some NICs or not.
            public System.Array NetworkAdapters;           //stores all the NICs on the servers.
            public string PnPCapabilities;      //Value from PnPCapabilities registry
            public bool SleepyNicDisabled;     //If the NIC can be in power saver mode by the OS.
        }

        public class PowerPlanInformation
        {
            public bool HighPerformanceSet;      // If the power plan is High Performance
            public string PowerPlanSetting;      //value for the power plan that is set
            public object PowerPlan;            //object to store the power plan information
        }

        public class OSRegistryValues
        {
            public int CurrentVersionUbr; // stores SOFTWARE\Microsoft\Windows NT\CurrentVersion\UBR
            public int LanManServerDisabledCompression; // stores SYSTEM\CurrentControlSet\Services\LanmanServer\Parameters\DisabledCompression
        }

        public class LmCompatibilityLevelInformation
        {
            public int RegistryValue;       //The LmCompatibilityLevel for the server (INT 1 - 5)
            public string Description;      //description of the LmCompat that the server is set to
        }

        public class TimeZoneInformation
        {
            public string CurrentTimeZone; //stores the value for the current time zone of the server.
            public int DynamicDaylightTimeDisabled; // the registry value for DynamicDaylightTimeDisabled.
            public string TimeZoneKeyName; // the registry value TimeZoneKeyName.
            public string StandardStart;   // the registry value for StandardStart.
            public string DaylightStart;   // the registry value for DaylightStart.
            public bool DstIssueDetected;  // Determines if there is a high chance of an issue.
            public System.Array ActionsToTake; //array of verbage of the issues detected.
        }

        public class ServerRebootInformation
        {
            public bool PendingFileRenameOperations;            //bool "HKLM:\SYSTEM\CurrentControlSet\Control\Session Manager\" item PendingFileRenameOperations.
            public object SccmReboot;                           // object to store CimMethod for class name CCM_ClientUtilities
            public bool SccmRebootPending;                      // SccmReboot has either PendingReboot or IsHardRebootPending is set to true.
            public bool ComponentBasedServicingPendingReboot;   // bool HKLM:\Software\Microsoft\Windows\CurrentVersion\Component Based Servicing\RebootPending
            public bool AutoUpdatePendingReboot;                // bool HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\WindowsUpdate\Auto Update\RebootRequired
            public bool PendingReboot;                         // bool if reboot types are set to true
            public bool UpdateExeVolatile;                      // bool HKLM:\Software\Microsoft\Updates\UpdateExeVolatile\Flags
        }

        public class InstalledUpdatesInformation
        {
            public System.Array HotFixes;     //array to keep all the hotfixes of the server
            public System.Array HotFixInfo;   //object to store hotfix information
            public System.Array InstalledUpdates; //store the install updates
        }

        public class ServerBootUpInformation
        {
            public string Days;
            public string Hours;
            public string Minutes;
            public string Seconds;
        }

        public class OSNetFrameworkInformation
        {
            public NetMajorVersion NetMajorVersion; //NetMajorVersion value
            public string FriendlyName;  //string of the friendly name
            public int RegistryValue; //store the registry value
            public Hashtable FileInformation; //stores Get-Item information for .NET Framework
        }

        //enum for the OSServerVersion that we are
        public enum OSServerVersion
        {
            Unknown,
            Windows2008,
            Windows2008R2,
            Windows2012,
            Windows2012R2,
            Windows2016,
            Windows2019,
            Windows2022,
            WindowsCore
        }

        //enum for the dword value of the .NET frame 4 that we are on
        public enum NetMajorVersion
        {
            Unknown = 0,
            Net4d5 = 378389,
            Net4d5d1 = 378675,
            Net4d5d2 = 379893,
            Net4d5d2wFix = 380035,
            Net4d6 = 393295,
            Net4d6d1 = 394254,
            Net4d6d1wFix = 394294,
            Net4d6d2 = 394802,
            Net4d7 = 460798,
            Net4d7d1 = 461308,
            Net4d7d2 = 461808,
            Net4d8 = 528040
        }
        // End OperatingSystemInformation

        // HardwareInformation
        public class HardwareInformation
        {
            public string Manufacturer; //String to display the hardware information
            public ServerType ServerType; //Enum to determine if the hardware is VMware, HyperV, Physical, or Unknown
            public System.Array MemoryInformation; //Detailed information about the installed memory
            public UInt64 TotalMemory; //Stores the total memory cooked value
            public object System;   //object to store the system information that we have collected
            public ProcessorInformation Processor;   //Detailed processor Information
            public bool AutoPageFile; //True/False if we are using a page file that is being automatically set
            public string Model; //string to display Model
        }

        //enum for the type of computer that we are
        public enum ServerType
        {
            VMWare,
            AmazonEC2,
            HyperV,
            Physical,
            Unknown
        }

        public class ProcessorInformation
        {
            public string Name;    //String of the processor name
            public int NumberOfPhysicalCores;    //Number of Physical cores that we have
            public int NumberOfLogicalCores;  //Number of Logical cores that we have presented to the os
            public int NumberOfProcessors; //Total number of processors that we have in the system
            public int MaxMegacyclesPerCore; //Max speed that we can get out of the cores
            public int CurrentMegacyclesPerCore; //Current speed that we are using the cores at
            public bool ProcessorIsThrottled;  //True/False if we are throttling our processor
            public bool DifferentProcessorsDetected; //true/false to detect if we have different processor types detected
            public bool DifferentProcessorCoreCountDetected; //detect if there are a different number of core counts per Processor CPU socket
            public int EnvironmentProcessorCount; //[system.environment]::processorcount
            public object ProcessorClassObject;        // object to store the processor information
        }

        //HTML & display classes
        public class HtmlServerValues
        {
            public System.Array OverviewValues;
            public System.Array ActionItems;   //use HtmlServerActionItemRow
            public System.Array ServerDetails;    // use HtmlServerInformationRow
        }

        public class HtmlServerActionItemRow
        {
            public string Setting;
            public string DetailValue;
            public string RecommendedDetails;
            public string MoreInformation;
            public string Class;
        }

        public class HtmlServerInformationRow
        {
            public string Name;
            public string DetailValue;
            public object TableValue;
            public string Class;
        }

        public class DisplayResultsLineInfo
        {
            public string DisplayValue;
            public string Name;
            public string TestingName; // Used for pestering testing
            public int TabNumber;
            public object TestingValue; //Used for pester testing down the road.
            public object OutColumns; //used for colorized format table option.
            public string WriteType;

            public string Line
            {
                get
                {
                    if (String.IsNullOrEmpty(this.Name))
                    {
                        return this.DisplayValue;
                    }

                    return String.Concat(this.Name, ": ", this.DisplayValue);
                }
            }
        }

        public class DisplayResultsGroupingKey
        {
            public string Name;
            public int DefaultTabNumber;
            public bool DisplayGroupName;
            public int DisplayOrder;
        }

        public class AnalyzedInformation
        {
            public HealthCheckerExchangeServer HealthCheckerExchangeServer;
            public Hashtable HtmlServerValues = new Hashtable();
            public Hashtable DisplayResults = new Hashtable();
        }
    }
"@

try {
    #Enums and custom data types
    if (-not($ScriptUpdateOnly)) {
        Add-Type -TypeDefinition $healthCheckerCustomClass -ErrorAction Stop
    }
} catch {
    Write-Warning "There was an error trying to add custom classes to the current PowerShell session. You need to close this session and open a new one to have the script properly work."
    exit
}

function Write-ResultsToScreen {
    param(
        [Hashtable]$ResultsToWrite
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $indexOrderGroupingToKey = @{}

    foreach ($keyGrouping in $ResultsToWrite.Keys) {
        $indexOrderGroupingToKey[$keyGrouping.DisplayOrder] = $keyGrouping
    }

    $sortedIndexOrderGroupingToKey = $indexOrderGroupingToKey.Keys | Sort-Object

    foreach ($key in $sortedIndexOrderGroupingToKey) {
        Write-Verbose "Working on Key: $key"
        $keyGrouping = $indexOrderGroupingToKey[$key]
        Write-Verbose "Working on Key Group: $($keyGrouping.Name)"
        Write-Verbose "Total lines to write: $($ResultsToWrite[$keyGrouping].Count)"

        if ($keyGrouping.DisplayGroupName) {
            Write-Grey($keyGrouping.Name)
            $dashes = [string]::empty
            1..($keyGrouping.Name.Length) | ForEach-Object { $dashes = $dashes + "-" }
            Write-Grey($dashes)
        }

        foreach ($line in $ResultsToWrite[$keyGrouping]) {
            $tab = [string]::Empty

            if ($line.TabNumber -ne 0) {
                1..($line.TabNumber) | ForEach-Object { $tab = $tab + "`t" }
            }

            $writeValue = "{0}{1}" -f $tab, $line.Line
            switch ($line.WriteType) {
                "Grey" { Write-Grey($writeValue) }
                "Yellow" { Write-Yellow($writeValue) }
                "Green" { Write-Green($writeValue) }
                "Red" { Write-Red($writeValue) }
                "OutColumns" { Write-OutColumns($line.OutColumns) }
            }
        }

        Write-Grey("")
    }
}


<#
.SYNOPSIS
    Outputs a table of objects with certain values colorized.
.EXAMPLE
    PS C:\> <example usage>
    Explanation of what the example does
.INPUTS
    Inputs (if any)
.OUTPUTS
    Output (if any)
.NOTES
    General notes
#>
function Out-Columns {
    [CmdletBinding()]
    param (
        [Parameter(ValueFromPipeline = $true)]
        [object[]]
        $InputObject,

        [Parameter(Mandatory = $false, Position = 0)]
        [string[]]
        $Properties,

        [Parameter(Mandatory = $false, Position = 1)]
        [scriptblock[]]
        $ColorizerFunctions = @(),

        [Parameter(Mandatory = $false)]
        [int]
        $IndentSpaces = 0,

        [Parameter(Mandatory = $false)]
        [int]
        $LinesBetweenObjects = 0,

        [Parameter(Mandatory = $false)]
        [ref]
        $StringOutput
    )

    begin {
        function WrapLine {
            param([string]$line, [int]$width)
            if ($line.Length -le $width -and $line.IndexOf("`n") -lt 0) {
                return $line
            }

            $lines = New-Object System.Collections.ArrayList

            $noLF = $line.Replace("`r", "")
            $lineSplit = $noLF.Split("`n")
            foreach ($l in $lineSplit) {
                if ($l.Length -le $width) {
                    [void]$lines.Add($l)
                } else {
                    $split = $l.Split(" ")
                    $sb = New-Object System.Text.StringBuilder
                    for ($i = 0; $i -lt $split.Length; $i++) {
                        if ($sb.Length -eq 0 -and $sb.Length + $split[$i].Length -lt $width) {
                            [void]$sb.Append($split[$i])
                        } elseif ($sb.Length -gt 0 -and $sb.Length + $split[$i].Length + 1 -lt $width) {
                            [void]$sb.Append(" " + $split[$i])
                        } elseif ($sb.Length -gt 0) {
                            [void]$lines.Add($sb.ToString())
                            [void]$sb.Clear()
                            $i--
                        } else {
                            if ($split[$i].Length -le $width) {
                                [void]$lines.Add($split[$i])
                            } else {
                                [void]$lines.Add($split[$i].Substring(0, $width))
                                $split[$i] = $split[$i].Substring($width)
                                $i--
                            }
                        }
                    }

                    if ($sb.Length -gt 0) {
                        [void]$lines.Add($sb.ToString())
                    }
                }
            }

            return $lines
        }

        function GetLineObjects {
            param($obj, $props, $colWidths)
            $linesNeededForThisObject = 1
            $multiLineProps = @{}
            for ($i = 0; $i -lt $props.Length; $i++) {
                $p = $props[$i]
                $val = $obj."$p"

                if ($val -isnot [array]) {
                    $val = WrapLine -line $val -width $colWidths[$i]
                } elseif ($val -is [array]) {
                    $val = $val | Where-Object { $null -ne $_ }
                    $val = $val | ForEach-Object { WrapLine -line $_ -width $colWidths[$i] }
                }

                if ($val -is [array]) {
                    $multiLineProps[$p] = $val
                    if ($val.Length -gt $linesNeededForThisObject) {
                        $linesNeededForThisObject = $val.Length
                    }
                }
            }

            if ($linesNeededForThisObject -eq 1) {
                $obj
            } else {
                for ($i = 0; $i -lt $linesNeededForThisObject; $i++) {
                    $lineProps = @{}
                    foreach ($p in $props) {
                        if ($null -ne $multiLineProps[$p] -and $multiLineProps[$p].Length -gt $i) {
                            $lineProps[$p] = $multiLineProps[$p][$i]
                        } elseif ($i -eq 0) {
                            $lineProps[$p] = $obj."$p"
                        } else {
                            $lineProps[$p] = $null
                        }
                    }

                    [PSCustomObject]$lineProps
                }
            }
        }

        function GetColumnColors {
            param($obj, $props, $funcs)

            $consoleHost = (Get-Host).Name -eq "ConsoleHost"
            $colColors = New-Object string[] $props.Count
            for ($i = 0; $i -lt $props.Count; $i++) {
                if ($consoleHost) {
                    $fgColor = (Get-Host).ui.rawui.ForegroundColor
                } else {
                    $fgColor = "White"
                }
                foreach ($func in $funcs) {
                    $result = $func.Invoke($obj, $props[$i])
                    if (-not [string]::IsNullOrEmpty($result)) {
                        $fgColor = $result
                        break # The first colorizer that takes action wins
                    }
                }

                $colColors[$i] = $fgColor
            }

            $colColors
        }

        function GetColumnWidths {
            param($objects, $props)

            $colWidths = New-Object int[] $props.Count

            # Start with the widths of the property names
            for ($i = 0; $i -lt $props.Count; $i++) {
                $colWidths[$i] = $props[$i].Length
            }

            # Now check the widths of the widest values
            foreach ($thing in $objects) {
                for ($i = 0; $i -lt $props.Count; $i++) {
                    $val = $thing."$($props[$i])"
                    if ($null -ne $val) {
                        $width = 0
                        if ($val -isnot [array]) {
                            $val = $val.ToString().Split("`n")
                        }

                        $width = ($val | ForEach-Object {
                                if ($null -ne $_) { $_.ToString() } else { "" }
                            } | Sort-Object Length -Descending | Select-Object -First 1).Length

                        if ($width -gt $colWidths[$i]) {
                            $colWidths[$i] = $width
                        }
                    }
                }
            }

            # If we're within the window width, we're done
            $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
            $windowWidth = (Get-Host).UI.RawUI.WindowSize.Width
            if ($windowWidth -lt 1 -or $totalColumnWidth -lt $windowWidth) {
                return $colWidths
            }

            # Take size away from one or more columns to make them fit
            while ($totalColumnWidth -ge $windowWidth) {
                $startingTotalWidth = $totalColumnWidth
                $widest = $colWidths | Sort-Object -Descending | Select-Object -First 1
                $newWidest = [Math]::Floor($widest * 0.95)
                for ($i = 0; $i -lt $colWidths.Length; $i++) {
                    if ($colWidths[$i] -eq $widest) {
                        $colWidths[$i] = $newWidest
                        break
                    }
                }

                $totalColumnWidth = $colWidths.Length * $padding + ($colWidths | Measure-Object -Sum).Sum + $IndentSpaces
                if ($totalColumnWidth -ge $startingTotalWidth) {
                    # Somehow we didn't reduce the size at all, so give up
                    break
                }
            }

            return $colWidths
        }

        $objects = New-Object System.Collections.ArrayList
        $padding = 2
        $stb = New-Object System.Text.StringBuilder
    }

    process {
        foreach ($thing in $InputObject) {
            [void]$objects.Add($thing)
        }
    }

    end {
        if ($objects.Count -gt 0) {
            $props = $null

            if ($null -ne $Properties) {
                $props = $Properties
            } else {
                $props = $objects[0].PSObject.Properties.Name
            }

            $colWidths = GetColumnWidths $objects $props

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i]) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $props[$i])
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            Write-Host (" " * $IndentSpaces) -NoNewline
            [void]$stb.Append(" " * $IndentSpaces)

            for ($i = 0; $i -lt $props.Count; $i++) {
                Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length)) -NoNewline
                [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f ("-" * $props[$i].Length))
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            foreach ($o in $objects) {
                $colColors = GetColumnColors -obj $o -props $props -funcs $ColorizerFunctions
                $lineObjects = @(GetLineObjects -obj $o -props $props -colWidths $colWidths)
                foreach ($lineObj in $lineObjects) {
                    Write-Host (" " * $IndentSpaces) -NoNewline
                    [void]$stb.Append(" " * $IndentSpaces)
                    for ($i = 0; $i -lt $props.Count; $i++) {
                        $val = $o."$($props[$i])"
                        Write-Host ("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])") -NoNewline -ForegroundColor $colColors[$i]
                        [void]$stb.Append("{0,$(-1 * ($colWidths[$i] + $padding))}" -f $lineObj."$($props[$i])")
                    }

                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }

                for ($i = 0; $i -lt $LinesBetweenObjects; $i++) {
                    Write-Host
                    [void]$stb.Append([System.Environment]::NewLine)
                }
            }

            Write-Host
            [void]$stb.Append([System.Environment]::NewLine)

            if ($null -ne $StringOutput) {
                $StringOutput.Value = $stb.ToString()
            }
        }
    }
}
function Write-Red($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Red
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Yellow($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Yellow
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Green($message) {
    Write-DebugLog $message
    Write-Host $message -ForegroundColor Green
    $message | Out-File ($OutputFullPath) -Append
}

function Write-Grey($message) {
    Write-DebugLog $message
    Write-Host $message
    $message | Out-File ($OutputFullPath) -Append
}

function Write-DebugLog($message) {
    if (![string]::IsNullOrEmpty($message)) {
        $Script:Logger = $Script:Logger | Write-LoggerInstance $message
    }
}

function Write-OutColumns($OutColumns) {
    if ($null -ne $OutColumns) {
        $stringOutput = $null
        $OutColumns.DisplayObject |
            Out-Columns -Properties $OutColumns.SelectProperties `
                -ColorizerFunctions $OutColumns.ColorizerFunctions `
                -IndentSpaces $OutColumns.IndentSpaces `
                -StringOutput ([ref]$stringOutput)
        $stringOutput | Out-File ($OutputFullPath) -Append
        Write-DebugLog $stringOutput
    }
}

function Write-Break {
    Write-Host ""
}

function Get-HtmlServerReport {
    param(
        [Parameter(Mandatory = $true)][array]$AnalyzedHtmlServerValues
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"

    function GetOutColumnHtmlTable {
        param(
            [object]$OutColumn
        )
        # this keeps the order of the columns
        $headerValues = $OutColumn[0].PSObject.Properties.Name
        $htmlTableValue = "<table>"

        foreach ($header in $headerValues) {
            $htmlTableValue += "<th>$header</th>"
        }

        foreach ($dataRow in $OutColumn) {
            $htmlTableValue += "$([System.Environment]::NewLine)<tr>"

            foreach ($header in $headerValues) {
                $htmlTableValue += "<td class=`"$($dataRow.$header.DisplayColor)`">$($dataRow.$header.Value)</td>"
            }
            $htmlTableValue += "$([System.Environment]::NewLine)</tr>"
        }
        $htmlTableValue += "</table>"
        return $htmlTableValue
    }

    $htmlHeader = "<html>
        <style>
        BODY{font-family: Arial; font-size: 8pt;}
        H1{font-size: 16px;}
        H2{font-size: 14px;}
        H3{font-size: 12px;}
        TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
        TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
        TD{border: 1px solid black; padding: 5px; }
        td.Green{background: #7FFF00;}
        td.Yellow{background: #FFE600;}
        td.Red{background: #FF0000; color: #ffffff;}
        td.Info{background: #85D4FF;}
        </style>
        <body>
        <h1 align=""center"">Exchange Health Checker v$($BuildVersion)</h1><br>
        <h2>Servers Overview</h2>"

    [array]$htmlOverviewTable += "<p>
        <table>
        <tr>$([System.Environment]::NewLine)"

    foreach ($tableHeaderName in $AnalyzedHtmlServerValues[0]["OverviewValues"].Name) {
        $htmlOverviewTable += "<th>{0}</th>$([System.Environment]::NewLine)" -f $tableHeaderName
    }

    $htmlOverviewTable += "</tr>$([System.Environment]::NewLine)"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        $htmlTableRow = @()
        [array]$htmlTableRow += "<tr>$([System.Environment]::NewLine)"
        foreach ($htmlTableDataRow in $serverHtmlServerValues["OverviewValues"]) {
            $htmlTableRow += "<td class=`"{0}`">{1}</td>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                $htmlTableDataRow.DetailValue
        }

        $htmlTableRow += "</tr>$([System.Environment]::NewLine)"
        $htmlOverviewTable += $htmlTableRow
    }

    $htmlOverviewTable += "</table>$([System.Environment]::NewLine)</p>$([System.Environment]::NewLine)"

    [array]$htmlServerDetails += "<p>$([System.Environment]::NewLine)<h2>Server Details</h2>$([System.Environment]::NewLine)<table>"

    foreach ($serverHtmlServerValues in $AnalyzedHtmlServerValues) {
        foreach ($htmlTableDataRow in $serverHtmlServerValues["ServerDetails"]) {
            if ($htmlTableDataRow.Name -eq "Server Name") {
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<th>{0}</th>$([System.Environment]::NewLine)<th>{1}</th>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            } elseif ($null -ne $htmlTableDataRow.TableValue) {
                $htmlTable = GetOutColumnHtmlTable $htmlTableDataRow.TableValue
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<td class=`"{0}`">{1}</td><td class=`"{0}`">{2}</td>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                    $htmlTableDataRow.Name, `
                    $htmlTable
            } else {
                $htmlServerDetails += "<tr>$([System.Environment]::NewLine)<td class=`"{0}`">{1}</td><td class=`"{0}`">{2}</td>$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)" -f $htmlTableDataRow.Class, `
                    $htmlTableDataRow.Name, `
                    $htmlTableDataRow.DetailValue
            }
        }
    }
    $htmlServerDetails += "$([System.Environment]::NewLine)</table>$([System.Environment]::NewLine)</p>$([System.Environment]::NewLine)"

    $htmlReport = $htmlHeader + $htmlOverviewTable + $htmlServerDetails + "</body>$([System.Environment]::NewLine)</html>"

    $htmlReport | Out-File $HtmlReportFile -Encoding UTF8
}

function Get-CASLoadBalancingReport {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $CASServers = @()

    if ($null -ne $CasServerList) {
        Write-Grey("Custom CAS server list is being used.  Only servers specified after the -CasServerList parameter will be used in the report.")
        $CASServers = Get-ExchangeServer | Where-Object { ($_.Name -in $CasServerList) -or ($_.FQDN -in $CasServerList) } | Sort-Object Name
    } elseif ($SiteName -ne [string]::Empty) {
        Write-Grey("Site filtering ON.  Only Exchange 2013/2016 CAS servers in {0} will be used in the report." -f $SiteName)
        $CASServers = Get-ExchangeServer | Where-Object {
            ($_.IsClientAccessServer -eq $true) -and
            ($_.AdminDisplayVersion -Match "^Version 15") -and
            ([System.Convert]::ToString($_.Site).Split("/")[-1] -eq $SiteName) } | Sort-Object Name
    } else {
        Write-Grey("Site filtering OFF.  All Exchange 2013/2016 CAS servers will be used in the report.")
        $CASServers = Get-ExchangeServer | Where-Object { ($_.IsClientAccessServer -eq $true) -and ($_.AdminDisplayVersion -Match "^Version 15") } | Sort-Object Name
    }

    if ($CASServers.Count -eq 0) {
        Write-Red("Error: No CAS servers found using the specified search criteria.")
        exit
    }

    function DisplayKeyMatching {
        param(
            [string]$CounterValue,
            [string]$DisplayValue
        )
        return [PSCustomObject]@{
            Counter = $CounterValue
            Display = $DisplayValue
        }
    }

    #Request stats from perfmon for all CAS
    $displayKeys = @{
        1 = DisplayKeyMatching "Default Web Site" "Load Distribution"
        2 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_Autodiscover" "AutoDiscover"
        3 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_EWS" "EWS"
        4 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_mapi" "MapiHttp"
        5 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync" "EAS"
        6 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_owa" "OWA"
        7 = DisplayKeyMatching "_LM_W3SVC_1_ROOT_Rpc" "RpcHttp"
    }
    $perServerStats = @{}
    $totalStats = @{}

    $currentErrors = $Error.Count
    $counterSamples = Get-LocalizedCounterSamples -MachineName $CASServers -Counter @(
        "\Web Service(*)\Current Connections",
        "\ASP.NET Apps v4.0.30319(*)\Requests Executing"
    ) `
        -CustomErrorAction "SilentlyContinue"

    Invoke-CatchActionErrorLoop $currentErrors ${Function:Invoke-CatchActions}

    foreach ($counterSample in $counterSamples) {
        $counterObject = Get-CounterFullNameToCounterObject -FullCounterName $counterSample.Path

        if (-not ($perServerStats.ContainsKey($counterObject.ServerName))) {
            $perServerStats.Add($counterObject.ServerName, @{})
        }

        if (-not ($perServerStats[$counterObject.ServerName].ContainsKey($counterObject.InstanceName))) {
            $perServerStats[$counterObject.ServerName].Add($counterObject.InstanceName, $counterSample.CookedValue)
        } else {
            Write-Verbose "This shouldn't occur...."
            $perServerStats[$counterObject.ServerName][$counterObject.InstanceName] += $counterSample.CookedValue
        }

        if (-not ($totalStats.ContainsKey($counterObject.InstanceName))) {
            $totalStats.Add($counterObject.InstanceName, 0)
        }

        $totalStats[$counterObject.InstanceName] += $counterSample.CookedValue
    }

    $keyOrders = $displayKeys.Keys | Sort-Object

    $htmlHeader = "<html>
    <style>
    BODY{font-family: Arial; font-size: 8pt;}
    H1{font-size: 16px;}
    H2{font-size: 14px;}
    H3{font-size: 12px;}
    TABLE{border: 1px solid black; border-collapse: collapse; font-size: 8pt;}
    TH{border: 1px solid black; background: #dddddd; padding: 5px; color: #000000;}
    TD{border: 1px solid black; padding: 5px; }
    td.Green{background: #7FFF00;}
    td.Yellow{background: #FFE600;}
    td.Red{background: #FF0000; color: #ffffff;}
    td.Info{background: #85D4FF;}
    </style>
    <body>
    <h1 align=""center"">Exchange Health Checker v$($BuildVersion)</h1>
    <h1 align=""center"">Domain : $(($(Get-ADDomain).DNSRoot).toUpper())</h1>
    <h2 align=""center"">Load balancer run finished : $((Get-Date).ToString("yyyy-MM-dd HH:mm"))</h2><br>"

    [array]$htmlLoadDetails += "<table>
    <tr><th>Server</th>
    <th>Site</th>
    "
    #Load the key Headers
    $keyOrders | ForEach-Object {
        $htmlLoadDetails += "$([System.Environment]::NewLine)<th><center>$($displayKeys[$_].Display) Requests</center></th>
        <th><center>$($displayKeys[$_].Display) %</center></th>"
    }
    $htmlLoadDetails += "$([System.Environment]::NewLine)</tr>$([System.Environment]::NewLine)"

    foreach ($server in $CASServers) {
        $serverKey = $server.Name.ToString()
        Write-Verbose "Working Server for HTML report $serverKey"
        $htmlLoadDetails += "<tr>
        <td>$($serverKey)</td>
        <td><center>$($server.Site)</center></td>"

        foreach ($key in $keyOrders) {
            $currentDisplayKey = $displayKeys[$key]
            $totalRequests = $totalStats[$currentDisplayKey.Counter]

            if ($perServerStats.ContainsKey($serverKey)) {
                $serverValue = $perServerStats[$serverKey][$currentDisplayKey.Counter]
                if ($null -eq $serverValue) { $serverValue = 0 }
            } else {
                $serverValue = 0
            }
            $percentageLoad = [math]::Round((($serverValue / $totalRequests) * 100))
            Write-Verbose "$($currentDisplayKey.Display) Server Value $serverValue Percentage usage $percentageLoad"

            $htmlLoadDetails += "$([System.Environment]::NewLine)<td><center>$($serverValue)</center></td>
            <td><center>$percentageLoad</center></td>"
        }
        $htmlLoadDetails += "$([System.Environment]::NewLine)</tr>"
    }

    # Totals
    $htmlLoadDetails += "$([System.Environment]::NewLine)<tr>
        <td><center>Totals</center></td>
        <td></td>"
    $keyOrders | ForEach-Object {
        $htmlLoadDetails += "$([System.Environment]::NewLine)<td><center>$($totalStats[(($displayKeys[$_]).Counter)])</center></td>
        <td></td>"
    }

    $htmlLoadDetails += "$([System.Environment]::NewLine)</table></p>"
    $htmlReport = $htmlHeader + $htmlLoadDetails + "</body></html>"
    $htmlFile = "$Script:OutputFilePath\HtmlLoadBalancerReport.html"
    $htmlReport | Out-File $htmlFile

    foreach ($key in $keyOrders) {
        $currentDisplayKey = $displayKeys[$key]
        $totalRequests = $totalStats[$currentDisplayKey.Counter]

        if ($totalRequests -le 0) { continue }

        Write-Grey ""
        Write-Grey "Current $($currentDisplayKey.Display) Per Server"
        Write-Grey "Total Requests: $totalRequests"

        foreach ($serverKey in $perServerStats.Keys) {
            if ($perServerStats.ContainsKey($serverKey)) {
                $serverValue = $perServerStats[$serverKey][$currentDisplayKey.Counter]
                Write-Grey "$serverKey : $serverValue Connections = $([math]::Round((([int]$serverValue / $totalRequests) * 100)))% Distribution"
            }
        }
    }

    Write-Grey "HTML File Report Written to $htmlFile"
}

function Get-ComputerCoresObject {
    param(
        [Parameter(Mandatory = $true)][string]$Machine_Name
    )
    Write-Verbose "Calling: $($MyInvocation.MyCommand) Passed: $Machine_Name"

    $returnObj = New-Object PSCustomObject
    $returnObj | Add-Member -MemberType NoteProperty -Name Error -Value $false
    $returnObj | Add-Member -MemberType NoteProperty -Name ComputerName -Value $Machine_Name
    $returnObj | Add-Member -MemberType NoteProperty -Name NumberOfCores -Value ([int]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name Exception -Value ([string]::empty)
    $returnObj | Add-Member -MemberType NoteProperty -Name ExceptionType -Value ([string]::empty)

    try {
        $wmi_obj_processor = Get-WmiObjectHandler -ComputerName $Machine_Name -Class "Win32_Processor" -CatchActionFunction ${Function:Invoke-CatchActions}

        foreach ($processor in $wmi_obj_processor) {
            $returnObj.NumberOfCores += $processor.NumberOfCores
        }

        Write-Grey("Server {0} Cores: {1}" -f $Machine_Name, $returnObj.NumberOfCores)
    } catch {
        Invoke-CatchActions
        $thisError = $Error[0]

        if ($thisError.Exception.Gettype().FullName -eq "System.UnauthorizedAccessException") {
            Write-Yellow("Unable to get processor information from server {0}. You do not have the correct permissions to get this data from that server. Exception: {1}" -f $Machine_Name, $thisError.ToString())
        } else {
            Write-Yellow("Unable to get processor information from server {0}. Reason: {1}" -f $Machine_Name, $thisError.ToString())
        }
        $returnObj.Exception = $thisError.ToString()
        $returnObj.ExceptionType = $thisError.Exception.Gettype().FullName
        $returnObj.Error = $true
    }

    return $returnObj
}

function Get-ExchangeDCCoreRatio {

    Invoke-ScriptLogFileLocation -FileName "HealthChecker-ExchangeDCCoreRatio"
    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    Write-Grey("Exchange Server Health Checker Report - AD GC Core to Exchange Server Core Ratio - v{0}" -f $BuildVersion)
    $coreRatioObj = New-Object PSCustomObject

    try {
        Write-Verbose "Attempting to load Active Directory Module"
        Import-Module ActiveDirectory
        Write-Verbose "Successfully loaded"
    } catch {
        Write-Red("Failed to load Active Directory Module. Stopping the script")
        exit
    }

    $ADSite = [System.DirectoryServices.ActiveDirectory.ActiveDirectorySite]::GetComputerSite().Name
    [array]$DomainControllers = (Get-ADForest).Domains |
        ForEach-Object { Get-ADDomainController -Server $_ } |
        Where-Object { $_.IsGlobalCatalog -eq $true -and $_.Site -eq $ADSite }

    [System.Collections.Generic.List[System.Object]]$DCList = New-Object System.Collections.Generic.List[System.Object]
    $DCCoresTotal = 0
    Write-Break
    Write-Grey("Collecting data for the Active Directory Environment in Site: {0}" -f $ADSite)
    $iFailedDCs = 0

    foreach ($DC in $DomainControllers) {
        $DCCoreObj = Get-ComputerCoresObject -Machine_Name $DC.Name
        $DCList.Add($DCCoreObj)

        if (-not ($DCCoreObj.Error)) {
            $DCCoresTotal += $DCCoreObj.NumberOfCores
        } else {
            $iFailedDCs++
        }
    }

    $coreRatioObj | Add-Member -MemberType NoteProperty -Name DCList -Value $DCList

    if ($iFailedDCs -eq $DomainControllers.count) {
        #Core count is going to be 0, no point to continue the script
        Write-Red("Failed to collect data from your DC servers in site {0}." -f $ADSite)
        Write-Yellow("Because we can't determine the ratio, we are going to stop the script. Verify with the above errors as to why we failed to collect the data and address the issue, then run the script again.")
        exit
    }

    [array]$ExchangeServers = Get-ExchangeServer | Where-Object { $_.Site -match $ADSite }
    $EXCoresTotal = 0
    [System.Collections.Generic.List[System.Object]]$EXList = New-Object System.Collections.Generic.List[System.Object]
    Write-Break
    Write-Grey("Collecting data for the Exchange Environment in Site: {0}" -f $ADSite)
    foreach ($svr in $ExchangeServers) {
        $EXCoreObj = Get-ComputerCoresObject -Machine_Name $svr.Name
        $EXList.Add($EXCoreObj)

        if (-not ($EXCoreObj.Error)) {
            $EXCoresTotal += $EXCoreObj.NumberOfCores
        }
    }
    $coreRatioObj | Add-Member -MemberType NoteProperty -Name ExList -Value $EXList

    Write-Break
    $CoreRatio = $EXCoresTotal / $DCCoresTotal
    Write-Grey("Total DC/GC Cores: {0}" -f $DCCoresTotal)
    Write-Grey("Total Exchange Cores: {0}" -f $EXCoresTotal)
    Write-Grey("You have {0} Exchange Cores for every Domain Controller Global Catalog Server Core" -f $CoreRatio)

    if ($CoreRatio -gt 8) {
        Write-Break
        Write-Red("Your Exchange to Active Directory Global Catalog server's core ratio does not meet the recommended guidelines of 8:1")
        Write-Red("Recommended guidelines for Exchange 2013/2016 for every 8 Exchange cores you want at least 1 Active Directory Global Catalog Core.")
        Write-Yellow("Documentation:")
        Write-Yellow("`thttps://aka.ms/HC-PerfSize")
        Write-Yellow("`thttps://aka.ms/HC-ADCoreCount")
    } else {
        Write-Break
        Write-Green("Your Exchange Environment meets the recommended core ratio of 8:1 guidelines.")
    }

    $XMLDirectoryPath = $OutputFullPath.Replace(".txt", ".xml")
    $coreRatioObj | Export-Clixml $XMLDirectoryPath
    Write-Grey("Output file written to {0}" -f $OutputFullPath)
    Write-Grey("Output XML Object file written to {0}" -f $XMLDirectoryPath)
}

function Get-MailboxDatabaseAndMailboxStatistics {

    Write-Verbose "Calling: $($MyInvocation.MyCommand)"
    $AllDBs = Get-MailboxDatabaseCopyStatus -server $Script:Server -ErrorAction SilentlyContinue
    $MountedDBs = $AllDBs | Where-Object { $_.ActiveCopy -eq $true }

    if ($MountedDBs.Count -gt 0) {
        Write-Grey("`tActive Database:")
        foreach ($db in $MountedDBs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating User Mailbox Total for Active Database: $_"; $TotalActiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Active User Mailboxes on server: " + $TotalActiveUserMailboxCount)
        $MountedDBs.DatabaseName | ForEach-Object { Write-Verbose "Calculating Public Mailbox Total for Active Database: $_"; $TotalActivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Active Public Folder Mailboxes on server: " + $TotalActivePublicFolderMailboxCount)
        Write-Grey("`tTotal Active Mailboxes on server " + $Script:Server + ": " + ($TotalActiveUserMailboxCount + $TotalActivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Active Mailbox Databases found on server " + $Script:Server + ".")
    }

    $HealthyDbs = $AllDBs | Where-Object { $_.Status -match 'Healthy' }

    if ($HealthyDbs.count -gt 0) {
        Write-Grey("`r`n`tPassive Databases:")
        foreach ($db in $HealthyDbs) {
            Write-Grey("`t`t" + $db.Name)
        }
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating User Mailbox Total for Passive Healthy Databases: $_"; $TotalPassiveUserMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited).Count }
        Write-Grey("`tTotal Passive user Mailboxes on Server: " + $TotalPassiveUserMailboxCount)
        $HealthyDbs.DatabaseName | ForEach-Object { Write-Verbose "`tCalculating Passive Mailbox Total for Passive Healthy Databases: $_"; $TotalPassivePublicFolderMailboxCount += (Get-Mailbox -Database $_ -ResultSize Unlimited -PublicFolder).Count }
        Write-Grey("`tTotal Passive Public Mailboxes on server: " + $TotalPassivePublicFolderMailboxCount)
        Write-Grey("`tTotal Passive Mailboxes on server: " + ($TotalPassiveUserMailboxCount + $TotalPassivePublicFolderMailboxCount).ToString())
    } else {
        Write-Grey("`tNo Passive Mailboxes found on server " + $Script:Server + ".")
    }
}


function Confirm-Administrator {
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal( [Security.Principal.WindowsIdentity]::GetCurrent() )

    return $currentPrincipal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
}

function Get-NewLoggerInstance {
    [CmdletBinding()]
    param(
        [string]$LogDirectory = (Get-Location).Path,

        [ValidateNotNullOrEmpty()]
        [string]$LogName = "Script_Logging",

        [bool]$AppendDateTime = $true,

        [bool]$AppendDateTimeToFileName = $true,

        [int]$MaxFileSizeMB = 10,

        [int]$CheckSizeIntervalMinutes = 10,

        [int]$NumberOfLogsToKeep = 10
    )

    $fileName = if ($AppendDateTimeToFileName) { "{0}_{1}.txt" -f $LogName, ((Get-Date).ToString('yyyyMMddHHmmss')) } else { "$LogName.txt" }
    $fullFilePath = [System.IO.Path]::Combine($LogDirectory, $fileName)

    if (-not (Test-Path $LogDirectory)) {
        try {
            New-Item -ItemType Directory -Path $LogDirectory -ErrorAction Stop | Out-Null
        } catch {
            throw "Failed to create Log Directory: $LogDirectory"
        }
    }

    return [PSCustomObject]@{
        FullPath                 = $fullFilePath
        AppendDateTime           = $AppendDateTime
        MaxFileSizeMB            = $MaxFileSizeMB
        CheckSizeIntervalMinutes = $CheckSizeIntervalMinutes
        NumberOfLogsToKeep       = $NumberOfLogsToKeep
        BaseInstanceFileName     = $fileName.Replace(".txt", "")
        Instance                 = 1
        NextFileCheckTime        = ((Get-Date).AddMinutes($CheckSizeIntervalMinutes))
        PreventLogCleanup        = $false
        LoggerDisabled           = $false
    } | Write-LoggerInstance -Object "Starting Logger Instance $(Get-Date)"
}

function Write-LoggerInstance {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance,

        [Parameter(Mandatory = $true, Position = 1)]
        [object]$Object
    )
    process {
        if ($LoggerInstance.LoggerDisabled) { return }

        if ($LoggerInstance.AppendDateTime -and
            $Object.GetType().Name -eq "string") {
            $Object = "[$([System.DateTime]::Now)] : $Object"
        }

        # Doing WhatIf:$false to support -WhatIf in main scripts but still log the information
        $Object | Out-File $LoggerInstance.FullPath -Append -WhatIf:$false

        #Upkeep of the logger information
        if ($LoggerInstance.NextFileCheckTime -gt [System.DateTime]::Now) {
            return
        }

        #Set next update time to avoid issues so we can log things
        $LoggerInstance.NextFileCheckTime = ([System.DateTime]::Now).AddMinutes($LoggerInstance.CheckSizeIntervalMinutes)
        $item = Get-ChildItem $LoggerInstance.FullPath

        if (($item.Length / 1MB) -gt $LoggerInstance.MaxFileSizeMB) {
            $LoggerInstance | Write-LoggerInstance -Object "Max file size reached rolling over" | Out-Null
            $directory = [System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)
            $fileName = "$($LoggerInstance.BaseInstanceFileName)-$($LoggerInstance.Instance).txt"
            $LoggerInstance.Instance++
            $LoggerInstance.FullPath = [System.IO.Path]::Combine($directory, $fileName)

            $items = Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*"

            if ($items.Count -gt $LoggerInstance.NumberOfLogsToKeep) {
                $item = $items | Sort-Object LastWriteTime | Select-Object -First 1
                $LoggerInstance | Write-LoggerInstance "Removing Log File $($item.FullName)" | Out-Null
                $item | Remove-Item -Force
            }
        }
    }
    end {
        return $LoggerInstance
    }
}

function Invoke-LoggerInstanceCleanup {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true, ValueFromPipeline = $true)]
        [object]$LoggerInstance
    )
    process {
        if ($LoggerInstance.LoggerDisabled -or
            $LoggerInstance.PreventLogCleanup) {
            return
        }

        Get-ChildItem -Path ([System.IO.Path]::GetDirectoryName($LoggerInstance.FullPath)) -Filter "*$($LoggerInstance.BaseInstanceFileName)*" |
            Remove-Item -Force
    }
}

function Write-Host {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'Proper handling of write host with colors')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [object]$Object,
        [switch]$NoNewLine,
        [string]$ForegroundColor
    )
    process {
        $consoleHost = $host.Name -eq "ConsoleHost"

        if ($null -ne $Script:WriteHostManipulateObjectAction) {
            $Object = & $Script:WriteHostManipulateObjectAction $Object
        }

        $params = @{
            Object    = $Object
            NoNewLine = $NoNewLine
        }

        if ([string]::IsNullOrEmpty($ForegroundColor)) {
            if ($null -ne $host.UI.RawUI.ForegroundColor -and
                $consoleHost) {
                $params.Add("ForegroundColor", $host.UI.RawUI.ForegroundColor)
            }
        } elseif ($ForegroundColor -eq "Yellow" -and
            $consoleHost -and
            $null -ne $host.PrivateData.WarningForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.WarningForegroundColor)
        } elseif ($ForegroundColor -eq "Red" -and
            $consoleHost -and
            $null -ne $host.PrivateData.ErrorForegroundColor) {
            $params.Add("ForegroundColor", $host.PrivateData.ErrorForegroundColor)
        } else {
            $params.Add("ForegroundColor", $ForegroundColor)
        }

        Microsoft.PowerShell.Utility\Write-Host @params

        if ($null -ne $Script:WriteHostDebugAction -and
            $null -ne $Object) {
            &$Script:WriteHostDebugAction $Object
        }
    }
}

function SetProperForegroundColor {
    $Script:OriginalConsoleForegroundColor = $host.UI.RawUI.ForegroundColor

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.WarningForegroundColor) {
        Write-Verbose "Foreground Color matches warning's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }

    if ($Host.UI.RawUI.ForegroundColor -eq $Host.PrivateData.ErrorForegroundColor) {
        Write-Verbose "Foreground Color matches error's color"

        if ($Host.UI.RawUI.ForegroundColor -ne "Gray") {
            $Host.UI.RawUI.ForegroundColor = "Gray"
        }
    }
}

function RevertProperForegroundColor {
    $Host.UI.RawUI.ForegroundColor = $Script:OriginalConsoleForegroundColor
}

function SetWriteHostAction ($DebugAction) {
    $Script:WriteHostDebugAction = $DebugAction
}

function SetWriteHostManipulateObjectAction ($ManipulateObject) {
    $Script:WriteHostManipulateObjectAction = $ManipulateObject
}

function Write-Verbose {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidOverwritingBuiltInCmdlets', '', Justification = 'In order to log Write-Verbose from Shared functions')]
    [CmdletBinding()]
    param(
        [Parameter(Position = 1, ValueFromPipeline)]
        [string]$Message
    )

    process {

        if ($null -ne $Script:WriteVerboseManipulateMessageAction) {
            $Message = & $Script:WriteVerboseManipulateMessageAction $Message
        }

        Microsoft.PowerShell.Utility\Write-Verbose $Message

        if ($null -ne $Script:WriteVerboseDebugAction) {
            & $Script:WriteVerboseDebugAction $Message
        }

        # $PSSenderInfo is set when in a remote context
        if ($PSSenderInfo -and
            $null -ne $Script:WriteRemoteVerboseDebugAction) {
            & $Script:WriteRemoteVerboseDebugAction $Message
        }
    }
}

function SetWriteVerboseAction ($DebugAction) {
    $Script:WriteVerboseDebugAction = $DebugAction
}

function SetWriteRemoteVerboseAction ($DebugAction) {
    $Script:WriteRemoteVerboseDebugAction = $DebugAction
}

function SetWriteVerboseManipulateMessageAction ($DebugAction) {
    $Script:WriteVerboseManipulateMessageAction = $DebugAction
}




function Confirm-ProxyServer {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $TargetUri
    )

    try {
        $proxyObject = ([System.Net.WebRequest]::GetSystemWebproxy()).GetProxy($TargetUri)
        if ($TargetUri -ne $proxyObject.OriginalString) {
            return $true
        } else {
            return $false
        }
    } catch {
        return $false
    }
}

function Invoke-WebRequestWithProxyDetection {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $Uri,

        [Parameter(Mandatory = $false)]
        [switch]
        $UseBasicParsing,

        [Parameter(Mandatory = $false)]
        [string]
        $OutFile
    )

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    if (Confirm-ProxyServer -TargetUri $Uri) {
        $webClient = New-Object System.Net.WebClient
        $webClient.Headers.Add("User-Agent", "PowerShell")
        $webClient.Proxy.Credentials = [System.Net.CredentialCache]::DefaultNetworkCredentials
    }

    $params = @{
        Uri     = $Uri
        OutFile = $OutFile
    }

    if ($UseBasicParsing) {
        $params.UseBasicParsing = $true
    }

    Invoke-WebRequest @params
}

<#
    Determines if the script has an update available.
#>
function Get-ScriptUpdateAvailable {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param (
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $BuildVersion = "22.11.14.1812"

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $result = [PSCustomObject]@{
        ScriptName     = $scriptName
        CurrentVersion = $BuildVersion
        LatestVersion  = ""
        UpdateFound    = $false
        Error          = $null
    }

    if ((Get-AuthenticodeSignature -FilePath $scriptFullName).Status -eq "NotSigned") {
        Write-Warning "This script appears to be an unsigned test build. Skipping version check."
    } else {
        try {
            $versionData = [Text.Encoding]::UTF8.GetString((Invoke-WebRequestWithProxyDetection $VersionsUrl -UseBasicParsing).Content) | ConvertFrom-Csv
            $latestVersion = ($versionData | Where-Object { $_.File -eq $scriptName }).Version
            $result.LatestVersion = $latestVersion
            if ($null -ne $latestVersion -and $latestVersion -ne $BuildVersion) {
                $result.UpdateFound = $true
            }

            Write-Verbose "Current version: $($result.CurrentVersion) Latest version: $($result.LatestVersion) Update found: $($result.UpdateFound)"
        } catch {
            Write-Verbose "Unable to check for updates: $($_.Exception)"
            $result.Error = $_
        }
    }

    return $result
}


function Confirm-Signature {
    [CmdletBinding()]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $true)]
        [string]
        $File
    )

    $IsValid = $false
    $MicrosoftSigningRoot2010 = 'CN=Microsoft Root Certificate Authority 2010, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'
    $MicrosoftSigningRoot2011 = 'CN=Microsoft Root Certificate Authority 2011, O=Microsoft Corporation, L=Redmond, S=Washington, C=US'

    try {
        $sig = Get-AuthenticodeSignature -FilePath $File

        if ($sig.Status -ne 'Valid') {
            Write-Warning "Signature is not trusted by machine as Valid, status: $($sig.Status)."
            throw
        }

        $chain = New-Object -TypeName System.Security.Cryptography.X509Certificates.X509Chain
        $chain.ChainPolicy.VerificationFlags = "IgnoreNotTimeValid"

        if (-not $chain.Build($sig.SignerCertificate)) {
            Write-Warning "Signer certificate doesn't chain correctly."
            throw
        }

        if ($chain.ChainElements.Count -le 1) {
            Write-Warning "Certificate Chain shorter than expected."
            throw
        }

        $rootCert = $chain.ChainElements[$chain.ChainElements.Count - 1]

        if ($rootCert.Certificate.Subject -ne $rootCert.Certificate.Issuer) {
            Write-Warning "Top-level certifcate in chain is not a root certificate."
            throw
        }

        if ($rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2010 -and $rootCert.Certificate.Subject -ne $MicrosoftSigningRoot2011) {
            Write-Warning "Unexpected root cert. Expected $MicrosoftSigningRoot2010 or $MicrosoftSigningRoot2011, but found $($rootCert.Certificate.Subject)."
            throw
        }

        Write-Host "File signed by $($sig.SignerCertificate.Subject)"

        $IsValid = $true
    } catch {
        $IsValid = $false
    }

    $IsValid
}

<#
.SYNOPSIS
    Overwrites the current running script file with the latest version from the repository.
.NOTES
    This function always overwrites the current file with the latest file, which might be
    the same. Get-ScriptUpdateAvailable should be called first to determine if an update is
    needed.

    In many situations, updates are expected to fail, because the server running the script
    does not have internet access. This function writes out failures as warnings, because we
    expect that Get-ScriptUpdateAvailable was already called and it successfully reached out
    to the internet.
#>
function Invoke-ScriptUpdate {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = 'High')]
    [OutputType([boolean])]
    param ()

    $scriptName = $script:MyInvocation.MyCommand.Name
    $scriptPath = [IO.Path]::GetDirectoryName($script:MyInvocation.MyCommand.Path)
    $scriptFullName = (Join-Path $scriptPath $scriptName)

    $oldName = [IO.Path]::GetFileNameWithoutExtension($scriptName) + ".old"
    $oldFullName = (Join-Path $scriptPath $oldName)
    $tempFullName = (Join-Path $env:TEMP $scriptName)

    if ($PSCmdlet.ShouldProcess("$scriptName", "Update script to latest version")) {
        try {
            Invoke-WebRequestWithProxyDetection "https://github.com/microsoft/CSS-Exchange/releases/latest/download/$scriptName" -OutFile $tempFullName
        } catch {
            Write-Warning "AutoUpdate: Failed to download update: $($_.Exception.Message)"
            return $false
        }

        try {
            if (Confirm-Signature -File $tempFullName) {
                Write-Host "AutoUpdate: Signature validated."
                if (Test-Path $oldFullName) {
                    Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                }
                Move-Item $scriptFullName $oldFullName
                Move-Item $tempFullName $scriptFullName
                Remove-Item $oldFullName -Force -Confirm:$false -ErrorAction Stop
                Write-Host "AutoUpdate: Succeeded."
                return $true
            } else {
                Write-Warning "AutoUpdate: Signature could not be verified: $tempFullName."
                Write-Warning "AutoUpdate: Update was not applied."
            }
        } catch {
            Write-Warning "AutoUpdate: Failed to apply update: $($_.Exception.Message)"
        }
    }

    return $false
}

<#
    Determines if the script has an update available. Use the optional
    -AutoUpdate switch to make it update itself. Pass -Confirm:$false
    to update without prompting the user. Pass -Verbose for additional
    diagnostic output.

    Returns $true if an update was downloaded, $false otherwise. The
    result will always be $false if the -AutoUpdate switch is not used.
#>
function Test-ScriptVersion {
    [Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSShouldProcess', '', Justification = 'Need to pass through ShouldProcess settings to Invoke-ScriptUpdate')]
    [CmdletBinding(SupportsShouldProcess)]
    [OutputType([bool])]
    param (
        [Parameter(Mandatory = $false)]
        [switch]
        $AutoUpdate,
        [Parameter(Mandatory = $false)]
        [string]
        $VersionsUrl = "https://github.com/microsoft/CSS-Exchange/releases/latest/download/ScriptVersions.csv"
    )

    $updateInfo = Get-ScriptUpdateAvailable $VersionsUrl
    if ($updateInfo.UpdateFound) {
        if ($AutoUpdate) {
            return Invoke-ScriptUpdate
        } else {
            Write-Warning "$($updateInfo.ScriptName) $BuildVersion is outdated. Please download the latest, version $($updateInfo.LatestVersion)."
        }
    }

    return $false
}

function Main {

    if (-not (Confirm-Administrator) -and
        (-not $AnalyzeDataOnly -and
        -not $BuildHtmlServersReport -and
        -not $ScriptUpdateOnly)) {
        Write-Warning "The script needs to be executed in elevated mode. Start the Exchange Management Shell as an Administrator."
        $Error.Clear()
        Start-Sleep -Seconds 2;
        exit
    }

    Invoke-ErrorMonitoring
    $Script:date = (Get-Date)
    $Script:dateTimeStringFormat = $date.ToString("yyyyMMddHHmmss")

    if ($BuildHtmlServersReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-HTMLServerReport"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths
        Get-HtmlServerReport -AnalyzedHtmlServerValues $importData.HtmlServerValues
        Start-Sleep 2;
        return
    }

    if ((Test-Path $OutputFilePath) -eq $false) {
        Write-Host "Invalid value specified for -OutputFilePath." -ForegroundColor Red
        return
    }

    if ($LoadBalancingReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-LoadBalancingReport"
        Write-Green("Client Access Load Balancing Report on " + $date)
        Get-CASLoadBalancingReport
        Write-Grey("Output file written to " + $OutputFullPath)
        Write-Break
        Write-Break
        return
    }

    if ($DCCoreRatio) {
        $oldErrorAction = $ErrorActionPreference
        $ErrorActionPreference = "Stop"
        try {
            Get-ExchangeDCCoreRatio
            return
        } finally {
            $ErrorActionPreference = $oldErrorAction
        }
    }

    if ($MailboxReport) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-MailboxReport" -IncludeServerName $true
        Get-MailboxDatabaseAndMailboxStatistics
        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        return
    }

    if ($AnalyzeDataOnly) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-Analyzer"
        $files = Get-HealthCheckFilesItemsFromLocation
        $fullPaths = Get-OnlyRecentUniqueServersXMLs $files
        $importData = Import-MyData -FilePaths $fullPaths

        $analyzedResults = @()
        foreach ($serverData in $importData) {
            $analyzedServerResults = Invoke-AnalyzerEngine -HealthServerObject $serverData.HealthCheckerExchangeServer
            Write-ResultsToScreen -ResultsToWrite $analyzedServerResults.DisplayResults
            $analyzedResults += $analyzedServerResults
        }

        Get-HtmlServerReport -AnalyzedHtmlServerValues $analyzedResults.HtmlServerValues
        return
    }

    if ($ScriptUpdateOnly) {
        Invoke-ScriptLogFileLocation -FileName "HealthChecker-ScriptUpdateOnly"
        switch (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/HC-VersionsUrl" -Confirm:$false) {
            ($true) { Write-Green("Script was successfully updated.") }
            ($false) { Write-Yellow("No update of the script performed.") }
            default { Write-Red("Unable to perform ScriptUpdateOnly operation.") }
        }
        return
    }

    Invoke-ScriptLogFileLocation -FileName "HealthChecker" -IncludeServerName $true
    $currentErrors = $Error.Count

    if ((-not $SkipVersionCheck) -and
        (Test-ScriptVersion -AutoUpdate -VersionsUrl "https://aka.ms/HC-VersionsUrl")) {
        Write-Yellow "Script was updated. Please rerun the command."
        return
    } else {
        $Script:DisplayedScriptVersionAlready = $true
        Write-Green "Exchange Health Checker version $BuildVersion"
    }

    Invoke-ErrorCatchActionLoopFromIndex $currentErrors
    Test-RequiresServerFqdn
    [HealthChecker.HealthCheckerExchangeServer]$HealthObject = Get-HealthCheckerExchangeServer
    $analyzedResults = Invoke-AnalyzerEngine -HealthServerObject $HealthObject
    Write-ResultsToScreen -ResultsToWrite $analyzedResults.DisplayResults
    $currentErrors = $Error.Count

    try {
        $analyzedResults | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction SilentlyContinue
    } catch {
        Write-Verbose "Failed to Export-Clixml. Converting HealthCheckerExchangeServer to json"
        $jsonHealthChecker = $analyzedResults.HealthCheckerExchangeServer | ConvertTo-Json

        $testOuputxml = [PSCustomObject]@{
            HealthCheckerExchangeServer = $jsonHealthChecker | ConvertFrom-Json
            HtmlServerValues            = $analyzedResults.HtmlServerValues
            DisplayResults              = $analyzedResults.DisplayResults
        }

        $testOuputxml | Export-Clixml -Path $OutXmlFullPath -Encoding UTF8 -Depth 6 -ErrorAction Stop
    } finally {
        Invoke-ErrorCatchActionLoopFromIndex $currentErrors

        Write-Grey("Output file written to {0}" -f $Script:OutputFullPath)
        Write-Grey("Exported Data Object Written to {0} " -f $Script:OutXmlFullPath)
    }
}

try {
    $Script:Logger = Get-NewLoggerInstance -LogName "HealthChecker-$($Script:Server)-Debug" `
        -LogDirectory $OutputFilePath `
        -AppendDateTime $false `
        -ErrorAction SilentlyContinue
    SetProperForegroundColor
    SetWriteVerboseAction ${Function:Write-DebugLog}
    Main
} finally {
    Get-ErrorsThatOccurred
    if ($Script:VerboseEnabled) {
        $Host.PrivateData.VerboseForegroundColor = $VerboseForeground
    }
    $Script:Logger | Invoke-LoggerInstanceCleanup
    if ($Script:Logger.PreventLogCleanup) {
        Write-Host("Output Debug file written to {0}" -f $Script:Logger.FullPath)
    }
    if (((Get-Date).Ticks % 2) -eq 1) {
        Write-Host("Do you like the script? Visit https://aka.ms/HC-Feedback to rate it and to provide feedback.") -ForegroundColor Green
        Write-Host
    }
    RevertProperForegroundColor
}

# SIG # Begin signature block
# MIIn4gYJKoZIhvcNAQcCoIIn0zCCJ88CAQExDzANBglghkgBZQMEAgEFADB5Bgor
# BgEEAYI3AgEEoGswaTA0BgorBgEEAYI3AgEeMCYCAwEAAAQQH8w7YFlLCE63JNLG
# KX7zUQIBAAIBAAIBAAIBAAIBADAxMA0GCWCGSAFlAwQCAQUABCBUqJ0Zdn8M59M/
# SL7KdTyiazLYQLFzUp2bXJb9XuDceqCCDYEwggX/MIID56ADAgECAhMzAAACzI61
# lqa90clOAAAAAALMMA0GCSqGSIb3DQEBCwUAMH4xCzAJBgNVBAYTAlVTMRMwEQYD
# VQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMTH01pY3Jvc29mdCBDb2RlIFNpZ25p
# bmcgUENBIDIwMTEwHhcNMjIwNTEyMjA0NjAxWhcNMjMwNTExMjA0NjAxWjB0MQsw
# CQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9u
# ZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMR4wHAYDVQQDExVNaWNy
# b3NvZnQgQ29ycG9yYXRpb24wggEiMA0GCSqGSIb3DQEBAQUAA4IBDwAwggEKAoIB
# AQCiTbHs68bADvNud97NzcdP0zh0mRr4VpDv68KobjQFybVAuVgiINf9aG2zQtWK
# No6+2X2Ix65KGcBXuZyEi0oBUAAGnIe5O5q/Y0Ij0WwDyMWaVad2Te4r1Eic3HWH
# UfiiNjF0ETHKg3qa7DCyUqwsR9q5SaXuHlYCwM+m59Nl3jKnYnKLLfzhl13wImV9
# DF8N76ANkRyK6BYoc9I6hHF2MCTQYWbQ4fXgzKhgzj4zeabWgfu+ZJCiFLkogvc0
# RVb0x3DtyxMbl/3e45Eu+sn/x6EVwbJZVvtQYcmdGF1yAYht+JnNmWwAxL8MgHMz
# xEcoY1Q1JtstiY3+u3ulGMvhAgMBAAGjggF+MIIBejAfBgNVHSUEGDAWBgorBgEE
# AYI3TAgBBggrBgEFBQcDAzAdBgNVHQ4EFgQUiLhHjTKWzIqVIp+sM2rOHH11rfQw
# UAYDVR0RBEkwR6RFMEMxKTAnBgNVBAsTIE1pY3Jvc29mdCBPcGVyYXRpb25zIFB1
# ZXJ0byBSaWNvMRYwFAYDVQQFEw0yMzAwMTIrNDcwNTI5MB8GA1UdIwQYMBaAFEhu
# ZOVQBdOCqhc3NyK1bajKdQKVMFQGA1UdHwRNMEswSaBHoEWGQ2h0dHA6Ly93d3cu
# bWljcm9zb2Z0LmNvbS9wa2lvcHMvY3JsL01pY0NvZFNpZ1BDQTIwMTFfMjAxMS0w
# Ny0wOC5jcmwwYQYIKwYBBQUHAQEEVTBTMFEGCCsGAQUFBzAChkVodHRwOi8vd3d3
# Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NlcnRzL01pY0NvZFNpZ1BDQTIwMTFfMjAx
# MS0wNy0wOC5jcnQwDAYDVR0TAQH/BAIwADANBgkqhkiG9w0BAQsFAAOCAgEAeA8D
# sOAHS53MTIHYu8bbXrO6yQtRD6JfyMWeXaLu3Nc8PDnFc1efYq/F3MGx/aiwNbcs
# J2MU7BKNWTP5JQVBA2GNIeR3mScXqnOsv1XqXPvZeISDVWLaBQzceItdIwgo6B13
# vxlkkSYMvB0Dr3Yw7/W9U4Wk5K/RDOnIGvmKqKi3AwyxlV1mpefy729FKaWT7edB
# d3I4+hldMY8sdfDPjWRtJzjMjXZs41OUOwtHccPazjjC7KndzvZHx/0VWL8n0NT/
# 404vftnXKifMZkS4p2sB3oK+6kCcsyWsgS/3eYGw1Fe4MOnin1RhgrW1rHPODJTG
# AUOmW4wc3Q6KKr2zve7sMDZe9tfylonPwhk971rX8qGw6LkrGFv31IJeJSe/aUbG
# dUDPkbrABbVvPElgoj5eP3REqx5jdfkQw7tOdWkhn0jDUh2uQen9Atj3RkJyHuR0
# GUsJVMWFJdkIO/gFwzoOGlHNsmxvpANV86/1qgb1oZXdrURpzJp53MsDaBY/pxOc
# J0Cvg6uWs3kQWgKk5aBzvsX95BzdItHTpVMtVPW4q41XEvbFmUP1n6oL5rdNdrTM
# j/HXMRk1KCksax1Vxo3qv+13cCsZAaQNaIAvt5LvkshZkDZIP//0Hnq7NnWeYR3z
# 4oFiw9N2n3bb9baQWuWPswG0Dq9YT9kb+Cs4qIIwggd6MIIFYqADAgECAgphDpDS
# AAAAAAADMA0GCSqGSIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMK
# V2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0
# IENvcnBvcmF0aW9uMTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0
# ZSBBdXRob3JpdHkgMjAxMTAeFw0xMTA3MDgyMDU5MDlaFw0yNjA3MDgyMTA5MDla
# MH4xCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdS
# ZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xKDAmBgNVBAMT
# H01pY3Jvc29mdCBDb2RlIFNpZ25pbmcgUENBIDIwMTEwggIiMA0GCSqGSIb3DQEB
# AQUAA4ICDwAwggIKAoICAQCr8PpyEBwurdhuqoIQTTS68rZYIZ9CGypr6VpQqrgG
# OBoESbp/wwwe3TdrxhLYC/A4wpkGsMg51QEUMULTiQ15ZId+lGAkbK+eSZzpaF7S
# 35tTsgosw6/ZqSuuegmv15ZZymAaBelmdugyUiYSL+erCFDPs0S3XdjELgN1q2jz
# y23zOlyhFvRGuuA4ZKxuZDV4pqBjDy3TQJP4494HDdVceaVJKecNvqATd76UPe/7
# 4ytaEB9NViiienLgEjq3SV7Y7e1DkYPZe7J7hhvZPrGMXeiJT4Qa8qEvWeSQOy2u
# M1jFtz7+MtOzAz2xsq+SOH7SnYAs9U5WkSE1JcM5bmR/U7qcD60ZI4TL9LoDho33
# X/DQUr+MlIe8wCF0JV8YKLbMJyg4JZg5SjbPfLGSrhwjp6lm7GEfauEoSZ1fiOIl
# XdMhSz5SxLVXPyQD8NF6Wy/VI+NwXQ9RRnez+ADhvKwCgl/bwBWzvRvUVUvnOaEP
# 6SNJvBi4RHxF5MHDcnrgcuck379GmcXvwhxX24ON7E1JMKerjt/sW5+v/N2wZuLB
# l4F77dbtS+dJKacTKKanfWeA5opieF+yL4TXV5xcv3coKPHtbcMojyyPQDdPweGF
# RInECUzF1KVDL3SV9274eCBYLBNdYJWaPk8zhNqwiBfenk70lrC8RqBsmNLg1oiM
# CwIDAQABo4IB7TCCAekwEAYJKwYBBAGCNxUBBAMCAQAwHQYDVR0OBBYEFEhuZOVQ
# BdOCqhc3NyK1bajKdQKVMBkGCSsGAQQBgjcUAgQMHgoAUwB1AGIAQwBBMAsGA1Ud
# DwQEAwIBhjAPBgNVHRMBAf8EBTADAQH/MB8GA1UdIwQYMBaAFHItOgIxkEO5FAVO
# 4eqnxzHRI4k0MFoGA1UdHwRTMFEwT6BNoEuGSWh0dHA6Ly9jcmwubWljcm9zb2Z0
# LmNvbS9wa2kvY3JsL3Byb2R1Y3RzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcmwwXgYIKwYBBQUHAQEEUjBQME4GCCsGAQUFBzAChkJodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpL2NlcnRzL01pY1Jvb0NlckF1dDIwMTFfMjAxMV8wM18y
# Mi5jcnQwgZ8GA1UdIASBlzCBlDCBkQYJKwYBBAGCNy4DMIGDMD8GCCsGAQUFBwIB
# FjNodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2RvY3MvcHJpbWFyeWNw
# cy5odG0wQAYIKwYBBQUHAgIwNB4yIB0ATABlAGcAYQBsAF8AcABvAGwAaQBjAHkA
# XwBzAHQAYQB0AGUAbQBlAG4AdAAuIB0wDQYJKoZIhvcNAQELBQADggIBAGfyhqWY
# 4FR5Gi7T2HRnIpsLlhHhY5KZQpZ90nkMkMFlXy4sPvjDctFtg/6+P+gKyju/R6mj
# 82nbY78iNaWXXWWEkH2LRlBV2AySfNIaSxzzPEKLUtCw/WvjPgcuKZvmPRul1LUd
# d5Q54ulkyUQ9eHoj8xN9ppB0g430yyYCRirCihC7pKkFDJvtaPpoLpWgKj8qa1hJ
# Yx8JaW5amJbkg/TAj/NGK978O9C9Ne9uJa7lryft0N3zDq+ZKJeYTQ49C/IIidYf
# wzIY4vDFLc5bnrRJOQrGCsLGra7lstnbFYhRRVg4MnEnGn+x9Cf43iw6IGmYslmJ
# aG5vp7d0w0AFBqYBKig+gj8TTWYLwLNN9eGPfxxvFX1Fp3blQCplo8NdUmKGwx1j
# NpeG39rz+PIWoZon4c2ll9DuXWNB41sHnIc+BncG0QaxdR8UvmFhtfDcxhsEvt9B
# xw4o7t5lL+yX9qFcltgA1qFGvVnzl6UJS0gQmYAf0AApxbGbpT9Fdx41xtKiop96
# eiL6SJUfq/tHI4D1nvi/a7dLl+LrdXga7Oo3mXkYS//WsyNodeav+vyL6wuA6mk7
# r/ww7QRMjt/fdW1jkT3RnVZOT7+AVyKheBEyIXrvQQqxP/uozKRdwaGIm1dxVk5I
# RcBCyZt2WwqASGv9eZ/BvW1taslScxMNelDNMYIZtzCCGbMCAQEwgZUwfjELMAkG
# A1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1vbmQx
# HjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEoMCYGA1UEAxMfTWljcm9z
# b2Z0IENvZGUgU2lnbmluZyBQQ0EgMjAxMQITMwAAAsyOtZamvdHJTgAAAAACzDAN
# BglghkgBZQMEAgEFAKCBxjAZBgkqhkiG9w0BCQMxDAYKKwYBBAGCNwIBBDAcBgor
# BgEEAYI3AgELMQ4wDAYKKwYBBAGCNwIBFTAvBgkqhkiG9w0BCQQxIgQgUAuN2stg
# QHxVdQq+GSHbFfhBtkQ4uKXBapVLqA4gJCwwWgYKKwYBBAGCNwIBDDFMMEqgGoAY
# AEMAUwBTACAARQB4AGMAaABhAG4AZwBloSyAKmh0dHBzOi8vZ2l0aHViLmNvbS9t
# aWNyb3NvZnQvQ1NTLUV4Y2hhbmdlIDANBgkqhkiG9w0BAQEFAASCAQBBeCUJelOO
# GPmWI6lCt8opwJksSzJMcosnU5Y+X5XPJs/Ofv9eYTfLyOiPdDxY3fmw0nT5+v7J
# oH9JPJmkA3T/Hj7LqV+j9qH5PzAKSdrHSwAdtH8e7gZpLXpYr8sSILh2rVrWcdR2
# AAoJi3dNsjzCs4pUhMWpHWi4R9pnE48io783AK2xjrLaRLa1fSaifU5FpBV9wRyU
# omEnZ75QDdU5Eyx/ylRh3hgW7N85xPwG7RerClD6ov0MpXGOUCumleLiM6G43l5O
# pfeXiNSSP6KbOoRSG2SWPwyYIbBYTFLBMY6MkQ3hrgmnLMi4ky5JqU3c22aNwdrP
# 3FWFZ9IHn8mPoYIXKTCCFyUGCisGAQQBgjcDAwExghcVMIIXEQYJKoZIhvcNAQcC
# oIIXAjCCFv4CAQMxDzANBglghkgBZQMEAgEFADCCAVkGCyqGSIb3DQEJEAEEoIIB
# SASCAUQwggFAAgEBBgorBgEEAYRZCgMBMDEwDQYJYIZIAWUDBAIBBQAEIEVmXaUf
# Y1QViebXfpzXAkbOByI4MyNj8bNTK3ET5yioAgZjYtdMM0oYEzIwMjIxMTE1MTg0
# NjAxLjc3MVowBIACAfSggdikgdUwgdIxCzAJBgNVBAYTAlVTMRMwEQYDVQQIEwpX
# YXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYDVQQKExVNaWNyb3NvZnQg
# Q29ycG9yYXRpb24xLTArBgNVBAsTJE1pY3Jvc29mdCBJcmVsYW5kIE9wZXJhdGlv
# bnMgTGltaXRlZDEmMCQGA1UECxMdVGhhbGVzIFRTUyBFU046MTc5RS00QkIwLTgy
# NDYxJTAjBgNVBAMTHE1pY3Jvc29mdCBUaW1lLVN0YW1wIFNlcnZpY2WgghF4MIIH
# JzCCBQ+gAwIBAgITMwAAAbWtGt/XhXBtEwABAAABtTANBgkqhkiG9w0BAQsFADB8
# MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQMA4GA1UEBxMHUmVk
# bW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9uMSYwJAYDVQQDEx1N
# aWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAxMDAeFw0yMjA5MjAyMDIyMTFaFw0y
# MzEyMTQyMDIyMTFaMIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3Rv
# bjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0
# aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25zIExpbWl0
# ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjE3OUUtNEJCMC04MjQ2MSUwIwYD
# VQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNlMIICIjANBgkqhkiG9w0B
# AQEFAAOCAg8AMIICCgKCAgEAlwsKuGVegsKNiYXFwU+CSHnt2a7PfWw2yPwiW+YR
# lEJsH3ibFIiPfk/yblMp8JGantu+7Di/+3e5wWN/nbJUIMUjEWJnc8JMjoPmHCWs
# MtJOuR/1Ru4aa1RrxQtIelq098TBl4k7NsEE87l7qKFmy8iwGNQjkwr0bMu4BJwy
# 7BUXiXHegOSU992rfQ4xNZoxznv42TLQsc9NmcBq5WslkqVATcc8PSfgBLEpdG1D
# p2wqNw4JrJFwJNA1bfzTScYABc5smRZBgsP4JiK/8CVrlocheEyQonjm3rFttroj
# AreSUnixALu9pDrsBI4DUPGG34oIbieI1oqFl/xk7A+7uM8k4o8ifMVWNTaczbPl
# dDYtn6hBre7r25RED4uecCxP8Dxy34YPUElWllPP3LAXp5cMwRjx+EWzjEtILEKX
# uAcfxrXCTwyYhm5XNzCCZYh4/gF2U2y/bYfekKpaoFYwkoZeT6ZxoQbX5Kftgj+t
# ZkFV21UvZIkJ6b34a/44dtrsK6diTmVnNTM9J6P6Ehlk2sfcUwbHIGL8mYqdKOiy
# d4RxOCmSvcFNkZEgrk548mHCbDbTyO9xSzN1EkWxbp8n/LHVnZ9fp5hILGntkMza
# D5aXRCQyHSIhsPtR7Q/rKoHyjFqgtGO9ftnxYvxzNrbKeMCzwmcqwMrX6Hcxe0Se
# KZ8CAwEAAaOCAUkwggFFMB0GA1UdDgQWBBRsUIbZgoZVXVXVWQX0Ok1VO2bHUzAf
# BgNVHSMEGDAWgBSfpxVdAF5iXYP05dJlpxtTNRnpcjBfBgNVHR8EWDBWMFSgUqBQ
# hk5odHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2NybC9NaWNyb3NvZnQl
# MjBUaW1lLVN0YW1wJTIwUENBJTIwMjAxMCgxKS5jcmwwbAYIKwYBBQUHAQEEYDBe
# MFwGCCsGAQUFBzAChlBodHRwOi8vd3d3Lm1pY3Jvc29mdC5jb20vcGtpb3BzL2Nl
# cnRzL01pY3Jvc29mdCUyMFRpbWUtU3RhbXAlMjBQQ0ElMjAyMDEwKDEpLmNydDAM
# BgNVHRMBAf8EAjAAMBYGA1UdJQEB/wQMMAoGCCsGAQUFBwMIMA4GA1UdDwEB/wQE
# AwIHgDANBgkqhkiG9w0BAQsFAAOCAgEAkFGOpyjKV2s2sA+wTqDwDdhp0mFrPtiU
# 4rN3OonTWqb85M6WH19c/P517xujLCih/HllP5xKWmXnAIRV1/NQDkJBLSdLTb/N
# QtcT1FWGQ7CMTnrn9tLZxqIFtKVylvQNyh31C/qkC8QmNpyzakO0G38uOGgOkJ9E
# q4nA+7QwVfobDlggWuEpzdFnRdyXL32gOqSvrLjFKpv4KEVqaBTiaxCWZDlIhG3Y
# gUza7cnG5Z2SA/feMq/IiV06AzUadZw6XgcTrqXmEmE0tMmdl44MMFC3wGU9AVeF
# CWKdD9WOnYA2zHg+XF2LQVto0VYtFLd6c6DQFcmB38GvPCKVYSn8r10EoXuRN+gQ
# 7hLcim12esOnW4F4bHCmHWTVWeAGgPiSItHHRfGKLEUZmotVOdFPR8wiuADT/fHS
# XBkkdpL12tvgEGELeTznzFulZ16b/Nv6dtbgSRZreesJBNKpTjdYju/GqnlAkpfl
# L6J0wxk957/UVYnmjjRY61jX90QGQmBzm9vs/+2bj02Xx/bXXy8vq57jmNXQ2ufO
# aJm3nAcD2qOaSyXEOj9mqhMt4tdvMjHhiNPldfj0Q7Kq1HgdRBrKWkzCQNi4ts8H
# RJBipNaVpWfU7BcRn8BeYzdLoIzwRLDtatz6aBho3oD/bXHrZagxprM5MsMB/rVf
# b5Xn1YS7/uEwggdxMIIFWaADAgECAhMzAAAAFcXna54Cm0mZAAAAAAAVMA0GCSqG
# SIb3DQEBCwUAMIGIMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2FzaGluZ3RvbjEQ
# MA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENvcnBvcmF0aW9u
# MTIwMAYDVQQDEylNaWNyb3NvZnQgUm9vdCBDZXJ0aWZpY2F0ZSBBdXRob3JpdHkg
# MjAxMDAeFw0yMTA5MzAxODIyMjVaFw0zMDA5MzAxODMyMjVaMHwxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQSAyMDEwMIICIjANBgkqhkiG9w0BAQEFAAOCAg8AMIICCgKC
# AgEA5OGmTOe0ciELeaLL1yR5vQ7VgtP97pwHB9KpbE51yMo1V/YBf2xK4OK9uT4X
# YDP/XE/HZveVU3Fa4n5KWv64NmeFRiMMtY0Tz3cywBAY6GB9alKDRLemjkZrBxTz
# xXb1hlDcwUTIcVxRMTegCjhuje3XD9gmU3w5YQJ6xKr9cmmvHaus9ja+NSZk2pg7
# uhp7M62AW36MEBydUv626GIl3GoPz130/o5Tz9bshVZN7928jaTjkY+yOSxRnOlw
# aQ3KNi1wjjHINSi947SHJMPgyY9+tVSP3PoFVZhtaDuaRr3tpK56KTesy+uDRedG
# bsoy1cCGMFxPLOJiss254o2I5JasAUq7vnGpF1tnYN74kpEeHT39IM9zfUGaRnXN
# xF803RKJ1v2lIH1+/NmeRd+2ci/bfV+AutuqfjbsNkz2K26oElHovwUDo9Fzpk03
# dJQcNIIP8BDyt0cY7afomXw/TNuvXsLz1dhzPUNOwTM5TI4CvEJoLhDqhFFG4tG9
# ahhaYQFzymeiXtcodgLiMxhy16cg8ML6EgrXY28MyTZki1ugpoMhXV8wdJGUlNi5
# UPkLiWHzNgY1GIRH29wb0f2y1BzFa/ZcUlFdEtsluq9QBXpsxREdcu+N+VLEhReT
# wDwV2xo3xwgVGD94q0W29R6HXtqPnhZyacaue7e3PmriLq0CAwEAAaOCAd0wggHZ
# MBIGCSsGAQQBgjcVAQQFAgMBAAEwIwYJKwYBBAGCNxUCBBYEFCqnUv5kxJq+gpE8
# RjUpzxD/LwTuMB0GA1UdDgQWBBSfpxVdAF5iXYP05dJlpxtTNRnpcjBcBgNVHSAE
# VTBTMFEGDCsGAQQBgjdMg30BATBBMD8GCCsGAQUFBwIBFjNodHRwOi8vd3d3Lm1p
# Y3Jvc29mdC5jb20vcGtpb3BzL0RvY3MvUmVwb3NpdG9yeS5odG0wEwYDVR0lBAww
# CgYIKwYBBQUHAwgwGQYJKwYBBAGCNxQCBAweCgBTAHUAYgBDAEEwCwYDVR0PBAQD
# AgGGMA8GA1UdEwEB/wQFMAMBAf8wHwYDVR0jBBgwFoAU1fZWy4/oolxiaNE9lJBb
# 186aGMQwVgYDVR0fBE8wTTBLoEmgR4ZFaHR0cDovL2NybC5taWNyb3NvZnQuY29t
# L3BraS9jcmwvcHJvZHVjdHMvTWljUm9vQ2VyQXV0XzIwMTAtMDYtMjMuY3JsMFoG
# CCsGAQUFBwEBBE4wTDBKBggrBgEFBQcwAoY+aHR0cDovL3d3dy5taWNyb3NvZnQu
# Y29tL3BraS9jZXJ0cy9NaWNSb29DZXJBdXRfMjAxMC0wNi0yMy5jcnQwDQYJKoZI
# hvcNAQELBQADggIBAJ1VffwqreEsH2cBMSRb4Z5yS/ypb+pcFLY+TkdkeLEGk5c9
# MTO1OdfCcTY/2mRsfNB1OW27DzHkwo/7bNGhlBgi7ulmZzpTTd2YurYeeNg2Lpyp
# glYAA7AFvonoaeC6Ce5732pvvinLbtg/SHUB2RjebYIM9W0jVOR4U3UkV7ndn/OO
# PcbzaN9l9qRWqveVtihVJ9AkvUCgvxm2EhIRXT0n4ECWOKz3+SmJw7wXsFSFQrP8
# DJ6LGYnn8AtqgcKBGUIZUnWKNsIdw2FzLixre24/LAl4FOmRsqlb30mjdAy87JGA
# 0j3mSj5mO0+7hvoyGtmW9I/2kQH2zsZ0/fZMcm8Qq3UwxTSwethQ/gpY3UA8x1Rt
# nWN0SCyxTkctwRQEcb9k+SS+c23Kjgm9swFXSVRk2XPXfx5bRAGOWhmRaw2fpCjc
# ZxkoJLo4S5pu+yFUa2pFEUep8beuyOiJXk+d0tBMdrVXVAmxaQFEfnyhYWxz/gq7
# 7EFmPWn9y8FBSX5+k77L+DvktxW/tM4+pTFRhLy/AsGConsXHRWJjXD+57XQKBqJ
# C4822rpM+Zv/Cuk0+CQ1ZyvgDbjmjJnW4SLq8CdCPSWU5nR0W2rRnj7tfqAxM328
# y+l7vzhwRNGQ8cirOoo6CGJ/2XBjU02N7oJtpQUQwXEGahC0HVUzWLOhcGbyoYIC
# 1DCCAj0CAQEwggEAoYHYpIHVMIHSMQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMS0wKwYDVQQLEyRNaWNyb3NvZnQgSXJlbGFuZCBPcGVyYXRpb25z
# IExpbWl0ZWQxJjAkBgNVBAsTHVRoYWxlcyBUU1MgRVNOOjE3OUUtNEJCMC04MjQ2
# MSUwIwYDVQQDExxNaWNyb3NvZnQgVGltZS1TdGFtcCBTZXJ2aWNloiMKAQEwBwYF
# Kw4DAhoDFQCNMJ9r11RZj0PWu3uk+aQHF3IsVaCBgzCBgKR+MHwxCzAJBgNVBAYT
# AlVTMRMwEQYDVQQIEwpXYXNoaW5ndG9uMRAwDgYDVQQHEwdSZWRtb25kMR4wHAYD
# VQQKExVNaWNyb3NvZnQgQ29ycG9yYXRpb24xJjAkBgNVBAMTHU1pY3Jvc29mdCBU
# aW1lLVN0YW1wIFBDQSAyMDEwMA0GCSqGSIb3DQEBBQUAAgUA5x3P0jAiGA8yMDIy
# MTExNTE2NDQwMloYDzIwMjIxMTE2MTY0NDAyWjB0MDoGCisGAQQBhFkKBAExLDAq
# MAoCBQDnHc/SAgEAMAcCAQACAjcZMAcCAQACAhIqMAoCBQDnHyFSAgEAMDYGCisG
# AQQBhFkKBAIxKDAmMAwGCisGAQQBhFkKAwKgCjAIAgEAAgMHoSChCjAIAgEAAgMB
# hqAwDQYJKoZIhvcNAQEFBQADgYEAeuXfuL0QLHxYO7YhZtJgCu7eV714lFpju7v4
# pcXAdTOtANjxYqcBZcGi+44otgXjfCwBFQDGddCGnFaQi8GNQTpFoFsGEdS5K8tv
# /7neU8f/4sW1Tr01DAX/AbaohCMEoMiYOYdKPwdHdKPkIfEbaYeoowJ2o8gsIzSf
# JQ7B9PsxggQNMIIECQIBATCBkzB8MQswCQYDVQQGEwJVUzETMBEGA1UECBMKV2Fz
# aGluZ3RvbjEQMA4GA1UEBxMHUmVkbW9uZDEeMBwGA1UEChMVTWljcm9zb2Z0IENv
# cnBvcmF0aW9uMSYwJAYDVQQDEx1NaWNyb3NvZnQgVGltZS1TdGFtcCBQQ0EgMjAx
# MAITMwAAAbWtGt/XhXBtEwABAAABtTANBglghkgBZQMEAgEFAKCCAUowGgYJKoZI
# hvcNAQkDMQ0GCyqGSIb3DQEJEAEEMC8GCSqGSIb3DQEJBDEiBCDprcLhYN5lXPyr
# VOK1Rth/AhKp+F/nsTJ2TaCS3hyQBTCB+gYLKoZIhvcNAQkQAi8xgeowgecwgeQw
# gb0EICfKDTUtaGcWifYc3OVnIpp7Ykn0S8JclVzrlAgF8ciDMIGYMIGApH4wfDEL
# MAkGA1UEBhMCVVMxEzARBgNVBAgTCldhc2hpbmd0b24xEDAOBgNVBAcTB1JlZG1v
# bmQxHjAcBgNVBAoTFU1pY3Jvc29mdCBDb3Jwb3JhdGlvbjEmMCQGA1UEAxMdTWlj
# cm9zb2Z0IFRpbWUtU3RhbXAgUENBIDIwMTACEzMAAAG1rRrf14VwbRMAAQAAAbUw
# IgQgRSAF/7ttMBkDwcw2bwop4lWkIygdMAEthRbbRiVEqXowDQYJKoZIhvcNAQEL
# BQAEggIAfHzcVofwpujCxGGmi5wVlqTX65GBmOn8aaWleNCpreT5virH6JwJP+yp
# xV9EAxSW4FWxHd4/mp912JVkWjA2btDSsLLfygq0IAKym/wStYRnD9NOpsvH7QLD
# Ira0Jm3Ng/RDP61YqJ3rOhVfGsWZHQTtE7K9GRpqHZow3l2TOBNBDz0Biii7cs6A
# qEbVtwHp/w1y2NwJxqOekNtHOzb6C/TvIIEgWo+Xu0RXONCRRKryQ31OXDjHX3Qx
# /N4xyvlsNL1IPX30+liwFPUl0Ur8Fancc+oYhP59ea0A9iYKdJiOF+KP5m9/OwDT
# jatnfIsGhDiCSe1TUgXF8afYmijh7lYDW9jQ9C47tCQ4txVpQ2jiG86rHReSpuos
# V6WAE4M4AWGY7A7oaaDlxUum70f1eM1+YrlBh/TWbcpe2jEmgPoEorblx7LxhXoR
# aj6dc+rIuSN49mjccOiGTwoKw1nHCf7WsPwpkhDa4zuwxlsYqM04OnJ9/Aif9o/9
# PQpggyt0xjyxn8cmGdiR7jE+XU36lZ3RdgSHRS2SeBuA9PmXbgxZS0Y/uMtRf9Nb
# TtJYyzgrPi391ZijNlD4r26hLe6HGKsJKUI14L2rXabqmGmGuyfXumRqri+Vm+HJ
# HvOoqt0poYVCQSUid42v8iHN//SPkuJVLiJw3TW0FaxXvYZTiTM=
# SIG # End signature block
