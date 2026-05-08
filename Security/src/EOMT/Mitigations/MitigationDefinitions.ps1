# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

<#
.DESCRIPTION
    Registry of all CVE mitigation definitions for EOMT.
    Provides a function to retrieve a mitigation definition in a standardized format
    so the orchestrator remains CVE-agnostic.

    Each CVE definition must return a PSCustomObject with these required properties:

    [string]   Id                 - The CVE identifier (e.g., "CVE-2022-41040")
    [int]      Priority           - Lower number = higher priority. The lowest Priority CVE
                                    is presented as the default selection when -CVE is not specified.
    [string]   Description        - Human-readable description of the vulnerability and mitigation
    [bool]     RequiresUrlRewrite - Whether the IIS URL Rewrite Module must be installed
    [string]   SiteName           - The IIS site where mitigations are applied (e.g., "Default Web Site")
    [ScriptBlock] TestMissingSecurityFix
                                    - Returns $true if the server's Exchange build is missing
                                      the security fix for this CVE. This checks the installed
                                      version/patch level only — it does not check whether
                                      the IIS mitigation is already applied.
                                      Should throw on unrecoverable errors (e.g., can't
                                      determine Exchange version).
    [ScriptBlock] GetActions      - Returns an array of PSCustomObject action definitions, each with:
                                        [string]    Cmdlet     - The IIS cmdlet to execute
                                        [hashtable] Parameters - Parameters to pass to the cmdlet
                                        [string]    RuleName   - (Required for Add-WebConfigurationProperty)
                                                                 The URL Rewrite rule name, used to build
                                                                 targeted Clear-WebConfiguration filter for rollback
#>

. $PSScriptRoot\CVE-2021-26855.ps1
. $PSScriptRoot\CVE-2022-41040.ps1

$script:MitigationDefinitionMap = @{
    "CVE-2021-26855" = { Get-CVE202126855-MitigationDefinition }
    "CVE-2022-41040" = { Get-CVE202241040-MitigationDefinition }
}

<#
.DESCRIPTION
    Returns the mitigation definition for the specified CVE.
#>
function Get-MitigationDefinition {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$CVE
    )

    if (-not $script:MitigationDefinitionMap.ContainsKey($CVE)) {
        throw "Unknown CVE identifier: $CVE. Available: $($script:MitigationDefinitionMap.Keys -join ', ')"
    }

    return (& $script:MitigationDefinitionMap[$CVE])
}
