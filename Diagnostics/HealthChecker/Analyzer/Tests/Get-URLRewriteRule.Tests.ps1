# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSAvoidUsingInvokeExpression', '', Justification = 'Pester testing file')]
[CmdletBinding()]
param()

BeforeAll {
    . $PSScriptRoot\..\..\..\..\Shared\PesterLoadFunctions.NotPublished.ps1
    $scriptContent = Get-PesterScriptContent -FilePath "$PSScriptRoot\..\Get-URLRewriteRule.ps1"
    Invoke-Expression $scriptContent
    function Invoke-CatchActions { throw "Called Invoke-CatchActions" }

    $Script:mockDataRoot = "$PSScriptRoot\..\..\Tests\DataCollection\E19\Exchange\IIS"
    [xml]$Script:appHost = Get-Content "$Script:mockDataRoot\applicationHost.config" -Raw -Encoding UTF8

    $Script:webConfigContent = @{
        "Default Web Site"      = (Get-Content "$Script:mockDataRoot\DefaultWebSite_web.config" -Raw -Encoding UTF8)
        "Default Web Site/owa"  = (Get-Content "$Script:mockDataRoot\DefaultWebSite-OWA_web.config" -Raw -Encoding UTF8)
        "Default Web Site/mapi" = (Get-Content "$Script:mockDataRoot\DefaultWebSite-MAPI_web.config" -Raw -Encoding UTF8)
        "Default Web Site/EWS"  = (Get-Content "$Script:mockDataRoot\DefaultWebSite-EWS_web.config" -Raw -Encoding UTF8)
    }

    $Script:result = Get-URLRewriteRule -ApplicationHostConfig $Script:appHost -WebConfigContent $Script:webConfigContent
}

Describe "Get-URLRewriteRule" {

    Context "Rule extraction from web.config" {

        It "Should find inbound rule from Default Web Site web.config" {
            $siteRules = $Script:result["Default Web Site"]
            $allRuleNames = @($siteRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Contain "CVE-2022-41040 Mitigation"
        }

        It "Should collect remove entry from MAPI web.config" {
            $mapiRules = $Script:result["Default Web Site/mapi"]
            # First entry is from web.config which contains the <remove> element
            $removeNames = @($mapiRules[0].remove.name)
            $removeNames | Should -Contain "Global Block Bad User Agents"
        }
    }

    Context "Rule extraction from applicationHost.config per-location" {

        It "Should find disabled inbound rule from appHost Default Web Site location" {
            $siteRules = $Script:result["Default Web Site"]
            $allRuleNames = @($siteRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Contain "Disable HTTP - Redirect to HTTPS"
        }

        It "Should preserve disabled attribute on appHost rule" {
            $siteRules = $Script:result["Default Web Site"]
            $disabledRule = $siteRules | ForEach-Object { $_.rule } |
                Where-Object { $_.name -eq "Disable HTTP - Redirect to HTTPS" }
            $disabledRule.enabled | Should -Be "false"
        }
    }

    Context "Rule extraction from applicationHost.config global section" {

        It "Should find global inbound rule" {
            $siteRules = $Script:result["Default Web Site"]
            $allRuleNames = @($siteRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Contain "Global Block Bad User Agents"
        }
    }

    Context "Inheritance walk-up" {

        It "Should collect rules from all 3 levels for Default Web Site" {
            # web.config (CVE-2022-41040) + appHost location (disabled HTTPS redirect) + global (Block Bad User Agents)
            $Script:result["Default Web Site"].Count | Should -Be 3
        }

        It "Should inherit parent and global rules for Default Web Site/owa" {
            # OWA web.config has no inbound rules (only outbound which the function cannot see yet)
            # OWA appHost location has no inbound rules (only outbound)
            # Walks up to Default Web Site: web.config has CVE-2022-41040, appHost has disabled HTTPS redirect
            # Then global has Block Bad User Agents
            $owaRules = $Script:result["Default Web Site/owa"]
            $allRuleNames = @($owaRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Contain "CVE-2022-41040 Mitigation"
            $allRuleNames | Should -Contain "Disable HTTP - Redirect to HTTPS"
            $allRuleNames | Should -Contain "Global Block Bad User Agents"
        }

        It "Should inherit rules for vDir with no rewrite config" {
            # EWS web.config has no rewrite section at all
            # Should inherit from parent Default Web Site and global
            $ewsRules = $Script:result["Default Web Site/EWS"]
            $allRuleNames = @($ewsRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Contain "CVE-2022-41040 Mitigation"
            $allRuleNames | Should -Contain "Global Block Bad User Agents"
        }
    }

    Context "Clear stops inheritance" {

        It "Should stop at clear in appHost location for Default Web Site/mapi" {
            # MAPI web.config has <remove> (collected but no clear)
            # MAPI appHost location has <clear/> which stops inheritance
            # Should NOT contain parent Default Web Site rules or global rules
            $mapiRules = $Script:result["Default Web Site/mapi"]
            $mapiRules.Count | Should -Be 2
            $allRuleNames = @($mapiRules.rule.name | Where-Object { $null -ne $_ })
            $allRuleNames | Should -Not -Contain "CVE-2022-41040 Mitigation"
            $allRuleNames | Should -Not -Contain "Global Block Bad User Agents"
        }

        It "Should stop at clear in web.config and not check appHost or parent" {
            # Edge case: clear in web.config is a different code path (line 52) than clear in appHost (line 75)
            # Our mock data only has clear in appHost, so use inline XML for this specific path
            $clearWebConfig = '<configuration><system.webServer><rewrite><rules><clear /></rules></rewrite></system.webServer></configuration>'
            $emptyWebConfig = '<configuration><system.webServer><modules /></system.webServer></configuration>'
            $webConfigs = @{
                "Default Web Site/owa" = $clearWebConfig
                "Default Web Site"     = $emptyWebConfig
            }

            $result = Get-URLRewriteRule -ApplicationHostConfig $Script:appHost -WebConfigContent $webConfigs

            # Should only have 1 entry (the clear node from web.config) - nothing from appHost or parent
            $result["Default Web Site/owa"].Count | Should -Be 1
            $null -ne $result["Default Web Site/owa"][0].clear | Should -Be $true
        }
    }

    Context "Empty rewrite sections" {

        It "Should return empty list when no rewrite rules exist at any level" {
            $emptyWebConfig = '<configuration><system.webServer><modules /></system.webServer></configuration>'
            [xml]$emptyAppHost = @"
<?xml version="1.0" encoding="UTF-8"?>
<configuration>
    <system.webServer />
    <location path="Default Web Site/test">
        <system.webServer />
    </location>
</configuration>
"@
            $webConfigs = @{
                "Default Web Site/test" = $emptyWebConfig
            }

            $result = Get-URLRewriteRule -ApplicationHostConfig $emptyAppHost -WebConfigContent $webConfigs

            $result.ContainsKey("Default Web Site/test") | Should -Be $true
            $result["Default Web Site/test"].Count | Should -Be 0
        }
    }
}
