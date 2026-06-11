# Copyright (c) Microsoft Corporation.
# Licensed under the MIT License.

# Generic testing for Exchange SE
[CmdletBinding()]
param()

Describe "Testing Health Checker by Mock Data Imports - Exchange SE" {

    BeforeAll {
        . $PSScriptRoot\HealthCheckerTests.ImportCode.NotPublished.ps1
        $Script:MockDataCollectionRoot = "$Script:parentPath\Tests\DataCollection\ExchangeSE"
        . $PSScriptRoot\HealthCheckerTest.CommonMocks.NotPublished.ps1
    }

    Context "Basic Exchange SE RTM Testing HyperV" {
        BeforeAll {
            # Windows Server 2025 ships with .NET 4.8.1 - override the CommonMocks default of 4.8
            Mock Get-NETFrameworkVersion {
                return [PSCustomObject]@{
                    FriendlyName  = "4.8.1"
                    RegistryValue = 533320
                    MinimumValue  = 533320
                }
            }

            SetDefaultRunOfHealthChecker "Debug_SE_HyperV_Results.xml"
        }

        It "Display Results - Exchange Information" {
            SetActiveDisplayGrouping "Exchange Information"

            TestObjectMatch "Name" "CONTOSO-EX1"
            TestObjectMatch "Version" "Exchange SE RTM"
            TestObjectMatch "Build Number" "15.02.2562.017"
            TestObjectMatch "Known Issue Detected" $true -WriteType "Yellow"
            TestObjectMatch "Server Role" "Mailbox"
            TestObjectMatch "Edition" "Warning - StandardEvaluation" -WriteType "Yellow"
            TestObjectMatch "Remaining Trial Period" "179.20:41:57.0813653"
            TestObjectMatch "DAG Name" "Standalone Server"
            TestObjectMatch "AD Site" "Default-First-Site-Name"
            TestObjectMatch "MRS Proxy Enabled" "False"
            TestObjectMatch "Exchange Server Maintenance" "Server is not in Maintenance Mode" -WriteType "Green"
            TestObjectMatch "Internet Web Proxy" "Not Set"
            TestObjectMatch "Extended Protection Enabled (Any VDir)" $true
            # Note: Baseline data was collected after scenario setup, so setting overrides are present
            TestObjectMatch "Setting Overrides Detected" $true
            TestObjectMatch "Monitoring Overrides Detected" $false
            TestObjectMatch "Exchange Server Membership" "Passed"
            TestObjectMatch "Exchange Server Token Groups" 6
            TestObjectMatch "Ring Level" 1
            TestObjectMatch "Features Enabled" "None Enabled"
            $Script:ActiveGrouping.Count | Should -Be 27
        }

        It "Display Results - Organization Information" {
            SetActiveDisplayGrouping "Organization Information"

            TestObjectMatch "MAPI/HTTP Enabled" "True"
            TestObjectMatch "Enable Download Domains" "False"
            TestObjectMatch "AD Split Permissions" "False"
            TestObjectMatch "Dynamic Distribution Group Public Folder Mailboxes Count" 1 -WriteType "Green"

            $Script:ActiveGrouping.Count | Should -Be 5
        }

        It "Display Results - Operating System Information" {
            SetActiveDisplayGrouping "Operating System Information"

            TestObjectMatch "Product Name" "Windows Server 2025 Datacenter"
            # Note: The display version comes from shared registry mocks in CommonMocks (ReleaseID=2009, CurrentBuild=26100, UBR=720).
            # These values happen to be correct for SE (Windows Server 2025) but are also used by E19/E16 tests.
            # All OS-version-dependent logic uses BuildVersion from Win32_OperatingSystem.xml (correctly per-version), not these registry values.
            # TODO: Once E19 data is removed, update CommonMocks registry values to be version-specific or sourced from mock data files.
            TestObjectMatch "Version" "2009 (OS Build: 26100.720)"
            TestObjectMatch "Time Zone" "Pacific Standard Time"
            TestObjectMatch "Dynamic Daylight Time Enabled" "True"
            TestObjectMatch ".NET Framework" "4.8.1" -WriteType "Green"
            TestObjectMatch "Power Plan" "Balanced --- Error" -WriteType "Red"
            $httpProxy = GetObject "Http Proxy Setting"
            $httpProxy.ProxyAddress | Should -Be "None"
            TestObjectMatch "Visual C++ 2012 x64" "11.0.61030 Version is current" -WriteType "Green"
            TestObjectMatch "Visual C++ 2013 x64" "12.0.40664 Version is current" -WriteType "Green"
            TestObjectMatch "Server Pending Reboot" "False"

            $pageFile = GetObject "PageFile Size 0"
            $pageFile.Name | Should -Be ""
            $pageFile.TotalPhysicalMemory | Should -Be 6144
            $pageFile.MaxPageSize | Should -Be 0
            $pageFile.MultiPageFile | Should -Be $false
            $pageFile.RecommendedPageFile | Should -Be 1536

            $pageFileAdditional = GetObject "PageFile Additional Information"
            $pageFileAdditional | Should -Be "Error: On Exchange SE RTM, the recommended PageFile size is 25% (1536MB) of the total system memory (6144MB)."
            $Script:ActiveGrouping.Count | Should -Be 14
        }

        It "Display Results - Process/Hardware Information" {
            SetActiveDisplayGrouping "Processor/Hardware Information"

            TestObjectMatch "Type" "HyperV"
            TestObjectMatch "Processor" "Intel(R) Xeon(R) CPU E5-2430 0 @ 2.20GHz"
            TestObjectMatch "Number of Processors" 1
            TestObjectMatch "Number of Physical Cores" 2 -WriteType "Green"
            TestObjectMatch "Number of Logical Cores" 4 -WriteType "Green"
            TestObjectMatch "All Processor Cores Visible" "Passed" -WriteType "Green"
            TestObjectMatch "Max Processor Speed" 2200
            TestObjectMatch "Physical Memory" 6 -WriteType "Yellow"
            TestObjectMatch "Dynamic Memory Detected" $false -WriteType "Green"

            $Script:ActiveGrouping.Count | Should -Be 11
        }

        It "Display Results - NIC Settings" {
            SetActiveDisplayGrouping "NIC Settings Per Active Adapter"

            TestObjectMatch "Interface Description" "Microsoft Hyper-V Network Adapter [Ethernet]"
            TestObjectMatch "Driver Date" "2006-06-21"
            TestObjectMatch "MTU Size" 1500
            TestObjectMatch "Max Processors" 2
            TestObjectMatch "Max Processor Number" 2
            TestObjectMatch "Number of Receive Queues" 2
            TestObjectMatch "RSS Enabled" "True" -WriteType "Green"
            TestObjectMatch "Link Speed" "10000 Mbps"
            TestObjectMatch "IPv6 Enabled" "True"
            TestObjectMatch "Address" "192.168.13.11/24 Gateway: 192.168.13.1"
            TestObjectMatch "Registered In DNS" "True"
            TestObjectMatch "Packets Received Discarded" 0 -WriteType "Green"

            $Script:ActiveGrouping.Count | Should -Be 17
        }

        It "Display Results - Frequent Configuration Issues" {
            SetActiveDisplayGrouping "Frequent Configuration Issues"

            TestObjectMatch "TCP/IP Settings" 90000 -WriteType "Yellow"
            TestObjectMatch "RPC Min Connection Timeout" 0
            TestObjectMatch "FIPS Algorithm Policy Enabled" 0
            TestObjectMatch "EnableEccCertificateSupport Registry Value" $false
            TestObjectMatch "CTS Processor Affinity Percentage" 0 -WriteType "Green"
            TestObjectMatch "Disable Async Notification" $false
            TestObjectMatch "Credential Guard Enabled" $false
            TestObjectMatch "Trusted Root Certificates Auto Update Disabled" $false -WriteType "Green"
            TestObjectMatch "EdgeTransport.exe.config Present" "True" -WriteType "Green"
            TestObjectMatch "NodeRunner.exe memory limit" "0 MB" -WriteType "Green"
            # Baseline data includes wildcard '*' InternalRelay accepted domain from lab setup
            TestObjectMatch "Open Relay Wild Card Domain" "Error --- Accepted Domain `"Problem Accepted Domain`" is set to a Wild Card (*) Domain Name with a domain type of InternalRelay. This is not recommended as this is an open relay for the entire environment.`r`n`t`tMore Information: https://aka.ms/HC-OpenRelayDomain" -WriteType "Red"
            # Now in the default data set
            TestObjectMatch "DisablePreservation" 0
            TestObjectMatch "EXO Connector Present" "True"
            # Baseline send connector created without RequireTLS/TlsAuthLevel triggers additional warnings
            $sendConnectorDetails = GetObject "Send Connector - Mail to O365"
            $sendConnectorDetails | Should -Contain "TlsAuthLevel not set to CertificateValidation or DomainValidation"
            TestObjectMatch "UnifiedContent Auto Cleanup Configured" $true -WriteType "Green"

            $Script:ActiveGrouping.Count | Should -Be 18
        }

        It "Display Results - Security Settings" {
            SetActiveDisplayGrouping "Security Settings"

            # TLS configuration - SE on Win2025 supports TLS 1.3
            TestObjectMatch "TLS 1.0" "Disabled" -WriteType "Green"
            TestObjectMatch "TLS 1.1" "Disabled" -WriteType "Green"
            TestObjectMatch "TLS 1.2" "Enabled" -WriteType "Green"
            TestObjectMatch "TLS 1.3" "Enabled" -WriteType "Green"
            TestObjectMatch "SecurityProtocol" "SystemDefault"
            TestObjectMatch "AllowInsecureRenegoClients Value" 0
            TestObjectMatch "AllowInsecureRenegoServers Value" 0

            TestObjectMatch "LmCompatibilityLevel Settings" 3
            TestObjectMatch "SMB1 Installed" "False" -WriteType "Green"
            TestObjectMatch "SMB1 Blocked" "True" -WriteType "Green"
            TestObjectMatch "Exchange Emergency Mitigation Service" "Enabled" -WriteType "Green"
            TestObjectMatch "Windows service" "Running"
            TestObjectMatch "Pattern service" "200 - Reachable"
            TestObjectMatch "Telemetry enabled" $true
            # This is disabled by default due to our overrides being included by default
            TestObjectMatch "AMSI Enabled" "false" -WriteType "Yellow"
            TestObjectMatch "AMSI Request Body Scanning" "False" -WriteType "Yellow"
            TestObjectMatch "AMSI Request Body Size Block" "False"
            TestObjectMatch "Strict Mode disabled" "False" -WriteType "Green"
            TestObjectMatch "BaseTypeCheckForDeserialization disabled" "False" -WriteType "Green"
            TestObjectMatch "Valid Internal Transport Certificate Found On Server" "True" -WriteType "Green"
            TestObjectMatch "AES256-CBC Protected Content Support" "Supported Build and Valid Configuration" -WriteType "Green"
            TestObjectMatch "SerializedDataSigning Enabled" "True" -WriteType "Green"
            # Auth Certificate (CN=Microsoft Exchange Server Auth Certificate) expires 05/15/2031.
            # If this mock data is still in use after that date, this assertion will need to change to "False"
            # and an expired warning line will appear.
            TestObjectMatch "Valid Auth Certificate Found On Server" "True" -WriteType "Green"

            $Script:ActiveGrouping.Count | Should -Be 80
        }

        It "Display Results - Security Vulnerability" {
            SetActiveDisplayGrouping "Security Vulnerability"

            $cveTests = GetObject "Security Vulnerability"
            $cveTests.Count | Should -Be 19

            $downloadDomains = GetObject "CVE-2021-1730"
            $downloadDomains.DownloadDomainsEnabled | Should -Be "False"
        }

        It "Display Results - Exchange IIS Information" {
            SetActiveDisplayGrouping "Exchange IIS Information"

            $tokenCacheModuleInformation = GetObject "TokenCacheModule loaded"
            $tokenCacheModuleInformation | Should -Be $null

            # Verify inbound URL rewrite rules are displayed (deduplicated across vDirs)
            $inboundRules = GetObject "Inbound URL Rewrite Rules"
            $inboundRules.Count | Should -Be 3
            $inboundRuleNames = $inboundRules.RewriteRuleName.Value
            $inboundRuleNames | Should -Contain "CVE-2022-41040 Mitigation"
            $inboundRuleNames | Should -Contain "Global Block Bad User Agents"
            $inboundRuleNames | Should -Contain "Negate Match Test Rule"

            # Verify rules excluded by <remove> in DWS web.config do not appear in detailed display
            $inboundRuleNames | Should -Not -Contain "AppHost Only Rule"

            # Verify match property resolves correctly when <match> has extra attributes like negate
            $negateRule = $inboundRules | Where-Object { $_.RewriteRuleName.Value -eq "Negate Match Test Rule" }
            $negateRule.MatchProperty.Value | Should -Be "url - .*"

            # Verify outbound URL rewrite rules are displayed (deduplicated across vDirs)
            $outboundRules = GetObject "Outbound URL Rewrite Rules"
            $outboundRules.Count | Should -Be 2
            $outboundRuleNames = $outboundRules.RewriteRuleName.Value
            $outboundRuleNames | Should -Contain "EOMT OWA CSP - outbound"
            $outboundRuleNames | Should -Contain "AppHost OWA Outbound Test"

            # Verify global IIS rewrite rules warning is displayed
            $globalRulesWarning = GetObject "Global IIS Rewrite Rules"
            $globalRulesWarning | Should -Not -BeNullOrEmpty
        }
    }

    Context "GetHtmlTextValue Unit Tests" {

        It "Should return null for null input" {
            $result = GetHtmlTextValue -OriginalValue $null
            $result | Should -BeNullOrEmpty
        }

        It "Should return empty string for empty input" {
            $result = GetHtmlTextValue -OriginalValue ""
            $result | Should -Be ""
        }

        It "Should return plain text unchanged" {
            $result = GetHtmlTextValue -OriginalValue "Exchange 2019 CU11"
            $result | Should -Be "Exchange 2019 CU11"
        }

        It "Should encode angle brackets for certificate SAN values" {
            $result = GetHtmlTextValue -OriginalValue "<SAN>CN=mail.contoso.com</SAN>"
            $result | Should -Be "&lt;SAN&gt;CN=mail.contoso.com&lt;/SAN&gt;"
        }

        It "Should encode greater-than sign" {
            $result = GetHtmlTextValue -OriginalValue "Value > 100"
            $result | Should -Be "Value &gt; 100"
        }

        It "Should encode less-than sign" {
            $result = GetHtmlTextValue -OriginalValue "Value < 100"
            $result | Should -Be "Value &lt; 100"
        }

        It "Should handle mixed content with angle brackets and normal text" {
            $result = GetHtmlTextValue -OriginalValue "Status: <Unknown> - Check docs"
            $result | Should -Be "Status: &lt;Unknown&gt; - Check docs"
        }

        It "Should preserve intentional br tags after encoding" {
            $testValue = "CVE-2020-1147<br>CVE-2023-36434<br>"
            $result = GetHtmlTextValue -OriginalValue $testValue
            $result | Should -Be "CVE-2020-1147<br>CVE-2023-36434<br>"
        }

        It "Should convert URLs to clickable hyperlinks" {
            $result = GetHtmlTextValue -OriginalValue "More Information: https://aka.ms/HC-ExBuilds"
            # cspell:ignore noopener noreferrer
            $result | Should -BeLike '*<a href="https://aka.ms/HC-ExBuilds"*>https://aka.ms/HC-ExBuilds</a>'
        }

        It "Should convert URLs with trailing sentence punctuation" {
            $result = GetHtmlTextValue -OriginalValue "See: https://portal.msrc.microsoft.com/security-guidance/advisory/CVE-2020-1147 for more information."
            $result | Should -BeLike '*<a href=*>https://portal.msrc.microsoft.com/security-guidance/advisory/CVE-2020-1147</a> for more information.'
        }

        It "Should handle br tags combined with URLs in security vulnerability summary" {
            $testValue = "CVE-2020-1147`r`n`t`tSee: https://portal.msrc.microsoft.com/security-guidance/advisory/CVE-2020-1147 for more information.<br>"
            $result = GetHtmlTextValue -OriginalValue $testValue
            $result | Should -BeLike "*<br>*"
            $result | Should -Not -BeLike "*&lt;br&gt;*"
            $result | Should -BeLike "*<a href=*>*</a>*"
        }
    }

    Context "Testing Throws" {
        BeforeAll {
            Mock Get-MailboxServer { throw "Pester testing" }

            SetDefaultRunOfHealthChecker "Debug_SE_TestingThrow_Results.xml"
        }

        It "Verify we still analyze the data from throw Get-MailboxServer" {
            SetActiveDisplayGrouping "Exchange Information"
            TestObjectMatch "DAG Name" "Standalone Server"
        }
    }
}
