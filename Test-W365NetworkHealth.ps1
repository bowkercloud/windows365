#Requires -Version 5.1
<#
.SYNOPSIS
    Windows 365 & AVD Network Health Check

.DESCRIPTION
    Tests network connectivity to all endpoints required for Windows 365 Cloud PCs and Azure Virtual Desktop.
    Can be run from the Cloud PC (Host Network) or from the physical client device (Client Network).

    Endpoints are loaded from a companion Endpoints.csv file (recommended) or from built-in defaults.

    Run directly from GitHub:
    powershell -ExecutionPolicy Bypass -Command "irm https://raw.githubusercontent.com/bowkercloud/windows365/main/Test-W365NetworkHealth.ps1 | iex"

.PARAMETER Mode
    1 = Cloud PC / Host Network
    2 = Client Device / Client Network
    3 = Both

.PARAMETER EndpointsCSV
    Path to the companion Endpoints.csv file. If not provided, the script will attempt to download it
    from the same GitHub location, then fall back to built-in defaults.

.PARAMETER OutputPath
    Optional path to export results to a CSV file.

.EXAMPLE
    .\Test-W365NetworkHealth.ps1
    .\Test-W365NetworkHealth.ps1 -Mode 1 -OutputPath C:\Temp\results.csv
    .\Test-W365NetworkHealth.ps1 -Mode 3 -EndpointsCSV .\Endpoints.csv

.NOTES
    Version:    1.1
    Blog:       https://bowker.cloud
    References:
        https://learn.microsoft.com/en-us/windows-365/enterprise/requirements-network
        https://learn.microsoft.com/en-us/azure/virtual-desktop/required-fqdn-endpoint
        https://learn.microsoft.com/en-us/intune/intune-service/fundamentals/intune-endpoints
    Inspired by: https://gist.github.com/shannonfritz/4c9f1cf800f3406729a58417639736f3

    NOTE on Intune endpoints:
    Microsoft have deprecated the Office 365 Endpoint API (endpoints.office.com) for retrieving
    Intune FQDNs. Per Microsoft's own caution on the Intune endpoints page:
    "The previously available PowerShell scripts for retrieving Microsoft Intune endpoint IP addresses
    and FQDNs no longer return accurate data from the Office 365 Endpoint service."
    This script uses the static consolidated list from the Microsoft documentation instead.
#>

[CmdletBinding()]
param(
    [ValidateSet(1, 2, 3)]
    [int]$Mode = 0,
    [string]$EndpointsCSV = '',
    [string]$OutputPath = ''
)

# ─────────────────────────────────────────────────────────────────────────────
# CONFIGURATION
# ─────────────────────────────────────────────────────────────────────────────
$ScriptName     = 'Test-W365NetworkHealth'
$ScriptVersion  = 'v1.1'
$CSVGitHubURL   = 'https://raw.githubusercontent.com/bowkercloud/windows365/main/Endpoints.csv'
$TimeoutSeconds = 5

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Banner
# ─────────────────────────────────────────────────────────────────────────────
function Write-Banner {
    $banner = @"

   ___   ___  __    __  _  __ _____  ____       ____  _     ___  _   _ ____
  | __ )/ _ \| |  / / | |/ /| ____|  _ \     / ___|| |   / _ \| | | |  _ \
  |  _ \ | | | | /  / | ' / |  _|  | |_) |   | |    | |  | | | | | | | | | |
  | |_) | |_| | |/ /  | . \ | |___ |  _ <    | |___ | |__| |_| | |_| | |_| |
  |____/ \___/|_/_/   |_|\_\|_____||_| \_\    \____||_____\___/ \___/|____/
                                                           https://bowker.cloud

"@
    Write-Host $banner -ForegroundColor Cyan
    Write-Host "  $ScriptName  $ScriptVersion" -ForegroundColor Blue
    Write-Host "  Windows 365 & AVD Network Health Check" -ForegroundColor Blue
    Write-Host ""
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Test a single host:port
# ─────────────────────────────────────────────────────────────────────────────
function Test-Endpoint {
    param(
        [string]$Hostname,
        [int]$Port,
        [string]$Notes = '',
        [string]$Category = ''
    )

    $result = [PSCustomObject]@{
        Category  = $Category
        Hostname  = $Hostname
        Port      = $Port
        Status    = 'UNKNOWN'
        Notes     = $Notes
        Timestamp = (Get-Date -Format 'yyyy-MM-dd HH:mm:ss')
    }

    # Wildcard endpoints cannot be TCP-tested
    if ($Hostname -match '^\*') {
        $result.Status = 'WILDCARD'
        $pad = [string]::new(' ', [Math]::Max(0, 55 - ($Hostname.Length + $Port.ToString().Length)))
        Write-Host "  [" -NoNewline
        Write-Host "SKIP" -ForegroundColor DarkYellow -NoNewline
        Write-Host "] $Hostname`:$Port$pad (wildcard - verify DNS resolution manually)" -ForegroundColor DarkYellow
        return $result
    }

    # IP ranges (CIDR) and UDP-only entries cannot be meaningfully TCP-tested
    if ($Hostname -match '/\d+$') {
        $result.Status = 'IPRANGE'
        $pad = [string]::new(' ', [Math]::Max(0, 55 - ($Hostname.Length + $Port.ToString().Length)))
        Write-Host "  [" -NoNewline
        Write-Host "INFO" -ForegroundColor DarkCyan -NoNewline
        Write-Host "] $Hostname  UDP:$Port$pad (IP range - verify firewall/NSG allows UDP $Port outbound)" -ForegroundColor DarkCyan
        return $result
    }

    # time.windows.com is UDP/123 - flag rather than TCP test
    if ($Port -eq 123) {
        $result.Status = 'UDPONLY'
        $pad = [string]::new(' ', [Math]::Max(0, 55 - ($Hostname.Length + $Port.ToString().Length)))
        Write-Host "  [" -NoNewline
        Write-Host "INFO" -ForegroundColor DarkCyan -NoNewline
        Write-Host "] $Hostname  UDP:$Port$pad (UDP only - NTP; verify firewall allows UDP 123 outbound)" -ForegroundColor DarkCyan
        return $result
    }

    $pad = [string]::new(' ', [Math]::Max(0, 55 - ($Hostname.Length + $Port.ToString().Length)))
    Write-Host "  [    ] $Hostname`:$Port$pad" -NoNewline

    try {
        $tcpClient = New-Object System.Net.Sockets.TcpClient
        $connect   = $tcpClient.BeginConnect($Hostname, $Port, $null, $null)
        $wait      = $connect.AsyncWaitHandle.WaitOne([TimeSpan]::FromSeconds($TimeoutSeconds), $false)

        if ($wait -and -not $tcpClient.Client.Poll(0, [System.Net.Sockets.SelectMode]::SelectError)) {
            $tcpClient.EndConnect($connect) 2>$null
            $result.Status = 'OK'
            Write-Host "`r  [" -NoNewline
            Write-Host " OK " -ForegroundColor Green -NoNewline
            Write-Host "] $Hostname`:$Port$pad"
        } else {
            $result.Status = 'FAIL'
            Write-Host "`r  [" -NoNewline
            Write-Host "FAIL" -ForegroundColor Red -NoNewline
            Write-Host "] $Hostname`:$Port$pad"
        }
        $tcpClient.Close()
    } catch {
        $result.Status = 'FAIL'
        Write-Host "`r  [" -NoNewline
        Write-Host "FAIL" -ForegroundColor Red -NoNewline
        Write-Host "] $Hostname`:$Port  ($_)"
    }

    return $result
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Test a list of endpoints from CSV data
# ─────────────────────────────────────────────────────────────────────────────
function Test-EndpointList {
    param(
        [array]$Endpoints,
        [string]$FilterMode   # 'CloudPC', 'Client', or 'Both'
    )

    $allResults = @()
    $categories = $Endpoints | Select-Object -ExpandProperty Category -Unique

    foreach ($cat in $categories) {
        $catEndpoints = $Endpoints | Where-Object { $_.Category -eq $cat }
        Write-Host ""
        Write-Host "  ── $cat ──" -ForegroundColor Cyan

        foreach ($ep in $catEndpoints) {
            $testMode = $ep.TestMode.Trim()
            if ($FilterMode -eq 'CloudPC' -and $testMode -eq 'Client')  { continue }
            if ($FilterMode -eq 'Client'  -and $testMode -eq 'CloudPC') { continue }

            $ports = $ep.Port -split ','
            foreach ($port in $ports) {
                $portNum = [int]$port.Trim()
                $r = Test-Endpoint -Hostname $ep.Endpoint.Trim() -Port $portNum -Notes $ep.Notes -Category $ep.Category
                $allResults += $r
            }
        }
    }

    return $allResults
}

# ─────────────────────────────────────────────────────────────────────────────
# HELPER: Print summary
# ─────────────────────────────────────────────────────────────────────────────
function Write-Summary {
    param([array]$Results)

    $ok       = ($Results | Where-Object { $_.Status -eq 'OK'      }).Count
    $fail     = ($Results | Where-Object { $_.Status -eq 'FAIL'    }).Count
    $wildcard = ($Results | Where-Object { $_.Status -eq 'WILDCARD'}).Count
    $udponly  = ($Results | Where-Object { $_.Status -in 'IPRANGE','UDPONLY' }).Count
    $total    = $Results.Count

    Write-Host ""
    Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  RESULTS SUMMARY" -ForegroundColor White
    Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray
    Write-Host "  Total entries  : $total"
    Write-Host "  OK             : $ok"       -ForegroundColor Green
    Write-Host "  FAILED         : $fail"     -ForegroundColor $(if ($fail -gt 0) { 'Red' } else { 'Green' })
    Write-Host "  Wildcards      : $wildcard  (verify DNS resolution manually)"         -ForegroundColor DarkYellow
    Write-Host "  UDP/IP Ranges  : $udponly   (verify firewall/NSG allows UDP outbound)" -ForegroundColor DarkCyan
    Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray

    if ($fail -gt 0) {
        Write-Host ""
        Write-Host "  FAILED ENDPOINTS:" -ForegroundColor Red
        $Results | Where-Object { $_.Status -eq 'FAIL' } | ForEach-Object {
            Write-Host "    $($_.Hostname):$($_.Port)  [$($_.Category)]" -ForegroundColor Red
        }
    }

    $wildcardItems = $Results | Where-Object { $_.Status -eq 'WILDCARD' }
    if ($wildcardItems) {
        Write-Host ""
        Write-Host "  WILDCARD ENDPOINTS (verify DNS resolution manually):" -ForegroundColor DarkYellow
        $wildcardItems | ForEach-Object {
            Write-Host "    $($_.Hostname)  [$($_.Category)]" -ForegroundColor DarkYellow
        }
    }

    $udpItems = $Results | Where-Object { $_.Status -in 'IPRANGE','UDPONLY' }
    if ($udpItems) {
        Write-Host ""
        Write-Host "  UDP / IP RANGE ENTRIES (verify firewall/NSG allows outbound UDP):" -ForegroundColor DarkCyan
        $udpItems | ForEach-Object {
            Write-Host "    $($_.Hostname)  UDP:$($_.Port)  [$($_.Category)]  - $($_.Notes)" -ForegroundColor DarkCyan
        }
    }
}

# ─────────────────────────────────────────────────────────────────────────────
# BUILT-IN FALLBACK ENDPOINT LIST
# Source: Microsoft documentation (hardcoded - do NOT use endpoints.office.com
# API for Intune as Microsoft have deprecated it for this purpose)
# ─────────────────────────────────────────────────────────────────────────────
function Get-BuiltInEndpoints {
    return @(
        # ── Windows 365 Service ──────────────────────────────────────────────
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='Registration';      Endpoint='login.microsoftonline.com';               Port=443;        TestMode='Both';    Notes='Entra ID authentication' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='Registration';      Endpoint='login.live.com';                         Port=443;        TestMode='Both';    Notes='Microsoft account authentication' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='Registration';      Endpoint='enterpriseregistration.windows.net';      Port=443;        TestMode='Both';    Notes='Device registration' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Provisioning';  Endpoint='global.azure-devices-provisioning.net';   Port='443,5671'; TestMode='CloudPC'; Notes='IoT Hub device provisioning' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-prod-prap01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='Asia Pacific' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-prod-prau01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='Australia' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-prod-preu01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='Europe' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-prod-prna01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='North America' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-prod-prna02.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='North America 2' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-2-prod-preu01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='Europe 2' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-2-prod-prna01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='North America 2b' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-3-prod-preu01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='Europe 3' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-3-prod-prna01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='North America 3' }
        [PSCustomObject]@{ Category='W365-CloudPC'; Subcategory='IoT Hubs';          Endpoint='hm-iot-in-4-prod-prna01.azure-devices.net'; Port='443,5671'; TestMode='CloudPC'; Notes='North America 4' }

        # ── AVD Session Host (required) ──────────────────────────────────────
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Core';             Endpoint='login.microsoftonline.com';               Port=443;  TestMode='CloudPC'; Notes='Authentication to Microsoft Online Services' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Core';             Endpoint='51.5.0.0/16';                             Port=3478; TestMode='CloudPC'; Notes='RDP Shortpath relayed connectivity (TURN/STUN). Service tag: WindowsVirtualDesktop' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Core';             Endpoint='catalogartifact.azureedge.net';            Port=443;  TestMode='CloudPC'; Notes='Azure Marketplace. Service tag: AzureFrontDoor.Frontend' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Core';             Endpoint='aka.ms';                                  Port=443;  TestMode='CloudPC'; Notes='Microsoft URL shortener' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Monitoring';       Endpoint='gcs.prod.monitoring.core.windows.net';     Port=443;  TestMode='CloudPC'; Notes='AVD agent traffic. Service tag: AzureMonitor' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Activation';       Endpoint='azkms.core.windows.net';                  Port=1688; TestMode='CloudPC'; Notes='Windows KMS activation' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Updates';          Endpoint='mrsglobalsteus2prod.blob.core.windows.net'; Port=443; TestMode='CloudPC'; Notes='AVD agent and SXS stack updates' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Portal';           Endpoint='wvdportalstorageblob.blob.core.windows.net'; Port=443; TestMode='CloudPC'; Notes='Azure portal support' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Azure';            Endpoint='169.254.169.254';                         Port=80;   TestMode='CloudPC'; Notes='Azure Instance Metadata Service (IMDS)' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Azure';            Endpoint='168.63.129.16';                           Port=80;   TestMode='CloudPC'; Notes='Session host health monitoring' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Certificates';     Endpoint='oneocsp.microsoft.com';                   Port=80;   TestMode='CloudPC'; Notes='OCSP certificate validation' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Certificates';     Endpoint='www.microsoft.com';                       Port=80;   TestMode='CloudPC'; Notes='Certificate chain' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Certificates';     Endpoint='azcsprodeusaikpublish.blob.core.windows.net'; Port=80; TestMode='CloudPC'; Notes='AIK certificate publishing' }
        [PSCustomObject]@{ Category='AVD-SessionHost'; Subcategory='Certificates';     Endpoint='ctldl.windowsupdate.com';                 Port=80;   TestMode='CloudPC'; Notes='Certificate Trust List download' }

        # ── AVD Session Host (optional) ──────────────────────────────────────
        [PSCustomObject]@{ Category='AVD-SessionHost-Optional'; Subcategory='Auth';    Endpoint='login.windows.net';                       Port=443;  TestMode='CloudPC'; Notes='Sign in to Microsoft Online Services and Microsoft 365' }
        [PSCustomObject]@{ Category='AVD-SessionHost-Optional'; Subcategory='Connectivity'; Endpoint='www.msftconnecttest.com';            Port=80;   TestMode='CloudPC'; Notes='Internet connectivity detection' }

        # ── Client / End User Device ─────────────────────────────────────────
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Auth';               Endpoint='login.microsoftonline.com';               Port=443;  TestMode='Client'; Notes='Authentication to Microsoft Online Services' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Navigation';         Endpoint='go.microsoft.com';                        Port=443;  TestMode='Client'; Notes='Microsoft FWLinks' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Navigation';         Endpoint='aka.ms';                                  Port=443;  TestMode='Client'; Notes='Microsoft URL shortener' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Docs';               Endpoint='learn.microsoft.com';                     Port=443;  TestMode='Client'; Notes='Microsoft documentation' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Legal';              Endpoint='privacy.microsoft.com';                   Port=443;  TestMode='Client'; Notes='Microsoft privacy statement' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Service';            Endpoint='graph.microsoft.com';                     Port=443;  TestMode='Client'; Notes='Microsoft Graph API' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Portal';             Endpoint='windows.cloud.microsoft';                 Port=443;  TestMode='Client'; Notes='Connection center' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Portal';             Endpoint='windows365.microsoft.com';                Port=443;  TestMode='Client'; Notes='Windows 365 service traffic' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Portal';             Endpoint='ecs.office.com';                          Port=443;  TestMode='Client'; Notes='Connection center configuration' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Certificates';       Endpoint='www.microsoft.com';                       Port=80;   TestMode='Client'; Notes='Certificate chain' }
        [PSCustomObject]@{ Category='Client-AVD'; Subcategory='Certificates';       Endpoint='azcsprodeusaikpublish.blob.core.windows.net'; Port=80; TestMode='Client'; Notes='AIK certificate publishing' }

        # ── Client - Azure CA Certificate checks (closed network) ───────────
        # Source: https://learn.microsoft.com/en-us/azure/security/fundamentals/azure-certificate-authority-details
        # Note: oneocsp.microsoft.com and www.microsoft.com already covered above in Client-AVD certs
        [PSCustomObject]@{ Category='Client-AVD-CertCA'; Subcategory='Certificate Authority'; Endpoint='cacerts.digicert.com'; Port=80; TestMode='Client'; Notes='DigiCert CA certificate downloads (AIA)' }
        [PSCustomObject]@{ Category='Client-AVD-CertCA'; Subcategory='Certificate Authority'; Endpoint='crl.digicert.com';    Port=80; TestMode='Client'; Notes='DigiCert CRL distribution point' }
        [PSCustomObject]@{ Category='Client-AVD-CertCA'; Subcategory='Certificate Authority'; Endpoint='crl3.digicert.com';   Port=80; TestMode='Client'; Notes='DigiCert CRL distribution point' }
        [PSCustomObject]@{ Category='Client-AVD-CertCA'; Subcategory='Certificate Authority'; Endpoint='crl4.digicert.com';   Port=80; TestMode='Client'; Notes='DigiCert CRL distribution point' }
        [PSCustomObject]@{ Category='Client-AVD-CertCA'; Subcategory='Certificate Authority'; Endpoint='ocsp.digicert.com';   Port=80; TestMode='Client'; Notes='DigiCert OCSP responder' }

        # ── Intune Core Service ──────────────────────────────────────────────
        # NOTE: Static list per Microsoft documentation. The endpoints.office.com API
        # no longer returns accurate Intune data and should NOT be used.
        # Source: https://learn.microsoft.com/en-us/intune/intune-service/fundamentals/intune-endpoints
        [PSCustomObject]@{ Category='Intune'; Subcategory='Core Service';           Endpoint='manage.microsoft.com';                    Port=443;  TestMode='CloudPC'; Notes='Intune client and host service' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Core Service';           Endpoint='EnterpriseEnrollment.manage.microsoft.com'; Port=443; TestMode='CloudPC'; Notes='Intune enterprise enrollment' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swda01-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swda02-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdb01-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdb02-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdc01-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdc02-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdd01-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdd02-mscdn.manage.microsoft.com';        Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdin01-mscdn.manage.microsoft.com';       Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN (India)' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Win32 Apps';             Endpoint='swdin02-mscdn.manage.microsoft.com';       Port=443;  TestMode='CloudPC'; Notes='Win32 app CDN (India)' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='login.microsoftonline.com';               Port=443;  TestMode='CloudPC'; Notes='Authentication and Identity (Entra ID)' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='graph.windows.net';                       Port=443;  TestMode='CloudPC'; Notes='Authentication and Identity' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='login.live.com';                          Port=443;  TestMode='CloudPC'; Notes='Consumer device auth and Microsoft account' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='account.live.com';                        Port=443;  TestMode='CloudPC'; Notes='Consumer Outlook.com and OneDrive device auth' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='enterpriseregistration.windows.net';      Port=443;  TestMode='CloudPC'; Notes='Device registration' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Auth';                   Endpoint='certauth.enterpriseregistration.windows.net'; Port=443; TestMode='CloudPC'; Notes='Certificate-based device registration' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Navigation';             Endpoint='go.microsoft.com';                        Port=443;  TestMode='CloudPC'; Notes='Endpoint discovery' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Config';                 Endpoint='config.edge.skype.com';                   Port=443;  TestMode='CloudPC'; Notes='Feature deployment dependency' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Config';                 Endpoint='ecs.office.com';                          Port=443;  TestMode='CloudPC'; Notes='Feature deployment dependency' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Config';                 Endpoint='fd.api.orgmsg.microsoft.com';             Port=443;  TestMode='CloudPC'; Notes='Organizational messages' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Config';                 Endpoint='config.office.com';                       Port=443;  TestMode='CloudPC'; Notes='Office Customization Service' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Org Messages';            Endpoint='ris.prod.api.personalization.ideas.microsoft.com'; Port=443; TestMode='CloudPC'; Notes='Organizational messages personalization service' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='WNS Push';               Endpoint='sinwns1011421.wns.windows.com';            Port=443;  TestMode='CloudPC'; Notes='Windows Push Notification - Singapore WNS node' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='WNS Push';               Endpoint='sin.notify.windows.com';                  Port=443;  TestMode='CloudPC'; Notes='Windows Push Notification - Singapore notify node' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Android AOSP';           Endpoint='intunecdnpeasd.azureedge.net';             Port=443;  TestMode='CloudPC'; Notes='Android AOSP - legacy domain (migrating to manage.microsoft.com)' }
        [PSCustomObject]@{ Category='Intune'; Subcategory='Android AOSP';           Endpoint='intunecdnpeasd.manage.microsoft.com';     Port=443;  TestMode='CloudPC'; Notes='Android AOSP device management' }

        # ── Intune Autopilot ─────────────────────────────────────────────────
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Windows Update'; Endpoint='tsfe.trafficshaping.dsp.mp.microsoft.com'; Port=443; TestMode='CloudPC'; Notes='Autopilot traffic shaping' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Windows Update'; Endpoint='adl.windows.com';                        Port=443; TestMode='CloudPC'; Notes='Autopilot Windows Update' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='NTP';            Endpoint='time.windows.com';                      Port=123;  TestMode='CloudPC'; Notes='NTP time sync (UDP only)' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='WNS';            Endpoint='clientconfig.passport.net';             Port=443;  TestMode='CloudPC'; Notes='Autopilot WNS dependency' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='WNS';            Endpoint='windowsphone.com';                      Port=443;  TestMode='CloudPC'; Notes='Autopilot WNS dependency' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='WNS';            Endpoint='c.s-microsoft.com';                     Port=443;  TestMode='CloudPC'; Notes='Autopilot WNS dependency (specific node alongside *.s-microsoft.com)' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='TPM';            Endpoint='ekop.intel.com';                        Port=443;  TestMode='CloudPC'; Notes='Intel TPM Endorsement Key certificate' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='TPM';            Endpoint='ekcert.spserv.microsoft.com';           Port=443;  TestMode='CloudPC'; Notes='Microsoft TPM EK certificate service' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='TPM';            Endpoint='ftpm.amd.com';                          Port=443;  TestMode='CloudPC'; Notes='AMD fTPM certificate' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Diagnostics';    Endpoint='lgmsapeweu.blob.core.windows.net';      Port=443;  TestMode='CloudPC'; Notes='Autopilot diagnostics - West Europe' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Diagnostics';    Endpoint='lgmsapewus2.blob.core.windows.net';     Port=443;  TestMode='CloudPC'; Notes='Autopilot diagnostics - West US2' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Diagnostics';    Endpoint='lgmsapesea.blob.core.windows.net';      Port=443;  TestMode='CloudPC'; Notes='Autopilot diagnostics - SE Asia' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Diagnostics';    Endpoint='lgmsapeaus.blob.core.windows.net';      Port=443;  TestMode='CloudPC'; Notes='Autopilot diagnostics - Australia' }
        [PSCustomObject]@{ Category='Intune-Autopilot'; Subcategory='Diagnostics';    Endpoint='lgmsapeind.blob.core.windows.net';      Port=443;  TestMode='CloudPC'; Notes='Autopilot diagnostics - India' }
    )
}

# ─────────────────────────────────────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────────────────────────────────────

Write-Banner

# ── Step 1: Resolve endpoint list ────────────────────────────────────────────
$endpointData = $null

if ($EndpointsCSV -and (Test-Path $EndpointsCSV)) {
    Write-Host "  Loading endpoints from: $EndpointsCSV" -ForegroundColor DarkGray
    $endpointData = Import-Csv -Path $EndpointsCSV
} else {
    Write-Host "  Attempting to download Endpoints.csv from GitHub..." -ForegroundColor DarkGray
    try {
        $tempCSV = Join-Path $env:TEMP 'W365Endpoints.csv'
        Invoke-WebRequest -Uri $CSVGitHubURL -OutFile $tempCSV -TimeoutSec 15 -ErrorAction Stop
        $endpointData = Import-Csv -Path $tempCSV
        Write-Host "  Downloaded successfully." -ForegroundColor Green
    } catch {
        Write-Host "  Could not download CSV. Using built-in endpoint defaults." -ForegroundColor Yellow
        $endpointData = Get-BuiltInEndpoints
    }
}

Write-Host ""
Write-Host "  NOTE: Intune endpoints use a static hardcoded list per Microsoft guidance." -ForegroundColor DarkGray
Write-Host "        The endpoints.office.com API is deprecated for Intune and returns inaccurate data." -ForegroundColor DarkGray

# ── Step 2: Prompt for mode if not supplied ───────────────────────────────────
if ($Mode -eq 0) {
    Write-Host ""
    Write-Host "  Which network do you want to test from?" -ForegroundColor Yellow
    Write-Host "    [1]  Cloud PC / Host Network  (run ON the Cloud PC or Azure VNet VM)" -ForegroundColor White
    Write-Host "    [2]  Client Device Network    (run on the physical device used to ACCESS the Cloud PC)" -ForegroundColor White
    Write-Host "    [3]  Both" -ForegroundColor White
    Write-Host ""
    $inputMode = Read-Host "  Enter choice [1]"
    if ([string]::IsNullOrWhiteSpace($inputMode)) { $inputMode = '1' }
    $Mode = [int]$inputMode
    if ($Mode -notin 1, 2, 3) { $Mode = 1 }
}

$modeLabel = switch ($Mode) {
    1 { 'Cloud PC / Host Network' }
    2 { 'Client Device Network' }
    3 { 'Both' }
}

Write-Host ""
Write-Host "  Mode      : $modeLabel" -ForegroundColor Blue
Write-Host "  Computer  : $env:COMPUTERNAME  |  User: $env:USERNAME" -ForegroundColor DarkGray
Write-Host "  Date/Time : $(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')" -ForegroundColor DarkGray
Write-Host ""
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray
Write-Host "  Legend:  [ OK ] Connected   [FAIL] Blocked   [SKIP] Wildcard   [INFO] UDP/IP Range" -ForegroundColor DarkGray
Write-Host "─────────────────────────────────────────────────────" -ForegroundColor DarkGray

$allResults = @()

# ── Step 3: Run tests ─────────────────────────────────────────────────────────
if ($Mode -eq 1 -or $Mode -eq 3) {
    Write-Host ""
    Write-Host "┌─ CLOUD PC / HOST NETWORK TESTS ─────────────────────────┐" -ForegroundColor Magenta
    $allResults += Test-EndpointList -Endpoints $endpointData -FilterMode 'CloudPC'
    Write-Host ""
    Write-Host "└─────────────────────────────────────────────────────────┘" -ForegroundColor Magenta
}

if ($Mode -eq 2 -or $Mode -eq 3) {
    Write-Host ""
    Write-Host "┌─ CLIENT DEVICE NETWORK TESTS ───────────────────────────┐" -ForegroundColor Blue
    $allResults += Test-EndpointList -Endpoints $endpointData -FilterMode 'Client'
    Write-Host ""
    Write-Host "└─────────────────────────────────────────────────────────┘" -ForegroundColor Blue
}

# ── Step 4: Summary ───────────────────────────────────────────────────────────
Write-Summary -Results $allResults

# ── Step 5: Export results (optional) ────────────────────────────────────────
if ($OutputPath) {
    try {
        $allResults | Export-Csv -Path $OutputPath -NoTypeInformation -Encoding UTF8
        Write-Host ""
        Write-Host "  Results exported to: $OutputPath" -ForegroundColor Green
    } catch {
        Write-Host "  [WARN] Could not export results: $_" -ForegroundColor Yellow
    }
}

Write-Host ""
Write-Host "  Done." -ForegroundColor Blue
Write-Host ""
