###################
#################

# MAIN

###################
#################

Function Set-RpcOverHttps {

# Parameters passed to the script:
Param(

    # Fully qualified DNS of the M-Files server. Example: "mfiles.customer.com"
    [Parameter(Mandatory=$true)]
    [string]$mfilesDns,


    # M-Files version for registry entries. Example: "11.3.4330.196"
    [Parameter(Mandatory=$true)]
    [string]$mfilesVersion  

    )

# Process:

# "1. Install the RPC over HTTP Proxy feature"
Write-Host "1. Install the RPC over HTTP Proxy feature"
Install-WindowsFeature RPC-over-HTTP-Proxy

# 2. Set RpcProxy registry values.
$Key = "HKEY_LOCAL_MACHINE\Software\Microsoft\Rpc\RpcProxy"
If  ( -Not ( Test-Path "Registry::$Key")){New-Item -Path "Registry::$Key" -ItemType RegistryKey -Force}
Set-ItemProperty -path "Registry::$Key" -Name "AllowAnonymous" -Type "DWord" -Value 1

$Key = "HKEY_LOCAL_MACHINE\Software\Microsoft\Rpc\RpcProxy"
If  ( -Not ( Test-Path "Registry::$Key")){New-Item -Path "Registry::$Key" -ItemType RegistryKey -Force}
Set-ItemProperty -path "Registry::$Key" -Name "ValidPorts" -Type "String" -Value ("{0}:4466" -f  $mfilesDns)
	
# 3. Map localhost to the public DNS in the hosts file to allow the RPC calls through. 
Write-Host "3. Add host entry"
$hosts_entries = @{}
$hosts_entries.Add($mfilesDns, "127.0.0.1")
setHostEntries($hosts_entries)

# 4a. Allow anonymous & Disable basic authentication for Default Web Site:
Set-WebConfigurationProperty `
-pspath 'MACHINE/WEBROOT/APPHOST' `
-location 'Default Web Site' `
-filter "system.webServer/security/authentication/anonymousAuthentication" `
-name "enabled" `
-value "True"

Set-WebConfigurationProperty `
-pspath 'MACHINE/WEBROOT/APPHOST' `
-location 'Default Web Site' `
-filter "system.webServer/security/authentication/basicAuthentication" `
-name "enabled" `
-value "False"

# 4b. Allow anonymous & basic authentication for Rpc virtual directory:
Set-WebConfigurationProperty `
-pspath 'MACHINE/WEBROOT/APPHOST' `
-location 'Default Web Site/Rpc' `
-filter "system.webServer/security/authentication/anonymousAuthentication" `
-name "enabled" `
-value "True"

Set-WebConfigurationProperty `
-pspath 'MACHINE/WEBROOT/APPHOST' `
-location 'Default Web Site/Rpc' `
-filter "system.webServer/security/authentication/basicAuthentication" `
-name "enabled" `
-value "True"


# 5. Allow incoming firewall traffic from TCP 443, 4466
New-NetFirewallRule -DisplayName "Allow inbound M-Files HTTPS/SSL traffic (TCP 443 and 4466)" -Direction Inbound -Action Allow -EdgeTraversalPolicy Allow -Protocol TCP -LocalPort 443,4466

# 6. Restart IIS to activate settings
invoke-command -scriptblock {iisreset}

# 7. Set registry keys that enable RPC over HTTPS for M-Files.
$Key = "HKEY_LOCAL_MACHINE\Software\Motive\M-Files\{0}\Server\MFServer" -f $mfilesVersion
If  ( -Not ( Test-Path "Registry::$Key")){New-Item -Path "Registry::$Key" -ItemType RegistryKey -Force}
Set-ItemProperty -path "Registry::$Key" -Name "EnableRPCOverHTTP" -Type "DWord" -Value 1
	
# 8. Restart the M-Files server service running on the computer.
Write-Host "8. Restart the M-Files server service running on the computer."
Restart-Service -Name ("MFServer {0}" -f  $mfilesVersion)


}



###################

# FUNCTIONS

#################


function setHostEntries([hashtable] $entries) {
    $hostsFile = "$env:windir\System32\drivers\etc\hosts"
    $newLines = @()

    $c = Get-Content -Path $hostsFile
    foreach ($line in $c) {
        $bits = [regex]::Split($line, "\s+")
        if ($bits.count -eq 2) {
            $match = $NULL
            ForEach($entry in $entries.GetEnumerator()) {
                if($bits[1] -eq $entry.Key) {
                    $newLines += ($entry.Value + '     ' + $entry.Key)
                    Write-Host Replacing HOSTS entry for $entry.Key
                    $match = $entry.Key
                    break
                }
            }
            if($match -eq $NULL) {
                $newLines += $line
            } else {
                $entries.Remove($match)
            }
        } else {
            $newLines += $line
        }
    }

    foreach($entry in $entries.GetEnumerator()) {
        Write-Host Adding HOSTS entry for $entry.Key
        $newLines += $entry.Value + '     ' + $entry.Key
    }

    Write-Host Saving $hostsFile
    Clear-Content $hostsFile
    foreach ($line in $newLines) {
        $line | Out-File -encoding ASCII -append $hostsFile
    }
}
