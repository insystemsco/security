netsh advfirewall set currentprofile state on
Write-Host Firewall enabled...

Set-ItemProperty -Path "HKLM:\System\CurrentControlSet\Control\Terminal Server" -Name "fDenyTSConnections" –Value 1
# Firewall rules
netsh advfirewall firewall set rule group="remote desktop" new enable=No
New-NetFirewallRule -DisplayName "Disabling Port 21" -Action Block -Direction Outbound -DynamicTarget Any -EdgeTraversalPolicy Block -Profile Any -Protocol tcp -RemotePort 21
New-NetFirewallRule -DisplayName "Disabling Port 20" -Action Block -Direction Outbound -DynamicTarget Any -EdgeTraversalPolicy Block -Profile Any -Protocol tcp -RemotePort 20
New-NetFirewallRule -DisplayName "Disabling Port 23 (Telnet)" -Action Block -Direction Outbound -DynamicTarget Any -EdgeTraversalPolicy Block -Profile Any -Protocol tcp -RemotePort 23
New-NetFirewallRule -DisplayName "Disabling Port 80 " -Action Block -Direction Outbound -DynamicTarget Any -EdgeTraversalPolicy Block -Profile Any -Protocol tcp -RemotePort 80
New-NetFirewallRule -DisplayName "Disabling Port 25 " -Action Block -Direction Outbound -DynamicTarget Any -EdgeTraversalPolicy Block -Profile Any -Protocol tcp -RemotePort 25

#Disable-PSRemoting -force

# If DNS exists, turn off zone updates and transfers 
if (Get-Service -Name DnsServer -ErrorAction SilentlyContinue -ErrorVariable WindowsServiceExistsError){
       Get-DNSServerZone | Set-DNSServerPrimaryZone -DynamicUpdate None -SecureSecondaries TransferToZoneNameServer -Notify NotifyServers
   }

Write-Host Install Sysinternals
Write-Host Disable remote assistance in sysdm.cpl

Write-Host ------------ Active Services ---------------
Get-Service | Where {$_.status –eq 'running'}
Write-Host ------------ Active Services ---------------

Set-ExecutionPolicy Restricted
