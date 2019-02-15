#=======================================================================================
# Workstation Quarantine
#=======================================================================================

if (-Not ($lockdown)) {
} else {
if ($lockdown -eq $true) {

    Write-Host "Locking down endpoint: $computerName - $ip"

# Lockdown
    Function Invoke-Lockdown{

        # Disable Network Interfaces
        $wirelessNic = Get-WmiObject -Class Win32_NetworkAdapter -filter "Name LIKE '%Wireless%'"
        $wirelessNic.disable()
        $localNic = Get-WmiObject -Class Win32_NetworkAdapter -filter "Name LIKE '%Intel%'"
        $localNic.disable()
        Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 1101 -Message "Lockdown : Network Interface Cards Disabled"

        $WmiHash = @{}
        if($Private:Credential){
            $WmiHash.Add('Credential',$credential)
        }
        Try{
            $Validate = (Get-WmiObject -Class Win32_OperatingSystem -ComputerName $C -ErrorAction Stop @WmiHash).Win32Shutdown('0x0')
        } Catch [System.Management.Automation.MethodInvocationException] {
            Write-Error 'No user session found to log off.'
            Exit 1
        } Catch {
            Throw
        }
        if($Validate.ReturnValue -ne 0){
            Write-Error "User could not be logged off, return value: $($Validate.ReturnValue)"
            Exit 1
        }
        Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 1102 -Message "Lockdown : All Local Users Logged Out"

    # Lock Workstation
    rundll32.exe user32.dll,LockWorkStation > $null 2>&1
    Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 1103 -Message "Lockdown : System Locked"
    }

} else {
        Write-Host "Missing Required Parameter [lockdown]"
        Write-Host "     This option was specified "
        Write-Host "PS C:\> .\TR3PS.ps1 -lockdown"
        Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 34404 -Message "Forensic Data Acquisition Failure : Missing Required Parameter"
        Exit 1
}
}

# Lock out the user's AD account
if (-Not ($adLock)) {
} else {
if ($adLock -eq $true) {
    function get-dn () {
    $root = New-Object System.DirectoryServices.DirectoryEntry
    $searcher = new-object System.DirectoryServices.DirectorySearcher($root)
    $searcher.filter = "(&(objectClass=user)(sAMAccountName= $accountNameAD))"
    $user = $searcher.findall()
        if ($user.count -gt 1) {
            $count = 0
                foreach($i in $user) {
                    write-host $count ": " $i.path
                    $count = $count + 1
                }
            $selection = Read-Host "Please select item: "
            return $user[$selection].path
          } else {
          return $user[0].path
          }
    }
    $path = get-dn $accountNameAD
    if ($path -ne $null)    {
        $account=[ADSI]$path
        $account.psbase.invokeset("AccountDisabled", "True")
        $account.setinfo()
    Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 2101 -Message "AD Lockout : User $account Disabled within Active Directory"
  } else {
        write-host "No user account found!"
        Write-Host "Please specify a user account with the following command line switch:"
        Write-Host "PS C:\> .\TR3PS.ps1 -adLock [username]"
        Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 34404 -Message "Forensic Data Acquisition Failure : Username Not Found"
        Exit 1
  }
}
}
}
if (-Not ($remote)) {
Invoke-Recon
} Else {
    if ($remote -eq $true) {
        $hostnameCheck = "^(?=.{1,255}$)[0-9A-Za-z](?:(?:[0-9A-Za-z]|-){0,61}[0-9A-Za-z])?(?:\.[0-9A-Za-z](?:(?:[0-9A-Za-z]|-){0,61}[0-9A-Za-z])?)*\.?$"
        if (-not ($target -match $hostnameCheck)) {
            Write-Host "That's not a hostname..."
            Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 34405 -Message "Potential Attack Detected via hostname parameter : $target"
            Exit 1
        }
        if ($sendEmail -eq $false) {
            Write-Host ""
            Write-Host "You must get the data off of the remote host."
            Write-Host "Try using the -sendEmail parameter."
            Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 34404 -Message "Forensic Data Acquisition Failure : Missing Parameter"
            Exit 1
        }
        try {
            if (-Not ($password)) {
                $cred = Get-Credential
            } Else {
                $securePass = ConvertTo-SecureString -string $password -AsPlainText -Force
                $cred = New-Object -typename System.Management.Automation.PSCredential -argumentlist $username, $securePass
            }
            $scriptName = $MyInvocation.MyCommand.Name
            $content = type $scriptName

            # send email
            if ($sendEmail -eq $true) {

                # extract client email data (send contents via email)
                if ($email -eq $true) {
                    Invoke-Command -ScriptBlock {
                        param ($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$email)
                        if (Test-Path \TR3PS.ps1) {
                            rm \TR3PS.ps1
                        }
                        $content >> \TR3PS.ps1
                        C:\TR3PS.ps1 -sendEmail -email -smtpServer $smtpServer -emailFrom $emailFrom -emailTo $emailTo
                        rm C:\TR3PS.ps1
                    } -ArgumentList @($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$email) -ComputerName $target -Credential $cred
                } Else {

                # Lockdown the endpoint (disable NIC's, log user out, lock workstation, and send results via email)
                    if ($lockdown -eq $true) {
                        Invoke-Command -ScriptBlock {
                            param ($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$lockdown)
                            if (Test-Path \TR3PS.ps1) {
                                rm \TR3PS.ps1
                            }
                            $content >> \TR3PS.ps1
                            C:\TR3PS.ps1 -sendEmail -smtpServer $smtpServer -emailFrom $emailFrom -emailTo $emailTo -lockdown
                            rm C:\TR3PS.ps1
                        } -ArgumentList @($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$lockdown) -ComputerName $target -Credential $cred
                    } Else {

                # lock out an account in AD (send results via email)
                    if ($adlock -eq $true) {
                        Invoke-Command -ScriptBlock {
                            param ($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$adlock,$user,$accountNameAD,$account)
                            if (Test-Path \TR3PS.ps1) {
                                rm \TR3PS.ps1
                            }
                            $content >> \TR3PS.ps1
                            C:\TR3PS.ps1 -sendEmail -smtpServer $smtpServer -emailFrom $emailFrom -emailTo $emailTo -adlock $account
                            rm C:\TR3PS.ps1
                        } -ArgumentList @($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo,$adlock,$user,$accountNameAD,$account) -ComputerName $target -Credential $cred
                    } Else {

                # default execution (send results via email)
                    Invoke-Command -ScriptBlock {
                        param ($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo)
                        if (Test-Path \TR3PS.ps1) {
                            rm \TR3PS.ps1
                        }
                        $content >> \TR3PS.ps1
                        C:\TR3PS.ps1 -sendEmail -smtpServer $smtpServer -emailFrom $emailFrom -emailTo $emailTo
                        rm \TR3PS.ps1
                    } -ArgumentList @($content,$scriptName,$sendEmail,$smtpServer,$emailFrom,$emailTo) -ComputerName $target -Credential $cred
                }
            }}}

            # push data to share (due to security concerns)
            if ($share -eq $true) {
                $banner
                Write-Host "currently pushing to a share from a remote host is not supported."
                Write-Host "Please use -sendEmail for now unless executing locally..."
                Exit 1
            }

      } Catch {
        Write-Host "Access Denied..."
        Write-EventLog -LogName Application -Source "TR3PS" -EntryType Information -EventId 34404 -Message "Forensic Data Acquisition Failure : Access Denied"
        Exit 1
      }
    }