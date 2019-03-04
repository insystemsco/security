#used for an initial audit of AD
#only collects and exports information
param(
    [string] $exportPath,
    [int] $daysInactive
)
Import-Module ActiveDirectory
function folderCheck ($path)
	{
    $pathTest = Test-Path -Path $path
    if ($pathTest -eq $true)
        {
            echo "Verified $path exists"
        }
    elseif ($pathTest -ne $true)
        {
            echo "$path does not exisit. Creating $path now"
            New-Item -ItemType Directory -Path $path
        }
	}
#all variables created here
$time = (Get-Date).AddDays(-($daysInactive))
   #gather all Groups
echo "Gathering list of AD Groups"
$adGroupList = Get-ADGroup -Filter * -Properties *
   #gather all enabled users in AD
echo "Gathering list of all enabled users"
$userList = Get-ADUser -Filter {enabled -eq $true} -Properties *
   #get list of GPO
echo "Gathering list of all GPOs"
$gpos = Get-GPO -All
   #get inactive items
echo "Gathering list of all inactive users"
$inactiveUsers = Get-ADUser -Filter{LastLogonTimeStamp -le $time -and enabled -eq $true} -Properties *
   #get inactive computers
echo "Gathering list of all inactive computers"
$inactiveComputers = Get-ADComputer -Filter {LastLogonDate -le $time} -Properties *
   #get disabled items
echo "Gathering all disabled users"
$disabledUsers = Get-ADUser -Filter {enabled -eq $false} -Properties *
echo "Gathering all disabled computers"
$disabledComputers = Get-ADComputer -Filter {enabled -eq $false} -Properties *
#check for directories
echo "Checking directories now"
folderCheck -path $exportPath
folderCheck -path "$exportPath\AD Groups"
folderCheck -path "$exportPath\Active AD Users"
folderCheck -path "$exportPath\AD GPOs"
folderCheck -path "$exportPath\AD GPOs\GPO Reports"
folderCheck -path "$exportPath\Inactive Items"
folderCheck -path "$exportPath\Disabled Items"
folderCheck -path "$exportPath\DC Information"
#export group lists
echo "Exporting all AD Groups"
$adGroupList|select name,groupcategory,samaccountname| Export-Csv -path "$exportPath\AD Groups\All Groups.csv" -NoTypeInformation
#gather all users in groups
echo "Starting export of members in AD Groups"
foreach ($group in $adGroupList)
	{
    $groupName = $group.samaccountname
    $fileName = $group.name
    echo "Exporting all group members for $groupName"
    Get-ADGroupMember -Identity $groupName| select name,samaccountname,objectclass|Export-Csv -Path "$exportPath\AD Groups\$fileName.csv" -NoTypeInformation	
	}
echo "Exporting enabled user list"
$userList|select Name,SamAccountName,Description,CanonicalName,LastLogonDate|Export-Csv -Path "$exportPath\Active AD Users\All Active Users.csv" -NoTypeInformation
echo "Exporting GPOs"
$gpos|select DisplayName,Owner,GpoStatus|Export-Csv -Path "$exportPath\AD GPOs\AllGPOs.csv" -NoTypeInformation
#get GPO report
echo "Starting GPO Reports"
foreach ($gpo in $gpos)
	{
		$gpoName = $gpo.DisplayName
		echo "Running GPO Report on $gpoName"
		Get-GPOReport -Name $gpoName -ReportType XML -Path "$exportPath\AD GPOs\GPO Reports\$gpoName.xml"
	}
echo "Exporting inactive users"
$inactiveUsers|select givenname,surname,name,samaccountname,enabled,@{Name="Stamp"; expression={[DateTime]::FromFileTime($_.lastLogonTimestamp).ToString('yyyy-MM-dd_hh:mm:ss')}},DistinguishedName|Export-Csv -Path "$exportPath\Inactive Items\Inactive Users.csv" -NoTypeInformation
echo "Exporting inactive computers"
$inactiveComputers|select name,DistinguishedName,LastLogonDate| export-csv -path "$exportPath\Inactive Items\Inactive Computers.csv" -NoTypeInformation
echo "Exporting disabled users"
$disabledUsers|select givenname,surname,name,samaccountname,enabled|Export-Csv -Path "$exportPath\Disabled Items\Disabled Users.csv" -NoTypeInformation
echo "Exporting disabled computers"
$disabledComputers|select name,DistinguishedName,LastLogonDate,Enabled|Export-Csv -Path "$exportPath\Disabled Items\Disabled Computers.csv" -NoTypeInformation
echo "Gathering all Domain Controllers"
$dcs = (Get-ADDomain).ReplicaDirectoryServers
$dcs += (Get-ADDomain).ReadOnlyReplicaDirectoryServers
Foreach ($dc in $dcs)
	{
    echo "Gathering information for $dc"
    Get-ADDomainController -Identity $dc|Export-Csv "$exportPath\DC Information\DC Information.csv" -Append -NoTypeInformation
    echo "Running dcdiag on $dc"
    dcdiag /s:$dc > "$exportPath\DC Information\$dc.txt"
	}
echo "Gathering FSMO roles"
NetDOM /query FSMO > "$exportPath\DC Information\FSMO.txt"
echo "Gathering Replication Status for domain"
Get-ADReplicationFailure -Scope Domain|Export-Csv -Path "$exportPath\DC Information\Replication Status.csv" -NoTypeInformation