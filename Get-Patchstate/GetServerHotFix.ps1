#
# GetServerHotFix.ps1
#
<#
	This Script uses the Get-Hotfix function to loop through a server list and return a report on Installed Hotfixes
	Created 21Aug17
	By TankCR
#>
##Modified 24Aug17##

#Use a domain admin account
$RunAccount = get-Credential

#Creating $ServerList then limiting properties, for somereason doing it in one line errors out when you have a large number of servers 
$ServerList =  get-adcomputer -Filter {OperatingSystem -like "*Server*"} -Properties *
$Serverlist = $Serverlist|select Name,IPv4Address,OperatingSystem,HotfixesInstalled,LastInstalledDate
FOREACH($Server in $ServerList)
{
    $Hotfixes = $Null
    $Hotfixes = Get-HotFix -ComputerName $Server.Name -Credential $RunAccount|where HotFixID -NE "File 1"
    $server.Hotfixesinstalled = $Hotfixes.count
    $server.LastInstalledDate = ($Hotfixes|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
    $server = $null
}
$server = $null
FOREACH($Server in ($serverList|where{$_.LastInstalledDate -eq $Null}))
{
    try{$result = Invoke-Command -ComputerName $server.name -Credential $RunAccount -ArgumentList $Server{
    $server = $server[0]
    $Hotfixes = $Null
    $Hotfixes = Get-HotFix -ComputerName $Server.Name -Credential $RunAccount|where HotFixID -NE "File 1"
    $server.Hotfixesinstalled = $Hotfixes.count
    $server.LastInstalledDate = ($Hotfixes|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
    
    }}
    Catch{write-host "error"}
}