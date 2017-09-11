##Created By Kristopher Roy##
##Create Date 15Aug17##
##Modified 24Aug17##

#Use a domain admin account
$RunAccount = get-Credential

#Creating $ServerList then limiting properties, for somereason doing it in one line errors out when you have a large number of servers 
$ServerList = $NULL
$ServerList =  get-adcomputer -Filter {OperatingSystem -like "*Server*"} -Properties IPv4Address,OperatingSystem
$Serverlist = $Serverlist|select Name,IPv4Address,OperatingSystem,HotfixesInstalled,LastInstalledDate
FOREACH($Server in $ServerList)
{
    $Hotfixes = $Null
    $Hotfixes = Get-HotFix -ComputerName $Server.Name -Credential $RunAccount|where HotFixID -NE "File 1"
    $server.Hotfixesinstalled = $Hotfixes.count
    $server.LastInstalledDate = ($Hotfixes|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
    IF($Server.LastInstalledDate -eq $Null -or $Server.LatInstalledDate -like "")
    {
        Invoke-Command -ComputerName $server.name -Credential $RunAccount{
        $Hotfixes = $Null
        $Hotfixes = Get-HotFix|where {$_.HotFixID -NE "File 1"}
        $Hotfixes
        }
    $server.Hotfixesinstalled = $result.count
    $server.LastInstalledDate = ($result|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
    }
}