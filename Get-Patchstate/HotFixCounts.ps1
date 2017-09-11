##Created By Kristopher Roy##
##Create Date 15Aug17##
##Modified 11Sep17##

#Use a domain admin account
#$RunAccount = get-Credential

#Creating $ServerList then limiting properties, for somereason doing it in one line errors out when you have a large number of servers 
$ServerList = $NULL
$ServerList =  get-adcomputer -Filter{OperatingSystem -like "*Server*"} -Properties OperatingSystem
#-Credential $RunAccount
$Serverlist = $Serverlist|select Name,IPv4Address,OperatingSystem,HotfixesInstalled,LastInstalledDate
$count = 0
FOREACH($Server in $ServerList)
{
    $count++
    Write-Progress -Activity "Gathering additional Server Details" -Status "$($count / $ServerList.Count * 100)% complete ($($count) of $($ServerList.count))" -CurrentOperation "processing Server '$($Server.Name)'" -PercentComplete $($count / $ServerList.Count * 100)
	
    write-host $server.Name"$count out of"$serverlist.count
    $name = $server.name
    #$xtravars = get-adcomputer -Filter{name -like $Name} -Properties IPv4Address 
    #-Credential $RunAccount
    #$Server.OperatingSystem = $xtravars.OperatingSystem
    $network = Test-Connection -ComputerName $server.name -Count 1 -ErrorAction SilentlyContinue
		##>>Check and write if good connection<<##
		IF($network.IPV4Address -eq $null){}Else{
			$server.IPv4Address = $network.IPV4Address
            $Hotfixes = $Null
            $Hotfixes = Get-HotFix -ComputerName $Server.Name|where {$_.HotFixID -NE "File 1"}
            #$Hotfixes.Description
            $server.Hotfixesinstalled = $Hotfixes.count
            $server.LastInstalledDate = ($Hotfixes|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
            IF($Hotfixes -eq $Null -or $Hotfixes.Count -like "")
            {
                Invoke-Command -ComputerName $name{
                #$Hotfixes = $Null
                $Hotfixes = Get-HotFix|where {$_.HotFixID -NE "File 1"}
                #$Hotfixes
            }
            $server.Hotfixesinstalled = $result.count
            $server.LastInstalledDate = ($result|where{$_.InstalledOn -ne "$Null" -and $_.InstalledOn -inotlike ""}|Sort-Object InstalledOn | Select-Object -Last 1).InstalledOn
        }
    }
    #$xtravars = $null
    $server = $null
}
$serverlist|export-csv c:\belltech\hotfixreport.csv