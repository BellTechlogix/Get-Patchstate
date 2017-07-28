<#
	This Function Gets the patch state of a PC
#>
function Get-Patchstate {
    param([string]$Name="Select Computer To Run Against",
		#Type True, False, or Both to return install, not-installed, or both
		[string]$installed="IsInstalled")

   If(!($Name)){$Name = $env:COMPUTERNAME}
   $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$Computername))
   $updatesearcher = $updatesession.CreateUpdateSearcher()
   # 0 = NotInstalled | 1 = Installed
   If($installed -eq 'True'){$searchresult = $updatesearcher.Search("IsInstalled=1 ")}
   If($installed -eq 'False'){$searchresult = $updatesearcher.Search("IsInstalled=0")}
   If($installed -eq 'Both'){$searchresult = $UpdateSearcher.Search("IsInstalled=0 or IsInstalled=1")}

   $Updates = If ($searchresult.Updates.Count  -gt 0) {

  #Updates are  waiting to be installed
  $count  = $searchresult.Updates.Count
  Write-Verbose  "Found $Count update\s!"

  #Cache the  count to make the For loop run faster   
  For ($i=0; $i -lt $Count; $i++) {
  #Create  object holding update

  $Update  = $searchresult.Updates.Item($i)
  [pscustomobject]@{
    Title =  $Update.Title
    KB =  $($Update.KBArticleIDs)
    SecurityBulletin = $($Update.SecurityBulletinIDs)
    MsrcSeverity = $Update.MsrcSeverity
    IsDownloaded = $Update.IsDownloaded
    IsInstalled = $Update.IsInstalled
    Url =  $Update.MoreInfoUrls
    Categories =  ($Update.Categories  | Select-Object  -ExpandProperty Name)
    BundledUpdates = @($Update.BundledUpdates)|ForEach{
    [pscustomobject]@{
      Title = $_.Title
      DownloadUrl = @($_.DownloadContents).DownloadUrl
      }
    }
   }
  }
 }
 Return $Updates 
}