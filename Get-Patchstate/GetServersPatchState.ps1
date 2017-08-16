<#
	This Script uses the Get-Patchstate function to loop through a server list and return a report on patch states
	Created 16Aug17
	By TankCR
#>
$RunAccount = get-Credential 'crsp\admin.kroy'
$Servers =  get-adcomputer -Filter {OperatingSystem -like "*Server*"} -Properties OperatingSystem|select Name
#Now Set your Output File Locations, use your temp folder
$File1 = "$env:TEMP\patchstate.xml"
$File2 = "$env:TEMP\patchstate.xlsx"
#Discard Old Copies
Remove-Item $file1,$file2

#create our initial file with the headers
(
 '<?xml version="1.0"?>
<?mso-application progid="Excel.Sheet"?>
<Workbook xmlns="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:o="urn:schemas-microsoft-com:office:office"
 xmlns:x="urn:schemas-microsoft-com:office:excel"
 xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet"
 xmlns:html="http://www.w3.org/TR/REC-html40">
 <DocumentProperties xmlns="urn:schemas-microsoft-com:office:office">
  <Author>Kristopher Roy</Author>
  <LastAuthor>'+$env:USERNAME+'</LastAuthor>
  <Created>'+(get-date)+'</Created>
   <Version>16.00</Version>
 </DocumentProperties>
 <OfficeDocumentSettings xmlns="urn:schemas-microsoft-com:office:office">
  <AllowPNG/>
 </OfficeDocumentSettings>
 <ExcelWorkbook xmlns="urn:schemas-microsoft-com:office:excel">
  <WindowHeight>8145</WindowHeight>
  <WindowWidth>20490</WindowWidth>
  <WindowTopX>0</WindowTopX>
  <WindowTopY>0</WindowTopY>
  <ActiveSheet>1</ActiveSheet>
  <ProtectStructure>False</ProtectStructure>
  <ProtectWindows>False</ProtectWindows>
 </ExcelWorkbook>
 <Styles>
  <Style ss:ID="Default" ss:Name="Normal">
   <Alignment ss:Vertical="Bottom"/>
   <Borders/>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"/>
   <Interior/>
   <NumberFormat/>
   <Protection/>
  </Style>
  <Style ss:ID="s90" ss:Name="Hyperlink">
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#0563C1"
    ss:Underline="Single"/>
  </Style>
  <Style ss:ID="s63">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
  </Style>
  <Style ss:ID="s67">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
   <Interior/>
   <NumberFormat ss:Format="0%"/>
  </Style>
  <Style ss:ID="s76">
   <Interior ss:Color="#FF0000" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s88">
   <Interior ss:Color="#00B050" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s89">
   <NumberFormat ss:Format="[ENG][$-409]mmmm\ d\,\ yyyy;@"/>
  </Style>
  <Style ss:ID="s91" ss:Parent="s90">
   <Alignment ss:Horizontal="Left" ss:Vertical="Bottom"/>
  </Style>
  <Style ss:ID="s92">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior ss:Color="#D0CECE" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s93">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior ss:Color="#D0CECE" ss:Pattern="Solid"/>
  </Style>
  <Style ss:ID="s94">
   <Borders>
    <Border ss:Position="Bottom" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Left" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Right" ss:LineStyle="Continuous" ss:Weight="2"/>
    <Border ss:Position="Top" ss:LineStyle="Continuous" ss:Weight="2"/>
   </Borders>
   <Font ss:FontName="Calibri" x:Family="Swiss" ss:Size="11" ss:Color="#000000"
    ss:Bold="1"/>
   <Interior ss:Color="#D0CECE" ss:Pattern="Solid"/>
  </Style>
 </Styles>'
 )> $file1

#$servers = 'CRSPDC2'

FOREACH($server in $servers.name){
$result = Invoke-Command -ComputerName $server -Credential $RunAccount{
function Get-Patchstate {
    param([string]$computer,
		#Type True, False, or Both to return install, not-installed, or both
		[string]$installed="Both")

   If(!($computer)){
   $Name = $env:COMPUTERNAME
   $AutoUpdate = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.AutoUpdate"))
   $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session"))
   }
   ELSE{
   $AutoUpdate = [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.AutoUpdate",$computer))
   $updatesession =  [activator]::CreateInstance([type]::GetTypeFromProgID("Microsoft.Update.Session",$computer))}
   $updatesearcher = $updatesession.CreateUpdateSearcher()
   # 0 = NotInstalled | 1 = Installed
   If($installed.ToUpper() -eq 'True'){$searchresult = $updatesearcher.Search("IsInstalled=1 ")}
   If($installed.ToUpper() -eq 'False'){$searchresult = $updatesearcher.Search("IsInstalled=0")}
   If($installed.ToUpper() -eq 'Both' -or !($installed)){$searchresult = $UpdateSearcher.Search("IsInstalled=0 or IsInstalled=1")}

   $Updates = If ($searchresult.Updates.Count  -gt 0) {

  #Updates are  waiting to be installed
  $count  = $searchresult.Updates.Count
  Write-Verbose  "Found $Count update\s!"

  #Header Objects
  [pscustomobject]@{
    Updates = @( 
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
 )
    LastSearchSuccessDate = $AutoUpdate.results.LastSearchSuccessDate
    LastInstallationSuccessDate = $AutoUpdate.results.LastInstallationSuccessDate
 }
 }
 Return $Updates 
}
Get-Patchstate	}
$updates = $result.Updates|select Title,KB,SecurityBulletin,MsrcSeverity,IsDownloaded,IsInstalled,@{N = 'URL'; E = {$_.url|Out-String}},@{N = 'Categories'; E = {$_.Categories|Out-String}},@{N = 'LastSearchSuccessDate'; E = {($result.LastSearchSuccessDate|Out-String).Trim()}},@{N = 'LastInstallationSuccessDate'; E = {($result.LastInstallationSuccessDate|Out-String).Trim()}}
$Number = $Null
$Number = $updates.count
add-content $file1('<Worksheet ss:Name="'+$server+'">
<Table ss:ExpandedColumnCount="10" ss:ExpandedRowCount="'+($Number+1)+'" x:FullColumns="1"
   x:FullRows="1" ss:DefaultRowHeight="15">
   <Column ss:Width="65.25"/>
   <Column ss:Width="93"/>
   <Column ss:Width="54.75"/>
   <Column ss:Width="78.75"/>
   <Column ss:Width="95.25"/>
   <Column ss:AutoFitWidth="0" ss:Width="84.75"/>
   <Column ss:AutoFitWidth="0" ss:Width="90.75"/>
   <Column ss:Width="70.5"/>
   <Column ss:Width="113.25"/>
   <Column ss:Width="135.75"/>
   <Row ss:AutoFitHeight="0" ss:Height="15.75">
    <Cell ss:StyleID="s92"><Data ss:Type="String">Title</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">KB</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">SecurityBulletin</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">MSRCSeverity</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">ISDownloaded</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">IsInstalled</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">URL</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">Categories</Data></Cell>
    <Cell ss:StyleID="s93"><Data ss:Type="String">LastSearchSuccessDate</Data></Cell>
    <Cell ss:StyleID="s94"><Data ss:Type="String">LastInstallationSuccessDate</Data></Cell>
   </Row>')
	FOREACH($Update in $Updates)
	{
		Add-Content $file1 ('<Row ss:AutoFitHeight="0">')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.Title+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.KB+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.SecurityBulletin+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.MsrcSeverity+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.IsDownloaded+'</Data></Cell>')
		IF($update.IsInstalled.tostring() -match "True"){Add-Content $file1 ('<Cell ss:StyleID="s88"><Data ss:Type="String">'+$update.IsInstalled+'</Data></Cell>')}
		IF($update.IsInstalled.tostring() -match "False"){Add-Content $file1 ('<Cell ss:StyleID="s76"><Data ss:Type="String">'+$update.IsInstalled+'</Data></Cell>')}
		Add-Content $file1 ('<Cell ss:StyleID="s91" ss:HRef="'+$update.url+'"><Data ss:Type="String">'+$update.url+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s63"><Data ss:Type="String">'+$update.Categories+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s89"><Data ss:Type="DateTime">'+(get-date ($update.LastSearchSuccessDate) -Format yyyy-MM-dd)+'</Data></Cell>')
		Add-Content $file1 ('<Cell ss:StyleID="s89"><Data ss:Type="DateTime">'+(get-date $update.LastInstallationSuccessDate -Format yyyy-MM-dd)+'</Data></Cell>')
		Add-Content $file1 ('</Row>')	
	}
Add-Content $File1(
	  '</Table>
  <WorksheetOptions xmlns="urn:schemas-microsoft-com:office:excel">
   <PageSetup>
    <Header x:Margin="0.3"/>
    <Footer x:Margin="0.3"/>
    <PageMargins x:Bottom="0.75" x:Left="0.7" x:Right="0.7" x:Top="0.75"/>
   </PageSetup>
   <Unsynced/>
   <Print>
    <ValidPrinterInfo/>
    <HorizontalResolution>600</HorizontalResolution>
    <VerticalResolution>600</VerticalResolution>
   </Print>
   <Panes>
    <Pane>
     <Number>3</Number>
     <RangeSelection>R1:R3</RangeSelection>
    </Pane>
   </Panes>
   <ProtectObjects>False</ProtectObjects>
   <ProtectScenarios>False</ProtectScenarios>
  </WorksheetOptions>
 </Worksheet>')
}
Add-Content $file1 ('</Workbook>')

