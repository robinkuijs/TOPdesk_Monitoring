# Dit script wordt gebruikt om de diagnostics van TOPdesk uit te lezen en vervolgens waarden hieruit weg te schrijven naar een monitoring database
# De resulterende rapportage staat op http://servernaam/Reports/Pages/Report.aspx?ItemPath=%2fDiagnostics 

# In dit script wordt gebruik gemaakt van het HTML Agility Pack. Deze moet aanwezig zijn op de volgende locatie:
add-type -Path 'D:\Program Files\Diagnostics\Script\HtmlPack\HtmlAgilityPack.dll'
$HTMLDocument = New-Object HtmlAgilityPack.HtmlDocument

# DE huidige datum/tijd is nodig om een timestamp in te vullen in de monitoring database
$currenthour = (Get-Date).ToString('HH')

 
###############  Mail parameters  ################
$to = "<mailadres>"
$from = "<mailadres>"
$smtpserver = "<mailserver>"
##################################################

###############  Database parameters  ############
$server = "<databaseserver>"
$username = "<username>"
$password = '<password>'
$databasename = "<databasename>"
##################################################

# De url van de TOPdesk diagnostics:
$URL = "http://<topdeskserver>/tas/secure/jsp/debug/diagnostic.jsp?j_username=<username>&j_password=<password>"

# Tijdelijke bestanden om bepaalde tabellen van de diagnostics pagina weg te schrijven:
$OutputFile = "d:\Program Files\Diagnostics\tempfile_diagnostics1.html"
$OutputFile2 = "d:\Program Files\Diagnostics\tempfile_diagnostics2.html"
$OutputFile3 = "d:\Program Files\Diagnostics\tempfile_diagnostics3.html"

# De eerste tabel wordt naar een tijdelijk bestand geschreven:
$data = Invoke-WebRequest -Uri $URL
@($data.ParsedHtml.getElementsByTagName("table"))[0].OuterHTML | 
Set-Content -Path $OutputFile

# De tweede tabel wordt naar een tijdelijk bestand geschreven:
$data = Invoke-WebRequest -Uri $URL
@($data.ParsedHtml.getElementsByTagName("table"))[1].OuterHTML | 
Set-Content -Path $OutputFile2

# De derde tabel wordt naar een tijdelijk bestand geschreven:
$data = Invoke-WebRequest -Uri $URL
@($data.ParsedHtml.getElementsByTagName("table"))[2].OuterHTML | 
Set-Content -Path $OutputFile3

# Bepaalde waarden worden uit de eerste tabel gehaald:
$result1 = $HTMLDocument.Load("d:\Program Files\Diagnostics\tempfile_diagnostics1.html")
$texts1 = $HTMLDocument.DocumentNode.SelectNodes("//table[1]//tr[2]")


$table1 = $texts1 | % {
					$securesessions = $_.SelectSingleNode("td[2]").Innertext
					$publicsessions = $_.SelectSingleNode("td[3]").Innertext
								
					}

# Bepaalde waarden worden uit de tweede tabel gehaald:					
$result2 = 	$HTMLDocument.Load("d:\Program Files\Diagnostics\tempfile_diagnostics2.html")				
$texts2 = $HTMLDocument.DocumentNode.SelectNodes("//table[1]//tr[2]")

$table2 = $texts2 | % {
					$heap_used = $_.SelectSingleNode("td[2]").Innertext
					$heap_committed = $_.SelectSingleNode("td[3]").Innertext
					$heap_max = $_.SelectSingleNode("td[4]").Innertext			
					}

# Bepaalde waarden worden uit de derde tabel gehaald:					
$result3 = 	$HTMLDocument.Load("d:\Program Files\Diagnostics\tempfile_diagnostics3.html")				
$texts3 = $HTMLDocument.DocumentNode.SelectNodes("//table[1]")

$table3 = $texts3 | % {
					$codecache_used = $_.SelectSingleNode("//tr[2]//td[3]").Innertext
					$codecache_max = $_.SelectSingleNode("//tr[2]//td[6]").Innertext
					$metaspace_used = $_.SelectSingleNode("//tr[3]//td[3]").Innertext
					$metaspace_max = $_.SelectSingleNode("//tr[3]//td[6]").Innertext
					$compressedclass_used = $_.SelectSingleNode("//tr[4]//td[3]").Innertext
					$compressedclass_max = $_.SelectSingleNode("//tr[4]//td[6]").Innertext
					$pareden_used = $_.SelectSingleNode("//tr[5]//td[3]").Innertext
					$pareden_max = $_.SelectSingleNode("//tr[5]//td[6]").Innertext
					$parsurvivor_used = $_.SelectSingleNode("//tr[6]//td[3]").Innertext
					$parsurvivor_max = $_.SelectSingleNode("//tr[6]//td[6]").Innertext
					$cmsoldgen_used = $_.SelectSingleNode("//tr[7]//td[3]").Innertext
					$cmsoldgen_max = $_.SelectSingleNode("//tr[7]//td[6]").Innertext
					}

# De waarden worden weg geschreven naar de monitoring database:					
Add-PSSnapin SqlServerCmdletSnapin100
Add-PSSnapin SqlServerProviderSnapin100					
$sqlquery = 'insert into diagnostics (timestamp,securesessions,publicsessions,heap_used,heap_committed,heap_max,codecache_used,codecache_max,metaspace_used,metaspace_max,compressedclass_used,compressedclass_max,pareden_used,pareden_max,parsurvivor_used,parsurvivor_max,cmsoldgen_used,cmsoldgen_max) values' + ' (' + 'getdate()' + ',' + $securesessions + ',' + $publicsessions + ',' + $heap_used + ',' + $heap_committed + ',' + $heap_max + ',' + $codecache_used + ',' + $codecache_max + ',' + $metaspace_used + ',' + $metaspace_max + ',' + $compressedclass_used + ',' + $compressedclass_max + ',' + $pareden_used + ',' + $pareden_max + ',' + $parsurvivor_used + ',' + $parsurvivor_max + ',' + $cmsoldgen_used + ',' + $cmsoldgen_max +');'
Invoke-Sqlcmd -Query $sqlquery -ServerInstance $server -Database $databasename -Username $username -Password $password

# mail sturen als geheugengebruik te groot wordt
# per uur kan er maar 1 keer een mail gestuurd worden
# Er worden lock bestanden weg geschreven zodat niet elke keer dat dit script uitgevoerd wordt een mail gestuurd wordt
$heapalarmlock = "d:\Program Files\Diagnostics\Heapalarm_lock.txt"
$metaspacealarmlock = "d:\Program Files\Diagnostics\Metaspacealarm_lock.txt"

# Alarm cutoff: Bij welk gedeelte van maximaal geheugen moet de mail verzonden worden?
$alarm_treshold = '0.8'
	
# heap geheugen		
if ([int]$heap_used -gt [int]$alarm_treshold*[int]$heap_max)
{
	if (Test-Path $heapalarmlock)
	{
	$heapalarm_content = Get-Content $heapalarmlock
		if ([int]$heapalarm_content -eq  [int]$currenthour)
		{
		$heap_alarm = $currenthour
		$heap_alarm | Set-Content -Path $heapalarmlock
		}
		else
		{
		$subject = "Geheugengebruik TOPdesk is nog steeds te hoog"
		$mailbody = "Heap geheugen in gebruik:" + ' ' + $heap_used
		send-mailmessage -to $to -from $from -subject $subject -smtpserver $smtpserver -body "$mailbody"

		$heap_alarm = $currenthour
		$heap_alarm | Set-Content -Path $heapalarmlock
		}
	}
	else 
	{
	$subject = "Storing TOPdesk: geheugengebruik"
	$mailbody = "Heap geheugen in gebruik:" + ' ' + $heap_used
	send-mailmessage -to $to -from $from -subject $subject -smtpserver $smtpserver -body "$mailbody"

	$heap_alarm = $currenthour
	$heap_alarm | Set-Content -Path $heapalarmlock
	}
}
else 
{
	if (Test-Path $heapalarmlock)
	{
	Remove-Item $heapalarmlock
	}
}	

# metaspace geheugen		
if ([int]$metaspace_used -gt [int]$alarm_treshold*[int]$metaspace_max)
{
	if (Test-Path $metaspacealarmlock)
	{
	$metaspacealarm_content = Get-Content $metaspacealarmlock
		if ([int]$metaspacealarm_content -eq  [int]$currenthour)
		{
		$metaspace_alarm = $currenthour
		$metaspace_alarm | Set-Content -Path $metaspacealarmlock
		}
		else
		{
		$subject = "Geheugengebruik TOPdesk is nog steeds te hoog"
		$mailbody = "Metaspace geheugen in gebruik:" + ' ' + $metaspace_used
		send-mailmessage -to $to -from $from -subject $subject -smtpserver $smtpserver -body "$mailbody"

		$metaspace_alarm = $currenthour
		$metaspace_alarm | Set-Content -Path $metaspacealarmlock
		}
	}
	else 
	{
	$subject = "Storing TOPdesk: geheugengebruik"
	$mailbody = "Metaspace geheugen in gebruik:" + ' ' + $metaspace_used
	send-mailmessage -to $to -from $from -subject $subject -smtpserver $smtpserver -body "$mailbody"

	$metaspace_alarm = $currenthour
	$metaspace_alarm | Set-Content -Path $metaspacealarmlock
	}
}
else 
{
	if (Test-Path $metaspacealarmlock)
	{
	Remove-Item $metaspacealarmlock
	}
}	




