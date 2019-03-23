
<#
    Script needs to be executed with farm admin rights, preferrably service account itself.
    Account should have full access on user profile service application
    Changes Required:


        Line 15
        Please change the path and other required parameters in Settings.xml
        Line 113: Put the applicable search string in where clause (as per your configuration values in Settings.xml)

#>

Add-PSSnapin Microsoft.SharePoint.PowerShell -ErrorAction SilentlyContinue
$xmlFile = ".\Settings.xml"
$xmlConfig = [System.Xml.XmlDocument](Get-Content $xmlFile)

#Variable to track audience counter
[int] $AudiCounter = 1
#Variable to track users counter
[int] $UsersCounter = 1

$date = (Get-Date).ToString('yyyy-MM-dd-HHmm')
$RootPath = $xmlConfig.Settings.LogsSettings.RootPath
$upsa = $xmlConfig.Settings.ConfigurationSettings.serviceapplication

#Number of rules per audience can only be 18 as per restriction from Microsoft
[int]$NumberOfUserPerAudience = 18
$CreatedAudiences = @()
$csvlocation = $RootPath + "Users.csv"
$LogsLocation = $RootPath + "Logs_$($date).txt"
$ErrorLogsLocation = $RootPath + "ErrorLogs_$($date).txt"

$TranscriptLogs = $RootPath + "TranscriptLogs_$($date).txt"

Start-Transcript -Path $TranscriptLogs -Append

#Import data from csv
$users = Import-csv -header UserID -Path $csvlocation

#Estimate the number of audiences
$CSVRowCount = $users.Count
$NoOfAudiences = ([math]::Round($CSVRowCount/$NumberOfUserPerAudience)+1)

## Settings you may want to change for Audience Name and Description ##
$upa = Get-SPServiceApplication | Where-Object {$_.DisplayName -eq $upsa}
$mySiteHostUrl = $xmlConfig.Settings.ConfigurationSettings.MySiteHostURL
$MySiteSPOUrl = $xmlConfig.Settings.ConfigurationSettings.MySiteSPOURL
$audienceDescription = $xmlConfig.Settings.ConfigurationSettings.AudienceDescription
$audienceName = $xmlConfig.Settings.ConfigurationSettings.Name


#Create audience, compile and apply profile redirection
function CreateAudience ([string] $audienceN, [string] $audienceDesc, $audienceR)
{
	try{
        #Get the My Site Host's SPSite object
	    $site = Get-SPSite $mySiteHostUrl
	    $ctx = [Microsoft.Office.Server.ServerContext]::GetContext($site)
	    $audMan = New-Object Microsoft.Office.Server.Audience.AudienceManager($ctx)
        $Audiences = $audMan.Audiences
        if(!($Audiences.AudienceExist($audienceN)))
        {
	        #Create a new audience object for the given Audience Manager
	        $aud = $audMan.Audiences.Create($audienceN, $audienceDesc)
	        $aud.AudienceRules = New-Object System.Collections.ArrayList            
            $audienceR | ForEach-Object { $aud.AudienceRules.Add($_) }
	        #Save the new Audience
	        $aud.Commit()
            $audJob = [Microsoft.Office.Server.Audience.AudienceJob]::RunAudienceJob(($upa.Id.Guid.ToString(), "1", "1", $aud.AudienceName))

            return $aud.AUdienceName
        }
        else
        {
             
            #Display a message - "Audience already exists in the system"
        }
    }
    catch
    {
        $audienceN, $audienceR, $_.Exception.Message | Out-File -FilePath $ErrorLogsLocation -Append

    }
}
   

try
{
	for($AudiCounter = 1; $AudiCounter -le $NoOfAudiences; $AudiCounter++)
	{
		$audienceNameFinal = $audienceName+$($AudiCounter+2)
		$audienceRules = @()

		for([int]$csvcounter=1; $csvcounter -le $NumberOfUserPerAudience;  $csvcounter++)     
		{ 
            if((($UsersCounter%$NumberOfUserPerAudience) -ne 0) -and ($users[$UsersCounter-1].UserID -ne $null))
            {
                $audienceRules += New-Object Microsoft.Office.Server.Audience.AudienceRuleComponent("AccountName", "=", "VCN\$($users[$UsersCounter-1].UserID)")
                #Create an OR group operator between the two audience rules.
                $audienceRules += New-Object Microsoft.Office.Server.Audience.AudienceRuleComponent("", "OR", "")
                $UsersCounter++
            }
            else
            {
                $audienceRules += New-Object Microsoft.Office.Server.Audience.AudienceRuleComponent("AccountName", "=", "VCN\$($users[$UsersCounter-1].UserID)")
                $AudName = CreateAudience $audienceNameFinal $audienceDescription $audienceRules
                $CreatedAudiences += $AudName
                $UsersCounter++
            } 
		}		
	}

	#setting redirection to all audiences
	$existingConfiguration = Get-SPO365LinkSettings	
    $CreatedAudiences = $CreatedAudiences | where {$_ -match "MySiteRedirection.*"}
    #Add the previously set audiences. If we don't do this then all the previously set audiences will get erased.
	$existingPlusCreatedAudiences = $existingConfiguration.Audiences + $CreatedAudiences		

    #Configure redirection
	Set-SPO365LinkSettings -MySiteHostUrl $MySiteSPOUrl -RedirectSites $true -Audiences $existingPlusCreatedAudiences -HybridAppLauncherEnabled $false -OnedriveDefaultToCloudEnabled $false
	Add-Content -Path $LogsLocation "Message: Profile redirection has been applied"
}
catch
{
	$_.Exception.Item, $_.Exception.Message | Out-File -FilePath $ErrorLogsLocation -Append
}    

Stop-Transcript


