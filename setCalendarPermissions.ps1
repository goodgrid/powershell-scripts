If (-Not (Get-Module -ListAvailable -Name ExchangeOnlineManagement )) {
    Write-Host "The Exchange Management module is not installed and will be installed before continuing"
    Install-Module -Name ExchangeOnlineManagement
} 

Import-Module ExchangeOnlineManagement
$Credential=Get-Credential
Connect-ExchangeOnline -Credential $Credential  

$Mailboxes = Get-Mailbox -RecipientTypeDetails UserMailbox 

ForEach($Mailbox in $Mailboxes) {

    $EmailAddress =  $Mailbox | Select-Object -ExpandProperty PrimarySmtpAddress

    $Calendar = Get-MailboxFolderStatistics -Identity $Mailbox -FolderScope Calendar | Where-Object {$_.FolderType -eq "Calendar"} | Select-Object @{n="Identity"; e={$_.Identity.ToString().Replace("\",":\")}}
	
	$PermissionLevel="LimitedDetails" #, AvailabilityOnly, FullDetails, Editor

	Write-Output "Setting calendar permissions for user $EmailAddress to $PermissionLevel"
	Set-MailboxFolderPermission -Identity $Calendar -User Default -AccessRights LimitedDetails

    

    $mailParams = @{
        SmtpServer                 = 'smtp.office365.com'
        Port                       = '587'
        UseSSL                     = $true
        Credential                 = $Credential
        From                       = 'XXX'
        To                         = $EmailAddress
        Bcc                        = 'XXX'
        Subject                    = "Je deelt nu beperkte details uit je agenda met collega's"
        BodyAsHtml                 = $true
        Body                       = "Hallo, <p>Vorige week kondigde Andr&eacute; aan dat het zinvol is om enig inzicht te hebben in agenda's bij het plannen van vergaderingen. Dit beleid is nu doorgevoerd in jouw agenda-instellingen en je deelt nu beperkte details met collega's. <P>Je kunt individuele afspraken als prive markeren als ze gevoelig of vertrouwelijk zijn. Achter de volgende link wordt uitgelegd hoe dat gaat: https://support.microsoft.com/nl-nl/office/een-vergadering-of-afspraak-privé-maken-dc3898f0-22f5-45c6-8cc8-b4d4db84111d.<p>Ik hoop je zo voldoende te hebben geinformeerd.<P>Hartelijke groet,<br>Koen Bonnet<p><i>PS. Email naar dit adres wordt niet gelezen. Als je vragen of opmerkingen hebt, stuur me dan een bericht op Slack</i>"
        DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
    }
    
    Send-MailMessage @mailParams
	
}
