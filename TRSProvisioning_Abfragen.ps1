###############################################################################
#
#    Prerequisites for Office 365
#
#    Connect to Office 365 using Windows PowerShell: http://go.microsoft.com/fwlink/p/?LinkID=614839 
#
#    Connect to Exchange Online using Windows PowerShell: http://go.microsoft.com/fwlink/p/?LinkId=396554
#
#    Connect to Teams using Windows PowerShell: https://docs.microsoft.com/de-de/MicrosoftTeams/teams-powershell-install#install-the-teams-powershell-module
#    
#    Modules Needed:
#    MicrosoftTeams install-module microsoftteams
#    EXO            install-Module ExchangeOnlineManagement
#    AAD            install-module AzureAd
#   
###############################################################################

#
# Global Constants
#

$script:strDevice = $null
$script:credAD = $null
$script:credExchange = $null
$script:credSkype = $null
$script:credNewAccount = $null
$script:strHybrid = $null
$script:strEasPolicy = $null
$script:strDatabase = $null
#$script:strUpn = $null

#Setzen der gewünschten Lizenzen, per Standart ist "Meeting_Room" voreingestellt
#Weiter gehen z.B. ENTERPRISEPREMIUM = O365 E5 oder SPE_E5 = M365 E5
$script:availableLicense = 'MEETING_ROOM'

$status = @{}



function Anfangen {
  Clear-Host
  Write-Host '*************************'
  Write-Host '*Möchten Sie mit der MTR*'
  Write-Host '* Einrichtung starten?  *' 
  Write-Host '*************************'
  $strProvisionMode = Read-Host -Prompt '1 für ja, 2 zum abbrechen'
  if ($strProvisionMode -eq 1)
    {
      CreateCloudAD
    }
    else
    {	
      if ($strProvisionMode -eq 2)
        {
          Clear-Host
          Write-Host 'Dann eben nicht, vielleicht beim nächsten mal :-)'
          Start-Sleep -Seconds 3
          Clear-Host
          CleanupAndFail
        }
       
          else
           {
                  Clear-Host
                  Write-Host 'Falsche Eingabe'
                  Start-Sleep -Seconds 1
                  Clear-Host
                  Anfangen
           }
            
        }
    }
		
	
function CountDown() {
  param($timeSpan)

  while ($timeSpan -gt 0)
{
  Write-Host '.' -NoNewline
  $timeSpan = $timeSpan - 1
  Start-Sleep -Seconds 1
}
}

function PrintAction {
    
  param
  (
    $strMsg
  )
  Write-Host $strMsg
}

function CleanupAndFail {
  # Cleans up and prints an error message
    
  param
  (
    $strMsg
  )
  if ($strMsg)
    {
        PrintError -strMsg ($strMsg)

    }
    Cleanup
    exit 1
}

function Cleanup () {
  # Cleans up set state such as remote powershell sessions
    if ($sessExchange)
    {
        Remove-PSSession -Id $sessExchange
    }
    if ($sessCS)
    {
        Remove-PSSession -Id $sessSkype
    }
}

function PrintError {
    
  param
  (
    $strMsg
  )
  Write-Host $strMsg
}

function RegistrarPool {
    try
    {
     Write-Host " "
      Write-Host -ForegroundColor Green "Warte 2 Minuten Auf RegistrarPool"

      Countdown -timeSpan 120

      #RegistrarPool auslesen
     
      $strRegPool = (Get-CsTenant).RegistrarPool

      

      #MeetingRoom Aktivieren
      Write-Host " "
      Enable-CsMeetingRoom -Identity $strUpn -RegistrarPool $strRegPool -SipAddressType EmailAddress
      }
      catch
      {
      }
       if ($Error)
    {
      $Error.Clear()
     Write-Host "."
     Write-Host -ForegroundColor Yellow "Der RegistrarPool ist noch nicht erstellt. Ich Versuche es weiter"
     Countdown -timeSpan 5
     RegistrarPool
    }
          else
      {
      Write-Host " "
      Write-Host -ForegroundColor Green "Der Registrierungspool lautet: $strRegPool"
         Write-Host " "
      Write-Host -ForegroundColor Green "Der Meeting Room $strDisplayName wurde erstellt"
      Pause
      Anfangen
      }
    }
function Connect2AzureAD {
  try
  {
    Clear-Host
    Write-Host '***************************'
    Write-Host '      UPN und Passwort     '
    Write-Host ' für Azure AD Admin Konto  '
    Write-Host '          eingeben         '
    Write-Host '***************************'
    
    $strAdmin = Read-Host -Prompt "Bitte geben Sie ihr Admin Konto an. Damit werden Sie an allen Konsolen angemeldet"
      

      Connect-AzureAD -AccountId $strAdmin
      Write-Host -ForegroundColor Green "Verbindung zu Azure AD Powershell hergestellt"
      Countdown -timeSpan 3
      
      #An Microsoft Teams PowerShell anmelden
      Connect-MicrosoftTeams -AccountId $strAdmin
      Write-Host -ForegroundColor Green 'Verbindung zu Teams Powershell hergestellt'
      Countdown -timeSpan 3

      #Verbindung zu Exchange Online herstellen
      Connect-ExchangeOnline -UserPrincipalName $strAdmin
      Write-Host -ForegroundColor Green 'Verbindung zu EXO hergestellt'
      Countdown -timespan 3

      #Verbindung zu CS Online herstellen
      Import-Module MicrosoftTeams
      $sfbSession = New-CsOnlineSession
      Import-PSSession $sfbSession
      Write-Host -ForegroundColor Green 'Verbindung zu CS Online hergestellt'
      Countdown -timespan 3

  }
  catch
  {
      CleanupAndFail -strMsg "Failed to connect to Azure Active Directory. Please check your credentials and try again. Error message: $_"
  }
}

function Licensecheck {

 $skus = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $availableLicense -EQ)
     $i = 1
    Foreach ($strSKU in $skus)
  {
    $iUnassigned = $strSKU.prepaidunits.Enabled - $strSKU.consumedunits
    	
	  Write-Host -NoNewLine $i
	  Write-Host -NoNewLine ': AccountSKUID: '
	  Write-Host -NoNewLine $strSKU.SkuPartNumber
	  Write-Host -NoNewLine ' Acctive Units: '
	  Write-Host -NoNewLine $strSKU.prepaidunits.Enabled
	  Write-Host -NoNewLine ' Unassigned Units '
	  Write-Host $iUnassigned
	
	
    $i++
  }
 If ($iUnassigned -gt 0)
   {
        Write-Host -ForegroundColor Green 'Super, MTR Lizenzen sind vorhanden!'
        Countdown -timeSpan 5
      }
        else
        {
        Write-Host -ForegroundColor Red 'Es sind leider keine Meeting Room Lizenzen vorhanden.'
        Pause
        CleanupAndFail

        }
}


function CreateCloudAD {
  



  Write-Host " "
  Write-Host " "
  Write-Host '***************************'
  Write-Host '      UPN und Passwort     '
  Write-Host ' für Azure AD Admin Konto  '
  Write-Host '          eingeben         '
  Write-Host '***************************'
  
  $strAdmin = Read-Host -Prompt "Bitte geben Sie ihr Admin Konto an. Damit werden Sie an allen Konsolen angemeldet"
    

    Connect-AzureAD -AccountId $strAdmin
    Write-Host -ForegroundColor Green "Verbindung zu Azure AD Powershell hergestellt"
    Countdown -timeSpan 3
    
    #An Microsoft Teams PowerShell anmelden
    Connect-MicrosoftTeams -AccountId $strAdmin
    Write-Host -ForegroundColor Green 'Verbindung zu Teams Powershell hergestellt'
    

    #Verbindung zu Exchange Online herstellen
    Connect-ExchangeOnline -UserPrincipalName $strAdmin
    Write-Host -ForegroundColor Green 'Verbindung zu EXO hergestellt'
    

    #Verbindung zu SFB Online herstellen
    Import-Module MicrosoftTeams
    $sfbSession = New-CsOnlineSession
    Import-PSSession $sfbSession
    Write-Host -ForegroundColor Green 'Verbindung zu SFB hergestellt'
  
  Write-Host '***************************'
  Write-Host '     Sind MTR Lizenzen     '
  Write-Host '        vorhanden?         '
  Write-Host '***************************'
  Licensecheck

  ## Collect account data ##
  Write-Host " "
  Write-Host " "
  Write-Host " "
  Write-Host '***************************'
  Write-Host '      UPN und Passwort     '
  Write-Host '    für neues MTR Konto    '
  Write-Host '          eingeben         '
  Write-Host '***************************'
  $script:credNewAccount = (Get-Credential -Message 'Bitte UPN und Passwort für das neue Raumsystem eingeben')
  $strUpn = $credNewAccount.UserName
  $strAlias = $credNewAccount.UserName.substring(0,$credNewAccount.UserName.indexOf('@'))

  

  $strDisplayName = Read-Host -Prompt "Bitte geben sie den Display Namen für $strUpn an"

  
  if (!$credNewAccount -Or [string]::IsNullOrEmpty($strDisplayName) -Or [string]::IsNullOrEmpty($credNewAccount.UserName) -Or $credNewAccount.Password.Length -le 7)
  {
    CleanupAndFail -strMsg 'Please enter all of the requested data to continue.'
    exit 1
  }
  if ($strProvisionMode -eq 1 -or $strProvisionMode -eq 2)
  {
    try 
    {
      $Error.Clear()

   $strPlainPass = $credNewAccount.GetNetworkCredential().Password  
      #Erstellen der TRS Mailbox

      
      

      #Erstellen einer neuen Mailbox für das Raumsystem
      New-Mailbox -Name $strDisplayName -Alias $strAlias -Room -EnableRoomMailboxAccount $true -MicrosoftOnlineServicesID $strUpn -RoomMailboxPassword (ConvertTo-SecureString -String $strPlainPass -AsPlainText -Force)
      
      Write-Host -ForegroundColor Green "Die Mailbox mit der Adresse $strupn wurde erstellt"

      #Konfiguration des Kalenders für das MTR
      Set-CalendarProcessing -Identity $strAlias -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "Das ist ein Microsoft Teams Meeting Raum System!“
     
      Write-Host -ForegroundColor Green "Kalendereinstellungen wurden gesetzt"


      
         
      #Erstellen und zuweisen der Passwort Policy
      
      $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
      $PasswordProfile.Password = $strPlainPass
      $PasswordProfile.ForceChangePasswordNextLogin = $false
      Write-Host -ForegroundColor Green 'Passwort Policy wurde erstellt'
      

      #Erstellen des AAD Accounts für das TRS
     
      Set-AzureADUser -objectid $strUpn -AccountEnabled $True -DisplayName $strDisplayName -PasswordProfile $PasswordProfile -MailNickName $strAlias -UserPrincipalName $credNewAccount.UserName -UsageLocation DE
      
      Write-Host -ForegroundColor Green "TRS Azure AD Konto wurde aktualisiert"

      #Zuweisen der MeetingRoom Lizenz

      $planName="MEETING_ROOM"
      $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
      $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
      $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
      $LicensesToAssign.AddLicenses = $License
      Set-AzureADUserLicense -ObjectId $strUpn -AssignedLicenses $LicensesToAssign
      Write-Host -ForegroundColor Green "Meetingroom Lizenz wurde zugewiesen"
     
      #Ab zur nächsten Funktion
      RegistrarPool


      #Raumliste erstellen
      #New-DistributionGroup -RoomList -Name 'Videoräume' 
      #Add-DistributionGroupMember –Identity Videoräume -Member Alias_Der_Mailbox
    }
    catch
    {
    }
    if ($Error)
    {
      $Error.Clear()
      $status['Azure Account Create'] = 'Failed to create Azure AD room account. Please validate if you have the appropriate access.'
    }
          else
      {
        $status['Azure Account Create'] = "Successfully added $strDisplayName to Azure AD"
      }
  }
}

function InstMods{
   $module=$_
   $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
   $admin=$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
   if ( $admin )
   {
   #Write-host -ForegroundColor green "Adminrechte vorhanden, Installation kann versucht werden"
      $moduleinstalled=Get-InstalledModule -Name $_
   If([string]::IsNullOrEmpty($moduleinstalled))
  {
  Write-Host -ForegroundColor Red "$module Modul ist nicht vorhanden"

   Write-Host "Soll es installiert werden?"

  $strInstModule = Read-Host -Prompt '1 für ja, 2 für nein'
 
  if ($strInstModule -eq 1)
    {
      install-module $module
      $moduleinstalled1=Get-InstalledModule -Name $_
      If([string]::IsNullOrEmpty($moduleinstalled1))
       { Write-Host -ForegroundColor Red "$module Installation fehlgeschlagen"
         Write-Host -ForegroundColor Red "Bitte manuell installieren"
         Write-host -ForegroundColor yellow "Install-Module $module"
         }
    }
    else
    {	
      if ($strInstModule -eq 2)
        {
          Write-Host 'PS Module nicht installiert'
          Start-Sleep -Seconds 3
          
        }
       
          else
           {
                  Clear-Host
                  Write-Host 'Falsche Eingabe'
                  Start-Sleep -Seconds 1
                  Clear-Host
                  Anfangen
           }
            
        }
    }
    else 
     { Write-Host -ForegroundColor Green "$module installiert"}
    }
    else
   {
   #Write-Host -ForegroundColor yellow "Keine Adminrechte, es wird nur geprüft ob die Module vorhanden sind"
   $moduleinstalled=Get-InstalledModule -Name $_
   If([string]::IsNullOrEmpty($moduleinstalled))
        {
        Write-Host -ForegroundColor Red "$module Modul ist nicht vorhanden"
        Write-Host -ForegroundColor Red "Keine Adminrechte vorhanden um die Installtion zu versuchen, bitte PS mit Adminrechten starten"
        }
    else 
        {
        Write-Host -ForegroundColor Green "$module installiert"
        }
}
    }

function modcheck{
  "AzureAD", "ExchangeOnlineManagement", "MicrosoftTeams" | ForEach-Object  {InstMods}
  }

$param1=$args[0]
 If([string]::IsNullOrEmpty($param1))
 {
 Countdown -timespan 2
 Anfangen
 }
 else {
    if ($param1 -eq "modcheck")
    {modcheck}
    else {
    Write-Host -ForegroundColor Red "Skript ohne Parameter starten oder mit Paramter modcheck um die PS Module zu prüfen"
    } 
  }
