###############################################################################
#    Powershell Script for creating Microsoft Teams Room accounts
#
#    Powershell Skript zur Erstellung von Microsoft Teams Room Konten
#
#    Prerequisites for Office 365:
#    Voraussetzungen für Office 365:
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
# Globale Variablen
#

$script:strDevice = $null
$script:credAD = $null
$script:credExchange = $null
$script:credSkype = $null
$script:credNewAccount = $null
$script:strHybrid = $null
$script:strEasPolicy = $null
$script:strDatabase = $null
$script:Text25Stars = '*************************'

#$script:strUpn = $null

#Setzen der gewünschten Lizenzen, per Standart ist "Meeting_Room" voreingestellt
#Weiter gehen z.B. ENTERPRISEPREMIUM = O365 E5 oder SPE_E5 = M365 E5
$script:availableLicense = 'MEETING_ROOM'

$status = @{}

#Verify the client language. If this is set to de-DE, all queries are placed in german otherwise in English
#Überprüfen der Clientsprache. Wenn diese auf de-DE Gesetzt ist werden alle Abfragen auf de-DE gestellt sonst auf Englisch

$Culture = Get-Culture
$Language = $Culture.Name


function Anfangen {
  Clear-Host

  if ($Language -eq "de-DE")
  {
    Write-Host $Text25Stars
    Write-Host '*Möchten Sie mit der MTR*'
    Write-Host '* Einrichtung starten?  *' 
    Write-Host $Text25Stars
  }
  else {
    Write-Host $Text25Stars
    Write-Host '* Do you want to start  *'
    Write-Host '*    the MTR setup?     *' 
    Write-Host $Text25Stars 
  }
  if ($Language -eq "de-DE")
  {
    $strProvisionMode = Read-Host -Prompt '1 für ja, 2 zum abbrechen'
  }
  else {
    $strProvisionMode = Read-Host -Prompt '1 yes, 2 no'
  }
  if ($strProvisionMode -eq 1)
    {
      CreateCloudAD
    }
    else
    {	
      if ($strProvisionMode -eq 2)
        {
          Clear-Host
          if ($Language -eq "de-DE")
          {
          Write-Host 'Dann eben nicht, vielleicht beim nächsten mal :-)'
          }
          else {
            Write-Host 'Maybe next time :-)'
          }
          Start-Sleep -Seconds 3
          Clear-Host
          CleanupAndFail
        }
       
          else
           {
                  Clear-Host
                  if ($Language -eq "de-DE")
                    {
                    Write-Host 'Falsche Eingabe'
                    }
                  else {
                    Write-Host 'Wrong input'
                  }
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
     if ($Language -eq "de-DE")
    {
        Write-Host -ForegroundColor Green "Warte 2 Minuten auf RegistrarPool"
    }

    else {
        Write-Host -ForegroundColor Green "Wait 2 minutes for RegistrarPool"
    }

      Countdown -timeSpan 120

      #RegistrarPool auslesen
      #Get RegistrarPool
     
      $strRegPool = (Get-CsTenant).RegistrarPool

      

      #MTR Aktivieren
      #Activate MTR
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
     if ($Language -eq "de-DE")
    {
        Write-Host -ForegroundColor Yellow "Der RegistrarPool ist noch nicht erstellt. Ich Versuche es weiter. Es kann bis zu 10 Minuten dauern."
    }
    else {
        Write-Host -ForegroundColor Yellow "The RegistrarPool is not yet created. I keep trying. Can take up to 5 cycles"
    } 
    Countdown -timeSpan 5
     RegistrarPool
    }
          else
      {
      Write-Host " "
        if ($Language -eq "de-DE")
        {
        Write-Host -ForegroundColor Green "Der Registrierungspool lautet: $strRegPool"
        Write-Host " "
        Write-Host -ForegroundColor Green "Der Meeting Room $strDisplayName wurde erstellt"
        }
        else {
        Write-Host -ForegroundColor Green "Registrierungspool name: $strRegPool"
        Write-Host " "
        Write-Host -ForegroundColor Green "The Meetingroom $strDisplayName has been created"           
        }
      Pause
      Anfangen
      }
    }
function Connect2AzureAD {
  try
  {
    Clear-Host
    if ($Language -eq "de-DE")
    {
    Write-Host '***************************'
    Write-Host '      UPN und Passwort     '
    Write-Host ' für Azure AD Admin Konto  '
    Write-Host '          eingeben         '
    Write-Host '***************************'
    
    $strAdmin = Read-Host -Prompt "Bitte geben Sie ihr Admin Konto an. Damit werden Sie an allen Konsolen angemeldet"
    }
    else {
    Write-Host '***************************'
    Write-Host '   Enter admin UPN and     '
    Write-Host '   password credentials    '
    Write-Host '       to sign in          '
    Write-Host '***************************'
    
    $strAdmin = Read-Host -Prompt "Please enter your admin credentials. This will sign you in to all consoles"    
    }  

    #Mit Azure AD Powershell verbinden
    #Connect to Azure AD PowerShell

      Connect-AzureAD -AccountId $strAdmin
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "Verbindung zu Azure AD Powershell hergestellt"
      }
      else {
        Write-Host -ForegroundColor Green "Connected to Azure AD Powershell"
      }
      Countdown -timeSpan 3
      
      #An Microsoft Teams PowerShell anmelden
      #Connect to Microsoft Teams PowerShell

      Connect-MicrosoftTeams -AccountId $strAdmin
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green 'Verbindung zu Microsoft Teams Powershell hergestellt'
      }
      else {
      Write-Host -ForegroundColor Green 'Connected to Microsoft Teams Powershell'
      }
      Countdown -timeSpan 3

      #Verbindung zu Exchange Online herstellen
      #Connect to Exchange Online PowerShell

      Connect-ExchangeOnline -UserPrincipalName $strAdmin
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green 'Verbindung zu Exchange Online PowerShell hergestellt'
      }
      else {
      Write-Host -ForegroundColor Green 'Connected to Exchange Online PowerShell'
      }
      Countdown -timespan 3

      #Verbindung zu CS Online PowerShell herstellen
      #Connect to CS online PowerShell

      Import-Module MicrosoftTeams
      $sfbSession = New-CsOnlineSession
      Import-PSSession $sfbSession
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green 'Verbindung zu CS Online hergestellt'
      }
      else {
    Write-Host -ForegroundColor Green 'Connected to CS Online PowerShell'
      }
      Countdown -timespan 3

  }
  catch
  {
    if ($Language -eq "de-DE"){
        CleanupAndFail -strMsg "Fehler beim Herstellen einer Verbindung mit den PowerShell Konsolen. Bitte die Zugangsdaten überprüfen. Fehlermeldung: $_"
    }
    else {
        CleanupAndFail -strMsg "Failed to connect to PowerShell Consoles. Please check your credentials and try again. Error message: $_"
    }
  }
}

function Licensecheck {

  #Überprüfung ob MTR Lizenzen im Tenant vorhanden sind
  #Check availability of MTR Licenses in Tenant
  
  if ($Language -eq "de-DE"){
    Write-Host '***************************'
    Write-Host '     Sind MTR Lizenzen     '
    Write-Host '        vorhanden?         '
    Write-Host '***************************'
  }
  else {
    Write-Host '***************************'
    Write-Host '     Are MTR licenses      '
    Write-Host '        available?         '
    Write-Host '***************************'    
  }
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
    if ($Language -eq "de-DE"){
        Write-Host -ForegroundColor Green 'Super, MTR Lizenzen sind vorhanden!'
    }
    else {
        Write-Host -ForegroundColor Green 'Perfect, MTR licenses are available!'
    }
    Countdown -timeSpan 5
      }
        else
        {
            if ($Language -eq "de-DE"){
            Write-Host -ForegroundColor Red 'Es sind leider keine Meeting Room Lizenzen vorhanden.'
            }
            else {
            Write-Host -ForegroundColor Red 'Sorry, there are No MTR licenses available.'
            }
            Pause
        CleanupAndFail

        }
}


function CreateCloudAD {
  
Connect2AzureAD



Licensecheck
  #Abfrage der MTR Daten
  # Collect account data 
  Write-Host " "
  Write-Host " "
  Write-Host " "
  Write-Host '***************************'
  if ($Language -eq "de-DE"){
  Write-Host '      UPN und Passwort     '
  Write-Host '    für neues MTR Konto    '
  Write-Host '          eingeben         '
  Write-Host '***************************'
  }
  else {
    Write-Host '       Enter UPN and       '
    Write-Host '      for the new MTR      '
    Write-Host '          account          '
    Write-Host '***************************'
  }
  if ($Language -eq "de-DE"){
  $script:credNewAccount = (Get-Credential -Message 'Bitte UPN und Passwort für das neue Raumsystem eingeben:')
  }
  else {
    $script:credNewAccount = (Get-Credential -Message 'Please enter UPN and Password for the new MTR:') 
  }
  $strUpn = $credNewAccount.UserName
  $strAlias = $credNewAccount.UserName.substring(0,$credNewAccount.UserName.indexOf('@'))

  
  if ($Language -eq "de-DE"){
  $strDisplayName = Read-Host -Prompt "Bitte geben sie den Display Namen für $strUpn an"
  }
  else {
    $strDisplayName = Read-Host -Prompt "Please enter the Display Name for $strUpn"
  }
  
  if (!$credNewAccount -Or [string]::IsNullOrEmpty($strDisplayName) -Or [string]::IsNullOrEmpty($credNewAccount.UserName) -Or $credNewAccount.Password.Length -le 7)
  {
    if ($Language -eq "de-DE"){
    CleanupAndFail -strMsg 'Please enter all of the requested data to continue.'
    }
    else {
        CleanupAndFail -strMsg 'Bitte geben Sie alle benötigten Informationen ein um fortzufahren'
    }
    exit 1
  }


   $strPlainPass = $credNewAccount.GetNetworkCredential().Password  
      #Erstellen der TRS Mailbox
      #Creation of MTR Mailbox Account      
      
      New-Mailbox -Name $strDisplayName -Alias $strAlias -Room -EnableRoomMailboxAccount $true -MicrosoftOnlineServicesID $strUpn -RoomMailboxPassword (ConvertTo-SecureString -String $strPlainPass -AsPlainText -Force)
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "Die Mailbox mit der Adresse $strupn wurde erstellt"
      }
      else {
        Write-Host -ForegroundColor Green "Mailbox for $strupn has been created"
      }
      #Konfiguration des Kalenders für das MTR
      #Calendar Configuration for MTR
      Set-CalendarProcessing -Identity $strAlias -AutomateProcessing AutoAccept -AddOrganizerToSubject $false -DeleteComments $false -DeleteSubject $false -RemovePrivateProperty $false -AddAdditionalResponse $true -AdditionalResponse "Das ist ein Microsoft Teams Meeting Raum System!"
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "Kalendereinstellungen wurden gesetzt"
      }
      else {
        Write-Host -ForegroundColor Green "Calendar configuration has been set"
      }
         
      #Erstellen und zuweisen der Passwort Policy
      #Password policy creation
      
      $PasswordProfile = New-Object -TypeName Microsoft.Open.AzureAD.Model.PasswordProfile
      $PasswordProfile.Password = $strPlainPass
      $PasswordProfile.ForceChangePasswordNextLogin = $false
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green 'Passwort Policy wurde erstellt'
      }
      else {
        Write-Host -ForegroundColor Green 'Password Policy has been created'
      }
      #Erstellen des AAD Accounts für das TRS
      #Creation and configuration for Azure AD account
      if ($Language -eq "de-DE"){
        Write-Host -ForegroundColor Green 'Aktualisierung des Azure AD Kontos'
        }
        else {
          Write-Host -ForegroundColor Green 'Update Azure AD account'
        }
        if ($Language -eq "de-DE"){
            Write-Host "In welchem Land wird der Meetingraum verwendet?"
            }
            else {
              Write-Host "In which country is the Meetingroom deployed?"
            }
            if ($Language -eq "de-DE"){
                $strUsageLocation = Read-Host -Prompt "Bitte geben sie den Ländercode für $strUpn an (z.B. US, DE, UK FR...)"
                }
                else {
                  $strUsageLocation = Read-Host -Prompt "Please enter the country code for $strUpn (e.g. US, DE, UK FR...)"
                }
                
      Set-AzureADUser -objectid $strUpn -AccountEnabled $True -DisplayName $strDisplayName -PasswordProfile $PasswordProfile -MailNickName $strAlias -UserPrincipalName $credNewAccount.UserName -UsageLocation $strUsageLocation
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "TRS Azure AD Konto wurde aktualisiert"
      }
      else {
        Write-Host -ForegroundColor Green "Azure AD account for MTR has been configured"
      }

      #Zuweisen der MeetingRoom Lizenz
      #Assign MTR license

      $planName="MEETING_ROOM"
      $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
      $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
      $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
      $LicensesToAssign.AddLicenses = $License
      Set-AzureADUserLicense -ObjectId $strUpn -AssignedLicenses $LicensesToAssign
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "Meetingroom Lizenz wurde zugewiesen"
      }
      else {
        Write-Host -ForegroundColor Green "MTR license has been assigned"
      }
      #Ab zur nächsten Funktion
      #Calling next function
      RegistrarPool
      #Raumliste erstellen
      #New-DistributionGroup -RoomList -Name 'Videoräume' 
      #Add-DistributionGroupMember –Identity Videoräume -Member Alias_Der_Mailbox

    if ($Error)
    {
      $Error.Clear()
      if ($Language -eq "de-DE"){
      $status['Azure Account Create'] = 'Fehler beim Erstellen eines MTR Kontos. Bitte überprüfen Sie, ob Sie über die entsprechenden Berechtigungen verfügen.'
      }
      else {
        $status['Azure Account Create'] = 'Failed to create Azure AD room account. Please validate if you have the appropriate access.'
      }
    }
          else
      {
        if ($Language -eq "de-DE"){
        $status['Azure Account Create'] = "$strDisplayName erfolgreich zum Azure AD hinzugefügt"
        }
        else {
            $status['Azure Account Create'] = "Successfully added $strDisplayName to Azure AD"
        }
      }

      #Zuweisen der MeetingRoom Lizenz
      #Assign MTR license

      $planName="MEETING_ROOM"
      $License = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicense
      $License.SkuId = (Get-AzureADSubscribedSku | Where-Object -Property SkuPartNumber -Value $planName -EQ).SkuID
      $LicensesToAssign = New-Object -TypeName Microsoft.Open.AzureAD.Model.AssignedLicenses
      $LicensesToAssign.AddLicenses = $License
      Set-AzureADUserLicense -ObjectId $strUpn -AssignedLicenses $LicensesToAssign
      if ($Language -eq "de-DE"){
      Write-Host -ForegroundColor Green "Meetingroom Lizenz wurde zugewiesen"
      }
      else {
        Write-Host -ForegroundColor Green "MTR license has been assigned"
      }
      #Ab zur nächsten Funktion
      #Calling next function
      RegistrarPool
      #Raumliste erstellen
      #New-DistributionGroup -RoomList -Name 'Videoräume' 
      #Add-DistributionGroupMember –Identity Videoräume -Member Alias_Der_Mailbox
    }

    if ($Error)
    {
      $Error.Clear()
      if ($Language -eq "de-DE"){
      $status['Azure Account Create'] = 'Fehler beim Erstellen eines MTR Kontos. Bitte überprüfen Sie, ob Sie über die entsprechenden Berechtigungen verfügen.'
      }
      else {
        $status['Azure Account Create'] = 'Failed to create Azure AD room account. Please validate if you have the appropriate access.'
      }
    }
          else
      {
        if ($Language -eq "de-DE"){
        $status['Azure Account Create'] = "$strDisplayName erfolgreich zum Azure AD hinzugefügt"
        }
        else {
            $status['Azure Account Create'] = "Successfully added $strDisplayName to Azure AD"
        }
      }
  

function InstMods{
   $module=$_
   $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
   $admin=$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
   if ( $admin )
   {
    if ($Language -eq "de-DE"){
    Write-host -ForegroundColor green "Adminrechte vorhanden, Installation kann versucht werden."
    }
    else {
        Write-host -ForegroundColor green "Admin rights checked. Try to install Modules."
    }
      $moduleinstalled=Get-InstalledModule -Name $_
   If([string]::IsNullOrEmpty($moduleinstalled))
  {
    if ($Language -eq "de-DE"){
        Write-Host -ForegroundColor Red "$module Modul ist nicht vorhanden"
        Write-Host "Soll es installiert werden?"
        $strInstModule = Read-Host -Prompt '1 für ja, 2 für nein'
    }
    else {
        Write-Host -ForegroundColor Red "$module Module is not installed"
        Write-Host "Do you want to install it?"
        $strInstModule = Read-Host -Prompt '1 yes, 2 no'
    }
 
  if ($strInstModule -eq 1)
    {
      install-module $module
      $moduleinstalled1=Get-InstalledModule -Name $_
      If([string]::IsNullOrEmpty($moduleinstalled1)){ 
        if ($Language -eq "de-DE"){
          Write-Host -ForegroundColor Red "$module Installation fehlgeschlagen"
          Write-Host -ForegroundColor Red "Bitte manuell installieren"
          Write-host -ForegroundColor yellow "Install-Module $module"
        }
        else {
            Write-Host -ForegroundColor Red "$module Installation failed"
            Write-Host -ForegroundColor Red "Please install manually"
            Write-host -ForegroundColor yellow "Install-Module $module"
        }
         }
    }
    else
    {	
      if ($strInstModule -eq 2)
        {
            if ($Language -eq "de-DE"){
                Write-Host 'PS Module nicht installiert'
            }
            else {
                Write-Host 'PS Module not installed'
            }
          Start-Sleep -Seconds 3
          
        }
       
          else
           {
            if ($Language -eq "de-DE"){   
                  Clear-Host
                  Write-Host 'Falsche Eingabe'
            }
            else {
                Clear-Host
                Write-Host 'Invalid input'
            }
                  Start-Sleep -Seconds 1
                  Clear-Host
                  Anfangen
           }
            
        }
    }
    else { 
        if ($Language -eq "de-DE"){
            Write-Host -ForegroundColor Green "$module installiert"
        }
        else {
            Write-Host -ForegroundColor Green "$module installed"
        }
    }
    }
    else
   {
    if ($Language -eq "de-DE"){
        Write-Host -ForegroundColor yellow "Keine Adminrechte, es wird nur geprüft ob die Module vorhanden sind"
    }
    else {
        Write-Host -ForegroundColor yellow "No Admin access, now checking if modules are installed"
    }
   $moduleinstalled=Get-InstalledModule -Name $_
   If([string]::IsNullOrEmpty($moduleinstalled))
        {
            if ($Language -eq "de-DE"){
        Write-Host -ForegroundColor Red "$module Modul ist nicht vorhanden"
        Write-Host -ForegroundColor Red "Keine Adminrechte vorhanden um die Installtion zu versuchen, bitte PS mit Adminrechten starten"
            }
            else {
                Write-Host -ForegroundColor Red "$module not installed"
                Write-Host -ForegroundColor Red "No admin access to install modules, pleas start console with admin rights"
            }
        }
    else 
        {
            if ($Language -eq "de-DE"){
                Write-Host -ForegroundColor Green "$module installiert"
            }
            else {
                Write-Host -ForegroundColor Green "$module installed"
            }
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
        if ($Language -eq "de-DE"){
            Write-Host -ForegroundColor Red "Skript ohne Parameter starten oder mit Paramter modcheck um die PS Module zu prüfen"
        }
        else {
            Write-Host -ForegroundColor Red "Start script without parameters or with Paramter modcheck to check the PS modules"
        }
    } 
  }
