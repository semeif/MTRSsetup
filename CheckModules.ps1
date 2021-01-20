  function InstMods{
   $module=$_
   $ErrorActionPreference = "SilentlyContinue"
   $moduleinstalled=Get-InstalledModule -Name $_
   $ErrorActionPreference = "Continue"
   If([string]::IsNullOrEmpty($moduleinstalled))
  {
  Write-Host -ForegroundColor Red "$module Modul ist nicht vorhanden"

   Write-Host "Soll es installiert werden?"

  $strInstModule = Read-Host -Prompt '1 für ja, 2 zum abbrechen'
 
  if ($strInstModule -eq 1)
    {
      install-module $module
      $moduleinstalled1=Get-InstalledModule -Name $_
      If([string]::IsNullOrEmpty($moduleinstalled1))
       { Write-Host -ForegroundColor Red "$module Installtion fehlgeschlagen"
         Write-Host -ForegroundColor Red "Bitte manuell installieren"
         Write-host -ForegroundColor yellow "Install-Module $module"
         }
    }
    else
    {	
      if ($strInstModule -eq 2)
        {
          Clear-Host
          Write-Host 'Dann eben nicht, vieleicht beim nächsten mal :-)'
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
    else 
     { Write-Host -ForegroundColor Green "$module installiert"}
    }

  "AzureAD", "ExchangeOnlineManagement", "MicrosoftTeams" | ForEach-Object  {InstMods}
