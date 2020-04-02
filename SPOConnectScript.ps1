# This is a basic Powershell script using 'Client Side Object Model (CSOM)' 
#      which reads the configuration values and establishes a connection to 
#      Sharepoint Onine. Use this script from other scripts to establish
#      SP Online connection.
#      Input Parameters are 2 
#      1) Environment Variable: Dev/QA/Prod
#      2) URL Selection: 1/2/3
# Written By :      Kalyan, Yalamanchili, 
# Initial Creation: 10/25/2018

Write-Host -ForegroundColor Green "---- STARTED running deployment scripts ----"  

# START Functions

# Parses the Deployment INI file for wsps and features
    Function Parse-IniFile ($file) 
    {
      $global:ConfigSettings = @{}
      switch -regex -file $file {
        "^\[(.+)\]$" {
          $section = $matches[1].Trim()      
          $global:ConfigSettings[$section] = @{}
        }
        "^\s*([^#].+?)\s*=\s*(.*)" {
          $name,$value = $matches[1..2]
          $global:ConfigSettings[$section][$name] = $value.Trim()
        }
      }
      $global:ConfigSettings
    }
    

# End Functions
#----------------------------------------------------------------------------------------------------------------------------


# Region Global Variables

    $ConfigSettings    
    $solutionPath=[Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath
    
    Write-Host "---- ADD CSOM Client Components PowerShell SnapIn ---- "
    Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll"
    Add-Type -Path "C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll"         
    #[System.Reflection.Assembly]::LoadFrom("C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.dll")
    #[System.Reflection.Assembly]::LoadFrom("C:\Program Files\SharePoint Online Management Shell\Microsoft.Online.SharePoint.PowerShell\Microsoft.SharePoint.Client.Runtime.dll")
    Write-Host -ForegroundColor Green "---- SUCCESSFULLY loaded CSOM Client Components PowerShell SnapIn ----"
    
# Get to the ini file and read it in.
    $IniFile = $MyInvocation.MyCommand.Path -replace  "ps1","ini"    
    Parse-IniFile $IniFile

#Create Directory and File for the log - Open the log in append mode.
    $LogDirectory = [Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath 
    [IO.Directory]::CreateDirectory($LogDirectory)
    $currentDate = Get-Date
    $Transcript = [IO.Path]::Combine($LogDirectory, $currentDate.Year.ToString()+$currentDate.Month.ToString()+$currentDate.Day.ToString() + " - SPOConnectScript.log")
        
    Start-Transcript $Transcript -Append  
    Write-Host "$(Get-Date -Format o) GETTING Environment variables --"
    Write-Host

    #write-host 'Args: ' $args[0] '; Args 1: '$args[1]
    
# ALL 'URLS' HERE
    # ALL 'FILTERED URLS' HERE
    IF($args[0] -eq 'Dev') 
     {      
      $webApplnURL = $ConfigSettings['Environment'].DevWebAppln           # URL of the 'Web Application'
      $siteURL = $ConfigSettings['Environment'].DevSiteColln              # URL of the 'Site Collection'  
      $webURL = $ConfigSettings['Environment'].DevWeb                     # URL of the 'Web Site' in a site collection
      $loginID = $ConfigSettings['Environment'].DevAdminLogin             # Admin Login UserName
     }
    ELSEIF ($args[0] -eq 'QA') 
     {
      $webApplnURL = $ConfigSettings['Environment'].QAWebAppln          # URL of the 'Web Application'
      $siteURL = $ConfigSettings['Environment'].QASiteColln             # URL of the 'Site Collection'  
      $webURL = $ConfigSettings['Environment'].QAWeb                    # URL of the 'Web Site' in a site collection
      $loginID = $ConfigSettings['Environment'].QAAdminLogin            # Admin Login UserName
     }
    ELSEIF ($args[0] -eq 'Uat') 
     {
      $webApplnURL = $ConfigSettings['Environment'].UatWebAppln           # URL of the 'Web Application'
      $siteURL = $ConfigSettings['Environment'].UatSiteColln              # URL of the 'Site Collection'  
      $webURL = $ConfigSettings['Environment'].UatWeb                     # URL of the 'Web Site' in a site collection
      $loginID = $ConfigSettings['Environment'].UatAdminLogin             # Admin Login UserName
     }
    ELSEIF ($args[0] -eq 'Prod') 
     {
      $webApplnURL = $ConfigSettings['Environment'].ProdWebAppln          # URL of the 'Web Application'
      $siteURL = $ConfigSettings['Environment'].ProdSiteColln             # URL of the 'Site Collection'  
      $webURL = $ConfigSettings['Environment'].ProdWeb                    # URL of the 'Web Site' in a site collection
      $loginID = $ConfigSettings['Environment'].ProdAdminLogin             # Admin Login UserName
     }
    ELSE  
     {
      Write-Host " "$(Get-Date -Format o) " Environment variable not passed - please pass appropriate environment for deployment -"
      exit
     }

    IF(($siteURL -eq '') -and ($webApplnURL -eq '') -and ($webURL -eq ''))
    {
     Write-Host " "$(Get-Date -Format o) " Unknown Environment - please pass 'Dev/QA/Uat/Prod' -"
     exit
    }

# ALL FILTERED Addin's HERE

    # All Add-in Names
    $globalAddins =  $ConfigSettings['Application'].Raw_SPHosted_Addins.Trim().Split(",") 


# Find which URL user wants to use now
  If($args[1] -eq '1')      { $url = $webApplnURL  }
  ELSEIF($args[1] -eq '2') { $url = $siteURL  }
  ELSEIF($args[1] -eq '3') { $url = $webURL  }
  
  Write-Host -ForegroundColor Gray 'Default URLs: '
  Write-Host ' Web Appln URL:           '$webApplnURL 
  Write-Host ' Site URL:                '$siteURL
  Write-Host ' Web URL:                 '$webURL 
  Write-Host ' URL you selected to use: '$url
  Write-Host ' Login ID:                '$loginID
       
# get login password
  $password = Read-Host -Prompt ' Enter password for the service account ' -AsSecureString
  Write-Host
  
# Hardcoded id, password
  #$loginID="shpqa_mh@kalyan.onmicrosoft.com"
  #$password = Read-Host -Prompt "Enter administrator password" -AsSecureString
     
# Build the Client Context
  $clientContext = New-Object Microsoft.SharePoint.Client.ClientContext($url)
  $clientContext.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($loginID,$password)

#----------------------------------------------------------------------------------------------------------------------------
#----------------------------------------------------------------------------------------------------------------------------

    Stop-Transcript   

    Write-Host "End Connecting Script "