#   Sample script which connects to sharepoint online and retrieves some basic info.
#   Use SPOConnect.ps1 to retrieve the parameters followed by execute and custom logic
#   Written By :      Kalyan, Yalamanchili
#   Initial Creation: 6/09/2019

#Create Directory and File for the log - Open the log in append mode.
    $LogDirectory = [Environment]::CurrentDirectory=(Get-Location -PSProvider FileSystem).ProviderPath 
    [IO.Directory]::CreateDirectory($LogDirectory)
    $currentDate = Get-Date
    $Transcript = [IO.Path]::Combine($LogDirectory, $currentDate.Year.ToString()+$currentDate.Month.ToString()+$currentDate.Day.ToString() + " - SPODoWork.log")
        
    Start-Transcript $Transcript -Append  

#----------------------------------------------------------------------------------------------------------------------------

#Script Starts here
    $envVariable=Read-Host -Prompt "Provide the environment here like Dev QA Prod?"
    $siteVariable=Read-Host -Prompt "Which URL would you like to use 1)WebApplication 2)SiteCollection 3)Web ? Enter 1/2/3"

    IF($envVariable -ne $null -and $siteVariable -ne $null){

       # Loads all the necessary parameters like urls, credentials to authenticate
       . .\SPOConnectScript.ps1 $envVariable $siteVariable
        Write-Host -ForegroundColor Magenta ' Authenticated and returned Context.'
        Write-Host
 
        $rootWeb = $clientContext.Web 
        $childWebs = $rootWeb.Webs 
        $clientContext.Load($rootWeb) 
        $clientContext.Load($childWebs) 
        $clientContext.ExecuteQuery()  

        Write-host ' After Client Context Execute '     
 
        Write-Host
        foreach ($WebAddIn in $childWebs) 
        { 
        
         Write-Host -ForegroundColor Magenta ' Child Site Title: ' $WebAddIn.Title        
         If($WebAddIn.Title -eq 'Investments'){
                 
             Write-Host -ForegroundColor Green "Subsite URL is " $WebAddIn.Url 
             Write-Host -ForegroundColor Green "Subsite ID is  " $WebAddIn.ID 
             Write-Host
             $lists = $WebAddIn.Lists 
             $clientContext.Load($lists) 
             $clientContext.ExecuteQuery() 

             Foreach($list in $lists)
             {
               Write-Host 'List Title: ' $list.Title '; ID: '$list.ID
             }

         } # WebAddIn
        } # Foreach WebAddIn
      }

  #----------------------------------------------------------------------------------------------------------------------------

    Stop-Transcript   

    Write-Host "End Connecting Script "