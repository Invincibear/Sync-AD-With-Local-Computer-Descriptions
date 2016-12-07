#requires -RunAsAdministrator
[CmdletBinding()]param()


<#
    .Synopsis
    This script will facilitate synchronizing computer descriptions in Active Directory with the system descriptions stored on each individual computer which by design, are not used by AD.

    .DESCRIPTION
    Enter credentials used to search AD, and same or different credentials to connect to each resulting computer.
    Searches AD for computer objects in AD matching the search term input, connects to each of the resulting computers to pull their system descriptions and compares them in a displayed table.
    
    Options
    [1] Update AD with ALL non-empty local system descriptions matching the entered AD computer name search term
        This will update in one batch all AD search results with their corresponding local system descriptions using the AD credentials supplied earlier. 

    [2] Manually approve each AD update one at a time 
        This will prompt you for approval to update AD with the local system description of each AD computer search result.
#>

## Set this to $False to live the script and remove the -WhatIf flag from AD commands
$Testing        = $True
$LogFile        = "$($PSScriptRoot)\Sync computer description with AD-$(Get-Date -format "yyyy-MM-dd-HHmmss").log"
$EmailResults   = $False
## End of user-configurable variables



###                                 ###
###                                 ###
### Edit the rest at your own risk! ###
###                                 ###
###                                 ###



## Logging functionality
$ErrorActionPreference = "SilentlyContinue"
Stop-Transcript | Out-Null
$ErrorActionPreference = "Continue"
Start-Transcript -Path $LogFile -Force


## Load ActiveDirectory module if it isn't already loaded
If ($(Get-Module | ? {$_.Name -eq "ActiveDirectory"} | Measure).Count -eq 0) {
    Write-Host "Loading the Active Directory PowerShell module"

    Import-Module ActiveDirectory
}


Function Set-ADComputerDescription {
    [CmdletBinding()] Param
    (
        [Parameter(Mandatory=$True)]          $updateList,
        [Parameter(Mandatory=$False)] [String]$activity = "Updating AD computer descriptions"
    )
    
    ForEach ($computer In $updateList) {
        If ($computer[0] -eq $Null) {
            Continue
        }

        If ($Testing) {## Simulate the AD update
            Write-Host "Attempting to update on AD the description of $($computer[0]) to '$($computer[2])'" -ForegroundColor Yellow -BackgroundColor Black
            Write-Progress -Activity $activity -Status "Attempting to update on AD the description of $($computer[0]) to '$($computer[2])'" -PercentComplete ($i / $updateCount * 100)

            Set-ADComputer $computer[0] -Description "$($computer[2])" -Verbose -WhatIf
        } Else {## Perform the AD update
            Write-Host "Attempting to update on AD the description of $($computer[0]) to '$($computer[2])'" -ForegroundColor Yellow -BackgroundColor Black
            Write-Progress -Activity $activity -Status "Attempting to update on AD the description of $($computer[0]) to '$($computer[2])'" -PercentComplete ($i / $updateCount * 100)

            Set-ADComputer $computer[0] -Description "$($computer[2])" -Verbose
        }

        ## Check last operation succeeded and pull AD description to compare to what we just changed it to
        If (($?) -And ($(Get-ADComputer -Identity $computer[0] -Property Description) -eq $computer[2])) {
            Write-Host "Successfully updated AD description of $($computer[0])`n" -ForegroundColor Green -BackgroundColor Black
        } ElseIf ($Testing) {
            Write-Host 'Testing mode enabled, no changes to AD were performed' -ForegroundColor Yellow -BackgroundColor Black
        } Else {
            Write-Host "WARNING: Failed to Update AD description of $($computer[0])`n" -ForegroundColor Red -BackgroundColor Black
        }

        Write-Progress -Activity $activity -PercentComplete ($i++ / $updateCount * 100)
    }
}



Write-Host ''
Write-Host "Search for computers in AD named: (case insensitive, use * for wildcards) eg: *NTBK*, SURWKS*, LAPTOP001" -NoNewline -ForegroundColor Yellow
$searchTerm  = Read-Host -Prompt ' '

If (-Not $searchTerm) {## A search term is required
    Write-Host ''
    Write-Host 'ERROR: Please enter a computer name search term' -ForegroundColor Red -BackgroundColor Black
    Write-Host ''
    Stop-Transcript

    Exit
}


$CredentialAD   = Get-Credential -ErrorAction SilentlyContinue -Message "Enter the credentials that will be used to search AD.`
Your credentials will not be saved.`
Press Esc to use the current credentials used to execute this script."
$CredentialLocal= Get-Credential -ErrorAction SilentlyContinue -Message "Enter the credentials that will be used to connect to search AD for the computers and to connect to the computers to retrieve their local system descriptions.`
Your credentials will not be saved.`
Press Esc to use the current credentials used to execute this script."

If ($CredentialAD) {## Uses the user-entered credentials
    $Computers  = Get-ADComputer -Credential $CredentialAD -Filter "Name -like '$($searchTerm)'" -Property Description
} Else {## Uses the POSH RunAs credentials
    $Computers  = Get-ADComputer -Filter "Name -like '$($searchTerm)'" -Property Description
}


$ComputersCount = $($Computers | Measure).Count

If ($ComputersCount -eq 0) {
    Write-Host "There were no computers found in AD matching computer name search term '$searchTerm'" -ForegroundColor Red -BackgroundColor Black
    Write-Host ''
    Stop-Transcript

    Exit
}


$compareList = @()
$compareList+= ,('Computer Name', 'AD Description', 'Local Description')## Insert table headers

ForEach ($computer in $Computers) {
    Write-Debug "Iterating `$computer=$($computer.Name) in `$Computers"

    If ($Testing) {## Dummy data for testing
        $localComputer = @{Description='Test AD description data'}
    } Else {## Get the system description from the computer. Requires an AD account with sufficient permissions.
        $localComputer = Get-WmiObject -class Win32_OperatingSystem -Credential $CredentialLocal -ComputerName $computer.Name -Property Description -ErrorAction Continue
    }

    If ($localComputer -eq $Null) {## Couldn't connect to local computer to retrieve description, skip this computer
        Write-Warning "Unable to connect to $($computer.Name) to retrieve it's system description, skipping this record"

        Continue
    }

    If (-Not $computer.Description) {
        Write-Warning "$($computer.Name) has no AD description"
    } Else {
        $computer.Description = $computer.Description.Trim()
    }

    If (-Not $localComputer.Description) {
        Write-Warning "$($computer.Name) has no local description"
    } Else {
        $localComputer.Description = $localComputer.Description.Trim()
    }

    If ($computer.Description -ne $localComputer.Description) {
        Write-Warning "$($computer.Name) needs AD description updated to match local description"
    }

    ## Add the computer to the overall list for review
    $compareList += ,($computer.Name, $computer.Description, $localComputer.Description)
}

## Output list of computres in a pretty-looking table
Write-Host ''
Write-Host "Displaying $ComputersCount search result" -NoNewline
If ($ComputersCount -ne 1) {
    Write-Host "s" -NoNewline
}
Write-Host " for '$searchTerm'"
$compareList | % {$_ -Join '|'} | ConvertFrom-Csv -Delimiter '|' | Format-Table -AutoSize -Wrap

## Build new array from previous computer list, Skip first row (contained table headers and not actual computer data), Select only computers with non-empty local system description fields that don't match the AD description
$updateList     = @($compareList | Select -Skip 1 | ? {(($_[2]) -and ($_[1] -notmatch $_[2]))})
$updateCount    = $($updateList | Measure).Count

If ($updateCount -eq 0) {
    Write-Host "The resulting computers' descriptions are already synchronized with AD. No further action is needed." -ForegroundColor Green
    Write-Host ''
    Stop-Transcript

    Exit
}

Write-Host "How would you like to update $updateCount AD system descriptions? (" -NoNewline
Write-Host '* denotes default option' -NoNewline -ForegroundColor Green
Write-Host ')'
Write-Host '  [1] Update AD with ALL of the above non-empty local system descriptions' -ForegroundColor Yellow
Write-Host '  [2] Manually approve each update one at a time' -ForegroundColor Yellow
Write-Host ' *[Q] ' -NoNewline -ForegroundColor Green
Write-Host 'Quit--update nothing and exit this script' -ForegroundColor Yellow
Write-Host 'Select Option' -NoNewline -ForegroundColor Yellow
$updateMethod = Read-Host -Prompt ' '

Switch ($updateMethod) {
    Default {
        Write-Host '[Q]uitting script' -ForegroundColor Red -BackgroundColor Black
        Write-Host ''
        Stop-Transcript

        Exit
    }
    '1' {
        Write-Host 'Selected option [1]'

        $i = 1;
        $activity = "Updating $($updateCount) computers' AD description entries with non-empty and non-matching local system descriptions"
        Write-Progress -Activity $activity -Status "Progress:" -PercentComplete ($i / $updateCount * 100)
        Write-Host "$activity`n"

        Set-ADComputerDescription $updateList $activity

        Write-Host "Finished iterating through list of local computer descriptions with which to update AD`n" -ForegroundColor Green
    }
    '2' {
        Write-Host 'Selected option [2]'

        $i = 1;
        $activity = "Updating $($updateCount) computers' AD description entries with non-empty and non-matching local system descriptions one at a time requiring manual input"
        Write-Progress -Activity $activity -Status "Progress:" -PercentComplete ($i / $updateCount * 100)
        Write-Host "$activity`n"

        ForEach ($computer In $updateList) {
            $fields = @()
            $fields+= ,('Computer Name', 'AD Description', 'Local Description')
            $fields+= ,($computer[0], $computer[1], $computer[2])
            $fields | % {$_ -Join '|'} | ConvertFrom-Csv -Delimiter '|' | Format-Table -AutoSize -Wrap

            Write-Host "Do you want to update the AD description of $($computer[0]) to '$($computer[2])'?" -NoNewline
            Write-Host ' [Y]es' -ForegroundColor Yellow -NoNewline
            Write-Host ', default: ' -NoNewline
            Write-Host '[N]o' -NoNewline -ForegroundColor Green
            Write-Host ', or' -NoNewline
            Write-Host ' [Q]uit' -ForegroundColor Red -NoNewline
            $updateComputer = Read-Host -Prompt ' '

            If ($updateComputer -like 'q*') {
                Write-Host '[Q]uiting script' -ForegroundColor Red -BackgroundColor Black
                Write-Host ''
                Stop-Transcript

                Exit
            }
            If ($updateComputer -notmatch 'y|yes') {# Skip this entry if they answered anything except yes
                Write-Host "`n`n`n"

                Continue
            }

            Set-ADComputerDescription @(@($updateList[$i][0], $updateList[$i][1], $updateList[$i][2]), @()) $activity

            Write-Host "`n`n`n"
        }

        Write-Host "Finished iterating through list of local computer descriptions with which to update AD`n" -ForegroundColor Green
    }
}

Stop-Transcript