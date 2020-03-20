#
# Import required module
#
Install-Module MicrosoftGraph -Scope CurrentUser
Import-Module MicrosoftGraph

#
# Required variables 
#
$clientID = "xxxxx"
$clientSecret = "xxxxx"
$tenantID = "xxxxx"
$redirectURI = "http://localhost"

#
# CSV format, required fields:
# - DisplayName
# - Owner
# - Visibility
#
$pathToCSV = "xxxxx"

# DO NOT EDIT BELOW

#
# Get access token to perform queries on Microsoft Graph
#
$authorizationCode = Get-GraphAuthorizationCode -tenantID $tenantID -clientID $clientID -redirectURI $redirectURI
$accessToken = Get-GraphAccessTokenByCode -tenantID $tenantID -clientID $clientID -clientSecret $clientSecret -redirectURI $redirectURI -authorizationCode $authorizationCode.AuthCodeCredential.GetNetworkCredential().password

$importCsv = Import-Csv "$($pathToCSV)"

foreach($rowInCsv in $importCsv)
{
    Write-Output "Looking for existing group with display name '$($rowInCsv.DisplayName)'"
    $existingTeam = Find-GraphGroupByName -accessToken $accessToken -displayName "$($rowInCsv.DisplayName)"

    #
    # If team doesn't exist, create it
    #
    if($existingTeam.Count -eq 0)
    {
        Write-Output "Group doesn't exist yet. Looking for owner with username '$($rowInCsv.Owner)'"
        $teamOwner = ""
        $teamOwner = Get-GraphUser -accessToken $accessToken -username "$($rowInCsv.Owner)"

        #
        # Only proceed if owner exists in AAD
        # 
        if($teamOwner -ne "" -and $rowInCsv.Owner -ne "")
        {
            Write-Output "Owner found, going to create the new team."

            #
            # Construct the request to create a new team in Teams based on a template
            #
            $jsonTeamToCreate = @"
            {
                "template@odata.bind": "https://graph.microsoft.com/beta/teamsTemplates('educationClass')",
                "displayName": "$($rowInCsv.DisplayName)",
                "description": "$($rowInCsv.DisplayName)",
                "owners@odata.bind": [
                "https://graph.microsoft.com/v1.0/users/$($teamOwner.id)"
                ]
            }
"@

            #
            # Create new team
            #
            $createdTeam = Invoke-RestMethod -Method POST -Uri "https://graph.microsoft.com/beta/teams" -Body $jsonTeamToCreate -Headers @{"Authorization" = "Bearer $($accessToken.AccessTokenCredential.GetNetworkCredential().password)"} -ContentType "application/json"
            
            #
            # Wait 15 seconds and Check if team was created successfully
            #
            Write-Output "Team created, waiting 15 seconds to let Microsoft do their work."
            Start-Sleep -Seconds 15
            $createdTeamLookup = Find-GraphGroupByName -accessToken $accessToken -displayName "$($rowInCsv.DisplayName)"

            if($createdTeamLookup.Count -gt 0)
            {
                Write-Output "Created team $($rowInCsv.DisplayName) with ID $($createdTeamLookup.id)"
            }
            else
            {
                Write-Output "Error creating team $($rowInCsv.DisplayName)"
            }
        }
    }
    else 
    {
        Write-Output "Skipped $($rowInCsv.DisplayName); group already exists"
    }
}
