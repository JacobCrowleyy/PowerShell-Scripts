<#
.SYNOPSIS
    Audits all Microsoft Teams in a tenant to identify candidates for deletion.

.DESCRIPTION
    This script connects to Microsoft Graph using an App Registration (Client ID & Secret)
    to perform a comprehensive audit of all Teams. It gathers information on membership,
    owner details, member departments, last user activity date, and SharePoint storage usage.

    The script analyzes each Team and its channels to determine if it contains any files
    and calculates a "Safe to Delete" recommendation based on user-defined inactivity
    thresholds, member counts, and file presence.

    The final output is a multi-sheet Excel file containing:
    1. A detailed per-channel report.
    2. A high-level per-team summary and deletion recommendation.

.NOTES
    Required Microsoft Graph API Permissions (Application Permissions):
    - Channel.ReadBasic.All
    - Group.Read.All
    - Reports.Read.All
    - Sites.Read.All
    - TeamMember.Read.All
    - User.Read.All
#>

# --- Stuff to Configure ---
# App registration details from Azure/Entra ID.
$tenantId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$clientId = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'
$clientSecret = 'xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx'

# Defines the inactivity threshold for the "Safe to Delete" recommendation.
$safeToDeleteThresholdDays = 90

# --- Output File Configuration ---
$outputFile = "C:\Temp\TeamsAudit_Report_$(Get-Date -Format 'yyyy-MM-dd_HHmm').xlsx"

# --- Script Initialization ---
# Ensures output directory exists.
$outputDir = Split-Path -Path $outputFile -Parent
if (-not (Test-Path -Path $outputDir)) {
    New-Item -Path $outputDir -ItemType Directory -Force | Out-Null
}

# --- 1. Authentication ---
Write-Host "Connecting to Microsoft Graph..." -ForegroundColor Cyan
try {
    # Using client secret to authenticate.
    $secureSecret = ConvertTo-SecureString -String $clientSecret -AsPlainText -Force
    $credential = New-Object -TypeName System.Management.Automation.PSCredential -ArgumentList $clientId, $secureSecret

    # Connects to my graph.
    Connect-MgGraph -TenantId $tenantId -Credential $credential -NoWelcome

    # Confirms connection details.
    $context = Get-MgContext
    Write-Host "Connection successful to tenant '$($context.TenantId)'." -ForegroundColor Green
    Write-Host "   Granted Scopes: $($context.Scopes -join ', ')"
}
catch {
    Write-Host "Error: Failed to connect to Microsoft Graph." -ForegroundColor Red
    Write-Host "   Details: $($_.Exception.Message)"
    return
}

# --- 2. Main Processing ---
$detailedResults = [System.Collections.Generic.List[object]]::new()
$summaryResults = [System.Collections.Generic.List[object]]::new()
$safeToDeleteCutoffDate = (Get-Date).AddDays(-$safeToDeleteThresholdDays)

# Fetch the activity report once for all teams.
Write-Host "`nFetching Teams activity report..." -ForegroundColor Cyan
$activityData = @{}
$tempReportFile = Join-Path -Path $env:TEMP -ChildPath "temp_activity_report.csv"
try {
    # Hides the progress bar to prevent the cosmetic bug.
    $ProgressPreference = 'SilentlyContinue'
    Get-MgReportTeamActivityDetail -Period "D180" -OutFile $tempReportFile
    $ProgressPreference = 'Continue'
    
    if ((Test-Path $tempReportFile) -and (Get-Item $tempReportFile).Length -gt 0) {
        $activityReport = Import-Csv -Path $tempReportFile
        foreach($entry in $activityReport){
            if ($entry."Team Id") {
                $activityData[$entry."Team Id"] = $entry."Last Activity Date"
            }
        }
        Write-Host "Activity report loaded successfully." -ForegroundColor Green
    } else {
        Write-Host "Warning: The activity report was empty or could not be created." -ForegroundColor Yellow
    }
}
catch {
    Write-Host "Error: Could not get the Teams activity report. This may be due to tenant licensing." -ForegroundColor Red
    Write-Host "Full Error Details: $($_.Exception.Message)" -ForegroundColor Yellow
}
finally {
    if (Test-Path $tempReportFile) { Remove-Item $tempReportFile -Force }
}

# Gets all M365 Groups that are Teams.
Write-Host "`nFetching all Teams..." -ForegroundColor Cyan
try {
    $ProgressPreference = 'SilentlyContinue'
    $teams = Get-MgGroup -Filter "resourceProvisioningOptions/Any(x:x eq 'Team')" -All -ErrorAction Stop
    $ProgressPreference = 'Continue'
    $totalTeams = $teams.Count
    Write-Host "Found $totalTeams Teams to process." -ForegroundColor Green
}
catch {
    Write-Host "Error: Failed to retrieve Teams list." -ForegroundColor Red
    Write-Host "   Details: $($_.Exception.Message)"
    return
}

# Loops through teams list.
$currentTeamIndex = 0
foreach ($team in $teams) {
    $currentTeamIndex++
    Write-Host "`nProcessing Team $currentTeamIndex of ${totalTeams}: '$($team.DisplayName)'"

    try {
        # Look up last activity date
        $lastActivityDateVal = $null
        $isTeamInactiveForDeletion = $true
        if ($activityData.ContainsKey($team.Id)) {
            $dateValue = $activityData.Get_Item($team.Id)
            if (-not [string]::IsNullOrWhiteSpace($dateValue)) {
                $lastActivityDateVal = [datetime]$dateValue
                if ($lastActivityDateVal -ge $safeToDeleteCutoffDate) {
                    $isTeamInactiveForDeletion = $false
                }
            }
        }
        $lastActivityDateStr = if ($lastActivityDateVal -is [datetime]) { $lastActivityDateVal.ToString("yyyy-MM-dd") } else { "No Activity in Period" }

        # Get Live Owner/Member details from API
        Write-Host "  -> Getting member and owner details..."
        $liveOwners = @(Get-MgGroupOwner -GroupId $team.Id -All)
        $liveMembers = @(Get-MgTeamMember -TeamId $team.Id -All)
        
        # ** THE FIX IS HERE **
        # Correctly get Owner Names
        $ownerNames = ""
        if ($liveOwners.Count -gt 0) {
            # The DisplayName for owner objects is in AdditionalProperties. This robustly gets names and filters out any empty ones.
            $ownerNameList = $liveOwners | ForEach-Object { $_.AdditionalProperties.displayName } | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            $ownerNames = $ownerNameList -join '; '
        }
        
        # Correctly get Member Names and Departments
        $departmentList = ""
        $memberNames = ""
        if ($liveMembers.Count -gt 0) {
            # Get member display names
            $memberNameList = $liveMembers.DisplayName | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
            
            # Format long member lists
            if ($memberNameList.Count -gt 5) {
                $formattedNames = foreach ($name in $memberNameList) {
                    $nameParts = $name.Split(' ')
                    $firstName = $nameParts[0]
                    $lastNameInitial = if ($nameParts.Count -gt 1) { $nameParts[-1][0] } else { "" }
                    "$(if($firstName.Length -gt 4){$firstName.Substring(0,4)}else{$firstName})$lastNameInitial"
                }
                $memberNames = $formattedNames -join '; '
            } else {
                $memberNames = $memberNameList -join '; '
            }

            # Get department property with stricter validation and batching
            # Robustly collect user IDs, ignoring non-user members that cause errors.
            $potentialUserIds = [System.Collections.Generic.List[string]]::new()
            foreach ($member in $liveMembers) {
                if ($member.AdditionalProperties.ContainsKey('userId')) {
                    $userId = $member.AdditionalProperties.userId
                    # Ensure the retrieved ID is a non-empty string before adding.
                    if (-not [string]::IsNullOrWhiteSpace($userId)) {
                        $potentialUserIds.Add($userId)
                    }
                }
            }

            # Final validation that we only have GUIDs
            $validMemberUserIds = $potentialUserIds | Where-Object { $_ -match '^[0-9a-fA-F]{8}-([0-9a-fA-F]{4}-){3}[0-9a-fA-F]{12}$' }

            if ($validMemberUserIds.Count -gt 0) {
                try {
                    $allUsers = [System.Collections.Generic.List[object]]::new()
                    $batchSize = 15 # Max IDs allowed in a single 'in' filter

                    for ($i = 0; $i -lt $validMemberUserIds.Count; $i += $batchSize) {
                        # Take a slice of the array for the current batch
                        $idBatch = $validMemberUserIds | Select-Object -Skip $i -First $batchSize
                        $filterString = "id in ('$($idBatch -join "','")')"
                        
                        # The @() ensures the result is always an array, even if only one user is returned.
                        $userBatch = @(Get-MgUser -Filter $filterString -Property "id,department")
                        
                        # AddRange is now safe because $userBatch is guaranteed to be a collection.
                        $allUsers.AddRange($userBatch)
                    }
                    
                    $deptList = $allUsers.Department | Where-Object { -not [string]::IsNullOrWhiteSpace($_) }
                    $departmentList = ($deptList | Get-Unique) -join '; '
                } catch {
                    Write-Warning "Could not retrieve department info for members of team $($team.DisplayName). Error: $($_.Exception.Message)"
                    $departmentList = "Error retrieving"
                }
            }
        }
        
        $members = $liveMembers.Count
        $owners = $liveOwners.Count
        $isTeamEmptyOfPeople = ($members -eq 0 -and $owners -eq 0)

        # Gets the primary SharePoint Site and its total size.
        Write-Host "  -> Finding SharePoint site..."
        $sharePointSite = Get-MgGroupSite -GroupId $team.Id -SiteId "root" -ErrorAction Stop
        $defaultDrive = Get-MgSiteDefaultDrive -SiteId $sharePointSite.Id -ErrorAction Stop
        $rootFolder = Get-MgDriveRoot -DriveId $defaultDrive.Id -ErrorAction Stop
        $totalLibrarySize = $rootFolder.Size
        $totalLibrarySizeFormatted = "N/A"
        if ($totalLibrarySize -ge 1GB) { $totalLibrarySizeFormatted = "$([math]::Round($totalLibrarySize/1GB, 2)) GB" }
        elseif ($totalLibrarySize -ge 1MB) { $totalLibrarySizeFormatted = "$([math]::Round($totalLibrarySize/1MB, 2)) MB" }
        elseif ($totalLibrarySize -ge 1KB) { $totalLibrarySizeFormatted = "$([math]::Round($totalLibrarySize/1KB, 2)) KB" }
        else { $totalLibrarySizeFormatted = "$totalLibrarySize Bytes" }
        Write-Host "  -> Found site. Total library size: $totalLibrarySizeFormatted."
        
        # Gets official list of channels and all folders from SharePoint.
        $officialChannels = @(Get-MgTeamChannel -TeamId $team.Id -All)
        $rootItems = @(Get-MgDriveItemChild -DriveId $defaultDrive.Id -DriveItemId $rootFolder.Id -All)
        $sharepointFolders = $rootItems | Where-Object { $_.Folder -ne $null }
        $processedFolderNames = [System.Collections.Generic.List[string]]::new()
        $teamHasAnyData = $false

        # Checks each official channel.
        Write-Host "  -> Cross-referencing $($officialChannels.Count) official channel(s)..."
        foreach ($channel in $officialChannels) {
            $hasFiles = "No"; $statusMessage = "Channel folder is empty."; $itemCountDetails = "0 of 0"; $channelName = $channel.DisplayName; $channelSizeFormatted = "0 KB"
            $processedFolderNames.Add($channelName)

            try {
                if ($channel.MembershipType -eq 'private') {
                    # Private channels have their own SharePoint sites; we need to query them separately.
                    try {
                        $folderInfo = Get-MgTeamChannelFileFolder -TeamId $team.Id -ChannelId $channel.Id
                        if ($folderInfo.ParentReference.DriveId) {
                            $privateDriveId = $folderInfo.ParentReference.DriveId
                            $privateRoot = Get-MgDriveRoot -DriveId $privateDriveId
                            
                            # Get size and format it
                            $channelSize = $privateRoot.Size
                            if ($channelSize -ge 1GB) { $channelSizeFormatted = "$([math]::Round($channelSize/1GB, 2)) GB" }
                            elseif ($channelSize -ge 1MB) { $channelSizeFormatted = "$([math]::Round($channelSize/1MB, 2)) MB" }
                            elseif ($channelSize -ge 1KB) { $channelSizeFormatted = "$([math]::Round($channelSize/1KB, 2)) KB" } 
                            else { $channelSizeFormatted = "$channelSize Bytes" }

                            # Check for files
                            $itemCount = $privateRoot.Folder.ChildCount
                            $itemCountDetails = "$itemCount item(s)"
                            
                            if ($itemCount -gt 0) {
                                $hasFiles = "Yes"
                                $statusMessage = "Private Channel contains content."
                            } else {
                                $hasFiles = "No"
                                $statusMessage = "Private Channel folder is empty."
                            }
                        } else {
                            $hasFiles = "Needs Manual Review"; $statusMessage = "Could not determine drive for private channel."
                        }
                    } catch {
                        $hasFiles = "Error"
                        $statusMessage = "Could not access private channel files. API Error: $($_.Exception.Message)"
                        $channelSizeFormatted = "Error"
                    }
                } else {
                    # Logic for Standard and Shared channels
                    $channelFolder = $sharepointFolders | Where-Object { $_.Name -eq $channelName } | Select-Object -First 1
                    if ($channelFolder) {
                        $channelSize = $channelFolder.Size
                        if ($channelSize -ge 1GB) { $channelSizeFormatted = "$([math]::Round($channelSize/1GB, 2)) GB" }
                        elseif ($channelSize -ge 1MB) { $channelSizeFormatted = "$([math]::Round($channelSize/1MB, 2)) MB" }
                        elseif ($channelSize -ge 1KB) { $channelSizeFormatted = "$([math]::Round($channelSize/1KB, 2)) KB" } else { $channelSizeFormatted = "$channelSize Bytes" }
                        $itemsInChannel = @(Get-MgDriveItemChild -DriveId $defaultDrive.Id -DriveItemId $channelFolder.Id -All -ErrorAction Stop)
                        $filesInChannel = $itemsInChannel | Where-Object { $_.Folder -eq $null }
                        $foldersInChannel = $itemsInChannel | Where-Object { $_.Folder -ne $null }
                        $apiItemCount = $channelFolder.Folder.ChildCount; $foundItemCount = $itemsInChannel.Count; $itemCountDetails = "$foundItemCount of $apiItemCount"
                        if ($filesInChannel.Count -gt 0) { $hasFiles = "Yes"; $statusMessage = "Channel has $($filesInChannel.Count) file(s)." } 
                        elseif ($foldersInChannel.Count -gt 0) {
                            if ($foldersInChannel | Where-Object { $_.Folder.ChildCount -gt 0 }) { $hasFiles = "Yes"; $statusMessage = "Contains files in sub-folders." } 
                            else { $hasFiles = "No"; $statusMessage = "Contains only empty sub-folders." }
                        }
                        if ($hasFiles -eq "No" -and $channelFolder.Size > 0) { $hasFiles = "Needs Manual Review"; $statusMessage = "Contains metadata or hidden items." }
                        elseif ($foundItemCount -ne $apiItemCount) { $hasFiles = "Needs Manual Review"; $statusMessage = "Item count mismatch." }
                    } elseif ($channel.DisplayName -eq "General") {
                        $rootFiles = $rootItems | Where-Object { $_.Folder -eq $null }
                        if ($rootFiles.Count -gt 0) {
                            $generalChannelSize = ($rootFiles | Measure-Object -Property Size -Sum).Sum
                            $channelSizeFormatted = if ($generalChannelSize -ge 1KB) { "$([math]::Round($generalChannelSize/1KB, 2)) KB" } else { "$generalChannelSize Bytes" }
                            $hasFiles = "Yes"; $itemCountDetails = "$($rootFiles.Count) of $($rootFiles.Count)"; $statusMessage = "Found files at root."
                        }
                    } else { $statusMessage = "Channel folder not found in SharePoint." }
                }
            } catch { $hasFiles = "Needs Manual Review"; $statusMessage = "Error processing channel."; $channelSizeFormatted = "Error" }

            if ($hasFiles -ne "No") { $teamHasAnyData = $true }
            
            $detailedResults.Add([pscustomobject]@{
                TeamName                     = $team.DisplayName
                ChannelName                  = $channelName
                Members                      = $members
                Owners                       = $owners
                LastActivityDate             = $lastActivityDateStr
                TotalTeamSize                = $totalLibrarySizeFormatted
                ChannelSize                  = $channelSizeFormatted
                SharePointUrl                = $sharePointSite.WebUrl
                'ItemCount (Found of Total)' = $itemCountDetails
                ContainsFiles                = $hasFiles
                Status                       = $statusMessage
                TeamGroupId                  = $team.Id
            })
        }
        
        $orphanedFolders = ($sharepointFolders | Select-Object -ExpandProperty Name) | Where-Object { $processedFolderNames -notcontains $_ }
        if ($orphanedFolders.Count > 0) { $teamHasAnyData = $true } 

        # Deletion Recommendation Algorithm for the Summary Report
        $teamSafeToDelete = "No"; $summaryStatus = "Keep"
        if ($isTeamEmptyOfPeople) { $teamSafeToDelete = "Yes"; $summaryStatus = "Safe to Delete: Team has no members or owners." } 
        elseif ($isTeamInactiveForDeletion -and (-not $teamHasAnyData)) { $teamSafeToDelete = "Yes"; $summaryStatus = "Safe to Delete: Team is inactive and contains no files." }
        else {
            if (-not $isTeamInactiveForDeletion) { $summaryStatus = "Keep: Team has recent activity."}
            elseif ($teamHasAnyData) { $summaryStatus = "Keep: Team contains files or needs review."}
            else { $summaryStatus = "Keep: Team has members but is inactive and empty." }
        }
        
        $summaryResults.Add([pscustomobject]@{
            TeamName         = $team.DisplayName
            Members          = $members
            Owners           = $owners
            SafeToDelete     = $teamSafeToDelete
            OwnerNames       = $ownerNames
            MemberNames      = $memberNames
            Departments      = $departmentList
            LastActivityDate = $lastActivityDateStr
            TotalTeamSize    = $totalLibrarySizeFormatted
            ContainsFiles    = if($teamHasAnyData){"Yes"}else{"No"}
            Status           = $summaryStatus
            TeamGroupId      = $team.Id
            SharePointUrl    = $sharePointSite.WebUrl
        })
        
        foreach ($orphan in $orphanedFolders) {
            $orphanFolderObject = $sharepointFolders | Where-Object { $_.Name -eq $orphan } | Select-Object -First 1
            $orphanSize = $orphanFolderObject.Size
            $orphanSizeFormatted = if ($orphanSize -ge 1KB) { "$([math]::Round($orphanSize/1KB, 2)) KB" } else { "$orphanSize Bytes" }
            $detailedResults.Add([pscustomobject]@{
                TeamName                     = $team.DisplayName
                ChannelName                  = "$($orphan) (Orphaned)"
                Members                      = $members
                Owners                       = $owners
                LastActivityDate             = $lastActivityDateStr
                TotalTeamSize                = $totalLibrarySizeFormatted
                ChannelSize                  = $orphanSizeFormatted
                SharePointUrl                = $sharePointSite.WebUrl
                'ItemCount (Found of Total)' = "N/A"
                ContainsFiles                = "Needs Manual Review"
                Status                       = "Orphaned Folder"
                TeamGroupId                  = $team.Id
            })
        }
    }
    catch {
        $errorMessage = $_.Exception.Message -replace "[\r\n]"," "; Write-Host "  Error processing team: $errorMessage" -ForegroundColor Red
        $detailedResults.Add([pscustomobject]@{
            TeamName                     = $team.DisplayName
            ChannelName                  = "N/A"
            Members                      = "Error"
            Owners                       = "Error"
            LastActivityDate             = "Error"
            TotalTeamSize                = "Error"
            ChannelSize                  = "Error"
            SharePointUrl                = "N/A"
            'ItemCount (Found of Total)' = "N/A"
            ContainsFiles                = "Error"
            Status                       = $errorMessage
            TeamGroupId                  = $team.Id
        })
        $summaryResults.Add([pscustomobject]@{
            TeamName         = $team.DisplayName
            Members          = "Error"
            Owners           = "Error"
            SafeToDelete     = "No"
            LastActivityDate = "Error"
            TotalTeamSize    = "Error"
            ContainsFiles    = "Error"
            Status           = "Error processing team"
            TeamGroupId      = $team.Id
            SharePointUrl    = "N/A"
        })
    }
}

# --- 3. Exporting Reports to a Single Excel File ---
if ($detailedResults.Count -gt 0) {
    Write-Host "`nExporting reports to a single Excel file with two tabs..." -ForegroundColor Cyan
    try {
        # Create the first worksheet with the detailed results
        $detailedResults | Export-Excel -Path $outputFile -WorksheetName "Detailed Per-Channel Report" -AutoSize -FreezeTopRow
        
        # Add the second worksheet to the SAME file with the summary results
        $summaryResults | Export-Excel -Path $outputFile -WorksheetName "Per-Team Summary" -AutoSize -FreezeTopRow

        Write-Host "Export complete. Excel file saved successfully to $outputFile" -ForegroundColor Green
    }
    catch { Write-Host "Error: Failed to save the Excel file." -ForegroundColor Red; Write-Host "   Details: $($_.Exception.Message)" }
} else { Write-Host "`nNo results were generated to export." -ForegroundColor Yellow }

# --- 4. Disconnecting ---
Write-Host "`nDisconnecting from Microsoft Graph." -ForegroundColor Cyan
Disconnect-MgGraph