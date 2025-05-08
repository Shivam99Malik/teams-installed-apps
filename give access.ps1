# Import the CSV file
$data = Import-CSV -Path "C:\Users\RWA\OneDrive - Contoso\Documents\TeamsAPP\TEAMWHICHLEFT.csv"
 
# Loop through each row and update app permissions
foreach ($row in $data) {
    $appId = $row.AppId
    $TeamsIds = $row.TeamsGroupIDs -split ", "
 
    if ([string]::IsNullOrEmpty($appId)) {
        Write-Host "Error: AppId is empty for TeamsGroupIDs: $TeamsIds"
        continue
    }
 
    try {
        # Update the app permissions for the groups
        Update-M365TeamsApp -Id $appId -IsBlocked $false -AppAssignmentType UsersAndGroups -OperationType Add -Groups $TeamsIds
        Write-Host "Successfully updated app permissions for App ID: $appId"
    } catch {
        Write-Host "Error updating app permissions for App ID: $appId"
        Write-Host $_.Exception.Message
    }
}