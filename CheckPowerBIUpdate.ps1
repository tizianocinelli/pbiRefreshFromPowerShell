# Install Power BI module if not already installed
# Install-Module -Name MicrosoftPowerBIMgmt -Scope CurrentUser -Force -AllowClobber
Import-Module -Name MicrosoftPowerBIMgmt

# Authenticate with Power BI
# Login-PowerBIServiceAccount

# SQL Server Connection Settings
$SqlServer = "localhost"  # e.g., "localhost" or "your.database.windows.net"
$SqlDatabase = "databasename"
$SqlUser = "databaseuser"  # Use "sa" for local SQL Server
$SqlPassword =  Get-Content "C:\Scripts\powerbi_cred_3.txt" | ConvertTo-SecureString


# Recupera le credenziali da file
$securePassword = Get-Content "C:\Scripts\powerbi_cred_2.txt" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential ("powerbiuser@email.email", $securePassword)

# Login a Power BI
try {
    Login-PowerBIServiceAccount -Credential $credential
    Write-Host "✅ Accesso a Power BI riuscito"
} catch {
    Write-Host "❌ Errore durante l'autenticazione a Power BI"
    exit 1
}



# Define SQL connection string
$ConnectionString = "Server=$SqlServer;Database=$SqlDatabase;Integrated Security = True;"

# Get all workspaces the user has access to
$Workspaces = Get-PowerBIWorkspace -Scope Organization

foreach ($workspace in $Workspaces) {

    # Get datasets that are refreshable
    $DataSets = Get-PowerBIDataset -WorkspaceId $workspace.Id | Where-Object { $_.isRefreshable -eq $true }

    foreach ($dataset in $DataSets) {

        # Construct API request for dataset refresh history
        $URI = "groups/$($workspace.Id)/datasets/$($dataset.Id)/refreshes"
        
        # Call the Power BI REST API
        $Results = Invoke-PowerBIRestMethod -Url $URI -Method Get | ConvertFrom-Json

        # Check if there are refresh records
        if ($Results.value.Count -gt 0) {
            foreach ($refresh in $Results.value) {
                
                # Format SQL-compatible timestamps
                $StartTimeFormatted = [datetime]::Parse($refresh.startTime).ToString("yyyy-MM-dd HH:mm:ss")
                $EndTimeFormatted = [datetime]::Parse($refresh.endTime).ToString("yyyy-MM-dd HH:mm:ss")

                # SQL Query to insert/update the refresh history
                $SqlQuery = @"
                IF NOT EXISTS (SELECT 1 FROM PowerBI_Refresh_Status WHERE refresh_id = '$($refresh.id)')
                BEGIN
                    INSERT INTO PowerBI_Refresh_Status 
                    (refresh_id, workspace_name, workspace_id, dataset_name, dataset_id, status, start_time, end_time, type)
                    VALUES 
                    ('$($refresh.id)', '$($workspace.Name)', '$($workspace.Id)', '$($dataset.Name)', '$($dataset.Id)', 
                     '$($refresh.status)', '$StartTimeFormatted', '$EndTimeFormatted', '$($refresh.type)')
                END
                ELSE
                BEGIN
                    UPDATE PowerBI_Refresh_Status 
                    SET status = '$($refresh.status)', start_time = '$StartTimeFormatted', end_time = '$EndTimeFormatted'
                    WHERE refresh_id = '$($refresh.id)'
                END
"@

                # Execute SQL query
                $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
                $SqlConnection.ConnectionString = $ConnectionString
                $SqlConnection.Open()

                $SqlCommand = $SqlConnection.CreateCommand()
                $SqlCommand.CommandText = $SqlQuery
                $SqlCommand.ExecuteNonQuery()

                $SqlConnection.Close()

                Write-Host "✅ Inserted/Updated Refresh: $($refresh.id) for Dataset: $($dataset.Name)"
            }
        } else {
            Write-Host "⚠️ No refresh history found for dataset: $($dataset.Name) | Dataset ID: $($dataset.Id)"
        }
    }
}
