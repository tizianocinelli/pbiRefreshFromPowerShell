# ============================== #
# 📌 COME CREARE I FILE DI CREDENZIALI
# ============================== #
# Per Power BI:
# "powerbi_cred_2.txt" deve contenere **solo** la password in chiaro o criptata con SecureString.
# Esempio per crearlo:
# "mypassword" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\Scripts\powerbi_cred_2.txt"

# Per SQL Server (se usato):
# "powerbi_cred_3.txt" deve contenere **solo** la password SQL.
# Esempio:
# "sqlpassword" | ConvertTo-SecureString -AsPlainText -Force | ConvertFrom-SecureString | Out-File "C:\Scripts\powerbi_cred_3.txt"


# ============================== #
# 📌 CONFIGURAZIONE INIZIALE
# ============================== #

# Percorso modulo Power BI
$PowerBIModulePath = "C:\Program Files\WindowsPowerShell\Modules\MicrosoftPowerBIMgmt"

# Credenziali Power BI
$PowerBICredentialFile = "C:\Scripts\powerbi_cred_2.txt"
$PowerBIUsername = "powerbiuser@email.email"

# Credenziali SQL Server (opzionale)
$SqlCredentialFile = "C:\Scripts\powerbi_cred_3.txt"
$SqlServer = "localhost"
$SqlDatabase = "databasename"
$SqlUser = "databaseuser"  # es: "sa"

# URL Healthchecks.io
$HealthCheckBaseURL = "ULR PING HEATHCHECKS.io"

# Dataset da aggiornare
$DatasetsToRefresh = @(
    @{ WorkspaceID = "workspaceid1"; DatasetID = "datasetid1"; Name = "Dataset 1" },
    @{ WorkspaceID = "workspaceid2"; DatasetID = "datasetid2"; Name = "Dataset 2" }
)

# ============================== #
# 🚀 AVVIO SCRIPT
# ============================== #

# Importa modulo Power BI
if (Test-Path $PowerBIModulePath) {
    Import-Module -Name MicrosoftPowerBIMgmt -Force -ErrorAction Stop
    Write-Host "✅ Modulo Power BI caricato correttamente."
} else {
    Write-Host "❌ Errore: Modulo MicrosoftPowerBIMgmt non trovato in $PowerBIModulePath"
    exit 1
}

# Healthcheck START
try {
    Invoke-RestMethod -Uri "$HealthCheckBaseURL/start" -Method Get
    Write-Host "📡 Healthcheck: Avvio segnalato con successo"
} catch {
    Write-Host "⚠️ Errore nel segnalare l'inizio a Healthchecks.io"
}

# Autenticazione Power BI
$securePassword = Get-Content $PowerBICredentialFile | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential ($PowerBIUsername, $securePassword)

try {
    Login-PowerBIServiceAccount -Credential $credential
    Write-Host "✅ Accesso a Power BI riuscito"
} catch {
    Write-Host "❌ Errore durante l'autenticazione a Power BI"
    Invoke-RestMethod -Uri "$HealthCheckBaseURL/1" -Method Get
    exit 1
}

# ============================== #
# 🔄 AGGIORNAMENTO DEI DATASET
# ============================== #

$scriptError = $false

foreach ($item in $DatasetsToRefresh) {
    $workspaceId = $item.WorkspaceID
    $datasetId = $item.DatasetID
    $datasetName = $item.Name

    Write-Host "🔄 Avvio aggiornamento per Dataset: $datasetName (ID: $datasetId) nel Workspace ID: $workspaceId"

    $URI = "groups/$workspaceId/datasets/$datasetId/refreshes"

    try {
        Invoke-PowerBIRestMethod -Url $URI -Method Post
        Write-Host "✅ Aggiornamento avviato con successo per Dataset: $datasetName"
    } catch {
        Write-Host "❌ Errore nell'avvio del refresh per Dataset: $datasetName"
        $scriptError = $true
        continue
    }

    $refreshComplete = $false
    $maxWaitTime = 900
    $elapsedTime = 0
    $checkInterval = 30

    Write-Host "⏳ Attesa del completamento del refresh per Dataset: $datasetName..."

    while (-not $refreshComplete -and $elapsedTime -lt $maxWaitTime) {
        Start-Sleep -Seconds $checkInterval
        $elapsedTime += $checkInterval

        $refreshStatusUri = "groups/$workspaceId/datasets/$datasetId/refreshes"
        $refreshHistory = Invoke-PowerBIRestMethod -Url $refreshStatusUri -Method Get | ConvertFrom-Json

        if ($refreshHistory.value.Count -gt 0) {
            $latestRefresh = $refreshHistory.value[0]

            Write-Host "🔎 Stato attuale: $($latestRefresh.status) | Avviato: $($latestRefresh.startTime)"

            if ($latestRefresh.status -eq "Completed") {
                Write-Host "✅ Refresh completato con successo per Dataset: $datasetName"
                $refreshComplete = $true
            } elseif ($latestRefresh.status -eq "Failed") {
                Write-Host "❌ Il refresh è FALLITO per Dataset: $datasetName!"
                $scriptError = $true
                break
            }
        }
    }

    if (-not $refreshComplete) {
        Write-Host "⚠️ Timeout: Il refresh per Dataset: $datasetName non si è completato entro il tempo massimo."
        $scriptError = $true
    }

    try {
        $logData = @{ message = "Aggiornamento completato per Dataset: $datasetName" } | ConvertTo-Json
        Invoke-RestMethod -Uri "$HealthCheckBaseURL/log" -Method Post -Body $logData -ContentType "application/json"
        Write-Host "📜 Healthcheck Log: Aggiornamento registrato per Dataset: $datasetName"
    } catch {
        Write-Host "⚠️ Errore nel log dell'aggiornamento su Healthchecks.io per Dataset: $datasetName"
    }
}

# ============================== #
# 🧾 REGISTRAZIONE STORICO SU SQL
# ============================== #

$SqlPassword = Get-Content $SqlCredentialFile | ConvertTo-SecureString
$ConnectionString = "Server=$SqlServer;Database=$SqlDatabase;Integrated Security = True;"

$Workspaces = Get-PowerBIWorkspace -Scope Organization

foreach ($workspace in $Workspaces) {
    $DataSets = Get-PowerBIDataset -WorkspaceId $workspace.Id | Where-Object { $_.isRefreshable -eq $true }

    foreach ($dataset in $DataSets) {
        $URI = "groups/$($workspace.Id)/datasets/$($dataset.Id)/refreshes"
        $Results = Invoke-PowerBIRestMethod -Url $URI -Method Get | ConvertFrom-Json

        if ($Results.value.Count -gt 0) {
            foreach ($refresh in $Results.value) {
                $StartTimeFormatted = [datetime]::Parse($refresh.startTime).ToString("yyyy-MM-dd HH:mm:ss")
                $EndTimeFormatted = [datetime]::Parse($refresh.endTime).ToString("yyyy-MM-dd HH:mm:ss")

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
            Write-Host "⚠️ No refresh history found for dataset: $($dataset.Name)"
        }
    }
}

# ============================== #
# 🔐 LOGOUT E CONCLUSIONE
# ============================== #

Disconnect-PowerBIServiceAccount
Write-Host "🔒 Disconnessione da Power BI completata."

try {
    if ($scriptError) {
        Invoke-RestMethod -Uri "$HealthCheckBaseURL/1" -Method Get
        Write-Host "❌ Healthcheck: Errore segnalato"
    } else {
        Invoke-RestMethod -Uri "$HealthCheckBaseURL" -Method Get
        Write-Host "✅ Healthcheck: Conclusione segnalata con successo"
    }
} catch {
    Write-Host "⚠️ Errore nel segnalare lo stato finale a Healthchecks.io"
}
