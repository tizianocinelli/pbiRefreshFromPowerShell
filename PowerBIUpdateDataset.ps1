# Importa il modulo Power BI
$modulePath = "C:\Program Files\WindowsPowerShell\Modules\MicrosoftPowerBIMgmt"
if (Test-Path $modulePath) {
    Import-Module -Name MicrosoftPowerBIMgmt -Force -ErrorAction Stop
    Write-Host "✅ Modulo Power BI caricato correttamente."
} else {
    Write-Host "❌ Errore: Modulo MicrosoftPowerBIMgmt non trovato in $modulePath"
    exit 1
}

# URL di monitoraggio Healthchecks.io
$HealthCheckBaseURL = "ULR PING HEATHCHECKS"

# Segnala l'inizio del processo a Healthchecks.io
try {
    Invoke-RestMethod -Uri "$HealthCheckBaseURL/start" -Method Get
    Write-Host "📡 Healthcheck: Avvio segnalato con successo"
} catch {
    Write-Host "⚠️ Errore nel segnalare l'inizio a Healthchecks.io"
}

# Recupera le credenziali da file e autentica su Power BI
$securePassword = Get-Content "C:\Scripts\powerbi_cred_2.txt" | ConvertTo-SecureString
$credential = New-Object System.Management.Automation.PSCredential ("powerbiuser@email.email", $securePassword)

try {
    Login-PowerBIServiceAccount -Credential $credential
    Write-Host "✅ Accesso a Power BI riuscito"
} catch {
    Write-Host "❌ Errore durante l'autenticazione a Power BI"
    Invoke-RestMethod -Uri "$HealthCheckBaseURL/1" -Method Get
    exit 1
}

# Definizione dei workspace e dataset da aggiornare
$DatasetsToRefresh = @(
    @{ WorkspaceID = "workspaceid1"; DatasetID = "datasetid1" },
    @{ WorkspaceID = "workspaceid2"; DatasetID = "datasetid2" }
)

$scriptError = $false  # Flag per segnalare errori nel processo

# Avvia il refresh per ciascun dataset specificato
foreach ($item in $DatasetsToRefresh) {
    $workspaceId = $item.WorkspaceID
    $datasetId = $item.DatasetID
    $datasetName = $item.Name

    Write-Host "🔄 Avvio aggiornamento per Dataset: $datasetName (ID: $datasetId) nel Workspace ID: $workspaceId"

    # Costruisce l'URL per la chiamata API
    $URI = "groups/$workspaceId/datasets/$datasetId/refreshes"

    # Esegue la richiesta API per avviare il refresh
    try {
        Invoke-PowerBIRestMethod -Url $URI -Method Post
        Write-Host "✅ Aggiornamento avviato con successo per Dataset: $datasetName"
    } catch {
        Write-Host "❌ Errore nell'avvio del refresh per Dataset: $datasetName"
        $scriptError = $true
        continue
    }

    # 🔍 Controlla lo stato del refresh
    $refreshComplete = $false
    $maxWaitTime = 900  # 15 minuti (900 secondi)
    $elapsedTime = 0
    $checkInterval = 30  # Controlla ogni 30 secondi

    Write-Host "⏳ Attesa del completamento del refresh per Dataset: $datasetName..."

    while (-not $refreshComplete -and $elapsedTime -lt $maxWaitTime) {
        Start-Sleep -Seconds $checkInterval
        $elapsedTime += $checkInterval

        # Recupera lo stato dell'ultimo refresh
        $refreshStatusUri = "groups/$workspaceId/datasets/$datasetId/refreshes"
        $refreshHistory = Invoke-PowerBIRestMethod -Url $refreshStatusUri -Method Get | ConvertFrom-Json

        if ($refreshHistory.value.Count -gt 0) {
            $latestRefresh = $refreshHistory.value[0]  # Ultimo tentativo di refresh

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

    # 📡 Log su Healthchecks.io con il nome del dataset
    try {
        $logData = @{ message = "Aggiornamento completato per Dataset: $datasetName" } | ConvertTo-Json
        Invoke-RestMethod -Uri "$HealthCheckBaseURL/log" -Method Post -Body $logData -ContentType "application/json"
        Write-Host "📜 Healthcheck Log: Aggiornamento registrato per Dataset: $datasetName"
    } catch {
        Write-Host "⚠️ Errore nel log dell'aggiornamento su Healthchecks.io per Dataset: $datasetName"
    }
}

# Logout da Power BI
Disconnect-PowerBIServiceAccount
Write-Host "🔒 Disconnessione da Power BI completata."

# Segnala lo stato finale a Healthchecks.io
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
