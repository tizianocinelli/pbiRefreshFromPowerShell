## 📄 UpdatePowerBIDataset

Questo script è pensato per essere eseguito tramite un **task schedulato di Windows** oppure tramite un **SQL Server Agent Job**.  
Il suo scopo principale è:

- Aggiornare uno o più **dataset** pubblicati nel **portale Power BI**.
- Registrare lo **storico delle esecuzioni** in una **tabella SQL Server locale**.
- Monitorare tutto il processo tramite **[Healthchecks.io](https://healthchecks.io/)**.

---

### 🔐 Gestione delle credenziali

Le credenziali vengono archiviate **in locale** sotto forma di **SecureString**, tramite file `.txt`.

- **Power BI:**  
  È necessario creare un file contenente la password dell’utente di servizio per accedere a Power BI.  
  ⚠️ **Attenzione alla MFA**: se l’utente dispone di autenticazione a più fattori, è necessario disabilitarla oppure usare un **account di servizio configurato in Entra ID** con accesso alle risorse di Power BI.

- **SQL Server:**  
  Se il logging su SQL è attivo, serve un secondo file con la password del database.

---

### ✅ Funzionalità principali

- Caricamento del modulo Power BI.
- Login al servizio Power BI.
- Refresh automatizzato di dataset specificati.
- Verifica del completamento o errore del refresh.
- Log del risultato via Healthchecks.io.
- Salvataggio delle informazioni di refresh in una tabella SQL Server.
- Logout sicuro a fine processo.

---

