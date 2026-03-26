# ============================================================
# Run-DailySalesReport.ps1
# Automates: Access macro -> Excel refresh/save -> Email report
# ============================================================

# --- Configuration -----------------------------------------------------------
$AccessDbPath   = "C:\DBs\BenefitsMeDB.accdb"
$AccessMacro    = "_mcrSALESONLY"

$ExcelTemplate  = "C:\Users\JesseSpencer\BenefitsMe, LLC\BenefitsMe - Business Intelligence\KPIs\Sales Reporting\Daily Sales Pivot Reporting\Sales Pivot (Template).xlsx"

$ExcelOutputDir = Split-Path $ExcelTemplate -Parent

$EmailTo        = "jessespencer@benefitsme.com"
$SmtpServer     = "smtp.office365.com"
$SmtpPort       = 587
$SmtpFrom       = "jessespencer@benefitsme.com"

# smtp.cred lives in the same folder as this script
$CredFile       = Join-Path $ExcelOutputDir "smtp.cred"

# --- Derived values ----------------------------------------------------------
$DateStamp      = Get-Date -Format "yyyy-MM-dd"
$OutputFileName = "Sales Pivot ($DateStamp).xlsx"
$OutputFilePath = Join-Path $ExcelOutputDir $OutputFileName

# --- Logging -----------------------------------------------------------------
$LogFile = Join-Path $ExcelOutputDir "DailySalesReport_$DateStamp.log"

function Write-Log {
    param(
        [string]$Message,
        [string]$Level = "INFO"
    )
    $entry = "[$(Get-Date -Format 'HH:mm:ss')] [$Level] $Message"
    Write-Host $entry
    Add-Content -Path $LogFile -Value $entry
}

Write-Log "========== Daily Sales Report started =========="

# =============================================================================
# STEP 1 - Open Access DB and run macro
# =============================================================================
Write-Log "Opening Access database: $AccessDbPath"

$access = $null
try {
    $access = New-Object -ComObject Access.Application
    $access.Visible = $false
    $access.OpenCurrentDatabase($AccessDbPath)
    Write-Log "Database opened. Running macro: $AccessMacro"
    $access.DoCmd.RunMacro($AccessMacro)
    Write-Log "Macro completed successfully."
}
catch {
    Write-Log "ERROR running Access macro: $_" "ERROR"
    throw
}
finally {
    if ($null -ne $access) {
        try { $access.CloseCurrentDatabase() } catch { }
        try { $access.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($access) | Out-Null
        Remove-Variable access
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-Log "Access database closed."
    }
}

# =============================================================================
# STEP 2 - Open Excel template, refresh all data connections, save with date
# =============================================================================
Write-Log "Opening Excel template: $ExcelTemplate"

$excel    = $null
$workbook = $null
try {
    $excel = New-Object -ComObject Excel.Application
    $excel.Visible          = $false
    $excel.DisplayAlerts    = $false
    $excel.AskToUpdateLinks = $false

    $workbook = $excel.Workbooks.Open($ExcelTemplate, $false, $false)

    # Give Excel a moment to fully settle after opening
    Start-Sleep -Seconds 4

    Write-Log "Refreshing all data connections..."

    # Retry loop - Excel can return RPC_E_CALL_REJECTED (0x80010001) if still busy
    $refreshAttempts = 0
    $refreshSuccess  = $false
    while (-not $refreshSuccess -and $refreshAttempts -lt 5) {
        try {
            $workbook.RefreshAll()
            $refreshSuccess = $true
        }
        catch {
            $refreshAttempts++
            Write-Log "RefreshAll attempt $refreshAttempts failed (Excel busy), retrying in 3 seconds..." "WARN"
            Start-Sleep -Seconds 3
            if ($refreshAttempts -ge 5) {
                throw
            }
        }
    }

    # Wait for background queries to finish (max 120 seconds)
    $maxWait  = 120
    $elapsed  = 0
    $interval = 2
    while ($excel.CalculationState -ne 0 -and $elapsed -lt $maxWait) {
        Start-Sleep -Seconds $interval
        $elapsed += $interval
    }

    if ($elapsed -ge $maxWait) {
        Write-Log "WARNING: Data refresh wait timed out after $maxWait seconds." "WARN"
    }
    else {
        Write-Log "Data refresh completed in ~$elapsed seconds."
    }

    # 51 = xlOpenXMLWorkbook (.xlsx)
    $workbook.SaveAs($OutputFilePath, 51)
    Write-Log "Saved as: $OutputFilePath"
}
catch {
    Write-Log "ERROR during Excel processing: $_" "ERROR"
    throw
}
finally {
    if ($null -ne $workbook) {
        try { $workbook.Close($false) } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($workbook) | Out-Null
        Remove-Variable workbook
    }
    if ($null -ne $excel) {
        try { $excel.Quit() } catch { }
        [System.Runtime.InteropServices.Marshal]::ReleaseComObject($excel) | Out-Null
        Remove-Variable excel
        [GC]::Collect()
        [GC]::WaitForPendingFinalizers()
        Write-Log "Excel closed."
    }
}

# =============================================================================
# STEP 3 - Send email with the saved file attached
# =============================================================================
Write-Log "Sending email to $EmailTo with attachment: $OutputFileName"

try {
    if (Test-Path $CredFile) {
        $smtpCred = Import-Clixml -Path $CredFile
        Write-Log "Loaded SMTP credential from: $CredFile"
    }
    else {
        Write-Log "No stored credential found at $CredFile - prompting interactively." "WARN"
        $smtpCred = Get-Credential -Message "Enter SMTP credentials for $SmtpFrom"
    }

    $mailBody = "Please find today's sales pivot report attached.`n`nFile: $OutputFileName`nGenerated: $(Get-Date -Format 'dddd, MMMM d, yyyy')"

    Send-MailMessage `
        -From        $SmtpFrom `
        -To          $EmailTo `
        -Subject     $OutputFileName `
        -Body        $mailBody `
        -Attachments $OutputFilePath `
        -SmtpServer  $SmtpServer `
        -Port        $SmtpPort `
        -UseSsl `
        -Credential  $smtpCred

    Write-Log "Email sent successfully."
}
catch {
    Write-Log "ERROR sending email: $_" "ERROR"
    throw
}

Write-Log "========== Daily Sales Report completed successfully =========="
