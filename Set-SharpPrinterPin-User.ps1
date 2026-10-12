param(
    [ValidateSet('Logon','User','Check')][string]$Mode = 'Logon',
    [string]$ConfigPath
)

if ($Mode -ne 'User') {
    # ===================================================================
    # SHARED (Logon + Check), runs as SYSTEM.
    # Check  = one-shot preflight; prints CLEAN / NEEDED / SKIP / UNKNOWN.
    # Logon  = waits for an unlocked desktop, runs the preflight, then
    #          launches the user-mode part in the user's session if needed.
    # ===================================================================
    $ErrorActionPreference = 'Stop'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

    $Root    = Join-Path $env:ProgramData 'PrinterPinAuto'
    $LogPath = Join-Path $Root 'log.txt'

    $PrinterNames = @(
        'Printer | ES Office 2025 - Sharp BP-70M31'
        'Printer | HS Office 2025- Sharp BP-70C31'
        'Printer | HS Staff - Sharp 2025 BP-70M31'
        'Printer | MS Office 2025 - Sharp BP-70M31'
        'Printer | MS Staff 2025 - Sharp BP-70C31'
        'Printer | ES Staff 2025 - Sharp BP-70C31'
    )

    if (-not ('ConsoleSession' -as [type])) {
        Add-Type -TypeDefinition 'using System.Runtime.InteropServices; public class ConsoleSession { [DllImport("kernel32.dll")] public static extern uint WTSGetActiveConsoleSessionId(); }'
    }

    function Write-Log([string]$Text) {
        try {
            Add-Content -LiteralPath $LogPath -Value ('{0:u} {1}' -f (Get-Date), $Text)
            $lines = @(Get-Content -LiteralPath $LogPath -ErrorAction SilentlyContinue)
            if ($lines.Count -gt 400) { $lines[-200..-1] | Set-Content -LiteralPath $LogPath }
        } catch {}
    }

    function Open-Hive64($Hive) {
        [Microsoft.Win32.RegistryKey]::OpenBaseKey($Hive, [Microsoft.Win32.RegistryView]::Registry64)
    }

    function Get-ActiveUserSid {
        $console = [ConsoleSession]::WTSGetActiveConsoleSessionId()
        if ($console -eq [uint32]::MaxValue) { return $null }
        foreach ($p in @(Get-CimInstance Win32_Process -Filter "Name='explorer.exe' AND SessionId=$console")) {
            $r = Invoke-CimMethod -InputObject $p -MethodName GetOwnerSid
            if ($r.ReturnValue -eq 0 -and $r.Sid) { return [string]$r.Sid }
        }
        return $null
    }

    # Identifies one sign-in session (the console session's earliest explorer start time)
    function Get-SessionKey([string]$Sid) {
        $console = [ConsoleSession]::WTSGetActiveConsoleSessionId()
        $p = @(Get-Process -Name explorer -ErrorAction SilentlyContinue | Where-Object { $_.SessionId -eq $console } |
               Sort-Object StartTime) | Select-Object -First 1
        if ($p) { return ('{0}|{1}' -f $Sid, $p.StartTime.ToUniversalTime().Ticks) }
        return $Sid
    }

    function Test-ScreenLocked {
        $console = [ConsoleSession]::WTSGetActiveConsoleSessionId()
        return [bool](Get-Process -Name LogonUI -ErrorAction SilentlyContinue |
            Where-Object { $_.SessionId -eq $console })
    }

    function Get-GcpwEmail([string]$Sid) {
        $k = (Open-Hive64 'LocalMachine').OpenSubKey("SOFTWARE\Google\GCPW\Users\$Sid")
        if (-not $k) { return $null }
        $email = $k.GetValue('email'); $k.Close(); return $email
    }

    # Throws on network errors; returns $null if the user is not in the roster.
    function Get-UserPin([string]$Email) {
        $body = @{ key = $LookupKey; email = $Email } | ConvertTo-Json -Compress
        $r = Invoke-RestMethod -Uri $LookupUrl -Method Post -ContentType 'text/plain' -Body $body -UseBasicParsing -TimeoutSec 30
        if ($r -is [string]) { $r = $r | ConvertFrom-Json }
        if (-not $r.ok -or -not $r.pin -or -not $r.hex) { return $null }
        return $r
    }

    function Test-PrinterPin([string]$Sid, [string]$Printer, [string]$ExpectedHex) {
        try {
            $k = (Open-Hive64 'Users').OpenSubKey("$Sid\Software\SHARP\$Printer\printer_ui\job_control")
            if (-not $k) { return $false }
            $use = $k.GetValue('use_account_number'); $acct = $k.GetValue('account_number'); $k.Close()
            if ($null -eq $use -or [int]$use -ne 1) { return $false }
            if ($acct -isnot [byte[]] -or $acct.Length -lt 16) { return $false }
            return ((($acct[0..15]) | ForEach-Object { $_.ToString('X2') }) -join '') -eq $ExpectedHex
        }
        catch { return $false }
    }

    function New-State($State, $Reason, $Sid = $null, $User = $null, $ToDo = @()) {
        [pscustomobject]@{ State = $State; Reason = $Reason; Sid = $Sid; User = $User; ToDo = $ToDo }
    }

    # CLEAN = verified correct or nothing applicable | NEEDED | SKIP = can't act for this user | UNKNOWN = can't verify
    function Get-DeploymentState([string]$Sid, [int]$Attempts = 3, [int]$RetrySeconds = 30) {
        $installed = @(Get-CimInstance Win32_Printer | Select-Object -ExpandProperty Name)
        $targets   = @($PrinterNames | Where-Object { $installed -contains $_ })
        if ($targets.Count -eq 0) { return New-State 'CLEAN' 'No target printers installed.' }

        if (-not $Sid) { $Sid = Get-ActiveUserSid }
        if (-not $Sid) { return New-State 'UNKNOWN' 'No signed-in user.' }

        $email = Get-GcpwEmail $Sid
        if (-not $email) { return New-State 'SKIP' 'No GCPW email for the active user.' }

        $user = $null
        $lookupOk = $false
        for ($i = 0; $i -lt $Attempts -and -not $lookupOk; $i++) {
            try { $user = Get-UserPin $email; $lookupOk = $true }
            catch { if ($i -lt $Attempts - 1) { Start-Sleep -Seconds $RetrySeconds } }
        }
        if (-not $lookupOk) { return New-State 'UNKNOWN' 'PIN lookup failed (network?).' }
        if (-not $user)     { return New-State 'SKIP' 'User is not in the roster.' }

        $toDo = @($targets | Where-Object { -not (Test-PrinterPin $Sid $_ $user.hex) })
        if ($toDo.Count -eq 0) { return New-State 'CLEAN' 'All installed printers already correct.' $Sid $user }
        return New-State 'NEEDED' ('{0} printer(s) need configuration.' -f $toDo.Count) $Sid $user $toDo
    }

    # ---------------- CHECK MODE (called by the Worker) ----------------
    if ($Mode -eq 'Check') {
        try {
            $cfg       = Get-Content -LiteralPath (Join-Path $Root 'config.json') -Raw | ConvertFrom-Json
            $LookupUrl = [string]$cfg.LookupUrl
            $LookupKey = [string]$cfg.LookupKey
            Write-Output (Get-DeploymentState -Attempts 2 -RetrySeconds 5).State
        }
        catch { Write-Output 'UNKNOWN' }
        exit 0
    }

    # ---------------- LOGON MODE ----------------
    $JobDir  = $null
    $JobTask = $null

    try {
        $cfg       = Get-Content -LiteralPath (Join-Path $Root 'config.json') -Raw | ConvertFrom-Json
        $LookupUrl = [string]$cfg.LookupUrl
        $LookupKey = [string]$cfg.LookupKey

        # Wait for an unlocked desktop (up to 10 minutes)
        $sid = $null
        $deadline = (Get-Date).AddMinutes(10)
        while ((Get-Date) -lt $deadline) {
            $candidate = Get-ActiveUserSid
            if ($candidate -and -not (Test-ScreenLocked)) { $sid = $candidate; break }
            Start-Sleep -Seconds 10
        }
        if (-not $sid) { Write-Log 'No unlocked desktop within 10 minutes; exiting.'; exit 0 }

        $attemptFile = Join-Path $Root "attempt_$sid.txt"

        $state = Get-DeploymentState -Sid $sid
        if ($state.State -ne 'NEEDED') {
            if ($state.State -eq 'CLEAN') { Remove-Item -LiteralPath $attemptFile -Force -ErrorAction SilentlyContinue }
            Write-Log ('{0} Exiting.' -f $state.Reason)
            exit 0
        }
        $user  = $state.User
        $toDo  = @($state.ToDo)

        # Once per sign-in session until it succeeds. The marker is written BEFORE the
        # attempt, so a crash or a "Run Later" also counts; it is cleared only on success.
        $sessionKey = Get-SessionKey $sid
        $previous   = ([string](Get-Content -LiteralPath $attemptFile -Raw -ErrorAction SilentlyContinue)).Trim()
        if ($previous -eq $sessionKey) { Write-Log 'Already attempted in this sign-in session; will retry at next logon.'; exit 0 }
        Set-Content -LiteralPath $attemptFile -Value $sessionKey

        # Locked-down job folder: SYSTEM + this user only
        $stamp   = [guid]::NewGuid().ToString('N')
        $JobDir  = Join-Path $env:ProgramData ('PrinterPinJob_' + $stamp)
        $JobTask = 'PrinterPinJob_' + $stamp
        New-Item -ItemType Directory -Path $JobDir | Out-Null

        $userSid = New-Object Security.Principal.SecurityIdentifier($sid)
        $sysSid  = New-Object Security.Principal.SecurityIdentifier('S-1-5-18')
        $inherit = 'ContainerInherit,ObjectInherit'
        $acl = New-Object Security.AccessControl.DirectorySecurity
        $acl.SetAccessRuleProtection($true, $false)
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($sysSid, 'FullControl', $inherit, 'None', 'Allow')))
        $acl.AddAccessRule((New-Object Security.AccessControl.FileSystemAccessRule($userSid, 'Modify', $inherit, 'None', 'Allow')))
        Set-Acl -Path $JobDir -AclObject $acl

        $UserScript  = Join-Path $JobDir 'PrinterPin-User.ps1'
        $JobConfig   = Join-Path $JobDir 'config.json'
        $ResultsPath = Join-Path $JobDir 'results.json'

        Copy-Item -LiteralPath $PSCommandPath -Destination $UserScript
        @{ Pin = $user.pin; Hex = $user.hex; Printers = $toDo; ResultsPath = $ResultsPath } |
            ConvertTo-Json -Compress | Set-Content -LiteralPath $JobConfig -Encoding UTF8

        # Run the user-mode part as the logged-in user, in their session
        $userName  = $userSid.Translate([Security.Principal.NTAccount]).Value
        $psExe     = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
        $action    = New-ScheduledTaskAction -Execute $psExe -Argument ("-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File `"$UserScript`" -Mode User -ConfigPath `"$JobConfig`"")
        $principal = New-ScheduledTaskPrincipal -UserId $userName -LogonType Interactive -RunLevel Limited
        $settings  = New-ScheduledTaskSettingsSet -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)

        Register-ScheduledTask -TaskName $JobTask -Action $action -Principal $principal -Settings $settings -Force | Out-Null
        Write-Log ('Starting user session job for {0} printer(s).' -f $toDo.Count)
        Start-ScheduledTask -TaskName $JobTask

        Start-Sleep -Seconds 5
        $deadline = (Get-Date).AddMinutes(12)
        while ((Get-Date) -lt $deadline) {
            if (Test-Path -LiteralPath $ResultsPath) { break }
            if ((Get-ScheduledTask -TaskName $JobTask).State -ne 'Running') { break }
            Start-Sleep -Seconds 3
        }
        Start-Sleep -Seconds 2

        $failed = 0; $deferred = 0
        if (Test-Path -LiteralPath $ResultsPath) {
            foreach ($r in @(Get-Content -LiteralPath $ResultsPath -Raw | ConvertFrom-Json)) {
                Write-Log ('{0}: {1} - {2}' -f $r.Printer, $r.Status, $r.Details)
                if     ($r.Status -eq 'DEFERRED') { $deferred++ }
                elseif ($r.Status -ne 'SUCCESS')  { $failed++ }
            }
        }
        else {
            Write-Log ('No results file produced. Task last result: {0}' -f (Get-ScheduledTaskInfo -TaskName $JobTask).LastTaskResult)
            $failed = 1
        }

        # Clear the session marker only on full success; Panel does the final verification and cleanup
        if ($failed -eq 0 -and $deferred -eq 0) { Remove-Item -LiteralPath $attemptFile -Force -ErrorAction SilentlyContinue }
    }
    catch { Write-Log ('ERROR: ' + $_.Exception.Message) }
    finally {
        if ($JobTask) { Unregister-ScheduledTask -TaskName $JobTask -Confirm:$false -ErrorAction SilentlyContinue }
        if ($JobDir)  { Remove-Item -LiteralPath $JobDir -Recurse -Force -ErrorAction SilentlyContinue }
    }
    exit 0
}

# ===================================================================
# USER MODE: runs as the logged-in user
# ===================================================================

$ErrorActionPreference = 'Stop'

$Config       = Get-Content -LiteralPath $ConfigPath -Raw | ConvertFrom-Json
$TestPin      = [string]$Config.Pin
$ExpectedHex  = [string]$Config.Hex
$ResultsPath  = [string]$Config.ResultsPath
$PrinterNames = @($Config.Printers)
$Results      = @()

function Get-PrinterPinStatus {
    param([Parameter(Mandatory)][string]$PrinterName)

    $RegistryKey = "HKCU:\Software\SHARP\$PrinterName\printer_ui\job_control"

    try {
        if (-not (Test-Path -LiteralPath $RegistryKey)) {
            return [PSCustomObject]@{ IsConfigured = $false; Details = 'Printer settings registry key was not found.' }
        }

        $Settings = Get-ItemProperty -LiteralPath $RegistryKey

        if ($null -eq $Settings.use_account_number -or [int]$Settings.use_account_number -ne 1) {
            return [PSCustomObject]@{ IsConfigured = $false; Details = 'User Number authentication is not enabled.' }
        }

        $AccountNumber = $Settings.account_number

        if ($AccountNumber -isnot [byte[]] -or $AccountNumber.Length -lt 16) {
            return [PSCustomObject]@{ IsConfigured = $false; Details = 'Stored User Number value is missing or too short.' }
        }

        $ActualHex = ($AccountNumber[0..15] | ForEach-Object { $_.ToString('X2') }) -join ''

        if ($ActualHex -ne $ExpectedHex) {
            return [PSCustomObject]@{ IsConfigured = $false; Details = 'Stored User Number does not match the required PIN.' }
        }

        return [PSCustomObject]@{ IsConfigured = $true; Details = 'User Number authentication and PIN are correct.' }
    }
    catch {
        return [PSCustomObject]@{ IsConfigured = $false; Details = "Could not validate registry settings: $($_.Exception.Message)" }
    }
}

Add-Type -AssemblyName UIAutomationClient
Add-Type -AssemblyName UIAutomationTypes
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

if (-not ('SharpPrinterPinMouse' -as [type])) {
    Add-Type -TypeDefinition @'
using System;
using System.Runtime.InteropServices;

public class SharpPrinterPinMouse {
    [DllImport("user32.dll")]
    public static extern bool SetCursorPos(int X, int Y);

    [DllImport("user32.dll")]
    public static extern void mouse_event(uint flags, uint dx, uint dy, uint data, UIntPtr extraInfo);

    public static void ClickAt(int x, int y) {
        SetCursorPos(x, y);
        mouse_event(0x0002, 0, 0, 0, UIntPtr.Zero);
        mouse_event(0x0004, 0, 0, 0, UIntPtr.Zero);
    }
}
'@
}

# -------------------------
# Prompt / progress window
# -------------------------

$Form = New-Object System.Windows.Forms.Form
$Form.Text = 'Printer Configuration'
$Form.Size = New-Object System.Drawing.Size(620, 285)
$Form.MinimumSize = $Form.Size
$Form.MaximumSize = $Form.Size
$Form.StartPosition = 'CenterScreen'
$Form.FormBorderStyle = 'FixedDialog'
$Form.MaximizeBox = $false
$Form.MinimizeBox = $false
$Form.ControlBox = $false
$Form.TopMost = $true
$Form.BackColor = [System.Drawing.Color]::White
$Form.Font = New-Object System.Drawing.Font('Segoe UI', 10)

$WarningPanel = New-Object System.Windows.Forms.Panel
$WarningPanel.Location = New-Object System.Drawing.Point(20, 15)
$WarningPanel.Size = New-Object System.Drawing.Size(565, 48)
$WarningPanel.BackColor = [System.Drawing.Color]::FromArgb(255, 240, 240)
$Form.Controls.Add($WarningPanel)

$WarningLabel = New-Object System.Windows.Forms.Label
$WarningLabel.Location = New-Object System.Drawing.Point(10, 5)
$WarningLabel.Size = New-Object System.Drawing.Size(545, 38)
$WarningLabel.Text = 'ONCE THE PROCESS STARTS, PLEASE DO NOT USE YOUR KEYBOARD OR MOUSE UNTIL IT FINISHES.'
$WarningLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$WarningLabel.ForeColor = [System.Drawing.Color]::Red
$WarningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$WarningPanel.Controls.Add($WarningLabel)

$InfoLabel = New-Object System.Windows.Forms.Label
$InfoLabel.Location = New-Object System.Drawing.Point(22, 72)
$InfoLabel.Size = New-Object System.Drawing.Size(560, 52)
$InfoLabel.Text = @(
    'This process takes about one minute.',
    'Choose Run Later if you are busy; you will be asked again at your next sign-in.',
    'If it does not finish promptly, please contact your IT department.'
) -join "`r`n"
$InfoLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$InfoLabel.ForeColor = [System.Drawing.Color]::FromArgb(65, 65, 65)
$Form.Controls.Add($InfoLabel)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(22, 130)
$StatusLabel.Size = New-Object System.Drawing.Size(560, 35)
$StatusLabel.Text = 'Your printer settings need to be updated.'
$StatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 45, 65)
$Form.Controls.Add($StatusLabel)

# Run Now / Run Later buttons (shown first, in the progress bar's spot)
$script:Choice = $null

$RunNowButton = New-Object System.Windows.Forms.Button
$RunNowButton.Text = 'Run Now'
$RunNowButton.Location = New-Object System.Drawing.Point(22, 170)
$RunNowButton.Size = New-Object System.Drawing.Size(150, 36)
$RunNowButton.Add_Click({ $script:Choice = 'Now' })
$Form.Controls.Add($RunNowButton)

$RunLaterButton = New-Object System.Windows.Forms.Button
$RunLaterButton.Text = 'Run Later'
$RunLaterButton.Location = New-Object System.Drawing.Point(182, 170)
$RunLaterButton.Size = New-Object System.Drawing.Size(150, 36)
$RunLaterButton.Add_Click({ $script:Choice = 'Later' })
$Form.Controls.Add($RunLaterButton)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(22, 173)
$ProgressBar.Size = New-Object System.Drawing.Size(560, 23)
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = $PrinterNames.Count
$ProgressBar.Style = 'Continuous'
$ProgressBar.Visible = $false
$Form.Controls.Add($ProgressBar)

$ProgressDetailLabel = New-Object System.Windows.Forms.Label
$ProgressDetailLabel.Location = New-Object System.Drawing.Point(22, 204)
$ProgressDetailLabel.Size = New-Object System.Drawing.Size(560, 35)
$ProgressDetailLabel.Text = ''
$ProgressDetailLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$ProgressDetailLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$ProgressDetailLabel.Visible = $false
$Form.Controls.Add($ProgressDetailLabel)

$Form.Show()
$Form.Activate()
$Form.BringToFront()
[System.Windows.Forms.Application]::DoEvents()

function Update-Progress {
    param([string]$Message, [int]$Step = -1, [string]$Detail = '')

    if ($Form.IsDisposed) { return }

    $StatusLabel.Text = $Message
    if ($Detail) { $ProgressDetailLabel.Text = $Detail }

    if ($Step -ge 0) {
        $ProgressBar.Value = [Math]::Min($ProgressBar.Maximum, [Math]::Max($ProgressBar.Minimum, $Step))
    }

    $Form.Refresh()
    [System.Windows.Forms.Application]::DoEvents()
}

# -------------------------
# UI Automation helpers
# -------------------------

function Find-PrinterWindow {
    param([string]$PrinterName)

    $ExpectedTitle = "$PrinterName Printing Preferences"
    $Desktop = [System.Windows.Automation.AutomationElement]::RootElement

    $Windows = $Desktop.FindAll(
        [System.Windows.Automation.TreeScope]::Children,
        [System.Windows.Automation.Condition]::TrueCondition
    )

    foreach ($Window in $Windows) {
        try {
            if ($Window.Current.Name -eq $ExpectedTitle) { return $Window }
        }
        catch {}
    }

    return $null
}

function Wait-ForPrinterWindow {
    param([string]$PrinterName, [int]$TimeoutSeconds = 30)

    $Deadline = [DateTime]::UtcNow.AddSeconds($TimeoutSeconds)

    while ([DateTime]::UtcNow -lt $Deadline) {
        $Window = Find-PrinterWindow -PrinterName $PrinterName
        if ($Window) { return $Window }
        [System.Windows.Forms.Application]::DoEvents()
    }

    throw "Timed out waiting for Printing Preferences for '$PrinterName'."
}

function Find-ControlById {
    param([System.Windows.Automation.AutomationElement]$Root, [string]$AutomationId)

    $Condition = [System.Windows.Automation.PropertyCondition]::new(
        [System.Windows.Automation.AutomationElement]::AutomationIdProperty,
        $AutomationId
    )

    return $Root.FindFirst([System.Windows.Automation.TreeScope]::Descendants, $Condition)
}

function Click-UiElement {
    param([System.Windows.Automation.AutomationElement]$Element, [string]$Description)

    if (-not $Element) { throw "$Description was not found." }

    $Rect = $Element.Current.BoundingRectangle

    if ($Rect.IsEmpty -or $Rect.Width -le 0 -or $Rect.Height -le 0) {
        throw "$Description has no usable screen position."
    }

    $X = [int]($Rect.Left + ($Rect.Width / 2))
    $Y = [int]($Rect.Top + ($Rect.Height / 2))

    [SharpPrinterPinMouse]::ClickAt($X, $Y)
}

function Close-PrinterDialog {
    param([string]$PrinterName)

    try {
        $Window = Find-PrinterWindow -PrinterName $PrinterName

        if ($Window) {
            $CancelButton = Find-ControlById -Root $Window -AutomationId '2'

            if ($CancelButton) {
                $Invoke = $null

                if ($CancelButton.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$Invoke)) {
                    $Invoke.Invoke()
                }
            }
        }
    }
    catch {}
}

# -------------------------
# Apply and verify one printer
# -------------------------

function Set-PrinterPin {
    param([string]$PrinterName)

    try {
        Update-Progress -Message 'Configuring printer...' -Detail $PrinterName

        $Arguments = "printui.dll,PrintUIEntry /e /n `"$PrinterName`""
        Start-Process -FilePath "$env:WINDIR\System32\rundll32.exe" -ArgumentList $Arguments

        $Window = Wait-ForPrinterWindow -PrinterName $PrinterName

        $JobHandling = Find-ControlById -Root $Window -AutomationId '10026'
        if (-not $JobHandling) { throw 'Job Handling tab was not found.' }

        Click-UiElement -Element $JobHandling -Description 'Job Handling tab'

        $Window = Wait-ForPrinterWindow -PrinterName $PrinterName

        $Authentication = Find-ControlById -Root $Window -AutomationId '20340'
        if (-not $Authentication) { throw 'Authentication dropdown was not found.' }

        $Expand = $null

        if ($Authentication.TryGetCurrentPattern([System.Windows.Automation.ExpandCollapsePattern]::Pattern, [ref]$Expand)) {
            if ($Expand.Current.ExpandCollapseState -eq [System.Windows.Automation.ExpandCollapseState]::Collapsed) {
                $Expand.Expand()
            }
        }
        else {
            throw 'Authentication dropdown cannot be expanded through UI Automation.'
        }

        $Window = Wait-ForPrinterWindow -PrinterName $PrinterName

        $UserNumberOption = $Window.FindFirst(
            [System.Windows.Automation.TreeScope]::Descendants,
            [System.Windows.Automation.PropertyCondition]::new(
                [System.Windows.Automation.AutomationElement]::NameProperty,
                'User Number'
            )
        )

        if (-not $UserNumberOption) { throw 'User Number option was not found in Authentication.' }

        $Selection = $null

        if ($UserNumberOption.TryGetCurrentPattern([System.Windows.Automation.SelectionItemPattern]::Pattern, [ref]$Selection)) {
            $Selection.Select()
        }
        else {
            $InvokeOption = $null

            if ($UserNumberOption.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$InvokeOption)) {
                $InvokeOption.Invoke()
            }
            else {
                throw 'User Number option cannot be selected.'
            }
        }

        $Window = Wait-ForPrinterWindow -PrinterName $PrinterName

        $UserNumberField = Find-ControlById -Root $Window -AutomationId '14572'

        if (-not $UserNumberField -or -not $UserNumberField.Current.IsEnabled) {
            throw 'User Number field was not found or is disabled.'
        }

        $ValuePattern = $null

        if ($UserNumberField.TryGetCurrentPattern([System.Windows.Automation.ValuePattern]::Pattern, [ref]$ValuePattern)) {
            $ValuePattern.SetValue($TestPin)
        }
        else {
            throw 'User Number field does not support direct value entry.'
        }

        $Window = Wait-ForPrinterWindow -PrinterName $PrinterName
        $OkButton = Find-ControlById -Root $Window -AutomationId '1'

        if (-not $OkButton) { throw 'OK button was not found.' }

        $InvokeOK = $null

        if ($OkButton.TryGetCurrentPattern([System.Windows.Automation.InvokePattern]::Pattern, [ref]$InvokeOK)) {
            $InvokeOK.Invoke()
        }
        else {
            throw 'OK button does not support Invoke.'
        }

        $PinStatus = Get-PrinterPinStatus -PrinterName $PrinterName

        if (-not $PinStatus.IsConfigured) {
            throw "Post-configuration verification failed: $($PinStatus.Details)"
        }

        return [PSCustomObject]@{ Printer = $PrinterName; Status = 'SUCCESS'; Details = 'PIN applied and verified.' }
    }
    catch {
        Close-PrinterDialog -PrinterName $PrinterName

        return [PSCustomObject]@{ Printer = $PrinterName; Status = 'FAILED'; Details = $_.Exception.Message }
    }
}

# -------------------------
# Main execution
# -------------------------

try {
    # Run Now / Run Later (no answer for 10 minutes counts as Run Later)
    $PromptDeadline = [DateTime]::UtcNow.AddMinutes(10)
    while (-not $script:Choice) {
        if ([DateTime]::UtcNow -gt $PromptDeadline) { $script:Choice = 'Later'; break }
        [System.Windows.Forms.Application]::DoEvents()
        Start-Sleep -Milliseconds 100
    }

    if ($script:Choice -eq 'Later') {
        $Results += [PSCustomObject]@{ Printer = '(all)'; Status = 'DEFERRED'; Details = 'User chose Run Later (or did not respond).' }
    }
    else {
        $RunNowButton.Visible = $false
        $RunLaterButton.Visible = $false
        $ProgressBar.Visible = $true
        $ProgressDetailLabel.Visible = $true

        Update-Progress -Message 'Starting printer configuration...' -Step 0 -Detail 'Please do not use your keyboard or mouse.'

        for ($Index = 0; $Index -lt $PrinterNames.Count; $Index++) {
            $PrinterName = $PrinterNames[$Index]
            $Step = $Index + 1

            Update-Progress -Message "Configuring printer $Step of $($PrinterNames.Count)..." -Step ($Step - 1) -Detail $PrinterName

            $Result = Set-PrinterPin -PrinterName $PrinterName
            $Results += $Result

            Update-Progress -Message "$($Result.Status): printer $Step of $($PrinterNames.Count)" -Step $Step -Detail $PrinterName
        }

        $SuccessCount = @($Results | Where-Object Status -eq 'SUCCESS').Count
        $FailedCount  = @($Results | Where-Object Status -eq 'FAILED').Count

        if ($FailedCount -gt 0) {
            Update-Progress -Message 'Printer setup completed with errors.' -Step $PrinterNames.Count `
                -Detail "$SuccessCount successful, $FailedCount failed. Contact IT if assistance is needed."
        }
        else {
            Update-Progress -Message 'Printer setup complete.' -Step $PrinterNames.Count `
                -Detail "$SuccessCount successful. You may now use your computer."
        }
    }
}
catch {
    $Results += [PSCustomObject]@{ Printer = '(all)'; Status = 'FAILED'; Details = $_.Exception.Message }
}
finally {
    if ($Form -and -not $Form.IsDisposed) {
        $Form.Close()
        $Form.Dispose()
    }
}

ConvertTo-Json -InputObject @($Results | Select-Object Printer, Status, Details) |
    Set-Content -LiteralPath $ResultsPath -Encoding UTF8
