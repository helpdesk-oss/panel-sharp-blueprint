param([Parameter(Mandatory)][string]$ConfigPath)

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
# Progress window
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
$WarningLabel.Text = 'PLEASE DO NOT USE YOUR KEYBOARD OR MOUSE WHILE THIS PROCESS IS RUNNING.'
$WarningLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9, [System.Drawing.FontStyle]::Bold)
$WarningLabel.ForeColor = [System.Drawing.Color]::Red
$WarningLabel.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
$WarningPanel.Controls.Add($WarningLabel)

$InfoLabel = New-Object System.Windows.Forms.Label
$InfoLabel.Location = New-Object System.Drawing.Point(22, 72)
$InfoLabel.Size = New-Object System.Drawing.Size(560, 48)
$InfoLabel.Text = @(
    'This process should take about one minute',
    'If it does not finish promptly, please contact your IT department',
    'This script runs once automatically to apply your User Number to office and staff room printers'
) -join "`r`n"
$InfoLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
$InfoLabel.ForeColor = [System.Drawing.Color]::FromArgb(65, 65, 65)
$Form.Controls.Add($InfoLabel)

$StatusLabel = New-Object System.Windows.Forms.Label
$StatusLabel.Location = New-Object System.Drawing.Point(22, 130)
$StatusLabel.Size = New-Object System.Drawing.Size(560, 35)
$StatusLabel.Text = 'Preparing printer configuration...'
$StatusLabel.Font = New-Object System.Drawing.Font('Segoe UI', 10, [System.Drawing.FontStyle]::Bold)
$StatusLabel.ForeColor = [System.Drawing.Color]::FromArgb(30, 45, 65)
$Form.Controls.Add($StatusLabel)

$ProgressBar = New-Object System.Windows.Forms.ProgressBar
$ProgressBar.Location = New-Object System.Drawing.Point(22, 173)
$ProgressBar.Size = New-Object System.Drawing.Size(560, 23)
$ProgressBar.Minimum = 0
$ProgressBar.Maximum = $PrinterNames.Count
$ProgressBar.Style = 'Continuous'
$Form.Controls.Add($ProgressBar)

$ProgressDetailLabel = New-Object System.Windows.Forms.Label
$ProgressDetailLabel.Location = New-Object System.Drawing.Point(22, 204)
$ProgressDetailLabel.Size = New-Object System.Drawing.Size(560, 35)
$ProgressDetailLabel.Text = 'Preparing to update printer settings...'
$ProgressDetailLabel.ForeColor = [System.Drawing.Color]::FromArgb(100, 100, 100)
$ProgressDetailLabel.Font = New-Object System.Drawing.Font('Segoe UI', 9)
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
    for ($Countdown = 5; $Countdown -ge 1; $Countdown--) {
        Update-Progress -Message "Starting in $Countdown seconds..." -Step 0 -Detail 'Please leave this window open.'
        Start-Sleep -Seconds 1
    }

    Update-Progress -Message 'Starting printer configuration...' -Step 0 -Detail 'Only printers needing changes will be opened.'

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
