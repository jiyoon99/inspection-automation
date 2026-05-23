Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing

[System.Windows.Forms.Application]::EnableVisualStyles()

$script:AppRoot = Split-Path -Parent $MyInvocation.MyCommand.Path
$script:DataDir = Join-Path $script:AppRoot "data"
$script:CsvPath = Join-Path $script:DataDir "inspections.csv"
$script:ReportDir = Join-Path $script:AppRoot "reports"
$script:BatteryHealthNeedsCheck = $false
$script:DiagnosticSummary = ""

if (-not (Test-Path -LiteralPath $script:DataDir)) {
    New-Item -ItemType Directory -Path $script:DataDir | Out-Null
}

if (-not (Test-Path -LiteralPath $script:ReportDir)) {
    New-Item -ItemType Directory -Path $script:ReportDir | Out-Null
}

function Get-InspectionResult {
    param(
        [string[]]$Statuses
    )

    if ($Statuses -contains "불량") {
        return "불합격"
    }

    if ($Statuses -contains "확인필요") {
        return "재검수"
    }

    return "합격"
}

function Get-ExistingRows {
    if (Test-Path -LiteralPath $script:CsvPath) {
        return Import-Csv -LiteralPath $script:CsvPath
    }

    return @()
}

function Save-InspectionRow {
    param(
        [pscustomobject]$Row
    )

    $rows = @()
    if (Test-Path -LiteralPath $script:CsvPath) {
        $rows = Import-Csv -LiteralPath $script:CsvPath
    }

    $rows += $Row
    $rows | Export-Csv -LiteralPath $script:CsvPath -NoTypeInformation -Encoding utf8
}

function Get-EdgePath {
    $candidates = @(
        "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Microsoft\Edge\Application\msedge.exe",
        "C:\Program Files\Google\Chrome\Application\chrome.exe"
    )

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate) {
            return $candidate
        }
    }

    return $null
}

function Convert-ToSafeFilePart {
    param(
        [string]$Value,
        [string]$Fallback = "report"
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Fallback
    }

    $safe = $Value -replace '[\\/:*?"<>|]', '_'
    $safe = $safe.Trim()
    if ([string]::IsNullOrWhiteSpace($safe)) {
        return $Fallback
    }

    return $safe
}

function Get-CurrentInspectionData {
    return [ordered]@{
        "검수일시" = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        "검수자" = $txtInspector.Text.Trim()
        "모델명" = $txtModel.Text.Trim()
        "시리얼" = $txtSerial.Text.Trim()
        "CPU" = $txtCpu.Text.Trim()
        "RAM/SSD" = $txtMemory.Text.Trim()
        "등급" = $cmbGrade.SelectedItem.ToString()
        "배터리 잔량" = ("{0}%" -f [int]$numBattery.Value)
        "배터리 효율" = $txtBatteryHealth.Text.Trim()
        "자동 판정" = $lblResult.Text
        "부팅" = $statusControls["boot"].SelectedItem.ToString()
        "액정" = $statusControls["display"].SelectedItem.ToString()
        "키보드" = $statusControls["keyboard"].SelectedItem.ToString()
        "터치패드" = $statusControls["touchpad"].SelectedItem.ToString()
        "외관" = $statusControls["body"].SelectedItem.ToString()
        "충전기" = $statusControls["adapter"].SelectedItem.ToString()
        "와이파이" = $statusControls["wifi"].SelectedItem.ToString()
        "카메라" = $statusControls["camera"].SelectedItem.ToString()
        "스피커" = $statusControls["speaker"].SelectedItem.ToString()
        "포트" = $statusControls["ports"].SelectedItem.ToString()
        "비고" = $txtNotes.Text.Trim()
        "자동 진단 리포트" = $script:DiagnosticSummary
    }
}

function New-InspectionHtmlReport {
    param(
        [hashtable]$Data
    )

    $rows = foreach ($entry in $Data.GetEnumerator()) {
        $label = [System.Net.WebUtility]::HtmlEncode([string]$entry.Key)
        $value = [System.Net.WebUtility]::HtmlEncode([string]$entry.Value)
        $value = $value -replace "(\r\n|\r|\n)", "<br/>"
        "<tr><th>$label</th><td>$value</td></tr>"
    }

    $resultClass = switch ($Data["자동 판정"]) {
        "합격" { "pass" }
        "재검수" { "check" }
        default { "fail" }
    }

    @"
<!DOCTYPE html>
<html lang="ko">
<head>
<meta charset="utf-8" />
<title>노트북 검수 결과서</title>
<style>
body { font-family: 'Malgun Gothic', sans-serif; margin: 32px; color: #222; }
h1 { margin-bottom: 8px; }
.meta { color: #666; margin-bottom: 24px; }
.result { display: inline-block; padding: 8px 16px; border-radius: 8px; font-weight: bold; margin-bottom: 20px; }
.pass { background: #e8f7ec; color: #1f7a39; }
.check { background: #fff4df; color: #b26b00; }
.fail { background: #fdeaea; color: #b42318; }
table { width: 100%; border-collapse: collapse; }
th, td { border: 1px solid #d9d9d9; padding: 10px; text-align: left; vertical-align: top; }
th { width: 180px; background: #f5f7fa; }
</style>
</head>
<body>
<h1>노트북 검수 결과서</h1>
<div class="meta">자동 생성 문서</div>
<div class="result $resultClass">최종 판정: $([System.Net.WebUtility]::HtmlEncode([string]$Data["자동 판정"]))</div>
<table>
$($rows -join [Environment]::NewLine)
</table>
</body>
</html>
"@
}

function Export-InspectionPdf {
    $edgePath = Get-EdgePath
    if (-not $edgePath) {
        throw "PDF 생성을 위한 Edge 또는 Chrome을 찾지 못했습니다."
    }

    $data = Get-CurrentInspectionData
    $stamp = Get-Date -Format "yyyyMMdd_HHmmss"
    $serialPart = Convert-ToSafeFilePart -Value $data["시리얼"] -Fallback "no-serial"
    $baseName = "{0}_{1}" -f $stamp, $serialPart
    $htmlPath = Join-Path $script:ReportDir ($baseName + ".html")
    $pdfPath = Join-Path $script:ReportDir ($baseName + ".pdf")
    $profileDir = Join-Path $script:ReportDir ("_browser_profile_" + $baseName)

    if (-not (Test-Path -LiteralPath $script:ReportDir)) {
        New-Item -ItemType Directory -Path $script:ReportDir | Out-Null
    }
    if (-not (Test-Path -LiteralPath $profileDir)) {
        New-Item -ItemType Directory -Path $profileDir | Out-Null
    }

    $html = New-InspectionHtmlReport -Data $data
    [System.IO.File]::WriteAllText($htmlPath, $html, [System.Text.UTF8Encoding]::new($true))

    $uri = [System.Uri]::new($htmlPath).AbsoluteUri
    $arguments = @(
        "--headless=new",
        "--disable-gpu",
        "--no-first-run",
        "--no-default-browser-check",
        "--user-data-dir=$profileDir",
        "--allow-file-access-from-files",
        "--print-to-pdf=$pdfPath",
        $uri
    )

    $process = Start-Process -FilePath $edgePath -ArgumentList $arguments -PassThru -Wait -WindowStyle Hidden
    Start-Sleep -Milliseconds 1500
    if (-not (Test-Path -LiteralPath $pdfPath)) {
        throw "PDF 파일 생성에 실패했습니다. HTML 결과서는 저장되었습니다: $htmlPath"
    }

    return [pscustomobject]@{
        HtmlPath = $htmlPath
        PdfPath = $pdfPath
    }
}

function Refresh-Grid {
    param(
        [System.Windows.Forms.DataGridView]$Grid
    )

    $Grid.DataSource = $null
    $data = Get-ExistingRows | Sort-Object inspected_at -Descending
    if ($data.Count -gt 0) {
        $Grid.DataSource = $data
        $Grid.AutoResizeColumns()
    }
}

function New-StatusCombo {
    param(
        [int]$X,
        [int]$Y
    )

    $combo = New-Object System.Windows.Forms.ComboBox
    $combo.Location = New-Object System.Drawing.Point($X, $Y)
    $combo.Size = New-Object System.Drawing.Size(120, 24)
    $combo.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
    [void]$combo.Items.AddRange(@("정상", "확인필요", "불량"))
    $combo.SelectedIndex = 0
    return $combo
}

function Invoke-SafeCommand {
    param(
        [scriptblock]$ScriptBlock
    )

    try {
        return & $ScriptBlock
    }
    catch {
        return $null
    }
}

function Get-FirstValue {
    param(
        [object[]]$Values
    )

    foreach ($value in $Values) {
        if ($null -ne $value -and -not [string]::IsNullOrWhiteSpace([string]$value)) {
            return [string]$value
        }
    }

    return ""
}

function Get-BatteryReportInfo {
    $reportPath = Join-Path $script:DataDir "battery-report.html"
    $generated = Invoke-SafeCommand {
        if (-not (Test-Path -LiteralPath $script:DataDir)) {
            New-Item -ItemType Directory -Path $script:DataDir | Out-Null
        }
        powercfg /batteryreport /output $reportPath | Out-Null
        Get-Content -LiteralPath $reportPath -Raw -Encoding UTF8
    }

    if ([string]::IsNullOrWhiteSpace($generated)) {
        return $null
    }

    $designMatch = [regex]::Match($generated, 'DESIGN CAPACITY</span></td><td>\s*([\d,]+)\s*mWh', 'IgnoreCase')
    $fullMatch = [regex]::Match($generated, 'FULL CHARGE CAPACITY</span></td><td>\s*([\d,]+)\s*mWh', 'IgnoreCase')
    $cycleMatch = [regex]::Match($generated, 'CYCLE COUNT</span></td><td>\s*([\d,]+)', 'IgnoreCase')

    $designCapacity = if ($designMatch.Success) { [int](($designMatch.Groups[1].Value) -replace ',', '') } else { 0 }
    $fullChargeCapacity = if ($fullMatch.Success) { [int](($fullMatch.Groups[1].Value) -replace ',', '') } else { 0 }
    $cycleCount = if ($cycleMatch.Success) { [int](($cycleMatch.Groups[1].Value) -replace ',', '') } else { $null }

    $healthPercent = $null
    if ($designCapacity -gt 0 -and $fullChargeCapacity -gt 0) {
        $healthPercent = [math]::Round(($fullChargeCapacity / $designCapacity) * 100, 1)
    }

    [pscustomobject]@{
        DesignCapacityMWh = $designCapacity
        FullChargeCapacityMWh = $fullChargeCapacity
        CycleCount = $cycleCount
        HealthPercent = $healthPercent
    }
}

function Show-KeyboardTestDialog {
    $testForm = New-Object System.Windows.Forms.Form
    $testForm.Text = "키보드 테스트"
    $testForm.StartPosition = "CenterParent"
    $testForm.Size = New-Object System.Drawing.Size(760, 520)
    $testForm.MinimumSize = New-Object System.Drawing.Size(760, 520)
    $testForm.KeyPreview = $true

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "키를 눌러 입력 상태를 확인하세요. 테스트한 키는 아래에 누적 표시됩니다."
    $infoLabel.Location = New-Object System.Drawing.Point(20, 20)
    $infoLabel.Size = New-Object System.Drawing.Size(680, 24)
    $testForm.Controls.Add($infoLabel)

    $pressedTitle = New-Object System.Windows.Forms.Label
    $pressedTitle.Text = "입력된 키"
    $pressedTitle.Location = New-Object System.Drawing.Point(20, 55)
    $pressedTitle.Size = New-Object System.Drawing.Size(120, 24)
    $testForm.Controls.Add($pressedTitle)

    $pressedBox = New-Object System.Windows.Forms.TextBox
    $pressedBox.Location = New-Object System.Drawing.Point(20, 82)
    $pressedBox.Size = New-Object System.Drawing.Size(700, 260)
    $pressedBox.Multiline = $true
    $pressedBox.ScrollBars = "Vertical"
    $pressedBox.ReadOnly = $true
    $pressedBox.BackColor = [System.Drawing.Color]::White
    $testForm.Controls.Add($pressedBox)

    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Text = "미입력 키나 불량 키가 있으면 아래에 직접 적어둘 수 있습니다."
    $noteLabel.Location = New-Object System.Drawing.Point(20, 355)
    $noteLabel.Size = New-Object System.Drawing.Size(420, 24)
    $testForm.Controls.Add($noteLabel)

    $issueBox = New-Object System.Windows.Forms.TextBox
    $issueBox.Location = New-Object System.Drawing.Point(20, 382)
    $issueBox.Size = New-Object System.Drawing.Size(700, 40)
    $testForm.Controls.Add($issueBox)

    $btnPass = New-Object System.Windows.Forms.Button
    $btnPass.Text = "정상 반영"
    $btnPass.Location = New-Object System.Drawing.Point(20, 435)
    $btnPass.Size = New-Object System.Drawing.Size(120, 34)
    $btnPass.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 230)
    $btnPass.ForeColor = [System.Drawing.Color]::White
    $btnPass.FlatStyle = "Flat"
    $testForm.Controls.Add($btnPass)

    $btnCheck = New-Object System.Windows.Forms.Button
    $btnCheck.Text = "확인필요 반영"
    $btnCheck.Location = New-Object System.Drawing.Point(155, 435)
    $btnCheck.Size = New-Object System.Drawing.Size(120, 34)
    $btnCheck.FlatStyle = "Flat"
    $testForm.Controls.Add($btnCheck)

    $btnFail = New-Object System.Windows.Forms.Button
    $btnFail.Text = "불량 반영"
    $btnFail.Location = New-Object System.Drawing.Point(290, 435)
    $btnFail.Size = New-Object System.Drawing.Size(120, 34)
    $btnFail.FlatStyle = "Flat"
    $testForm.Controls.Add($btnFail)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "닫기"
    $btnClose.Location = New-Object System.Drawing.Point(600, 435)
    $btnClose.Size = New-Object System.Drawing.Size(120, 34)
    $btnClose.FlatStyle = "Flat"
    $testForm.Controls.Add($btnClose)

    $pressedKeys = New-Object System.Collections.Generic.List[string]
    $resultStatus = $null

    $testForm.Add_KeyDown({
        param($sender, $e)

        $keyName = $e.KeyCode.ToString()
        if (-not $pressedKeys.Contains($keyName)) {
            $pressedKeys.Add($keyName)
            $pressedBox.Text = $pressedKeys -join ", "
        }
    })

    $btnPass.Add_Click({
        $script:KeyboardTestResult = [pscustomobject]@{
            Status = "정상"
            Summary = if ([string]::IsNullOrWhiteSpace($issueBox.Text)) { "키 입력 테스트 정상" } else { "키 입력 테스트 정상 / 메모: $($issueBox.Text.Trim())" }
        }
        $testForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $testForm.Close()
    })

    $btnCheck.Add_Click({
        $script:KeyboardTestResult = [pscustomobject]@{
            Status = "확인필요"
            Summary = if ([string]::IsNullOrWhiteSpace($issueBox.Text)) { "키 입력 테스트 중 일부 확인 필요" } else { "키 입력 테스트 확인 필요 / 문제: $($issueBox.Text.Trim())" }
        }
        $testForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $testForm.Close()
    })

    $btnFail.Add_Click({
        $script:KeyboardTestResult = [pscustomobject]@{
            Status = "불량"
            Summary = if ([string]::IsNullOrWhiteSpace($issueBox.Text)) { "키보드 불량 확인" } else { "키보드 불량 / 문제: $($issueBox.Text.Trim())" }
        }
        $testForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $testForm.Close()
    })

    $btnClose.Add_Click({
        $testForm.Close()
    })

    [void]$testForm.ShowDialog()
}

function Show-PortTestDialog {
    $testForm = New-Object System.Windows.Forms.Form
    $testForm.Text = "포트 테스트"
    $testForm.StartPosition = "CenterParent"
    $testForm.Size = New-Object System.Drawing.Size(760, 560)
    $testForm.MinimumSize = New-Object System.Drawing.Size(760, 560)

    $infoLabel = New-Object System.Windows.Forms.Label
    $infoLabel.Text = "각 포트의 테스트 결과를 선택하세요. 하나라도 불량이면 포트 항목이 불량으로 반영됩니다."
    $infoLabel.Location = New-Object System.Drawing.Point(20, 20)
    $infoLabel.Size = New-Object System.Drawing.Size(700, 24)
    $testForm.Controls.Add($infoLabel)

    $portItems = @(
        @{ Name = "usb_a"; Label = "USB-A" },
        @{ Name = "usb_c"; Label = "USB-C" },
        @{ Name = "hdmi"; Label = "HDMI" },
        @{ Name = "lan"; Label = "LAN" },
        @{ Name = "audio"; Label = "오디오잭" },
        @{ Name = "charger"; Label = "충전포트" },
        @{ Name = "sd"; Label = "SD카드" }
    )

    $portControls = @{}
    $y = 65
    foreach ($item in $portItems) {
        $label = New-Object System.Windows.Forms.Label
        $label.Text = $item.Label
        $label.Location = New-Object System.Drawing.Point(30, $y)
        $label.Size = New-Object System.Drawing.Size(120, 24)
        $testForm.Controls.Add($label)

        $combo = New-StatusCombo -X 170 -Y ($y - 2)
        $testForm.Controls.Add($combo)
        $portControls[$item.Name] = $combo

        $y += 42
    }

    $noteLabel = New-Object System.Windows.Forms.Label
    $noteLabel.Text = "특이사항"
    $noteLabel.Location = New-Object System.Drawing.Point(30, 370)
    $noteLabel.Size = New-Object System.Drawing.Size(100, 24)
    $testForm.Controls.Add($noteLabel)

    $issueBox = New-Object System.Windows.Forms.TextBox
    $issueBox.Location = New-Object System.Drawing.Point(30, 398)
    $issueBox.Size = New-Object System.Drawing.Size(680, 55)
    $issueBox.Multiline = $true
    $testForm.Controls.Add($issueBox)

    $btnApply = New-Object System.Windows.Forms.Button
    $btnApply.Text = "결과 반영"
    $btnApply.Location = New-Object System.Drawing.Point(30, 475)
    $btnApply.Size = New-Object System.Drawing.Size(120, 34)
    $btnApply.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 230)
    $btnApply.ForeColor = [System.Drawing.Color]::White
    $btnApply.FlatStyle = "Flat"
    $testForm.Controls.Add($btnApply)

    $btnClose = New-Object System.Windows.Forms.Button
    $btnClose.Text = "닫기"
    $btnClose.Location = New-Object System.Drawing.Point(590, 475)
    $btnClose.Size = New-Object System.Drawing.Size(120, 34)
    $btnClose.FlatStyle = "Flat"
    $testForm.Controls.Add($btnClose)

    $btnApply.Add_Click({
        $statuses = foreach ($control in $portControls.Values) { $control.SelectedItem.ToString() }
        $finalStatus = "정상"
        if ($statuses -contains "불량") {
            $finalStatus = "불량"
        }
        elseif ($statuses -contains "확인필요") {
            $finalStatus = "확인필요"
        }

        $details = foreach ($item in $portItems) {
            "{0}:{1}" -f $item.Label, $portControls[$item.Name].SelectedItem.ToString()
        }

        $summary = $details -join ", "
        if (-not [string]::IsNullOrWhiteSpace($issueBox.Text)) {
            $summary += " / 메모: " + $issueBox.Text.Trim()
        }

        $script:PortTestResult = [pscustomobject]@{
            Status = $finalStatus
            Summary = $summary
        }

        $testForm.DialogResult = [System.Windows.Forms.DialogResult]::OK
        $testForm.Close()
    })

    $btnClose.Add_Click({
        $testForm.Close()
    })

    [void]$testForm.ShowDialog()
}

function Get-HardwareSnapshot {
    $computerInfo = Invoke-SafeCommand { Get-ComputerInfo }
    $cs = Invoke-SafeCommand { Get-CimInstance Win32_ComputerSystem }
    $bios = Invoke-SafeCommand { Get-CimInstance Win32_BIOS }
    $processor = Invoke-SafeCommand { Get-CimInstance Win32_Processor | Select-Object -First 1 }
    $cpuCounter = Invoke-SafeCommand { (Get-Counter '\Processor(_Total)\% Processor Time').CounterSamples | Select-Object -First 1 }
    $battery = Invoke-SafeCommand { Get-CimInstance Win32_Battery | Select-Object -First 1 }
    $physicalMemory = Invoke-SafeCommand { Get-CimInstance Win32_PhysicalMemory }
    $diskDrives = Invoke-SafeCommand { Get-CimInstance Win32_DiskDrive }
    $camera = Invoke-SafeCommand { Get-PnpDevice -Class Camera -Status OK }
    $sound = Invoke-SafeCommand { Get-PnpDevice -Class Sound,AudioEndpoint -Status OK }
    $bluetooth = Invoke-SafeCommand { Get-PnpDevice -Class Bluetooth -Status OK }
    $network = Invoke-SafeCommand { Get-NetAdapter -Physical -ErrorAction Stop | Where-Object { $_.Status -eq "Up" -or $_.Status -eq "Disconnected" } }
    $batteryReport = Get-BatteryReportInfo
    $drives = Invoke-SafeCommand { Get-PSDrive -PSProvider FileSystem }

    $memoryBytes = 0
    if ($physicalMemory) {
        $memoryBytes = ($physicalMemory | Measure-Object -Property Capacity -Sum).Sum
    }
    elseif ($cs -and $cs.TotalPhysicalMemory) {
        $memoryBytes = [int64]$cs.TotalPhysicalMemory
    }
    elseif ($computerInfo -and $computerInfo.CsTotalPhysicalMemory) {
        $memoryBytes = [int64]$computerInfo.CsTotalPhysicalMemory
    }

    $diskBytes = 0
    if ($diskDrives) {
        $diskBytes = ($diskDrives | Measure-Object -Property Size -Sum).Sum
    }

    $batteryPercent = $null
    if ($battery -and $battery.EstimatedChargeRemaining -ne $null) {
        $batteryPercent = [int]$battery.EstimatedChargeRemaining
    }

    $cpuUsagePercent = $null
    if ($cpuCounter -and $cpuCounter.CookedValue -ne $null) {
        $cpuUsagePercent = [math]::Round([double]$cpuCounter.CookedValue, 1)
    }

    $notes = New-Object System.Collections.Generic.List[string]
    if (-not $cs -and -not $computerInfo) { $notes.Add("시스템 정보 자동 조회 실패") }
    if (-not $bios -and -not $computerInfo) { $notes.Add("시리얼 자동 조회 실패") }
    if (-not $processor -and -not $computerInfo) { $notes.Add("CPU 자동 조회 실패") }
    if (-not $physicalMemory -and -not $cs -and -not $computerInfo) { $notes.Add("메모리 자동 조회 실패") }
    if (-not $diskDrives) { $notes.Add("디스크 자동 조회 실패") }
    if (-not $battery) { $notes.Add("배터리 장치 조회 실패 또는 미장착") }

    [pscustomobject]@{
        Model = Get-FirstValue @(
            if ($cs) { $cs.Model }
            if ($computerInfo) { $computerInfo.CsModel }
        )
        Serial = Get-FirstValue @(
            if ($bios) { $bios.SerialNumber }
            if ($computerInfo) { $computerInfo.BiosSerialNumber }
        )
        Cpu = Get-FirstValue @(
            if ($processor) { $processor.Name }
            if ($computerInfo -and $computerInfo.CsProcessors) { ($computerInfo.CsProcessors | Select-Object -First 1).Name }
        )
        CpuUsagePercent = $cpuUsagePercent
        MemoryGb = if ($memoryBytes) { [math]::Round($memoryBytes / 1GB, 1) } else { 0 }
        DiskGb = if ($diskBytes) { [math]::Round($diskBytes / 1GB, 1) } else { 0 }
        BatteryPercent = $batteryPercent
        BatteryHealthPercent = if ($batteryReport) { $batteryReport.HealthPercent } else { $null }
        BatteryCycleCount = if ($batteryReport) { $batteryReport.CycleCount } else { $null }
        BatteryDesignCapacityMWh = if ($batteryReport) { $batteryReport.DesignCapacityMWh } else { 0 }
        BatteryFullChargeCapacityMWh = if ($batteryReport) { $batteryReport.FullChargeCapacityMWh } else { 0 }
        WifiDetected = [bool]($network | Where-Object { $_.InterfaceDescription -match "Wireless|Wi-Fi|802\.11" -or $_.Name -match "Wi-Fi|WLAN|Wireless" })
        BluetoothDetected = [bool]($bluetooth)
        CameraDetected = [bool]($camera)
        SpeakerDetected = [bool]($sound)
        FileSystemDrives = $drives
        RawNotes = ($notes -join ", ")
    }
}

function Get-DiagnosticReport {
    param(
        [pscustomobject]$Snapshot
    )

    $findings = New-Object System.Collections.Generic.List[string]
    $hasFail = $false
    $hasWarning = $false

    if (-not $Snapshot) {
        return [pscustomobject]@{
            Summary = "[불량] 자동 진단 데이터를 읽지 못했습니다."
            HasFail = $true
            HasWarning = $false
        }
    }

    if ($Snapshot.MemoryGb -ge 8) {
        $findings.Add("[정상] 메모리 용량: $($Snapshot.MemoryGb)GB")
    }
    elseif ($Snapshot.MemoryGb -gt 0) {
        $findings.Add("[주의] 메모리 용량 낮음: $($Snapshot.MemoryGb)GB")
        $hasWarning = $true
    }
    else {
        $findings.Add("[주의] 메모리 정보를 확인하지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.DiskGb -ge 237) {
        $findings.Add("[정상] 저장장치 총 용량: $($Snapshot.DiskGb)GB")
    }
    elseif ($Snapshot.DiskGb -gt 0) {
        $findings.Add("[주의] 저장장치 용량 낮음: $($Snapshot.DiskGb)GB")
        $hasWarning = $true
    }
    else {
        $findings.Add("[주의] 저장장치 정보를 확인하지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.BatteryHealthPercent -ge 80) {
        $findings.Add("[정상] 배터리 효율: $($Snapshot.BatteryHealthPercent)%")
    }
    elseif ($Snapshot.BatteryHealthPercent -gt 0) {
        $findings.Add("[주의] 배터리 효율 저하: $($Snapshot.BatteryHealthPercent)%")
        $hasWarning = $true
    }
    else {
        $findings.Add("[주의] 배터리 효율 정보를 확인하지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.BatteryCycleCount -ge 500) {
        $findings.Add("[주의] 배터리 사이클 수 높음: $($Snapshot.BatteryCycleCount)")
        $hasWarning = $true
    }
    elseif ($Snapshot.BatteryCycleCount -gt 0) {
        $findings.Add("[정상] 배터리 사이클 수: $($Snapshot.BatteryCycleCount)")
    }

    if (-not [string]::IsNullOrWhiteSpace($Snapshot.Cpu)) {
        $findings.Add("[정상] CPU 정보: $($Snapshot.Cpu)")
    }
    else {
        $findings.Add("[주의] CPU 정보를 확인하지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.CpuUsagePercent -ge 85) {
        $findings.Add("[주의] CPU 사용률 높음: $($Snapshot.CpuUsagePercent)%")
        $hasWarning = $true
    }
    elseif ($Snapshot.CpuUsagePercent -ge 0) {
        $findings.Add("[정상] CPU 사용률: $($Snapshot.CpuUsagePercent)%")
    }

    if ($Snapshot.WifiDetected) {
        $findings.Add("[정상] Wi-Fi 장치 인식")
    }
    else {
        $findings.Add("[불량] Wi-Fi 장치를 찾지 못했습니다.")
        $hasFail = $true
    }

    if ($Snapshot.CameraDetected) {
        $findings.Add("[정상] 카메라 장치 인식")
    }
    else {
        $findings.Add("[주의] 카메라 장치를 찾지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.SpeakerDetected) {
        $findings.Add("[정상] 오디오 장치 인식")
    }
    else {
        $findings.Add("[불량] 오디오 장치를 찾지 못했습니다.")
        $hasFail = $true
    }

    if (-not $Snapshot.BluetoothDetected) {
        $findings.Add("[주의] 블루투스 장치를 찾지 못했습니다.")
        $hasWarning = $true
    }

    if ($Snapshot.FileSystemDrives) {
        foreach ($drive in $Snapshot.FileSystemDrives) {
            $total = [double]($drive.Free + $drive.Used)
            if ($total -gt 0) {
                $freeRatio = [math]::Round(($drive.Free / $total) * 100, 1)
                if ($freeRatio -lt 10) {
                    $findings.Add("[주의] $($drive.Name): 남은 공간 부족 ($freeRatio%)")
                    $hasWarning = $true
                }
                else {
                    $findings.Add("[정상] $($drive.Name): 남은 공간 $freeRatio%")
                }
            }
        }
    }

    if (-not [string]::IsNullOrWhiteSpace($Snapshot.RawNotes)) {
        $findings.Add("[참고] $($Snapshot.RawNotes)")
    }

    [pscustomobject]@{
        Summary = ($findings -join [Environment]::NewLine)
        HasFail = $hasFail
        HasWarning = $hasWarning
    }
}

$form = New-Object System.Windows.Forms.Form
$form.Text = "노트북 검수 자동화"
$form.StartPosition = "CenterScreen"
$form.AutoScaleMode = [System.Windows.Forms.AutoScaleMode]::Dpi
$form.Size = New-Object System.Drawing.Size(1500, 920)
$form.MinimumSize = New-Object System.Drawing.Size(1400, 860)
$form.BackColor = [System.Drawing.Color]::FromArgb(245, 247, 250)

$rootPanel = New-Object System.Windows.Forms.TableLayoutPanel
$rootPanel.Dock = "Fill"
$rootPanel.Padding = New-Object System.Windows.Forms.Padding(24, 20, 24, 20)
$rootPanel.ColumnCount = 1
$rootPanel.RowCount = 8
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 320)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::AutoSize)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 64)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 130)))
$rootPanel.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$form.Controls.Add($rootPanel)

$title = New-Object System.Windows.Forms.Label
$title.Text = "노트북 검수 자동화 프로그램"
$title.Font = New-Object System.Drawing.Font("Malgun Gothic", 18, [System.Drawing.FontStyle]::Bold)
$title.AutoSize = $true
$title.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 6)
$rootPanel.Controls.Add($title, 0, 0)

$subtitle = New-Object System.Windows.Forms.Label
$subtitle.Text = "기본 정보 입력 후 검수 항목을 선택하면 자동으로 판정되고 CSV로 저장됩니다."
$subtitle.Font = New-Object System.Drawing.Font("Malgun Gothic", 9)
$subtitle.ForeColor = [System.Drawing.Color]::DimGray
$subtitle.AutoSize = $true
$subtitle.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 12)
$rootPanel.Controls.Add($subtitle, 0, 1)

$topLayout = New-Object System.Windows.Forms.TableLayoutPanel
$topLayout.Dock = "Fill"
$topLayout.ColumnCount = 2
$topLayout.RowCount = 1
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$topLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$topLayout.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$rootPanel.Controls.Add($topLayout, 0, 2)

$groupInfo = New-Object System.Windows.Forms.GroupBox
$groupInfo.Text = "기본 정보"
$groupInfo.Font = New-Object System.Drawing.Font("Malgun Gothic", 10, [System.Drawing.FontStyle]::Bold)
$groupInfo.Dock = "Fill"
$groupInfo.Padding = New-Object System.Windows.Forms.Padding(12)
$groupInfo.Margin = New-Object System.Windows.Forms.Padding(0, 0, 12, 0)
$topLayout.Controls.Add($groupInfo, 0, 0)

$infoLayout = New-Object System.Windows.Forms.TableLayoutPanel
$infoLayout.Dock = "Fill"
$infoLayout.ColumnCount = 4
$infoLayout.RowCount = 5
$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 95)))
$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 110)))
$infoLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($i = 0; $i -lt 4; $i++) {
    $infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Absolute, 46)))
}
$infoLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 100)))
$groupInfo.Controls.Add($infoLayout)

function New-FieldLabel {
    param([string]$Text)
    $label = New-Object System.Windows.Forms.Label
    $label.Text = $Text
    $label.Dock = "Fill"
    $label.TextAlign = [System.Drawing.ContentAlignment]::MiddleLeft
    $label.Font = New-Object System.Drawing.Font("Malgun Gothic", 9)
    return $label
}

$txtInspector = New-Object System.Windows.Forms.TextBox
$txtInspector.Dock = "Fill"
$txtModel = New-Object System.Windows.Forms.TextBox
$txtModel.Dock = "Fill"
$txtSerial = New-Object System.Windows.Forms.TextBox
$txtSerial.Dock = "Fill"
$txtCpu = New-Object System.Windows.Forms.TextBox
$txtCpu.Dock = "Fill"
$txtMemory = New-Object System.Windows.Forms.TextBox
$txtMemory.Dock = "Fill"

$cmbGrade = New-Object System.Windows.Forms.ComboBox
$cmbGrade.Dock = "Fill"
$cmbGrade.DropDownStyle = [System.Windows.Forms.ComboBoxStyle]::DropDownList
[void]$cmbGrade.Items.AddRange(@("A", "B", "C"))
$cmbGrade.SelectedIndex = 0

$numBattery = New-Object System.Windows.Forms.NumericUpDown
$numBattery.Dock = "Fill"
$numBattery.Minimum = 0
$numBattery.Maximum = 100
$numBattery.Value = 80

$txtBatteryHealth = New-Object System.Windows.Forms.TextBox
$txtBatteryHealth.Dock = "Fill"
$txtBatteryHealth.ReadOnly = $true
$txtBatteryHealth.BackColor = [System.Drawing.Color]::White

$txtNotes = New-Object System.Windows.Forms.TextBox
$txtNotes.Dock = "Fill"
$txtNotes.Multiline = $true
$txtNotes.ScrollBars = "Vertical"

$infoLayout.Controls.Add((New-FieldLabel "검수자"), 0, 0)
$infoLayout.Controls.Add($txtInspector, 1, 0)
$infoLayout.Controls.Add((New-FieldLabel "등급"), 2, 0)
$infoLayout.Controls.Add($cmbGrade, 3, 0)
$infoLayout.Controls.Add((New-FieldLabel "모델명"), 0, 1)
$infoLayout.Controls.Add($txtModel, 1, 1)
$infoLayout.Controls.Add((New-FieldLabel "배터리 잔량 (%)"), 2, 1)
$infoLayout.Controls.Add($numBattery, 3, 1)
$infoLayout.Controls.Add((New-FieldLabel "시리얼"), 0, 2)
$infoLayout.Controls.Add($txtSerial, 1, 2)
$infoLayout.Controls.Add((New-FieldLabel "배터리 효율"), 2, 2)
$infoLayout.Controls.Add($txtBatteryHealth, 3, 2)
$infoLayout.Controls.Add((New-FieldLabel "CPU"), 0, 3)
$infoLayout.Controls.Add($txtCpu, 1, 3)
$infoLayout.Controls.Add((New-FieldLabel "비고"), 2, 3)
$infoLayout.Controls.Add($txtNotes, 3, 3)
$infoLayout.SetRowSpan($txtNotes, 2)
$infoLayout.Controls.Add((New-FieldLabel "RAM/SSD"), 0, 4)
$infoLayout.Controls.Add($txtMemory, 1, 4)

$groupCheck = New-Object System.Windows.Forms.GroupBox
$groupCheck.Text = "검수 항목"
$groupCheck.Font = New-Object System.Drawing.Font("Malgun Gothic", 10, [System.Drawing.FontStyle]::Bold)
$groupCheck.Dock = "Fill"
$groupCheck.Padding = New-Object System.Windows.Forms.Padding(12)
$groupCheck.Margin = New-Object System.Windows.Forms.Padding(12, 0, 0, 0)
$topLayout.Controls.Add($groupCheck, 1, 0)

$checkLayout = New-Object System.Windows.Forms.TableLayoutPanel
$checkLayout.Dock = "Fill"
$checkLayout.ColumnCount = 4
$checkLayout.RowCount = 5
$checkLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
$checkLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
$checkLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Absolute, 90)))
$checkLayout.ColumnStyles.Add((New-Object System.Windows.Forms.ColumnStyle([System.Windows.Forms.SizeType]::Percent, 50)))
for ($i = 0; $i -lt 5; $i++) {
    $checkLayout.RowStyles.Add((New-Object System.Windows.Forms.RowStyle([System.Windows.Forms.SizeType]::Percent, 20)))
}
$groupCheck.Controls.Add($checkLayout)

$checkItems = @(
    @{ Name = "boot"; Label = "부팅"; Row = 0; Col = 0 },
    @{ Name = "display"; Label = "액정"; Row = 1; Col = 0 },
    @{ Name = "keyboard"; Label = "키보드"; Row = 2; Col = 0 },
    @{ Name = "touchpad"; Label = "터치패드"; Row = 3; Col = 0 },
    @{ Name = "body"; Label = "외관"; Row = 4; Col = 0 },
    @{ Name = "adapter"; Label = "충전기"; Row = 0; Col = 2 },
    @{ Name = "wifi"; Label = "와이파이"; Row = 1; Col = 2 },
    @{ Name = "camera"; Label = "카메라"; Row = 2; Col = 2 },
    @{ Name = "speaker"; Label = "스피커"; Row = 3; Col = 2 },
    @{ Name = "ports"; Label = "포트"; Row = 4; Col = 2 }
)

$statusControls = @{}
foreach ($item in $checkItems) {
    $label = New-FieldLabel $item.Label
    $combo = New-StatusCombo -X 0 -Y 0
    $combo.Dock = "Fill"
    $combo.Margin = New-Object System.Windows.Forms.Padding(0, 6, 0, 6)
    $checkLayout.Controls.Add($label, $item.Col, $item.Row)
    $checkLayout.Controls.Add($combo, $item.Col + 1, $item.Row)
    $statusControls[$item.Name] = $combo
}

$resultPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$resultPanel.Dock = "Fill"
$resultPanel.FlowDirection = "TopDown"
$resultPanel.WrapContents = $false
$resultPanel.AutoSize = $true
$resultPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 0)
$rootPanel.Controls.Add($resultPanel, 0, 3)

$lblResultTitle = New-Object System.Windows.Forms.Label
$lblResultTitle.Text = "자동 판정"
$lblResultTitle.Font = New-Object System.Drawing.Font("Malgun Gothic", 10, [System.Drawing.FontStyle]::Bold)
$lblResultTitle.AutoSize = $true
$resultPanel.Controls.Add($lblResultTitle)

$lblResult = New-Object System.Windows.Forms.Label
$lblResult.Text = "합격"
$lblResult.Font = New-Object System.Drawing.Font("Malgun Gothic", 20, [System.Drawing.FontStyle]::Bold)
$lblResult.AutoSize = $true
$lblResult.ForeColor = [System.Drawing.Color]::ForestGreen
$resultPanel.Controls.Add($lblResult)

$descPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$descPanel.Dock = "Fill"
$descPanel.FlowDirection = "TopDown"
$descPanel.WrapContents = $false
$descPanel.AutoSize = $true
$descPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 8)
$rootPanel.Controls.Add($descPanel, 0, 4)

$lblHint = New-Object System.Windows.Forms.Label
$lblHint.Text = "불량이 하나라도 있으면 불합격, 확인필요가 있으면 재검수, 배터리 80 미만도 재검수 기준입니다."
$lblHint.AutoSize = $true
$lblHint.ForeColor = [System.Drawing.Color]::DimGray
$descPanel.Controls.Add($lblHint)

$lblHardware = New-Object System.Windows.Forms.Label
$lblHardware.Text = "하드웨어 자동진단: 모델/시리얼/CPU/RAM/SSD/배터리/장치 인식 여부를 자동 수집합니다."
$lblHardware.AutoSize = $true
$lblHardware.ForeColor = [System.Drawing.Color]::DimGray
$descPanel.Controls.Add($lblHardware)

$buttonPanel = New-Object System.Windows.Forms.FlowLayoutPanel
$buttonPanel.Dock = "Fill"
$buttonPanel.FlowDirection = "LeftToRight"
$buttonPanel.WrapContents = $true
$buttonPanel.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$rootPanel.Controls.Add($buttonPanel, 0, 5)

$btnSave = New-Object System.Windows.Forms.Button
$btnSave.Text = "검수 저장"
$btnSave.Size = New-Object System.Drawing.Size(145, 40)
$btnSave.BackColor = [System.Drawing.Color]::FromArgb(34, 139, 230)
$btnSave.ForeColor = [System.Drawing.Color]::White
$btnSave.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnSave)

$btnReset = New-Object System.Windows.Forms.Button
$btnReset.Text = "입력 초기화"
$btnReset.Size = New-Object System.Drawing.Size(145, 40)
$btnReset.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnReset)

$btnReload = New-Object System.Windows.Forms.Button
$btnReload.Text = "이력 새로고침"
$btnReload.Size = New-Object System.Drawing.Size(145, 40)
$btnReload.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnReload)

$btnHardware = New-Object System.Windows.Forms.Button
$btnHardware.Text = "하드웨어 자동진단"
$btnHardware.Size = New-Object System.Drawing.Size(160, 40)
$btnHardware.BackColor = [System.Drawing.Color]::FromArgb(47, 79, 79)
$btnHardware.ForeColor = [System.Drawing.Color]::White
$btnHardware.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnHardware)

$btnKeyboardTest = New-Object System.Windows.Forms.Button
$btnKeyboardTest.Text = "키보드 테스트"
$btnKeyboardTest.Size = New-Object System.Drawing.Size(145, 40)
$btnKeyboardTest.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnKeyboardTest)

$btnDiagnostic = New-Object System.Windows.Forms.Button
$btnDiagnostic.Text = "전체 자동진단"
$btnDiagnostic.Size = New-Object System.Drawing.Size(145, 40)
$btnDiagnostic.BackColor = [System.Drawing.Color]::FromArgb(92, 124, 250)
$btnDiagnostic.ForeColor = [System.Drawing.Color]::White
$btnDiagnostic.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnDiagnostic)

$btnPortTest = New-Object System.Windows.Forms.Button
$btnPortTest.Text = "포트 테스트"
$btnPortTest.Size = New-Object System.Drawing.Size(145, 40)
$btnPortTest.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnPortTest)

$btnPdf = New-Object System.Windows.Forms.Button
$btnPdf.Text = "PDF 출력"
$btnPdf.Size = New-Object System.Drawing.Size(145, 40)
$btnPdf.BackColor = [System.Drawing.Color]::FromArgb(220, 53, 69)
$btnPdf.ForeColor = [System.Drawing.Color]::White
$btnPdf.FlatStyle = "Flat"
$buttonPanel.Controls.Add($btnPdf)

$groupReport = New-Object System.Windows.Forms.GroupBox
$groupReport.Text = "자동 진단 리포트"
$groupReport.Font = New-Object System.Drawing.Font("Malgun Gothic", 10, [System.Drawing.FontStyle]::Bold)
$groupReport.Dock = "Fill"
$groupReport.Padding = New-Object System.Windows.Forms.Padding(12)
$groupReport.Margin = New-Object System.Windows.Forms.Padding(0, 0, 0, 10)
$rootPanel.Controls.Add($groupReport, 0, 6)

$txtDiagnostic = New-Object System.Windows.Forms.TextBox
$txtDiagnostic.Dock = "Fill"
$txtDiagnostic.Multiline = $true
$txtDiagnostic.ScrollBars = "Vertical"
$txtDiagnostic.ReadOnly = $true
$txtDiagnostic.BackColor = [System.Drawing.Color]::White
$groupReport.Controls.Add($txtDiagnostic)

$grid = New-Object System.Windows.Forms.DataGridView
$grid.Dock = "Fill"
$grid.ReadOnly = $true
$grid.AllowUserToAddRows = $false
$grid.AllowUserToDeleteRows = $false
$grid.SelectionMode = "FullRowSelect"
$grid.AutoSizeColumnsMode = "DisplayedCells"
$rootPanel.Controls.Add($grid, 0, 7)

function Set-StatusIfExists {
    param(
        [string]$Name,
        [string]$Status
    )

    if ($statusControls.ContainsKey($Name)) {
        $statusControls[$Name].SelectedItem = $Status
    }
}

function Update-ResultLabel {
    $statuses = foreach ($control in $statusControls.Values) { $control.SelectedItem.ToString() }
    if ([int]$numBattery.Value -lt 80) {
        $statuses += "확인필요"
    }
    if ($script:BatteryHealthNeedsCheck) {
        $statuses += "확인필요"
    }

    $result = Get-InspectionResult -Statuses $statuses
    $lblResult.Text = $result

    switch ($result) {
        "합격" { $lblResult.ForeColor = [System.Drawing.Color]::ForestGreen }
        "재검수" { $lblResult.ForeColor = [System.Drawing.Color]::DarkOrange }
        "불합격" { $lblResult.ForeColor = [System.Drawing.Color]::Crimson }
    }
}

foreach ($control in $statusControls.Values) {
    $control.Add_SelectedIndexChanged({ Update-ResultLabel })
}

$numBattery.Add_ValueChanged({ Update-ResultLabel })

$btnReset.Add_Click({
    $txtInspector.Text = ""
    $txtModel.Text = ""
    $txtSerial.Text = ""
    $txtCpu.Text = ""
    $txtMemory.Text = ""
    $cmbGrade.SelectedIndex = 0
    $numBattery.Value = 80
    $txtBatteryHealth.Text = ""
    $txtNotes.Text = ""
    $script:BatteryHealthNeedsCheck = $false
    $script:DiagnosticSummary = ""
    $txtDiagnostic.Text = ""

    foreach ($control in $statusControls.Values) {
        $control.SelectedIndex = 0
    }

    Update-ResultLabel
})

$btnReload.Add_Click({
    Refresh-Grid -Grid $grid
})

$btnHardware.Add_Click({
    $snapshot = Get-HardwareSnapshot
    $autoNotes = New-Object System.Collections.Generic.List[string]
    $script:BatteryHealthNeedsCheck = $false

    if (-not [string]::IsNullOrWhiteSpace($snapshot.Model)) {
        $txtModel.Text = $snapshot.Model
    }

    if (-not [string]::IsNullOrWhiteSpace($snapshot.Serial)) {
        $txtSerial.Text = $snapshot.Serial
    }

    if (-not [string]::IsNullOrWhiteSpace($snapshot.Cpu)) {
        $txtCpu.Text = $snapshot.Cpu
    }

    if ($snapshot.MemoryGb -gt 0 -or $snapshot.DiskGb -gt 0) {
        $parts = @()
        if ($snapshot.MemoryGb -gt 0) { $parts += ("RAM {0}GB" -f $snapshot.MemoryGb) }
        if ($snapshot.DiskGb -gt 0) { $parts += ("SSD/HDD {0}GB" -f $snapshot.DiskGb) }
        $txtMemory.Text = $parts -join " / "
    }

    if ($snapshot.BatteryPercent -ne $null) {
        $safeBattery = [math]::Max(0, [math]::Min(100, [int]$snapshot.BatteryPercent))
        $numBattery.Value = $safeBattery
    }

    if ($snapshot.BatteryHealthPercent -ne $null) {
        $healthText = "{0}% (설계 {1} / 완충 {2} mWh" -f $snapshot.BatteryHealthPercent, $snapshot.BatteryDesignCapacityMWh, $snapshot.BatteryFullChargeCapacityMWh
        if ($snapshot.BatteryCycleCount -ne $null) {
            $healthText += ", 사이클 $($snapshot.BatteryCycleCount)"
        }
        $healthText += ")"
        $txtBatteryHealth.Text = $healthText
    }
    else {
        $txtBatteryHealth.Text = ""
    }

    if ($snapshot.MemoryGb -gt 0 -and $snapshot.MemoryGb -lt 8) {
        $autoNotes.Add("RAM 8GB 미만")
    }

    if ($snapshot.DiskGb -gt 0 -and $snapshot.DiskGb -lt 237) {
        $autoNotes.Add("저장장치 256GB 미만")
    }

    if ($snapshot.BatteryPercent -ne $null -and [int]$snapshot.BatteryPercent -lt 80) {
        $autoNotes.Add("배터리 잔량 80% 미만")
    }

    if ($snapshot.BatteryHealthPercent -ne $null -and [double]$snapshot.BatteryHealthPercent -lt 80) {
        $autoNotes.Add("배터리 효율 80% 미만")
        $script:BatteryHealthNeedsCheck = $true
    }

    if ($snapshot.BatteryCycleCount -ne $null -and [int]$snapshot.BatteryCycleCount -ge 500) {
        $autoNotes.Add("배터리 사이클 수 높음")
        $script:BatteryHealthNeedsCheck = $true
    }

    if ($snapshot.WifiDetected) {
        Set-StatusIfExists -Name "wifi" -Status "정상"
    }
    else {
        Set-StatusIfExists -Name "wifi" -Status "확인필요"
        $autoNotes.Add("와이파이 장치 미확인")
    }

    if ($snapshot.CameraDetected) {
        Set-StatusIfExists -Name "camera" -Status "정상"
    }
    else {
        Set-StatusIfExists -Name "camera" -Status "확인필요"
        $autoNotes.Add("카메라 장치 미확인")
    }

    if ($snapshot.SpeakerDetected) {
        Set-StatusIfExists -Name "speaker" -Status "정상"
    }
    else {
        Set-StatusIfExists -Name "speaker" -Status "확인필요"
        $autoNotes.Add("오디오 장치 미확인")
    }

    if (-not $snapshot.BluetoothDetected) {
        $autoNotes.Add("블루투스 장치 미확인")
    }

    if (-not [string]::IsNullOrWhiteSpace($snapshot.RawNotes)) {
        $autoNotes.Add($snapshot.RawNotes)
    }

    if ($autoNotes.Count -gt 0) {
        if ([string]::IsNullOrWhiteSpace($txtNotes.Text)) {
            $txtNotes.Text = "[자동진단] " + ($autoNotes -join ", ")
        }
        else {
            $txtNotes.Text = $txtNotes.Text.Trim() + [Environment]::NewLine + "[자동진단] " + ($autoNotes -join ", ")
        }
    }

    Update-ResultLabel
    [System.Windows.Forms.MessageBox]::Show("하드웨어 자동진단이 완료되었습니다.", "자동진단 완료", "OK", "Information") | Out-Null
})

$btnKeyboardTest.Add_Click({
    $script:KeyboardTestResult = $null
    Show-KeyboardTestDialog

    if ($null -ne $script:KeyboardTestResult) {
        Set-StatusIfExists -Name "keyboard" -Status $script:KeyboardTestResult.Status

        if ([string]::IsNullOrWhiteSpace($txtNotes.Text)) {
            $txtNotes.Text = "[키보드테스트] " + $script:KeyboardTestResult.Summary
        }
        else {
            $txtNotes.Text = $txtNotes.Text.Trim() + [Environment]::NewLine + "[키보드테스트] " + $script:KeyboardTestResult.Summary
        }

        Update-ResultLabel
    }
})

$btnDiagnostic.Add_Click({
    $snapshot = Get-HardwareSnapshot
    $report = Get-DiagnosticReport -Snapshot $snapshot

    $script:DiagnosticSummary = $report.Summary
    $txtDiagnostic.Text = $report.Summary

    if ($report.HasFail) {
        if ([string]::IsNullOrWhiteSpace($txtNotes.Text)) {
            $txtNotes.Text = "[자동리포트] 불량 징후 감지"
        }
        else {
            $txtNotes.Text = $txtNotes.Text.Trim() + [Environment]::NewLine + "[자동리포트] 불량 징후 감지"
        }
    }
    elseif ($report.HasWarning) {
        if ([string]::IsNullOrWhiteSpace($txtNotes.Text)) {
            $txtNotes.Text = "[자동리포트] 주의 항목 있음"
        }
        else {
            $txtNotes.Text = $txtNotes.Text.Trim() + [Environment]::NewLine + "[자동리포트] 주의 항목 있음"
        }
    }

    if ($report.HasFail) {
        $lblResult.Text = "불합격"
        $lblResult.ForeColor = [System.Drawing.Color]::Crimson
    }
    else {
        Update-ResultLabel
    }

    [System.Windows.Forms.MessageBox]::Show("전체 자동진단이 완료되었습니다.", "자동진단 완료", "OK", "Information") | Out-Null
})

$btnPortTest.Add_Click({
    $script:PortTestResult = $null
    Show-PortTestDialog

    if ($null -ne $script:PortTestResult) {
        Set-StatusIfExists -Name "ports" -Status $script:PortTestResult.Status

        if ([string]::IsNullOrWhiteSpace($txtNotes.Text)) {
            $txtNotes.Text = "[포트테스트] " + $script:PortTestResult.Summary
        }
        else {
            $txtNotes.Text = $txtNotes.Text.Trim() + [Environment]::NewLine + "[포트테스트] " + $script:PortTestResult.Summary
        }

        Update-ResultLabel
    }
})

$btnPdf.Add_Click({
    try {
        $exported = Export-InspectionPdf
        [System.Windows.Forms.MessageBox]::Show("PDF 저장 완료: $($exported.PdfPath)", "PDF 출력", "OK", "Information") | Out-Null
    }
    catch {
        [System.Windows.Forms.MessageBox]::Show("PDF 출력 실패: $($_.Exception.Message)", "PDF 출력", "OK", "Error") | Out-Null
    }
})

$btnSave.Add_Click({
    if ([string]::IsNullOrWhiteSpace($txtInspector.Text) -or [string]::IsNullOrWhiteSpace($txtModel.Text) -or [string]::IsNullOrWhiteSpace($txtSerial.Text)) {
        [System.Windows.Forms.MessageBox]::Show("검수자, 모델명, 시리얼은 필수 입력입니다.", "입력 확인", "OK", "Warning") | Out-Null
        return
    }

    $row = [pscustomobject]@{
        inspected_at = (Get-Date).ToString("yyyy-MM-dd HH:mm:ss")
        inspector    = $txtInspector.Text.Trim()
        model        = $txtModel.Text.Trim()
        serial       = $txtSerial.Text.Trim()
        cpu          = $txtCpu.Text.Trim()
        memory_ssd   = $txtMemory.Text.Trim()
        grade        = $cmbGrade.SelectedItem.ToString()
        battery      = [int]$numBattery.Value
        battery_health = $txtBatteryHealth.Text.Trim()
        boot         = $statusControls["boot"].SelectedItem.ToString()
        display      = $statusControls["display"].SelectedItem.ToString()
        keyboard     = $statusControls["keyboard"].SelectedItem.ToString()
        touchpad     = $statusControls["touchpad"].SelectedItem.ToString()
        body         = $statusControls["body"].SelectedItem.ToString()
        adapter      = $statusControls["adapter"].SelectedItem.ToString()
        wifi         = $statusControls["wifi"].SelectedItem.ToString()
        camera       = $statusControls["camera"].SelectedItem.ToString()
        speaker      = $statusControls["speaker"].SelectedItem.ToString()
        ports        = $statusControls["ports"].SelectedItem.ToString()
        result       = $lblResult.Text
        diagnostic_summary = $script:DiagnosticSummary
        notes        = $txtNotes.Text.Trim()
    }

    Save-InspectionRow -Row $row
    Refresh-Grid -Grid $grid
    [System.Windows.Forms.MessageBox]::Show("검수 결과가 저장되었습니다.", "저장 완료", "OK", "Information") | Out-Null
})

Refresh-Grid -Grid $grid
Update-ResultLabel

[void]$form.ShowDialog()
