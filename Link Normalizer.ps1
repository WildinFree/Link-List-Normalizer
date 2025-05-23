# Load required assemblies
try {
    Add-Type -AssemblyName PresentationFramework
    Add-Type -AssemblyName System.Windows.Forms
    Write-Host "✓ Loaded required assemblies" -ForegroundColor Green
}
catch {
    Write-Error "Failed to load assemblies: $($_.Exception.Message)"
    Write-ErrorLog "Failed to load assemblies: $($_.Exception.Message)"
    return
}

# Setup logging
$logFile = "$env:APPDATA\HostsConverter\debug.log"
$errorLogFile = "C:\Users\amant\Documents\Programming Related Stuff\Chrome Extentions\Porn Blocker - Project Midnight Arrow\Blocker Resources\Powershell Output\Logs\error.txt"
$logQueue = New-Object 'System.Collections.Concurrent.ConcurrentQueue[string]'
$logDir = Split-Path $logFile -Parent
$errorLogDir = Split-Path $errorLogFile -Parent
if (-not (Test-Path $logDir)) { New-Item -ItemType Directory -Path $logDir -Force | Out-Null }
if (-not (Test-Path $errorLogDir)) { New-Item -ItemType Directory -Path $errorLogDir -Force | Out-Null }

# Start async logging job
$logJob = Start-ThreadJob -ScriptBlock {
    param($queue, $logFile)
    while ($true) {
        $messages = @()
        while ($queue.TryDequeue([ref]$null)) {
            $messages += $null
        }
        if ($messages) {
            $messages | Out-File -FilePath $logFile -Append -Encoding UTF8
        }
        Start-Sleep -Milliseconds 100
    }
} -ArgumentList $logQueue, $logFile

function Write-Log {
    param($Message)
    $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
    $logQueue.Enqueue("$timestamp - $Message")
}

function Write-ErrorLog {
    param($Message)
    try {
        $timestamp = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
        "$timestamp - $Message" | Out-File -FilePath $errorLogFile -Append -Encoding UTF8
    }
    catch {
        Write-Host "✗ Failed to write to error log: $($_.Exception.Message)" -ForegroundColor Red
    }
}

Write-Log "Script started"
Write-Host "✓ Script initialized" -ForegroundColor Green

# Store last output folder
$configFile = "$env:APPDATA\HostsConverter\lastOutputFolder.txt"
$lastOutputFolder = if (Test-Path $configFile) { Get-Content $configFile -ErrorAction SilentlyContinue } else { $env:USERPROFILE }
Write-Log "Last output folder $lastOutputFolder"

# Optimize thread pool
$threadLimit = [math]::Min([Environment]::ProcessorCount * 2, 16)
$os = Get-CimInstance Win32_OperatingSystem
$freeMemory = $os.FreePhysicalMemory / 1MB
if ($freeMemory -lt 1000) { $threadLimit = [math]::Max(2, $threadLimit / 2) }
Write-Log "Thread pool set to $threadLimit threads based on $freeMemory MB free memory"

# Compile regex pattern
$combinedRegex = [regex]::new('^(?:(?:\|\||@@\|\||[\+\-0]+[\-\.]?|0\.0\.0\.0|127\.0\.0\.1|::1|\*)?\s*(?<domain1>[a-zA-Z0-9\-\.]+)(?:\^[\$\w,=~]*|$|\/.*\/)|https?://(?:www\.)?(?<domain2>[^\s/]+?)(?:[/?#].*|$)|\/[\w\-_\?]*(?<domain3>[a-zA-Z0-9\-\.]+?)[\w\-_\?]*\/|(?<domain4>[a-zA-Z0-9][a-zA-Z0-9\-\.]*[a-zA-Z0-9]))', 'Compiled')

# XAML for UI (unchanged)
$xaml = @"
<Window xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        Title="Hosts Converter" Height="200" Width="400" Background="#1E1E1E"
        WindowStyle="SingleBorderWindow" ResizeMode="NoResize" WindowStartupLocation="CenterScreen">
    <Window.Resources>
        <Style TargetType="Button">
            <Setter Property="Background" Value="#3C3C3C"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="10,5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Template">
                <Setter.Value>
                    <ControlTemplate TargetType="Button">
                        <Border Background="{TemplateBinding Background}" 
                                BorderBrush="{TemplateBinding BorderBrush}" 
                                BorderThickness="1" CornerRadius="5">
                            <ContentPresenter HorizontalAlignment="Center" VerticalAlignment="Center"/>
                        </Border>
                    </ControlTemplate>
                </Setter.Value>
            </Setter>
            <Style.Triggers>
                <Trigger Property="IsMouseOver" Value="True">
                    <Setter Property="Background" Value="#4A4A4A"/>
                </Trigger>
                <Trigger Property="IsEnabled" Value="False">
                    <Setter Property="Opacity" Value="0.5"/>
                </Trigger>
            </Style.Triggers>
        </Style>
        <Style TargetType="TextBox">
            <Setter Property="Background" Value="#2D2D2D"/>
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="BorderBrush" Value="#555555"/>
            <Setter Property="Padding" Value="5"/>
            <Setter Property="Margin" Value="5"/>
            <Setter Property="FontSize" Value="14"/>
        </Style>
        <Style TargetType="TextBlock">
            <Setter Property="Foreground" Value="#FFFFFF"/>
            <Setter Property="FontSize" Value="14"/>
            <Setter Property="Margin" Value="5"/>
        </Style>
    </Window.Resources>
    <Grid Margin="10">
        <Grid.RowDefinitions>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
            <RowDefinition Height="Auto"/>
        </Grid.RowDefinitions>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="Auto"/>
            <ColumnDefinition Width="*"/>
            <ColumnDefinition Width="Auto"/>
        </Grid.ColumnDefinitions>
        
        <TextBlock Grid.Row="0" Grid.Column="0" Text="Input File:" VerticalAlignment="Center"/>
        <TextBox x:Name="InputFileTextBox" Grid.Row="0" Grid.Column="1" Margin="5,0"/>
        <Button x:Name="BrowseInputButton" Grid.Row="0" Grid.Column="2" Content="Browse" Width="80"/>
        
        <TextBlock Grid.Row="1" Grid.Column="0" Text="Output Folder:" VerticalAlignment="Center"/>
        <TextBox x:Name="OutputFolderTextBox" Grid.Row="1" Grid.Column="1" Margin="5,0" Text="$lastOutputFolder"/>
        <Button x:Name="BrowseOutputButton" Grid.Row="1" Grid.Column="2" Content="Browse" Width="80"/>
        
        <Button x:Name="ConvertButton" Grid.Row="2" Grid.Column="1" Grid.ColumnSpan="2" Content="Convert" Background="#8B0000" Width="100" Margin="5,10"/>
        
        <TextBlock x:Name="StatusTextBlock" Grid.Row="3" Grid.Column="0" Grid.ColumnSpan="3" Foreground="#FF5555" Visibility="Hidden"/>
    </Grid>
</Window>
"@

# Main loop
do {
    # Create the WPF window
    try {
        $window = [Windows.Markup.XamlReader]::Parse($xaml)
        Write-Log "XAML parsed successfully"
        Write-Host "✓ UI loaded successfully" -ForegroundColor Green
    }
    catch {
        Write-Log "Failed to parse XAML: $($_.Exception.Message)"
        Write-Host "✗ Failed to parse XAML: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog "Failed to parse XAML: $($_.Exception.Message)"
        continue
    }

    # Find UI elements
    $inputFileTextBox = $window.FindName("InputFileTextBox")
    $browseInputButton = $window.FindName("BrowseInputButton")
    $outputFolderTextBox = $window.FindName("OutputFolderTextBox")
    $browseOutputButton = $window.FindName("BrowseOutputButton")
    $convertButton = $window.FindName("ConvertButton")
    $statusTextBlock = $window.FindName("StatusTextBlock")

    # Verify UI elements
    $uiElements = @{
        "InputFileTextBox" = $inputFileTextBox
        "BrowseInputButton" = $browseInputButton
        "OutputFolderTextBox" = $outputFolderTextBox
        "BrowseOutputButton" = $browseOutputButton
        "ConvertButton" = $convertButton
        "StatusTextBlock" = $statusTextBlock
    }
    foreach ($element in $uiElements.GetEnumerator()) {
        if (-not $element.Value) {
            Write-Log "UI element $($element.Key) not found"
            Write-Host "✗ UI element $($element.Key) not found" -ForegroundColor Red
            Write-ErrorLog "UI element $($element.Key) not found"
        } else {
            Write-Log "UI element $($element.Key) initialized"
        }
    }

    # Variables
    $script:inputFile = ""
    $script:outputFolder = $lastOutputFolder

    # Browse input file
    if ($browseInputButton) {
        $browseInputButton.Add_Click({
            $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog
            $openFileDialog.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
            if ($openFileDialog.ShowDialog() -eq 'OK') {
                $inputFileTextBox.Text = $openFileDialog.FileName
                $script:inputFile = $openFileDialog.FileName
                Write-Log "Input file selected: $script:inputFile"
                Write-Host "✓ Input file selected: $script:inputFile" -ForegroundColor Green
            }
        })
    }

    # Browse output folder
    if ($browseOutputButton) {
        $browseOutputButton.Add_Click({
            $folderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
            $folderBrowser.SelectedPath = $script:outputFolder
            if ($folderBrowser.ShowDialog() -eq 'OK') {
                $outputFolderTextBox.Text = $folderBrowser.SelectedPath
                $script:outputFolder = $folderBrowser.SelectedPath
                Write-Log "Output folder selected: $script:outputFolder"
                Write-Host "✓ Output folder selected: $script:outputFolder" -ForegroundColor Green
            }
        })
    }

    # Convert button
    if ($convertButton) {
        $convertButton.Add_Click({
            Write-Log "Convert button clicked"
            if (-not $inputFileTextBox.Text) {
                $statusTextBlock.Text = "Please select an input file."
                $statusTextBlock.Visibility = "Visible"
                Write-Log "Conversion failed: No input file selected"
                Write-Host "✗ No input file selected" -ForegroundColor Red
                Write-ErrorLog "No input file selected"
                return
            }
            if (-not $outputFolderTextBox.Text) {
                $statusTextBlock.Text = "Please select an output folder."
                $statusTextBlock.Visibility = "Visible"
                Write-Log "Conversion failed: No output folder selected"
                Write-Host "✗ No output folder selected" -ForegroundColor Red
                Write-ErrorLog "No output folder selected"
                return
            }
            $script:inputFile = $inputFileTextBox.Text
            $script:outputFolder = $outputFolderTextBox.Text
            try {
                Set-Content -Path $configFile -Value $script:outputFolder -ErrorAction Stop
                Write-Log "Saved output folder to config: $configFile"
            }
            catch {
                Write-Log "Failed to save config file: $($_.Exception.Message)"
                Write-Host "✗ Failed to save config file: $($_.Exception.Message)" -ForegroundColor Red
                Write-ErrorLog "Failed to save config file: $($_.Exception.Message)"
            }
            Write-Log "Closing UI to start conversion"
            Write-Host "✓ Starting conversion process" -ForegroundColor Green
            $window.Close()
        })
    }

    # Show the window
    try {
        Write-Log "Showing window"
        Write-Host "✓ Displaying UI" -ForegroundColor Green
        $window.ShowDialog() | Out-Null
    }
    catch {
        Write-Log "Failed to show window: $($_.Exception.Message)"
        Write-Host "✗ Failed to show window: $($_.Exception.Message)" -ForegroundColor Red
        Write-ErrorLog "Failed to show window: $($_.Exception.Message)"
        continue
    }

    # Validate inputs
    if (-not $script:inputFile) {
        Write-Host "✗ No input file provided. Skipping." -ForegroundColor Red
        Write-Log "No input file provided. Skipping."
        Write-ErrorLog "No input file provided"
        continue
    }
    if (-not (Test-Path $script:inputFile)) {
        Write-Host "✗ Input file does not exist: $script:inputFile" -ForegroundColor Red
        Write-Log "Input file does not exist: $script:inputFile"
        Write-ErrorLog "Input file does not exist: $script:inputFile"
        continue
    }
    if (-not $script:outputFolder) {
        Write-Host "✗ No output folder provided. Skipping." -ForegroundColor Red
        Write-Log "No output folder provided. Skipping."
        Write-ErrorLog "No output folder provided"
        continue
    }
    if (-not (Test-Path $script:outputFolder)) {
        try {
            New-Item -ItemType Directory -Path $script:outputFolder -Force | Out-Null
            Write-Log "Created output folder: $script:outputFolder"
            Write-Host "✓ Created output folder: $script:outputFolder" -ForegroundColor Green
        }
        catch {
            Write-Host "✗ Error creating output folder: $($_.Exception.Message)" -ForegroundColor Red
            Write-Log "Error creating output folder: $($_.Exception.Message)"
            Write-ErrorLog "Error creating output folder: $($_.Exception.Message)"
            continue
        }
    }

    # Create timestamped folder
    $timestamp = Get-Date -Format "yyyyMMdd"
    $inputFileName = [System.IO.Path]::GetFileNameWithoutExtension($script:inputFile)
    $outputDir = Join-Path $script:outputFolder "$inputFileName - $timestamp"
    try {
        New-Item -ItemType Directory -Path $outputDir -Force | Out-Null
        Write-Log "Created timestamped folder: $outputDir"
        Write-Host "✓ Created timestamped folder: $outputDir" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error creating timestamped folder: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error creating timestamped folder: $($_.Exception.Message)"
        Write-ErrorLog "Error creating timestamped folder: $($_.Exception.Message)"
        continue
    }

    # Set output file paths
    $txtFile = Join-Path $outputDir "$inputFileName - converted domains.txt"
    $csvFile = Join-Path $outputDir "$inputFileName - converted domains.csv"
    $jsonFile = Join-Path $outputDir "$inputFileName - converted domains.json"
    $stateFile = Join-Path $outputDir "conversion_state.xml"

    # Calculate batch size
    $batchSize = [math]::Max(50000, [math]::Min(200000, [math]::Floor($freeMemory * 1000)))
    Write-Log "Batch size set to $batchSize based on $freeMemory MB free memory"

    # Initialize collections
    $uniqueDomains = New-Object 'System.Collections.Generic.HashSet[string]'
    $duplicateCounts = @{}
    $totalLines = 0
    $invalidEntries = 0

    # Check for resume state
    $resume = $false
    if (Test-Path $stateFile) {
        try {
            $state = Import-Clixml -Path $stateFile
            foreach ($domain in $state.UniqueDomains) { $uniqueDomains.Add($domain) | Out-Null }
            foreach ($entry in $state.DuplicateCounts) { $duplicateCounts[$entry.Key] = $entry.Value }
            $resume = $true
            Write-Log "Resuming from state file: $stateFile"
            Write-Host "✓ Resuming from previous state" -ForegroundColor Green
        }
        catch {
            Write-Log "Error loading state file: $($_.Exception.Message)"
            Write-Host "✗ Error loading resume state, starting fresh: $($_.Exception.Message)" -ForegroundColor Yellow
            Write-ErrorLog "Error loading state file: $($_.Exception.Message)"
        }
    }

    # Define script block for processing
    $scriptBlock = {
        param($lines, $combinedRegex)
        $localDomains = New-Object 'System.Collections.Generic.HashSet[string]'
        $localDuplicates = @{}
        foreach ($line in $lines) {
            $line = $line.Trim()
            if (-not $line) { continue }
            $match = $combinedRegex.Match($line)
            $domain = $match.Groups['domain1'].Value + $match.Groups['domain2'].Value + $match.Groups['domain3'].Value + $match.Groups['domain4'].Value
            if ($domain) {
                $domain = $domain -replace '^\d+\.|^[\w\-]+\.\d+\.|^ads--', ''
                if ($domain -match '^[a-zA-Z0-9][a-zA-Z0-9\-\.]*[a-zA-Z0-9]$') {
                    if (-not $localDomains.Add($domain)) {
                        $localDuplicates[$domain] = ($localDuplicates[$domain] + 1) -as [int]
                    }
                } else {
                    $localDuplicates['Invalid'] = ($localDuplicates['Invalid'] + 1) -as [int]
                }
            } else {
                $localDuplicates['Invalid'] = ($localDuplicates['Invalid'] + 1) -as [int]
            }
        }
        return [PSCustomObject]@{
            Domains = $localDomains
            Duplicates = $localDuplicates
        }
    }

    # Process file in streaming mode
    Write-Host "⏳ Processing input file..." -ForegroundColor Cyan
    $batchBuffer = New-Object 'System.Collections.Generic.List[string]' -ArgumentList $batchSize
    $reader = $null
    $stream = $null
    $jobs = New-Object 'System.Collections.Generic.List[psobject]'
    $runspacePool = $null
    $useThreadJobs = $PSVersionTable.PSVersion.Major -ge 7

    try {
        if (-not $useThreadJobs) {
            $runspacePool = [runspacefactory]::CreateRunspacePool(1, $threadLimit)
            $runspacePool.Open()
        }
        $stream = [System.IO.FileStream]::new($script:inputFile, [System.IO.FileMode]::Open, [System.IO.FileAccess]::Read, [System.IO.FileShare]::Read, 8192)
        $reader = [System.IO.StreamReader]::new($stream, [System.Text.Encoding]::UTF8, $true, 8192)
        $batchIndex = 0
        $lineCount = 0
        while ($line = $reader.ReadLine()) {
            $batchBuffer.Add($line)
            $lineCount++
            $totalLines++
            if ($batchBuffer.Count -ge $batchSize) {
                if ($resume) {
                    Write-Log "Skipping batch $batchIndex due to resume"
                    $batchBuffer.Clear()
                    $batchIndex++
                    continue
                }
                if ($useThreadJobs) {
                    $job = Start-ThreadJob -ScriptBlock $scriptBlock -ArgumentList @($batchBuffer.ToArray(), $combinedRegex)
                    $jobs.Add($job)
                } else {
                    $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($batchBuffer.ToArray()).AddArgument($combinedRegex)
                    $powershell.RunspacePool = $runspacePool
                    $jobs.Add([PSCustomObject]@{
                        PowerShell = $powershell
                        Handle = $powershell.BeginInvoke()
                    })
                }
                $batchBuffer.Clear()
                $batchIndex++
                Write-Host "  ✓ Dispatched batch $batchIndex ($lineCount lines)" -ForegroundColor Green
                Write-Log "Dispatched batch $batchIndex"
            }
        }
        if ($batchBuffer.Count -gt 0 -and -not $resume) {
            if ($useThreadJobs) {
                $job = Start-ThreadJob -ScriptBlock $scriptBlock -ArgumentList @($batchBuffer.ToArray(), $combinedRegex)
                $jobs.Add($job)
            } else {
                $powershell = [powershell]::Create().AddScript($scriptBlock).AddArgument($batchBuffer.ToArray()).AddArgument($combinedRegex)
                $powershell.RunspacePool = $runspacePool
                $jobs.Add([PSCustomObject]@{
                    PowerShell = $powershell
                    Handle = $powershell.BeginInvoke()
                })
            }
            Write-Host "  ✓ Dispatched final batch ($lineCount lines)" -ForegroundColor Green
            Write-Log "Dispatched final batch"
        }
    }
    catch {
        Write-Host "✗ Error processing input file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error processing input file: $($_.Exception.Message)"
        Write-ErrorLog "Error processing input file: $($_.Exception.Message)"
        continue
    }
    finally {
        if ($reader) { $reader.Close() }
        if ($stream) { $stream.Close() }
    }

    # Collect results
    Write-Host "⏳ Collecting results..." -ForegroundColor Cyan
    try {
        if ($useThreadJobs) {
            foreach ($job in $jobs) {
                try {
                    $result = $job | Receive-Job -Wait -AutoRemoveJob
                    foreach ($domain in $result.Domains) { $uniqueDomains.Add($domain) | Out-Null }
                    foreach ($entry in $result.Duplicates.GetEnumerator()) {
                        $duplicateCounts[$entry.Key] = ($duplicateCounts[$entry.Key] + $entry.Value) -as [int]
                    }
                }
                catch {
                    Write-Log "Error processing job: $($_.Exception.Message)"
                    Write-ErrorLog "Error processing job: $($_.Exception.Message)"
                }
            }
        } else {
            while ($jobs.Count -gt 0) {
                foreach ($job in $jobs.ToArray()) {
                    if ($job.Handle.IsCompleted) {
                        try {
                            $result = $job.PowerShell.EndInvoke($job.Handle)
                            foreach ($domain in $result.Domains) { $uniqueDomains.Add($domain) | Out-Null }
                            foreach ($entry in $result.Duplicates.GetEnumerator()) {
                                $duplicateCounts[$entry.Key] = ($duplicateCounts[$entry.Key] + $entry.Value) -as [int]
                            }
                        }
                        catch {
                            Write-Log "Error processing runspace: $($_.Exception.Message)"
                            Write-ErrorLog "Error processing runspace: $($_.Exception.Message)"
                        }
                        finally {
                            $job.PowerShell.Dispose()
                            $jobs.Remove($job)
                        }
                    }
                }
                Start-Sleep -Milliseconds 10
            }
        }
        $invalidEntries = $duplicateCounts['Invalid'] -as [int]
        Write-Log "Processed $totalLines lines, found $($uniqueDomains.Count) unique domains"
        Write-Host "✓ Processed $totalLines lines, found $($uniqueDomains.Count) unique domains" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error collecting results: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error collecting results: $($_.Exception.Message)"
        Write-ErrorLog "Error collecting results: $($_.Exception.Message)"
    }
    finally {
        if ($runspacePool) {
            $runspacePool.Close()
            $runspacePool.Dispose()
        }
    }

    # Save state for resume
    try {
        $state = @{
            UniqueDomains = @($uniqueDomains)
            DuplicateCounts = @($duplicateCounts.GetEnumerator() | ForEach-Object { @{Key = $_.Key; Value = $_.Value} })
        }
        Export-Clixml -InputObject $state -Path $stateFile -Force
        Write-Log "Saved state to: $stateFile"
    }
    catch {
        Write-Log "Error saving state: $($_.Exception.Message)"
        Write-ErrorLog "Error saving state: $($_.Exception.Message)"
    }

    # Write output files
    Write-Host "⏳ Writing output files..." -ForegroundColor Cyan
    $txtWriter = $null
    $csvWriter = $null
    $jsonWriter = $null
    try {
        $txtWriter = [System.IO.StreamWriter]::new($txtFile, $false, [System.Text.Encoding]::UTF8, 1048576)
        $csvWriter = [System.IO.StreamWriter]::new($csvFile, $false, [System.Text.Encoding]::UTF8, 1048576)
        $jsonWriter = [System.IO.StreamWriter]::new($jsonFile, $false, [System.Text.Encoding]::UTF8, 1048576)
        $csvWriter.WriteLine("Domain,DuplicateCount")
        $jsonWriter.WriteLine('{"domains":[')
        $txtBuffer = New-Object 'System.Text.StringBuilder' -ArgumentList 1000000
        $csvBuffer = New-Object 'System.Text.StringBuilder' -ArgumentList 1000000
        $jsonBuffer = New-Object 'System.Text.StringBuilder' -ArgumentList 1000000
        $firstJson = $true
        $count = 0
        foreach ($domain in $uniqueDomains) {
            $txtBuffer.AppendLine($domain)
            $csvBuffer.AppendLine("`"$domain`",`$($duplicateCounts[$domain])")
            if (-not $firstJson) { $jsonBuffer.Append(",") }
            $jsonBuffer.AppendLine("`"$domain`"")
            $firstJson = $false
            $count++
            if ($count % $batchSize -eq 0) {
                $txtWriter.Write($txtBuffer.ToString())
                $csvWriter.Write($csvBuffer.ToString())
                $jsonWriter.Write($jsonBuffer.ToString())
                $txtBuffer.Clear()
                $csvBuffer.Clear()
                $jsonBuffer.Clear()
            }
        }
        if ($txtBuffer.Length -gt 0) {
            $txtWriter.Write($txtBuffer.ToString())
            $csvWriter.Write($csvBuffer.ToString())
            $jsonWriter.Write($jsonBuffer.ToString())
        }
        $jsonWriter.WriteLine('],"metadata":{"totalDomains":' + $uniqueDomains.Count + ',"duplicates":' + ($duplicateCounts.Values | Where-Object { $_ -gt 0 } | Measure-Object -Sum).Sum + '}}')
        Write-Log "Wrote outputs: TXT $txtFile, CSV $csvFile, JSON $jsonFile"
        Write-Host "✓ Final outputs saved:" -ForegroundColor Green
        Write-Host "  TXT: $txtFile ($($uniqueDomains.Count) domains)" -ForegroundColor Green
        Write-Host "  CSV: $csvFile ($($uniqueDomains.Count) domains)" -ForegroundColor Green
        Write-Host "  JSON: $jsonFile ($($uniqueDomains.Count) domains)" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error writing outputs: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error writing outputs: $($_.Exception.Message)"
        Write-ErrorLog "Error writing outputs: $($_.Exception.Message)"
    }
    finally {
        if ($txtWriter) { $txtWriter.Close() }
        if ($csvWriter) { $csvWriter.Close() }
        if ($jsonWriter) { $jsonWriter.Close() }
    }

    # Log statistics
    $duplicateReport = $duplicateCounts.GetEnumerator() | Where-Object { $_.Key -ne 'Invalid' -and $_.Value -gt 0 } | ForEach-Object { "$($_.Key) found $($_.Value) duplicates" }
    if ($duplicateReport) {
        Write-Host "`nDuplicate domains found:" -ForegroundColor Yellow
        $duplicateReport | ForEach-Object {
            Write-Host "  $_" -ForegroundColor Yellow
            Write-Log $_
        }
    } else {
        Write-Host "`n✓ No duplicate domains found." -ForegroundColor Green
        Write-Log "No duplicate domains found."
    }
    Write-Host "Summary Statistics:" -ForegroundColor Cyan
    Write-Host "  Total lines processed: $totalLines" -ForegroundColor Green
    Write-Host "  Total unique domains: $($uniqueDomains.Count)" -ForegroundColor Green
    Write-Host "  Total duplicates: $(($duplicateCounts.Values | Where-Object { $_ -gt 0 } | Measure-Object -Sum).Sum)" -ForegroundColor Yellow
    Write-Host "  Total invalid entries: $invalidEntries" -ForegroundColor Yellow
    Write-Log "Total unique domains: $($uniqueDomains.Count)"

    # Clean up state file
    try {
        if (Test-Path $stateFile) { Remove-Item -Path $stateFile -Force }
        Write-Log "Cleaned up state file: $stateFile"
        Write-Host "✓ Cleaned up state file" -ForegroundColor Green
    }
    catch {
        Write-Host "✗ Error cleaning up state file: $($_.Exception.Message)" -ForegroundColor Red
        Write-Log "Error cleaning up state file: $($_.Exception.Message)"
        Write-ErrorLog "Error cleaning up state file: $($_.Exception.Message)"
    }

    Write-Host "`n✓ Conversion complete!" -ForegroundColor Green
    Write-Log "Conversion completed successfully"

    # Prompt to convert another file
    Write-Host "`nWant to convert another file? (Y/N)" -ForegroundColor Cyan
    $response = Read-Host
    Write-Log "User response to convert another file: $response"
    $continue = $response -eq 'Y' -or $response -eq 'y'
} while ($continue)

# Final cleanup
Write-Log "Exiting script"
Write-Host "✓ Script exiting" -ForegroundColor Green
try { $logJob | Stop-Job -Force; $logJob | Remove-Job } catch {}