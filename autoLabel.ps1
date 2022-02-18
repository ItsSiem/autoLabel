# Auto Label v0.2.70
# Siem Gerritsen 2022

$ascii = "
            _        _           _          _ 
           | |      | |         | |        | |
 __ _ _   _| |_ ___ | |     __ _| |__   ___| |
/ _  | | | | __/ _ \| |    / _  | '_ \ / _ \ |
|(_| | |_| |  ||(_) | |___| (_| | |_) |  __/ |
\__,_|\__,_|\__\___/\_____/\__,_|_.__/ \___|_|
"
""
"==== Auto Label v0.2.70 ==== "
"Siem Gerritsen 2022"
Start-Sleep -Milliseconds 1000
Write-Host $ascii
"Dit script is gemaakt voor gebruik bij QueenSystems"
"Als je problemen tegenkomt laat deze dan achter in het Verbeterpunten.txt bestand"
Start-Sleep -Milliseconds 2000
""
"Het systeem wordt nu gescant..."

function throwError($msg){
    $msg
    Write-Host "Press any key to exit..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
}

function throwWarning($msg) {
    $msg
    Write-Host "Press any key to continue..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    return
}

function wrapText( $text, $width = 23 ) {
    $string = ""
    $col = 0
    $words = $text -split "/"
    foreach ( $word in $words ) {
        $col += $word.Length + 1
        if ( $col -gt $width ) {
            $string += "`n"
            $col = $word.Length + 1
        }
        $string += "$word "
    }
    return $string
}

function Get-CPUs {
    $processors = get-wmiobject win32_processor
    
    $sockets = @(@($processors) |
        ForEach-Object { $_.SocketDesignation } |
        select-object -unique).count;

    return [int]$sockets
}

# ==== MODEL RELATED OPERATIONS ====
$modelRegexes = @(
    "HP Z\w+( \d{2}\w? G\d)?", # HP Systems (HP Z840, HP ZBook 15 G3, HP ZBook 14U G5)
    "Precision \w* \w*"            # DELL Systems (Precision WorkStation T3500, Precision Tower 3620)
)

$model = (Get-WmiObject Win32_ComputerSystem).Model

for ($i = 0; $i -lt $modelRegexes.Count; $i++) {
    if ($model -match $modelRegexes[$i]) {
        $model = $Matches[0]  # Is the first index allways the correct one?
        break
    } 
}

# ==== MEMORY RELATED OPERATIONS ====
$ramSize = [string]((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb).tostring().Split(",")[0] + "GB"
$mem = Get-WmiObject Win32_PhysicalMemory

if ($mem.SMBIOSMemoryType -eq 26) {
    $memType = "DDR4"
}
else {
    $memType = "DDR3"
}

# ==== CPU RELATED OPERATIONS ====
$cpuRegexes = @(
    # regex voor vrijwel alle core i en xeon processoren vanaf 2006
    "(((Platinum)|(Gold)|(Silver)|(Bronze)) \w*-*\d+\w*)|(\w+-*\d{3,}\w*( v\d)*)"
    # TODO Meer regexes voor amd processoren en misschien andere niet compatibles
)

$cpu = Get-WmiObject -class win32_processor
for ($i = 0; $i -lt $cpuRegexes.Count; $i++) {
    if ((-join $cpu.Name) -match $cpuRegexes[$i]) {
        $cpuName = $Matches[0]  # Is the first index allways the correct one?
        break
    } 
}

$cpuAmount = if ((Get-CPUs) -eq 2) { "2x" } else { "" }

# ==== GPU RELATED OPERATIONS ====
$gpus = Get-WmiObject Win32_VideoController
$gpuNames = ""
$gpuRegexes = @(
    "\w{2,3} Graphics \w+", # Intel intergrated graphics (HD Graphics 405, Pro Graphics 600)
    "Quadro (RTX )*\w+",          # Quadro's (Quadro RTX 4000, Quadro K2200, Quadro M2000M)
    "GeForce \wTX \d{3,}( \w+)*"    # Nvidia GeForce GTX / RTX 3060 Ti
)

for ($i = 0; $i -lt $gpuRegexes.Count; $i++) {
    if (($gpus.Name -join "\") -match $gpuRegexes[$i]) {
        foreach ($match in $matches) {
            $gpuNames += $match[0] + "/"
        }
    }
}

# Check of een zbook wel intel graphics aan heeft staan
if ($model -match "HP Z\w+( \d{2}\w? G\d)" -and $gpuNames -match $gpuRegexes[1] -and $gpuNames -notmatch $gpuRegexes[0]) {
    throwWarning("ZBook met een enkele GPU gedetecteerd, staan de graphics in de bios op 'Hybrid'?")
}

if ($gpuNames -eq "" -or $null -eq $gpuNames) {
    throwError("Er is geen GPU gevonden, zijn de drivers wel geinstalleerd?")
}

# ==== DISK RELATED OPERATIONS ====
$osDiskID = (Get-WmiObject Win32_DiskPartition | Where-Object { $_.BootPartition -eq "true" } | Select-Object -first 1).deviceID.Substring(6, 1)
$language = (Get-WmiObject win32_operatingsystem).MUILanguages.Substring(3, 2)
$diskTable = @{}
$disks = Get-PhysicalDisk

foreach ($disk in $disks) {
    [string]$size = [math]::Round(($disk.size / 1000000000))
    $unit = "GB"
    if ($size.length -gt 3) {
        $size = $size.ToString().substring(0, ($size.Length - 3))
        $unit = "TB"
    }
    if ($disk.MediaType -eq "SSD") {
        if ($disk.busType -eq "SATA" -or $disk.busType -eq "RAID") {
            $type = "SSD"
        }
        elseif ($disk.busType -eq "NVMe") {
            $type = "NVMe"
        }
    }
    elseif ($disk.MediaType -eq "HDD") {
        $type = "HDD"
    }
    else {
        continue
    }
    $suffix = ""
    if ($disk.DeviceID -eq $osDiskID) {
        $suffix += " + W10P"
        if ($language -ne "NL") {
            $suffix += " $language"
        }
    }
    $instance = "$size$unit $type$suffix"
    
    foreach ($type in $diskTable.Keys) {
        if ($instance -eq $type) {
            $diskTable.$type++
            continue
        }
    }
    $diskTable.Add($instance, 1)
    continue
}
foreach ($entry in $diskTable.keys) {
    $multiplier = ""
    if ($($diskTable.$entry) -gt 1) {
       $multiplier = $($diskTable.$entry).ToString()+"x" 
    }
    $diskLines += $multiplier+$entry+"/"
}

# ==== OUTPUT TEXT OPERATIONS ====
[string]$outputText = "$Model/$cpuAmount$cpuName/$ramSize $memType/$diskLines$gpuNames"
[string]$wrappedText = wrapText($outputText)
"===================="
$wrappedText
"===================="

Write-Host "Controleer de bovenstaande systeem specificaties"
Write-Host "Druk op een knop om door te gaan..."
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")

# ==== PRINTING RELATED OPERATIONS ====
Add-Type -AssemblyName System.Drawing

(New-Object -ComObject WScript.Network).AddWindowsPrinterConnection("\\HP-Z400\labelPrinter")

$PrintDocument = New-Object System.Drawing.Printing.PrintDocument
$PrintDocument.PrinterSettings.PrinterName = '\\HP-Z400\DYMO LabelWriter 400 Turbo (Kopie 1)'
# $PrintDocument.PrinterSettings.PrinterName = 'Microsoft Print To PDF'
# $PrintDocument.PrinterSettings.PrinterName = 'Brother MFC-L2710DW series Printer'
$PrintDocument.DocumentName = "autoLabel automatic printjob"
$PrintDocument.DefaultPageSettings.PaperSize = $PrintDocument.PrinterSettings.PaperSizes | Where-Object { $_.PaperName -eq '11354 Multi-Purpose' }
$PrintDocument.DefaultPageSettings.Landscape = $false   # Unnececary?

$PrintDocument.add_PrintPage({
        # Create font and colors for text and background
        $Font = [System.Drawing.Font]::new('Arial', 12, [System.Drawing.FontStyle]::Bold)
        $BrushFG = [System.Drawing.SolidBrush]::new([System.Drawing.Color]::FromArgb(255, 0, 0, 0))

        $Width = 1
        # Draw text to the right of the image
        $_.Graphics.DrawString($wrappedText, $Font, $BrushFG, ($Width), 0)
    })

$PrintDocument.Print()
(New-Object -ComObject WScript.Network).RemovePrinterConnection("\\HP-Z400\labelPrinter")

Write-Host "Druk op een knop om te sluiten"
$Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")