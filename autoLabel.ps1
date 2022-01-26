# Auto Label v0.2.1
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
"==== Auto Label v0.2.1 ==== "
"Siem Gerritsen 2022"
Start-Sleep -Milliseconds 1000
Write-Host $ascii
"Dit script is gemaakt voor gebruik bij QueenSystems"
"Als je problemen tegenkomt laat deze dan achter in het Verbeterpunten.txt bestand"
Start-Sleep -Milliseconds 2000
""
"Het systeem wordt nu gescant..."

function throwError($msg = "Een of meerdere componenten in dit systeem worden nog niet ondersteund") {
    $msg
    Write-Host "Press any key to exit..."
    $Host.UI.RawUI.ReadKey("NoEcho,IncludeKeyDown")
    exit
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
    "HP Z\w+( \d{2}\w? G\d)?",     # HP Systems (HP Z840, HP ZBook 15 G3, HP ZBook 14U G5)
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
$ramSize = [string]((Get-CimInstance Win32_PhysicalMemory | Measure-Object -Property capacity -Sum).sum / 1gb).tostring().Split(".")[0] + "GB"
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

$cpuAmount = if((Get-CPUs) -eq 2) {"2x"} else {""}

# ==== GPU RELATED OPERATIONS ====
$gpus = Get-WmiObject Win32_VideoController
$gpuNames = ""
$gpuRegexes = @(
    "\w{2,3} Graphics \w+",      # Intel intergrated graphics (HD Graphics 405)
    "Quadro (RTX )*\w+"          # Quadro's (Quadro RTX 4000, Quadro K2200, Quadro M2000M)
)

# NEW CODE
for ($i = 0; $i -lt $gpuRegexes.Count; $i++) {
    if (($gpus.Name -join "\") -match $gpuRegexes[$i]) {
        foreach ($match in $matches) {
            $gpuNames += $match[0] + "/"
        }
    }
}

if ($gpuNames -eq "" -or $null -eq $gpuNames) {
    throwError("Er is geen GPU gevonden, zijn de drivers wel geinstalleerd?")
}

# ==== DISK RELATED OPERATIONS ====
$OS = Get-WmiObject -Class win32_operatingsystem
$osPartition = Get-WmiObject Win32_DiskPartition | Where-Object { $_.BootPartition -eq "true" }
$osDiskID = $osPartition.deviceID.Substring(6, 1)
$language = $OS.MUILanguages.Substring(3, 2)

$disks = Get-PhysicalDisk
$diskLines = @()

foreach ($disk in $disks) {
    $diskLine = ""
    $capacitySuffix = ""
    $win = ""
    $mediaType = $disk.MediaType
    $busType = $disk.BusType
    [string]$rawSize = [math]::Round($disk.Size / 1000000000, 2)

    if ($disk.Size.tostring().Length -eq 12) {
        $capacitySuffix = "GB"
    }
    elseif ($disk.Size.tostring().Length -eq 13) {
        $capacitySuffix = "TB"
    }

    if ($capacitySuffix -eq "GB") {
        $size = $rawSize.Substring(0, 3) + $capacitySuffix
    }
    elseif ($capacitySuffix -eq "TB") {
        if ($disk.Size.tostring().Length -eq 14) {
            $size = $rawSize.Substring(0, 2) + $capacitySuffix
        }
        else {
            $size = $rawSize.Substring(0, 1) + $capacitySuffix
        }
    }

    if ($mediaType -eq "SSD") {
        if ($busType -eq "SATA" -or $busType -eq "RAID") {
            $type = "SSD"
        }
        elseif ($busType -eq "NVMe") {
            $type = "NVMe"
        }
    }
    elseif ($mediaType -eq "HDD") {
        $type = "HDD"
    }
    else {
        Continue
    }

    if ($disk.DeviceID -eq $osDiskID) {
        $win = " + W10P"
        if ($language -ne "NL") {
            $win = " + W10P $language"
        }
    }
    $diskLine = $size + " " + $type + $win + "/"
    $diskLines += $diskLine
}

$trimmedDiskLines = -join $diskLines

# ==== OUTPUT TEXT OPERATIONS ====
[string]$outputText = "$Model/$cpuAmount$cpuName/$ramSize $memType/$trimmedDiskLines$gpuNames"
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