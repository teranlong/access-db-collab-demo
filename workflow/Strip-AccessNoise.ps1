<#
.SYNOPSIS
    Strips noise from MS Access VCS exported text files.

.DESCRIPTION
    Removes machine-specific, non-semantic content that MS Access VCS exports
    but which differs between machines/sessions and pollutes Git diffs.

    Targets:
    - PrtMip / PrtDevMode / PrtDevNames (printer blobs)
    - NameMap binary data
    - Checksum lines
    - dbLongBinary sections (SummaryInfo, DocumentMap)
    - DatasheetFont* resets
    - Trailing whitespace

.PARAMETER Path
    Path to a single file or directory to process recursively.

.PARAMETER WhatIf
    Show what would be changed without modifying files.
#>
[CmdletBinding(SupportsShouldProcess)]
param(
    [Parameter(Mandatory)]
    [string]$Path,

    [Parameter()]
    [string]$ConfigPath
)

$ErrorActionPreference = 'Stop'

# Load config if provided
$config = @{
    stripPrinterSettings = $true
    stripNameMap         = $true
    stripChecksums       = $true
    stripGuids           = $false
    stripSummaryInfo     = $true
    stripDatasheetFont   = $true
    customPatterns       = @()
}

if ($ConfigPath -and (Test-Path $ConfigPath)) {
    $jsonConfig = Get-Content $ConfigPath -Raw | ConvertFrom-Json
    if ($jsonConfig.noiseFilter) {
        $nf = $jsonConfig.noiseFilter
        if ($null -ne $nf.stripPrinterSettings) { $config.stripPrinterSettings = $nf.stripPrinterSettings }
        if ($null -ne $nf.stripNameMap)          { $config.stripNameMap = $nf.stripNameMap }
        if ($null -ne $nf.stripChecksums)        { $config.stripChecksums = $nf.stripChecksums }
        if ($null -ne $nf.stripGuids)            { $config.stripGuids = $nf.stripGuids }
        if ($null -ne $nf.stripSummaryInfo)      { $config.stripSummaryInfo = $nf.stripSummaryInfo }
        if ($null -ne $nf.stripDatasheetFont)    { $config.stripDatasheetFont = $nf.stripDatasheetFont }
        if ($nf.customPatterns)                  { $config.customPatterns = $nf.customPatterns }
    }
}

function Strip-NoiseFromContent {
    param([string[]]$Lines)

    $result = [System.Collections.Generic.List[string]]::new()
    $skipBlock = $false
    $blockIndent = 0

    for ($i = 0; $i -lt $Lines.Count; $i++) {
        $line = $Lines[$i]

        # --- Block-level stripping (multi-line sections) ---

        # PrtMip / PrtDevMode / PrtDevNames blocks
        if ($config.stripPrinterSettings -and $line -match '^\s*(PrtMip|PrtDevMode|PrtDevNames)\s*=\s*Begin') {
            $skipBlock = $true
            continue
        }

        # NameMap block
        if ($config.stripNameMap -and $line -match '^\s*NameMap\s*=\s*Begin') {
            $skipBlock = $true
            continue
        }

        # dbLongBinary blocks (SummaryInfo, DocumentMap)
        if ($config.stripSummaryInfo -and $line -match '^\s*dbLongBinary\s+"(SummaryInfo|DocumentMap)"') {
            $skipBlock = $true
            continue
        }

        # End of a block
        if ($skipBlock) {
            if ($line -match '^\s*End\s*$') {
                $skipBlock = $false
            }
            continue
        }

        # --- Line-level stripping ---

        # Checksum lines
        if ($config.stripChecksums -and $line -match '^\s*Checksum\s*=') {
            continue
        }

        # GUID lines (optional - disabled by default as some GUIDs are meaningful)
        if ($config.stripGuids -and $line -match '^\s*GUID\s*=') {
            continue
        }

        # DatasheetFont properties that reset per machine
        if ($config.stripDatasheetFont -and $line -match '^\s*Datasheet(FontHeight|FontWeight|FontName|FontUnderline|FontItalic)\s*=') {
            continue
        }

        # Custom patterns from config
        foreach ($pattern in $config.customPatterns) {
            if ($line -match $pattern) {
                continue
            }
        }

        # Strip trailing whitespace
        $result.Add($line.TrimEnd())
    }

    # Remove trailing empty lines
    while ($result.Count -gt 0 -and $result[$result.Count - 1] -eq '') {
        $result.RemoveAt($result.Count - 1)
    }

    return $result
}

function Process-File {
    param([string]$FilePath)

    $content = Get-Content $FilePath -Encoding UTF8
    $cleaned = Strip-NoiseFromContent -Lines $content

    # Compare to see if anything changed
    $original = ($content -join "`n").TrimEnd()
    $new = ($cleaned -join "`n").TrimEnd()

    if ($original -ne $new) {
        if ($PSCmdlet.ShouldProcess($FilePath, "Strip noise")) {
            $output = ($cleaned -join "`r`n") + "`r`n"
            [System.IO.File]::WriteAllText($FilePath, $output, [System.Text.UTF8Encoding]::new($false))
            Write-Host "  Cleaned: $FilePath" -ForegroundColor Green
            return $true
        }
    }
    return $false
}

# Main execution
$targetPath = Resolve-Path $Path -ErrorAction Stop

if (Test-Path $targetPath -PathType Container) {
    # Process all text files in directory (common VCS export extensions)
    $extensions = @('*.bas', '*.cls', '*.frm', '*.txt', '*.json', '*.xml', '*.sql')
    $files = Get-ChildItem -Path $targetPath -Recurse -Include $extensions -File
    $totalCleaned = 0

    Write-Host "Stripping noise from $($files.Count) files in: $targetPath" -ForegroundColor Cyan

    foreach ($file in $files) {
        if (Process-File -FilePath $file.FullName) {
            $totalCleaned++
        }
    }

    Write-Host "`nDone. Cleaned $totalCleaned of $($files.Count) files." -ForegroundColor Cyan
}
elseif (Test-Path $targetPath -PathType Leaf) {
    Process-File -FilePath $targetPath
}
else {
    Write-Error "Path not found: $targetPath"
}
