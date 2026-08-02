param(
    [Parameter(Mandatory = $true)]
    [string]$CsvPath,

    [string]$OutputPath,

    [switch]$InPlace,

    [switch]$KeepScheduleLabelRows
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

function Test-IsBlank {
    param([object]$Value)

    if ($null -eq $Value) {
        return $true
    }

    return [string]::IsNullOrWhiteSpace([string]$Value)
}

function Convert-ToInvariantNumberString {
    param([string]$Value)

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $clean = $Value.Trim() -replace ',', '.'
    return $clean
}

function Get-PrimaryRectPair {
    param([string]$SizeText)

    $matches = [regex]::Matches($SizeText, '(?i)(\d+(?:[\.,]\d+)?)\s*[x×]\s*(\d+(?:[\.,]\d+)?)')
    if ($matches.Count -eq 0) {
        return $null
    }

    $pairs = @()
    foreach ($m in $matches) {
        $wText = Convert-ToInvariantNumberString -Value $m.Groups[1].Value
        $hText = Convert-ToInvariantNumberString -Value $m.Groups[2].Value

        $wNum = 0.0
        $hNum = 0.0
        [void][double]::TryParse($wText, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$wNum)
        [void][double]::TryParse($hText, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$hNum)

        $pairs += [pscustomobject]@{
            WidthText  = $wText
            HeightText = $hText
            Area       = ($wNum * $hNum)
        }
    }

    # When multiple connector pairs exist, use the largest area pair as the primary body size.
    return $pairs | Sort-Object Area -Descending | Select-Object -First 1
}

function Get-PrimaryDiameter {
    param([string]$SizeText)

    $diamMatches = [regex]::Matches($SizeText, '(?i)(\d+(?:[\.,]\d+)?)\s*(?:ø|Ø|dia\b|diameter\b|dn\b)')
    if ($diamMatches.Count -eq 0) {
        $diamMatches = [regex]::Matches($SizeText, '(?i)(?:ø|Ø|dia\b|diameter\b|dn\s*)(\d+(?:[\.,]\d+)?)')
    }

    if ($diamMatches.Count -gt 0) {
        $max = $null
        foreach ($m in $diamMatches) {
            $raw = if ($m.Groups[1].Success) { $m.Groups[1].Value } else { $m.Groups[0].Value }
            $value = Convert-ToInvariantNumberString -Value $raw
            if ($null -eq $value) {
                continue
            }

            $num = 0.0
            if ([double]::TryParse($value, [System.Globalization.NumberStyles]::Any, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$num)) {
                if ($null -eq $max -or $num -gt $max.Num) {
                    $max = [pscustomobject]@{ Num = $num; Text = $value }
                }
            }
        }

        if ($null -ne $max) {
            return $max.Text
        }
    }

    return $null
}

function Parse-SizeToDims {
    param([string]$SizeText)

    $result = [ordered]@{
        Width    = $null
        Height   = $null
        Diameter = $null
    }

    if ([string]::IsNullOrWhiteSpace($SizeText)) {
        return [pscustomobject]$result
    }

    $s = $SizeText.Trim()

    $rect = Get-PrimaryRectPair -SizeText $s
    if ($null -ne $rect) {
        $result.Width = $rect.WidthText
        $result.Height = $rect.HeightText
    }

    $diam = Get-PrimaryDiameter -SizeText $s
    if ($null -ne $diam) {
        $result.Diameter = $diam
    }

    if ($null -eq $result.Diameter -and $null -eq $result.Width -and $null -eq $result.Height) {
        if ($s -match '^\s*(\d+(?:[\.,]\d+)?)\s*(?:mm)?\s*$') {
            $result.Diameter = Convert-ToInvariantNumberString -Value $Matches[1]
        }
    }

    return [pscustomobject]$result
}

function Test-IsScheduleLabelRow {
    param([psobject]$Row)

    $checks = @(
        @{ Name = 'Family'; Expected = 'Family' },
        @{ Name = 'Family and Type'; Expected = 'Family and Type' },
        @{ Name = 'NUMBER'; Expected = 'NUMBER' },
        @{ Name = 'NWB_AssetType'; Expected = 'NWB_AssetType' },
        @{ Name = 'NWB_AssetTypeCode'; Expected = 'NWB_AssetTypeCode' },
        @{ Name = 'NWB_Category'; Expected = 'NWB_Category' },
        @{ Name = 'NWB_DesignPkg'; Expected = 'NWB_DesignPkg' },
        @{ Name = 'NWB_DimDiameter'; Expected = 'NWB_DimDiameter' },
        @{ Name = 'NWB_DimHeight'; Expected = 'NWB_DimHeight' },
        @{ Name = 'NWB_DimWidth'; Expected = 'NWB_DimWidth' },
        @{ Name = 'Size'; Expected = 'Size' },
        @{ Name = 'DIAMETER'; Expected = 'DIAMETER' },
        @{ Name = 'LENGTH'; Expected = 'LENGTH' }
    )

    $matches = 0
    foreach ($c in $checks) {
        $name = $c.Name
        if ($null -eq $Row.PSObject.Properties[$name]) {
            continue
        }

        $value = [string]$Row.$name
        if ($value -eq $c.Expected) {
            $matches++
        }
    }

    return ($matches -ge 6)
}

if (-not (Test-Path -LiteralPath $CsvPath)) {
    throw "CSV file not found: $CsvPath"
}

$rows = Import-Csv -LiteralPath $CsvPath
if ($rows.Count -eq 0) {
    throw "CSV has no data rows: $CsvPath"
}

$stats = [ordered]@{
    TotalRows                  = 0
    DroppedScheduleLabelRows   = 0
    RowsWithSize               = 0
    SizeParsedRect             = 0
    SizeParsedRound            = 0
    UpdatedNWB_DimWidth        = 0
    UpdatedNWB_DimHeight       = 0
    UpdatedNWB_DimDiameter     = 0
    RowsWithAnyUpdate          = 0
}

$outputRows = New-Object System.Collections.Generic.List[object]
$updatedSamples = New-Object System.Collections.Generic.List[object]

foreach ($row in $rows) {
    $stats.TotalRows++

    if (Test-IsScheduleLabelRow -Row $row) {
        if (-not $KeepScheduleLabelRows) {
            $stats.DroppedScheduleLabelRows++
            continue
        }
        else {
            # Keep row but never parse/fill against schedule label values.
            $outputRows.Add($row)
            continue
        }
    }

    $sizeText = [string]$row.Size

    if ([string]::IsNullOrWhiteSpace($sizeText)) {
        $outputRows.Add($row)
        continue
    }

    $stats.RowsWithSize++
    $dims = Parse-SizeToDims -SizeText $sizeText

    if (-not (Test-IsBlank -Value $dims.Width) -or -not (Test-IsBlank -Value $dims.Height)) {
        $stats.SizeParsedRect++
    }
    if (-not (Test-IsBlank -Value $dims.Diameter)) {
        $stats.SizeParsedRound++
    }

    $rowChanged = $false

    if ((Test-IsBlank -Value $row.NWB_DimWidth) -and -not (Test-IsBlank -Value $dims.Width)) {
        $row.NWB_DimWidth = $dims.Width
        $stats.UpdatedNWB_DimWidth++
        $rowChanged = $true
    }

    if ((Test-IsBlank -Value $row.NWB_DimHeight) -and -not (Test-IsBlank -Value $dims.Height)) {
        $row.NWB_DimHeight = $dims.Height
        $stats.UpdatedNWB_DimHeight++
        $rowChanged = $true
    }

    if ((Test-IsBlank -Value $row.NWB_DimDiameter) -and -not (Test-IsBlank -Value $dims.Diameter)) {
        $row.NWB_DimDiameter = $dims.Diameter
        $stats.UpdatedNWB_DimDiameter++
        $rowChanged = $true
    }

    if ($rowChanged) {
        $stats.RowsWithAnyUpdate++
        if ($updatedSamples.Count -lt 8) {
            $updatedSamples.Add([pscustomobject]@{
                ElementId = [string]$row.__MHT_ElementId
                Size = [string]$row.Size
                NWB_DimWidth = [string]$row.NWB_DimWidth
                NWB_DimHeight = [string]$row.NWB_DimHeight
                NWB_DimDiameter = [string]$row.NWB_DimDiameter
            })
        }
    }

    $outputRows.Add($row)
}

if (-not $OutputPath) {
    $base = [System.IO.Path]::GetFileNameWithoutExtension($CsvPath)
    $dir = [System.IO.Path]::GetDirectoryName($CsvPath)
    $ext = [System.IO.Path]::GetExtension($CsvPath)

    if ($InPlace) {
        $OutputPath = $CsvPath
    }
    else {
        $OutputPath = Join-Path $dir ($base + '_nwb_dims_filled' + $ext)
    }
}

if ($InPlace) {
    $backupPath = "$CsvPath.bak_$(Get-Date -Format 'yyyyMMdd_HHmmss')"
    Copy-Item -LiteralPath $CsvPath -Destination $backupPath -Force
    Write-Host "Backup created: $backupPath"
}

$outputRows | Export-Csv -LiteralPath $OutputPath -NoTypeInformation -Encoding UTF8

Write-Host "Done. Output: $OutputPath"
Write-Host "Stats:"
$stats.GetEnumerator() | ForEach-Object { Write-Host ("- {0}: {1}" -f $_.Key, $_.Value) }

if ($updatedSamples.Count -gt 0) {
    Write-Host "Updated row samples:"
    $updatedSamples | Format-Table -AutoSize | Out-String | Write-Host
}