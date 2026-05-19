param(
    [Parameter(Mandatory)]
    [string]$FolderPath
)

# Create timestamped filename
$timestamp = Get-Date -Format "yyyy-MM-dd_HH-mm-ss"
$OutXlsx = Join-Path $FolderPath "Combined_Phases_$timestamp.xlsx"

# Ensure ImportExcel is available
if (-not (Get-Module -ListAvailable -Name ImportExcel)) {
    Install-Module ImportExcel -Scope CurrentUser -Force
}
Import-Module ImportExcel

# ---------------------------
# Load X-Dana mapping (CSV)
# ---------------------------
$CacheFile = Join-Path $FolderPath "xdana_cache.csv"
$XDanaMap = @{}

Write-Host "Loading Dana mapping..."

Import-Csv $CacheFile | ForEach-Object {

    if (-not $_.Phase) { continue }

    if (-not $_.X -or $_.X.Trim().ToLower() -ne "x") { continue }

    $nums = @(
        [int]$_.N1,
        [int]$_.N2,
        [int]$_.N3,
        [int]$_.N4,
        [int]$_.N5,
        [int]$_.N6
    )

    $sortKey = ($nums | ForEach-Object { "{0:D2}" -f $_ }) -join "."

    $XDanaMap[$_.Phase.Trim().ToLower()] = $sortKey
}

Write-Host "Loaded $($XDanaMap.Count) Dana entries"

$allRows = @()

# Process each txt file
Get-ChildItem -Path $FolderPath -Filter *.txt | ForEach-Object {

    $filePath = $_.FullName
    Write-Host "Processing $($_.Name)..."

    $lines = Get-Content $filePath

    if ($lines.Count -lt 10) { return }

    # ---------------------------
    # 1. Find ALL Format lines
    # ---------------------------
    $formatIndices = @()

    for ($i = 0; $i -lt $lines.Count; $i++) {
        if ($lines[$i] -match '^\s*Format:\s*$') {
            $formatIndices += $i
        }
    }

    if ($formatIndices.Count -eq 0) { return }

    $lastFormat = $formatIndices[-1]

    # ---------------------------
    # 2. Find previous Displacement BEFORE last Format
    # ---------------------------
    $startIndex = 0

    for ($i = $lastFormat - 1; $i -ge 0; $i--) {
        if ($lines[$i] -match '^\s*Displacement:') {
            $startIndex = $i + 1
            break
        }
    }

    # Special case: only one block
    if ($startIndex -eq 0) {
        for ($i = 0; $i -lt $lastFormat; $i++) {
            if ($lines[$i] -match '^\s*Format:') {
                $startIndex = $i + 1
            }
        }
    }

    if ($lastFormat -le $startIndex) { return }

    $phaseLines = $lines[$startIndex..($lastFormat - 1)]

    # ---------------------------
    # 3. Extract XRD file name
    # ---------------------------
    $xrdFile = $null

    for ($i = $lastFormat; $i -lt [Math]::Min($lastFormat + 25, $lines.Count); $i++) {
        if ($lines[$i] -match 'XRD file:\s*(.+)') {
            $xrdFile = $matches[1].Trim()
            break
        }
    }

    if (-not $xrdFile) {
        $xrdFile = $_.BaseName
    }

    # Remove extension (.raw)
    $xrdFile = [System.IO.Path]::GetFileNameWithoutExtension($xrdFile)

    # ---------------------------
    # 4. Parse phase rows
    # ---------------------------
    foreach ($line in $phaseLines) {

        $t = $line.Trim()

        if (-not $t) { continue }

        # Skip headers / metadata
        if ($t -match '^(Format|Phase\s+name|XRD file|Analyst|Rwp|GOF|Displacement)') {
            continue
        }

        # Split correctly (tabs OR multiple spaces)
        $parts = $t -split "\t+|\s{2,}" | Where-Object { $_ -ne "" }

        if ($parts.Count -ge 2) {

            if ($parts[1] -match '^[+-]?\d+(\.\d+)?$') {

                $phase = $parts[0].Trim()
                $value = $parts[1].Trim()

                $phaseBase = ($phase -replace '\s*<[^>]+>', '').Trim()
                $lookup = $phaseBase.ToLower()

                $xrd = $xrdFile

                $sortKey = if ($XDanaMap.ContainsKey($lookup)) {
                    $XDanaMap[$lookup]
                } else {
                    "99.99.99.99.99.99"
                }

                $allRows += [pscustomobject]@{
                    "Data File"         = $xrd
                    #PhaseName = $phase
                    "Phase Modelled"    = $phaseBase
                    "Phase Filter"      = $phaseBase
                    "wt%"               = [double]$value
                    "X-Dana#"           = $sortKey
                }
            }
        }

    }
}

# ---------------------------
# 5. Output: write data sheet + create pivot sheet
# ---------------------------
if ($allRows.Count -eq 0) {
    throw "No phase data found in any .txt files under: $FolderPath"
}

# Write the raw data to sheet "Phases" as an Excel Table, keep workbook open (-PassThru)
$excel = $allRows | Export-Excel `
    -Path $OutXlsx `
    -WorksheetName "Phases" `
    -AutoSize `
    -TableName "PhaseData" `
    -TableStyle Medium6 `
    -BoldTopRow `
    -FreezeTopRow `
    -PassThru

# Build a pivot from the table we just created
$sourceWs    = $excel.Workbook.Worksheets["Phases"]
$sourceRange = $sourceWs.Tables[0].Address   # uses the Excel Table as the pivot source

Add-PivotTable `
    -ExcelPackage $excel `
    -PivotTableName "PhasePivot" `
    -SourceRange $sourceRange `
    -PivotRows "X-Dana#","Phase Modelled" `
    -PivotColumns "Data File" `
    -PivotFilter "Phase Filter" `
    -PivotData @{ "wt%" = "Sum" } `
    -PivotDataToColumn `
    -PivotNumberFormat "0.00" `
    -PivotTotals None `
    -PivotTableStyle Medium6 `
    -Activate

# Do things to pivot table. 
$pivot = $excel.Workbook.Worksheets["PhasePivot"].PivotTables["PhasePivot"]
 
    #rename Sum of wt%
    $pivot.DataFields[0].Name = "wt%"

# Close and save
Close-ExcelPackage $excel

Write-Host "`n✅ SUCCESS: Data + Pivot written to $OutXlsx"