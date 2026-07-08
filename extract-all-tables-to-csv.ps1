# --- Configuration ---
$Server     = "GSSDB1"
$Database   = "GLOBALCMI"
$ExportDir  = "\\Cryomech\Documents$\MRaymond\Documents\psql-exports\exports"
$RowsPerFile = 0  # Split files after this many rows, 0 to disable splitting

if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir }

# --- Driver & Credentials ---
$SelectedDriver = "Pervasive ODBC Unicode Interface"
$Creds      = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User       = $Creds.UserName
$Pass       = $Creds.GetNetworkCredential().Password
$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)

# --- Helpers ---
# Force invariant culture so decimals always use . not ,
$InvariantCulture = [System.Globalization.CultureInfo]::InvariantCulture

function Format-CsvField {
    param($value)
    if ($null -eq $value) { return '\N' }
    $str = $value.ToString($InvariantCulture).Trim()
    # Replace embedded newlines and carriage returns with a space
    $str = $str -replace "`r`n", " " -replace "`r", " " -replace "`n", " "
    # CSV-escape: wrap in quotes if contains comma, quote, or was trimmed to empty
    $str = $str.Replace('"', '""')
    return """$str"""
}

function Format-DateField {
    param($value)
    if ($null -eq $value -or $value -is [DBNull]) { return '\N' }
    try {
        $dt = [datetime]$value
        if ($dt -lt [datetime]"0001-01-01") { return '\N' }
        return """$($dt.ToString('yyyy-MM-dd HH:mm:ss'))"""
    } catch {
        return '\N'
    }
}

function Get-SafeFileName {
    param($name)
    if ([string]::IsNullOrWhiteSpace([string]$name)) { return 'unnamed' }
    #$invalidChars = [System.IO.Path]::GetInvalidFileNameChars() + [char[]]'\/\\:*?"<>|'
    #$pattern = '[' + ([regex]::Escape(($invalidChars -join ''))) + ']'
    return [regex]::Replace($name, '_')
}

function Get-TableColumns {
    param(
        [System.Data.Odbc.OdbcConnection]$Connection,
        [string]$TableName
    )
    
    $cmd = New-Object System.Data.Odbc.OdbcCommand("SELECT * FROM `"$TableName`"", $Connection)
    $reader = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
    $schema = $reader.GetSchemaTable()
    $reader.Close()
    
    $columns = [System.Collections.Generic.List[hashtable]]::new()
    foreach ($row in $schema.Rows) {
        $dataType = $row["DataType"] -as [type]
        $dataTypeName = if ($null -ne $dataType) {
            $dataType.Name
        } else {
            ''
        }
        
        $columns.Add(@{
            Name       = [string]$row["ColumnName"]
            DotNet     = $dataTypeName
            IsDateTime = ($null -ne $dataType -and ($dataType -eq [datetime] -or $dataTypeName -match '(?i)datetime|timestamp|time'))
        })
    }
    
    return $columns
}

try {
    Write-Host "Connecting to $Database..." -ForegroundColor Gray
    $Connection.Open()

    # --- PHASE 1: DISCOVERY ---
    Write-Host "Discovering tables and checking permissions..." -ForegroundColor Cyan
    $AllTables = $Connection.GetSchema("Tables") |
        Where-Object { $_.TABLE_TYPE -eq "TABLE" -and $_.TABLE_NAME -notlike "f$*" }

    $AllowedTables = [System.Collections.Generic.List[string]]::new()
    foreach ($row in $AllTables) {
        $tName = $row.TABLE_NAME
        try {
            $checkCmd = New-Object System.Data.Odbc.OdbcCommand("SELECT TOP 0 * FROM ""$tName""", $Connection)
            $null = $checkCmd.ExecuteNonQuery()
            $AllowedTables.Add($tName)
        }
        catch {
            Write-Host "  [!] Skipping $tName (Access Denied)" -ForegroundColor DarkGray
        }
    }
    Write-Host "Found $($AllowedTables.Count) accessible tables.`n" -ForegroundColor Green

    # --- PHASE 2: EXPORT ---
    $TotalTables  = $AllowedTables.Count
    $TablesDone   = 0
    $FailedTables = [System.Collections.Generic.List[string]]::new()

    # foreach loop that goes through each table in allowedtables and exports it to CSV
    foreach ($TableName in $AllowedTables) {
        # increments the tables done counter and writes to the console which table is being exported
        $TablesDone++
        Write-Host "[$TablesDone/$TotalTables] Exporting: $TableName" -ForegroundColor Yellow

        # try-catch block to handle any errors that may occur during the export process
        try {
            $writer = $null
            $reader = $null

            $columns = Get-TableColumns -Connection $Connection -TableName $TableName
            if ($columns.Count -eq 0) {
                Write-Warning "  Skipping $TableName (no columns found)"
                continue
            }

            # Build SELECT with date guards for datetime-like columns
            $SqlColumns = $columns | ForEach-Object {
                if ($_.IsDateTime) {
                    "IF(`"$($_.Name)`" < '0001-01-01', NULL, `"$($_.Name)`") AS `"$($_.Name)`""
                } else {
                    "`"$($_.Name)`""
                }
            }
            $query = "SELECT " + ($SqlColumns -join ", ") + " FROM ""$TableName"""

            $cmd = New-Object System.Data.Odbc.OdbcCommand($query, $Connection)
            $cmd.CommandTimeout = 300
            $reader = $cmd.ExecuteReader()

            # Write header
            $header = ($columns | ForEach-Object { """$($_.Name)""" }) -join ","

            $fileIndex  = 1
            $rowCount   = 0
            $fileRows   = 0

            $filePath  = Join-Path $ExportDir "$($TableName).csv"
            $writer    = [System.IO.StreamWriter]::new($filePath, $false, [System.Text.Encoding]::UTF8)
            $writer.WriteLine($header)

            while ($reader.Read()) {
                $rowCount++
                $fileRows++

                # Build CSV line
                $fields = for ($i = 0; $i -lt $reader.FieldCount; $i++) {
                    if ($reader.IsDBNull($i)) {
                        '\N'
                    } elseif ($columns[$i]["DotNet"] -eq "DateTime") {
                        Format-DateField $reader.GetValue($i)
                    } else {
                        Format-CsvField $reader.GetValue($i)
                    }
                }
                $writer.WriteLine($fields -join ",")

                # Split file if threshold hit
                if ($RowsPerFile -gt 0 -and $fileRows -ge $RowsPerFile) {
                    $writer.Close()
                    $fileIndex++
                    $fileRows  = 0
                    $filePath  = Join-Path $ExportDir "$($TableName)_part$($fileIndex).csv"
                    $writer    = [System.IO.StreamWriter]::new($filePath, $false, [System.Text.Encoding]::UTF8)
                    $writer.WriteLine($header)
                    Write-Host "  -> Split: starting part $fileIndex" -ForegroundColor DarkCyan
                }

                if ($rowCount % 10000 -eq 0) {
                    Write-Host "  $rowCount rows written..." -ForegroundColor DarkGray
                }
            }

            $writer.Close()
            $reader.Close()
            Write-Host "  Done: $rowCount rows" -ForegroundColor Green

        } catch {
            if ($null -ne $writer) { try { $writer.Close() } catch {} }
            if ($null -ne $reader) { try { $reader.Close() } catch {} }
            Write-Error "  Failed: $TableName -- $($_.Exception.Message)"
            $FailedTables.Add($TableName)
        }
    }

    # --- SUMMARY ---
    Write-Host "`nExport complete." -ForegroundColor Green
    Write-Host "Tables exported : $($TotalTables - $FailedTables.Count)/$TotalTables"
    if ($FailedTables.Count -gt 0) {
        Write-Host "Failed tables:" -ForegroundColor Red
        $FailedTables | ForEach-Object { Write-Host "  $_" -ForegroundColor Red }
    }

} finally {
    $Connection.Close()
    Write-Host "Connection closed." -ForegroundColor Gray
}