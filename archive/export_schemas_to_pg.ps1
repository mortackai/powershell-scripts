param(
    [string]$Server = "GSSDB1",
    [string]$Database = "GLOBALCMI",
    [string]$OutputDir = "C:\Users\MRaymond\Documents\exports\Schemas",
    [string]$SelectedDriver = "Pervasive ODBC Unicode Interface"
)

# Prompt for credentials
$Creds = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User = $Creds.UserName
$Pass = $Creds.GetNetworkCredential().Password

$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)

function Convert-DataRowTypeToPg($row) {
    $dataType = $row["DataType"].ToString()
    $colSize = $row["ColumnSize"]
    $precision = $row["NumericPrecision"]
    $scale = $row["NumericScale"]

    switch ($dataType) {
        "System.String" {
            if ($colSize -and $colSize -gt 0 -and $colSize -lt 1000) { return "varchar($colSize)" } else { return "text" }
        }
        "System.Int16" { return "smallint" }
        "System.Int32" { return "integer" }
        "System.Int64" { return "bigint" }
        "System.Decimal" { if ($precision -and $scale -ne $null) { return "numeric($precision,$scale)" } else { return "numeric" } }
        "System.Double" { return "double precision" }
        "System.Single" { return "real" }
        "System.Boolean" { return "boolean" }
        "System.DateTime" { return "timestamp without time zone" }
        "System.Byte[]" { return "bytea" }
        default { return "text" }
    }
}

try {
    Write-Host "Connecting to $Server / $Database..." -ForegroundColor Gray
    $Connection.Open()

    # Get tables using the ADO.NET schema API
    $tablesDt = $Connection.GetSchema("Tables")

    if (-not (Test-Path $OutputDir)) { New-Item -ItemType Directory -Path $OutputDir -Force | Out-Null }
    $outFile = Join-Path $OutputDir "create_tables.sql"
    if (Test-Path $outFile) { Remove-Item $outFile -Force }

    $total = 0
    foreach ($trow in $tablesDt.Rows) {
        $tableType = $null
        if ($trow.Table.Columns.Contains('TABLE_TYPE')) { $tableType = $trow['TABLE_TYPE'] }
        elseif ($trow.Table.Columns.Contains('TABLETYPE')) { $tableType = $trow['TABLETYPE'] }
        if ($tableType -and $tableType -ne 'TABLE' -and $tableType -ne 'BASE TABLE') { continue }

        # Table name may be in different columns depending on provider
        $tableName = $null
        foreach ($col in 'TABLE_NAME','TABLE') { if ($trow.Table.Columns.Contains($col) -and $trow[$col]) { $tableName = $trow[$col]; break } }
        if (-not $tableName) { continue }

        $schemaName = $null
        if ($trow.Table.Columns.Contains('TABLE_SCHEMA')) { $schemaName = $trow['TABLE_SCHEMA'] }
        elseif ($trow.Table.Columns.Contains('OWNER')) { $schemaName = $trow['OWNER'] }

        # Get column metadata via a reader
        $safeTable = $tableName -replace '"','""'
        if ($schemaName) {
            $safeSchema = $schemaName -replace '"','""'
            $fullName = '"' + $safeSchema + '"."' + $safeTable + '"'
        } else {
            $fullName = '"' + $safeTable + '"'
        }
        $sql = 'SELECT TOP 1 * FROM ' + $fullName
        $cmd = New-Object System.Data.Odbc.OdbcCommand($sql, $Connection)
        $reader = $cmd.ExecuteReader()
        $schema = $reader.GetSchemaTable()

        $cols = @()
        $pkCols = @()
        foreach ($colRow in $schema.Rows) {
            $colName = $colRow["ColumnName"]
            $pgType = Convert-DataRowTypeToPg $colRow
            $nullable = ($colRow["AllowDBNull"] -eq $true)
            $isKey = $false
            if ($colRow.Table.Columns.Contains("IsKey")) { $isKey = $colRow["IsKey"] }

            $colDef = '"' + $colName + '" ' + $pgType
            if (-not $nullable) { $colDef += ' NOT NULL' }
            $cols += $colDef
            if ($isKey) { $pkCols += '"' + $colName + '"' }
        }
        $reader.Close()

        $create = 'CREATE TABLE IF NOT EXISTS "' + $tableName + '" (' + ($cols -join ", `n    ")
        if ($pkCols.Count -gt 0) {
            $create += ',`n    CONSTRAINT "PK_' + $tableName + '" PRIMARY KEY (' + ($pkCols -join ", ") + ')'
        }
        $create += ') ;`n`n'

        $create | Out-File -FilePath $outFile -Encoding UTF8 -Append
        Write-Host "Generated CREATE for: $tableName" -ForegroundColor Cyan
        $total++
    }

    Write-Host "Done. Generated $total CREATE statements in $outFile" -ForegroundColor Yellow
}
catch {
    Write-Error "Error: $($_.Exception.Message)"
}
finally {
    if ($Connection.State -eq 'Open') { $Connection.Close() }
}
