# --- Configuration ---
$Server = "GSSDB1" 
$Database = "GLOBALCMI"
$ExportDir = "\\Cryomech\Documents$\MRaymond\Documents\psql-exports\exports"
if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir }

# --- Type mapping: Actian/Pervasive -> PostgreSQL ---
function Get-PgType {
    param($OdbcTypeName, $Size, $Precision, $Scale)

    if ([string]::IsNullOrWhiteSpace([string]$OdbcTypeName)) {
        return "TEXT  -- TODO: review (source: <null>)"
    }

    $typeName = [string]$OdbcTypeName
    if ($typeName.StartsWith("System.", [System.StringComparison]::InvariantCultureIgnoreCase)) {
        $typeName = $typeName.Substring(7)
    }

    switch -Regex ($typeName.ToLower()) {
        "datetime"     { return "TIMESTAMP" }
        "bigint"       { return "BIGINT" }
        "int64"        { return "BIGINT" }
        "int32"        { return "INTEGER" }
        "int16"        { return "SMALLINT" }
        "tinyint"      { return "SMALLINT" }
        "\bint\b"    { return "INTEGER" }
        "decimal"      { return "NUMERIC($Precision, $Scale)" }
        "numeric"      { return "NUMERIC($Precision, $Scale)" }
        "money"        { return "NUMERIC(19, 4)" }
        "float"        { return "DOUBLE PRECISION" }
        "real"         { return "REAL" }
        "double"       { return "DOUBLE PRECISION" }
        "char"         {
            if ($Size -gt 0 -and $Size -le 10485760) { return "VARCHAR($Size)" }
            else { return "TEXT" }
        }
        "varchar"      {
            if ($Size -gt 0 -and $Size -le 10485760) { return "VARCHAR($Size)" }
            else { return "TEXT" }
        }
        "string"       {
            if ($Size -gt 0 -and $Size -le 10485760) { return "VARCHAR($Size)" }
            else { return "TEXT" }
        }
        "text"         { return "TEXT" }
        "longvar"      { return "TEXT" }
        "date"         { return "DATE" }
        "timestamp"    { return "TIMESTAMP" }
        "time"         { return "TIME" }
        "bit"          { return "BOOLEAN" }
        "bool"         { return "BOOLEAN" }
        "logical"      { return "BOOLEAN" }
        "guid"         { return "UUID" }
        "uniqueidentifier" { return "UUID" }
        "binary"       { return "BYTEA" }
        "blob"         { return "BYTEA" }
        "image"        { return "BYTEA" }
        "byte\[\]"   { return "BYTEA" }
        "long"         { return "BIGINT" }
        default         { return "TEXT  -- TODO: review (source: $typeName)" }
        # TEXT is a safe fallback but flag it for review
    }
}

# --- Driver & Credentials ---
$SelectedDriver = "Pervasive ODBC Unicode Interface"
$Creds = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User = $Creds.UserName
$Pass = $Creds.GetNetworkCredential().Password
$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)

$SummaryLines    = [System.Collections.Generic.List[string]]::new()
$CreateTableLines = [System.Collections.Generic.List[string]]::new()
$ReviewFlags     = [System.Collections.Generic.List[string]]::new()

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

    # --- PHASE 2: SCHEMA EXTRACTION ---
    foreach ($TableName in $AllowedTables) {
        Write-Host "Inspecting: $TableName" -ForegroundColor Yellow

        try {
            $cmd    = New-Object System.Data.Odbc.OdbcCommand("SELECT * FROM ""$TableName""", $Connection)
            $reader = $cmd.ExecuteReader([System.Data.CommandBehavior]::SchemaOnly)
            $dt     = $reader.GetSchemaTable()
            $reader.Close()

            # --- Human-readable summary ---
            $SummaryLines.Add("=== $TableName ===")
            $colDefs = [System.Collections.Generic.List[string]]::new()

            foreach ($col in $dt.Rows) {
                $cName = $col["ColumnName"]
                $cTypeRaw = if ($col.Table.Columns.Contains("DataTypeName") -and -not [string]::IsNullOrWhiteSpace([string]$col["DataTypeName"])) {
                    $col["DataTypeName"]
                } else {
                    $col["DataType"]
                }
                $cType = [string]$cTypeRaw
                $cSize     = $col["ColumnSize"]
                $cPrec     = if ($col["NumericPrecision"] -is [DBNull]) { 0 } else { [int]$col["NumericPrecision"] }
                $cScale    = if ($col["NumericScale"]     -is [DBNull]) { 0 } else { [int]$col["NumericScale"] }
                $nullable  = if ($col["AllowDBNull"] -eq $true) { "NULL" } else { "NOT NULL" }

                $pgType    = Get-PgType -OdbcTypeName $cType -Size $cSize -Precision $cPrec -Scale $cScale

                $SummaryLines.Add("  $cName | source: $cType($cSize) | pg: $pgType | $nullable")
                $colDefs.Add("    `"$cName`" $pgType $nullable")

                # Flag anything that fell through to the TEXT default for manual review
                if ($pgType -like "*TODO*") {
                    $ReviewFlags.Add("  [$TableName].$cName  ->  $cType")
                }
            }

            $SummaryLines.Add("")

            # --- PostgreSQL CREATE TABLE block ---
            $CreateTableLines.Add("-- $TableName")
            $CreateTableLines.Add("CREATE TABLE IF NOT EXISTS $TableName (")
            $CreateTableLines.Add(($colDefs -join ",`n"))
            $CreateTableLines.Add(");")
            $CreateTableLines.Add("")
        }
        catch {
            Write-Warning "  Could not inspect $TableName : $($_.Exception.Message)"
            $SummaryLines.Add("=== $TableName === [ERROR: $($_.Exception.Message)]`n")
        }
    }

    # --- Write outputs ---
    $SummaryPath = Join-Path $ExportDir "schema_summary.txt"
    $SqlPath     = Join-Path $ExportDir "create_tables.sql"
    $ReviewPath  = Join-Path $ExportDir "types_to_review.txt"

    $SummaryLines    | Out-File $SummaryPath -Encoding UTF8
    $CreateTableLines | Out-File $SqlPath    -Encoding UTF8

    if ($ReviewFlags.Count -gt 0) {
        @("Columns whose Actian type had no direct mapping and defaulted to TEXT:") + $ReviewFlags |
            Out-File $ReviewPath -Encoding UTF8
        Write-Host "`n[!] $($ReviewFlags.Count) column(s) need manual type review -> $ReviewPath" -ForegroundColor Red
    }

    Write-Host "`nSchema export complete:" -ForegroundColor Green
    Write-Host "  Summary  : $SummaryPath"
    Write-Host "  SQL DDL  : $SqlPath"
    if ($ReviewFlags.Count -gt 0) { Write-Host "  Review   : $ReviewPath" }
}
finally {
    $Connection.Close()
    Write-Host "Done." -ForegroundColor Yellow
}