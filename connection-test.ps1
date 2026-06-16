# --- Configuration ---
$Server = "GSSDB1" 
$Database = "GLOBALCMI"
$TargetTable = "ITEM_MASTER"
$ExportFile = "\\tsclient\H\psql-exports\$($TargetTable).csv"

# --- Driver & Credentials ---
$SelectedDriver = "Pervasive ODBC Unicode Interface"
$Creds = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User = $Creds.UserName
$Pass = $Creds.GetNetworkCredential().Password

$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)

try {
    Write-Host "Connecting via Unicode Interface..." -ForegroundColor Gray
    $Connection.Open()

    # 1. Get Column names and types
    $setupCommand = New-Object System.Data.Odbc.OdbcCommand("SELECT TOP 1 * FROM ""$TargetTable""", $Connection)
    $setupReader = $setupCommand.ExecuteReader()
    $dt = $setupReader.GetSchemaTable()
    
    # 2. Build the SQL Query
    # We use Actian SQL to handle the '0000-00-00' date issue internally
    $Columns = @()
    foreach ($row in $dt.Rows) {
        $colName = $row["ColumnName"]
        $dataType = $row["DataType"].ToString()

        if ($dataType -eq "System.DateTime") {
            # If the date is the Pervasive 'zero date', return NULL, otherwise format it
            $Columns += "IF(""$colName"" < '0001-01-01', NULL, ""$colName"") AS ""$colName"""
        } else {
            $Columns += """$colName"""
        }
    }
    $setupReader.Close()

    $Sql = "SELECT " + ($Columns -join ", ") + " FROM ""$TargetTable"""
    
    $Command = New-Object System.Data.Odbc.OdbcCommand($Sql, $Connection)
    $Command.CommandTimeout = 0
    $Reader = $Command.ExecuteReader()

    # 3. Write CSV with UTF8BOM (Best for maintaining character integrity)
    $HeaderRow = ($dt.Rows.ColumnName | ForEach-Object { """$_""" }) -join ","
    $HeaderRow | Out-File -FilePath $ExportFile -Encoding UTF8

    Write-Host "Exporting with Unicode support and SQL-side date handling..." -ForegroundColor Cyan
    $RowCount = 0

    while ($Reader.Read()) {
        $RowArray = New-Object System.Collections.Generic.List[string]
        for ($i = 0; $i -lt $Reader.FieldCount; $i++) {
            if ($Reader.IsDBNull($i)) {
                $RowArray.Add("")
            } else {
                # GetValue on a Unicode driver should now return the full string
                $val = $Reader.GetValue($i).ToString().Trim()
                $RowArray.Add("""$($val.Replace('"', '""'))""")
            }
        }
        $RowArray -join "," | Out-File -FilePath $ExportFile -Encoding UTF8 -Append
        
        $RowCount++
        if ($RowCount % 500 -eq 0) { Write-Host "`rRows processed: $RowCount" -NoNewline }
    }
    Write-Host "`nExport Complete! Total rows: $RowCount" -ForegroundColor Yellow
}
catch {
    Write-Error "Export Error: $($_.Exception.Message)"
}
finally {
    if ($Connection.State -eq "Open") { $Connection.Close() }
}