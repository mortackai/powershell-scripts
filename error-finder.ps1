# --- Configuration ---
$Server = "GSSDB1" 
$Database = "GLOBALCMI"
$TargetTable = "ITEM_MASTER" # <--- Change this to the table you want to test
$ExportFile = "C:\Users\MRaymond\Downloads\$($TargetTable).csv"

# --- Driver & Credentials ---
$SelectedDriver = "Pervasive ODBC Unicode Interface"
$Creds = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User = $Creds.UserName
$Pass = $Creds.GetNetworkCredential().Password

$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)
$Timer = [System.Diagnostics.Stopwatch]::StartNew()

try {
    Write-Host "Connecting to $Database..." -ForegroundColor Gray
    $Connection.Open()

    # 1. Get Structure
    $setupCommand = New-Object System.Data.Odbc.OdbcCommand("SELECT TOP 1 * FROM ""$TargetTable""", $Connection)
    $setupReader = $setupCommand.ExecuteReader()
    $dt = $setupReader.GetSchemaTable()
    $setupReader.Close()

    # 2. Build SQL with RTRIM CAST to minimize errors
    $SqlColumns = @()
    foreach ($col in $dt.Rows) {
        $cName = $col["ColumnName"]
        $SqlColumns += "RTRIM(CAST(""$cName"" AS VARCHAR)) AS ""$cName"""
    }
    $query = "SELECT " + ($SqlColumns -join ", ") + " FROM ""$TargetTable"""
    
    $cmd = New-Object System.Data.Odbc.OdbcCommand($query, $Connection)
    $cmd.CommandTimeout = 0
    $reader = $cmd.ExecuteReader()

    # 3. Prepare File
    $Header = ($dt.Rows.ColumnName | ForEach-Object { """$_""" }) -join ","
    $Header | Out-File $ExportFile -Encoding UTF8

    Write-Host "Starting trace on $TargetTable..." -ForegroundColor Cyan
    $RowCount = 0
    $ErrorCount = 0

    while ($reader.Read()) {
        $RowCount++
        $line = for ($i=0; $i -lt $reader.FieldCount; $i++) {
            $colName = $reader.GetName($i)
            try {
                $val = $reader.GetValue($i)
                if ($val -eq [DBNull]::Value -or $val -eq $null) { "" } 
                else { 
                    $strVal = $val.ToString().Trim()
                    if ($strVal -eq "0000-00-00" -or $strVal -eq "0001-01-01") { "" }
                    else { """$($strVal.Replace('"', '""'))""" }
                }
            }
            catch {
                $ErrorCount++
                Write-Host "`n[!] ERROR at Row: $RowCount | Column: $colName" -ForegroundColor Red
                Write-Host "    Message: $($_.Exception.Message)" -ForegroundColor White
                "NULL" # Write NULL to CSV to keep columns aligned
            }
        }
        $line -join "," | Out-File $ExportFile -Encoding UTF8 -Append
        
        if ($RowCount % 100 -eq 0) { 
            Write-Host "`rRows processed: $RowCount | Errors found: $ErrorCount" 
        }
    }

    $Timer.Stop()
    Write-Host "`n`nExport Finished!" -ForegroundColor Green
    Write-Host "Total Rows: $RowCount"
    Write-Host "Total Errors: $ErrorCount"
    Write-Host "Time Elapsed: $($Timer.Elapsed.ToString('mm\:ss'))"
    Write-Host "File saved to: $ExportFile" -ForegroundColor Yellow
}
catch {
    Write-Error "Critical Script Failure: $($_.Exception.Message)"
}
finally {
    if ($Connection.State -eq "Open") { $Connection.Close() }
}