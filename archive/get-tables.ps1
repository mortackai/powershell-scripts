# --- Configuration ---
$Server = "GSSDB1" 
$Database = "GLOBALCMI"
$ExportDir = "C:\Users\MRaymond\Downloads\exports"
if (!(Test-Path $ExportDir)) { New-Item -ItemType Directory -Path $ExportDir }

# --- Driver & Credentials ---
$SelectedDriver = "Pervasive ODBC Unicode Interface"
$Creds = Get-Credential -UserName "TestUser" -Message "Enter credentials for $Database"
$User = $Creds.UserName
$Pass = $Creds.GetNetworkCredential().Password

$ConnString = "Driver={$SelectedDriver};ServerName=$Server;DBQ=$Database;UID=$User;PWD=$Pass;"
$Connection = New-Object System.Data.Odbc.OdbcConnection($ConnString)

try {
    Write-Host "Connecting to $Database..." -ForegroundColor Gray
    $Connection.Open()

    # --- PHASE 1: DISCOVERY ---
    Write-Host "Discovering tables and checking permissions..." -ForegroundColor Cyan
    $AllTables = $Connection.GetSchema("Tables") | Where-Object { $_.TABLE_TYPE -eq "TABLE" -and $_.TABLE_NAME -notlike "f$*" }
    
    $AllowedTables = New-Object System.Collections.Generic.List[string]
    
    foreach ($row in $AllTables) {
        $tName = $row.TABLE_NAME
        try {
            # Attempt to grab 0 rows just to check SELECT permission
            $checkCmd = New-Object System.Data.Odbc.OdbcCommand("SELECT TOP 0 * FROM ""$tName""", $Connection)
            $null = $checkCmd.ExecuteNonQuery()
            $AllowedTables.Add($tName)
        }
        catch {
            # Silently skip tables where we get 'Permission Denied' or '[42000]'
            Write-Host "  [!] Skipping $tName (Access Denied)" -ForegroundColor DarkGray
        }
    }

    Write-Host "Found $($AllowedTables.Count) accessible tables.`n" -ForegroundColor Green

    # --- PHASE 2: EXPORT ---
    foreach ($TableName in $AllowedTables) {
        $FilePath = Join-Path $ExportDir "$($TableName).csv"
        Write-Host "Exporting: $TableName" -ForegroundColor Yellow
        
        try {
            # Get structure for the SQL-side Date handling we built earlier
            $setupCommand = New-Object System.Data.Odbc.OdbcCommand("SELECT TOP 1 * FROM ""$TableName""", $Connection)
            $setupReader = $setupCommand.ExecuteReader()
            $dt = $setupReader.GetSchemaTable()
            
            $SqlColumns = @()
            foreach ($col in $dt.Rows) {
                $cName = $col["ColumnName"]
                if ($col["DataTypeName"] -like "*date*" -or $col["DataTypeName"] -like "*timestamp*") {
                    $SqlColumns += "IF(""$cName"" < '0001-01-01', NULL, ""$cName"") AS ""$cName"""
                } else {
                    $SqlColumns += """$cName"""
                }
            }
            $setupReader.Close()

            # Execute full export
            $query = "SELECT " + ($SqlColumns -join ", ") + " FROM ""$TableName"""
            $cmd = New-Object System.Data.Odbc.OdbcCommand($query, $Connection)
            $reader = $cmd.ExecuteReader()

            # Write to CSV
            $Header = ($dt.Rows.ColumnName | ForEach-Object { """$_""" }) -join ","
            $Header | Out-File $FilePath -Encoding UTF8
            
            while ($reader.Read()) {
                $line = for ($i=0; $i -lt $reader.FieldCount; $i++) {
                    if ($reader.IsDBNull($i)) { "" } 
                    else { """$($reader.GetValue($i).ToString().Trim().Replace('"', '""'))""" }
                }
                $line -join "," | Out-File $FilePath -Encoding UTF8 -Append
            }
            $reader.Close()
        }
        catch {
            Write-Error "Failed to process $TableName : $($_.Exception.Message)"
        }
    }
}
finally {
    $Connection.Close()
    Write-Host "Export finished." -ForegroundColor Yellow
}