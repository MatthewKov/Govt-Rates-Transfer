$curDir = Get-Location
$pdf_extract = "DLARates.csv"

$fst = Get-Content -Path $pdf_extract -first 1
[void] ($fst -match "20[0-9][0-9]")
$year = [int]$matches[0]

$temp1 = New-Item -Name "temp1.csv" -ItemType "file"
Get-Content $pdf_extract | Select-Object -skip 2 | Set-Content $temp1

$temp2 = New-Item -Name "temp2.csv" -ItemType "file"
Import-csv $temp1 | Select-Object 'OrgID', 'Grade', 'EffDate', 'ExpDate', 'With-Dependent Rate', `
    'Without-Dependent Rate', 'DLAPrimary', 'LastUpdatedBy', 'LastUpdatedOn' |
Export-csv $temp2 -NoTypeInformation

if(!(Test-Path ./FinishedRates.csv)) {
    $finished_rates = New-Item -Name "FinishedRates.csv" -ItemType "file"
}
else { $finished_rates = "$curDir\FinishedRates.csv" }

$i = 0
$csv = import-csv $temp2
foreach($line in $csv) {
    $line.'With-Dependent Rate' = $line.'With-Dependent Rate'.Replace("$", "").Replace(",", "")
    $line.'Without-Dependent Rate' = $line.'Without-Dependent Rate'.Replace("$", "").Replace(",", "")
    if($i -lt 27){
        $line.DLAPrimary = 1
    }
    else { $line.DLAPrimary = 0 }
    $i++
}
$csv | export-csv $finished_rates -NoTypeInformation

remove-item $temp1
remove-item $temp2

do {
    $failed = $false
    $server = Read-Host -Prompt "Enter the server name"
    $database = Read-Host -Prompt "Enter the database name"
    $username = Read-Host -Prompt "Enter the SQL username"
    $password = Read-Host -Prompt "Enter the SQL password"
    $connectionString = 'Data Source={0};database={1};User ID={2};Password={3}' -f $server,$database,$username,$password
    $sqlConnection = New-Object System.Data.SqlClient.SqlConnection $connectionString
    $ErrorActionPreference = "Stop"
    try {
        $sqlConnection.open()
    }
    catch [System.Management.Automation.MethodInvocationException] {
        $failed = $true
        Write-Output "`nCould not connect: Reenter the correct connection details and credentials`n"
    }
} while ($failed)


$createtable = "IF NOT EXISTS (SELECT * FROM sysobjects WHERE name='NewDLARates' and xtype='U') 
    CREATE TABLE NewDLARates 
        (OrgID bigint,
        Grade varchar(25),
        EffDate smalldatetime,
        ExpDate smalldatetime,
        WithDep money,
        WithoutDep money,
        DLAPrimary bit,
        LastUpdatedBy varchar(50),
        LastUpdatedOn datetime)"
invoke-sqlcmd -SuppressProviderContextWarning -ServerInstance $server -Database $database -Query $createtable

Import-csv $finished_rates | ForEach-Object {Invoke-Sqlcmd -SuppressProviderContextWarning `
    -Database $database -ServerInstance $server `
    -Query "insert into NewDLARates VALUES ('$($_.OrgID)','$($_.Grade)','$($_.EffDate)',
    '$($_.ExpDate)','$($_.'With-Dependent Rate')','$($_.'Without-Dependent Rate')','$($_.DLAPrimary)',
    '$($_.LastUpdatedBy)','$($_.LastUpdateOn)')"
}

remove-item $finished_rates

$updatetable = "UPDATE NewDLARates 
                SET OrgID = 1, 
                EffDate = '$year-01-01 00:00:00', 
                ExpDate = '2049-12-31 00:00:00', 
                LastUpdatedBy = 'ReloAdmin', 
                LastUpdatedOn = CURRENT_TIMESTAMP;"
Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server -Query $updatetable

$rowcount = (Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server `
    -Query "SELECT COUNT(*) FROM DLARates").column1

Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server -Query "DBCC CHECKIDENT (DLARates, RESEED, $rowcount)"

Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server -Query "INSERT INTO DLARates SELECT * FROM NewDLARates"

$last_year = $year - 1 
Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server `
    -Query "UPDATE DLARates SET ExpDate = '$last_year-12-31 00:00:00' WHERE EffDate = '$last_year-01-01 00:00:00';"

Invoke-Sqlcmd -SuppressProviderContextWarning -Database $database -ServerInstance $server -Query "DROP TABLE NewDLARates"

Write-Output "`nScript Finished"

cd $curDir
$sqlConnection.Close()