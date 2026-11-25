
$MSX = "YOUR_MSX_SERVER"

# Get all registered target servers
$targets = Invoke-Sqlcmd -ServerInstance $MSX -Database msdb -Query "
    SELECT server_name
    FROM dbo.systargetservers
"

$results = foreach ($target in $targets) {

    $server = $target.server_name

    try {
        $details = Invoke-Sqlcmd -ServerInstance $server -Query "
            SELECT
                @@SERVERNAME AS ServerName,
                SERVERPROPERTY('InstanceName') AS InstanceName,
                SERVERPROPERTY('ProductVersion') AS ProductVersion,
                SERVERPROPERTY('ProductLevel') AS ProductLevel,
                SERVERPROPERTY('Edition') AS Edition
        "

        [PSCustomObject]@{
            TSX            = $server
            Instance       = $details.InstanceName
            Version        = $details.ProductVersion
            ProductLevel   = $details.ProductLevel
            Edition        = $details.Edition
            Status         = "Success"
        }
    }
    catch {
        [PSCustomObject]@{
            TSX            = $server
            Instance       = ""
            Version        = ""
            ProductLevel   = ""
            Edition        = ""
            Status         = "Failed: $($_.Exception.Message)"
        }
    }
}

$results | Format-Table -AutoSize

# Export report
$results | Export-Csv "C:\Temp\TSX_SQL_Versions.csv" -NoTypeInformation

$results | Format-Table TSX, Instance, Version, ProductLevel, Edition, Status -AutoSize
