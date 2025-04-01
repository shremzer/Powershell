function Get-UserIP {
    param (
        [string]$region,
        [string]$csvColumn,
        [string]$password,
        [switch]$NoCredential
    )

    $servers = Import-Csv E:\laptops\all.csv | ForEach-Object { $_.$csvColumn } | Where-Object { $_ }
    $results = @{}

    foreach ($s in $servers) {
        if ($NoCredential) {
            $result = Invoke-Command -ComputerName $s -ScriptBlock {
                          $name = if ($user = (Get-WmiObject -Class win32_computersystem).UserName) {
    if ($user -like "*manager*") {
        cd C:\users
        $m = ls | ? { $_.Name -match "manager" } | sort lastwritetime -descending
        cd $m[0]
        cd .\appdata\roaming\thunderbird\profiles
        $list = dir | sort lastwritetime -descending
        cd $list[0]
        $n = type prefs.js | ? { $_ -match "useremail" } | select -first 1
        $user=$n.Split().split("@")[1].trimstart("""")
$name.TrimStart('OP8\')
    } else {
        $user=$user.TrimStart('OP8\')
    }
} else {
                 cd C:\Users\
                 $m = Get-ChildItem | Sort-Object LastWriteTime -Descending
                 $m[0].Name
             }
$ip = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '169.*' -and $_.IPAddress -notlike '192.*'  -and $_.IPAddress -ne '127.0.0.1' }).IPAddress
                @{ $user = "$ip" }
            } -ErrorAction SilentlyContinue
        } else {
            $securePassword = ConvertTo-SecureString -String $password -AsPlainText -Force
            $cred = New-Object System.Management.Automation.PSCredential ("$s\admin", $securePassword)
            $result = Invoke-Command -ComputerName $s -ScriptBlock {
                cd C:\users;
                $m = ls | ? { $_.Name -match "manager" } | sort lastwritetime -descending;
                cd $m[0];
                cd .\appdata\roaming\thunderbird\profiles;
                $list = dir | sort lastwritetime -descending;
                cd $list[0];
                $n = type prefs.js | ? { $_ -match "useremail" } | select -first 1;
                $user = $n.Split().split("@")[1].trimstart("""")

$ip = (Get-NetIPAddress | Where-Object { $_.AddressFamily -eq 'IPv4' -and $_.IPAddress -notlike '169.*' -and $_.IPAddress -notlike '192.*'  -and $_.IPAddress -ne '127.0.0.1' }).IPAddress
                @{ $user = "$ip" }
            } -Credential $cred -ErrorAction SilentlyContinue
        }

        foreach ($key in $result.Keys) {
            $results[$key] = $result[$key]
        }
    }

    return $results
}

# Define the password
$password = "Fenix1985"

# Process each region/category
$results1 = Get-UserIP -region "REGIONS" -csvColumn "regions" -password $password
$results2 = Get-UserIP -region "MOSCOW" -csvColumn "moscow" -password $password
$results3 = Get-UserIP -region "TM" -csvColumn "TM" -password $password
$results4 = Get-UserIP -region "2floor" -csvColumn "Twofloor" -password $password
$results5 = Get-UserIP -region "Service" -csvColumn "service" -password $password
$results6 = Get-UserIP -region "IRBIS" -csvColumn "irbis_and_desktop" -NoCredential
$results7 = Get-UserIP -region "AD" -csvColumn "AD" -NoCredential

# Merge all results
$finalResults = @{}
$allResults = $results1, $results2, $results3, $results4, $results5, $results6, $results7

foreach ($result in $allResults) {
    foreach ($key in $result.Keys) {
        $finalResults[$key] = $result[$key]
    }
}

# Output the final results
$finalResults

Export-Clixml -Path "E:\hash.cli" -InputObject $finalResults

