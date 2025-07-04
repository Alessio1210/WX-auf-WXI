$ConsumerOUs = @(
    "OU=DE,OU=WX,OU=CONSUMER,DC=group,DC=pirelli,DC=com",   
    "OU=DE,OU=WXI,OU=CONSUMER,DC=group,DC=pirelli,DC=com"
)

$DeviceTypes = @("DESKTOPS", "LAPTOPS")


$Locations = @(
    "DriverDEHQ",
    "FactoryOffice",
    "FactoryRestricted",
    "HoechstWarehouse",
    "Hoechst",
    "Munich",
    "PCSDE",
    "Pneumobil",
    "SalesForceTyres"
)

$results = @()


function Test-OUExists {
    param([string]$OUPath)
    try {
        $testOU = Get-ADOrganizationalUnit -Identity $OUPath -ErrorAction Stop
        return $true
    }
    catch {
        return $false
    }
}

foreach ($consumerOU in $ConsumerOUs) {
    Write-Host "Processing OU: $consumerOU"
    

    if (-not (Test-OUExists -OUPath $consumerOU)) {
        Write-Host "  WARNING: OU does not exist or you don't have access: $consumerOU" -ForegroundColor Yellow
        $results += [PSCustomObject]@{
            OU = $consumerOU
            WindowsVersion = if ($consumerOU -like "*WXI*") { "Windows 11" } else { "Windows 10" }
            TotalComputers = 0
            Desktops = 0
            Laptops = 0
            Status = "OU not found or no access"
        }
        continue
    }
    

    try {
        Write-Host "  Attempting to get computers from OU..."
        $allComputers = Get-ADComputer -SearchBase $consumerOU -Filter * -ErrorAction Stop
        $totalComputers = $allComputers.Count
        
        Write-Host "  Total computers found: $totalComputers"
        
      
        $desktops = @()
        $laptops = @()
        
        foreach ($computer in $allComputers) {
            if ($computer.Name -like "*LAP*" -or $computer.Name -like "*LT*" -or $computer.Name -like "*NB*") {
                $laptops += $computer
            } else {
                $desktops += $computer
            }
        }
        
        $desktopCount = $desktops.Count
        $laptopCount = $laptops.Count
        
        Write-Host "  Desktops: $desktopCount"
        Write-Host "  Laptops: $laptopCount"
        
        
        $results += [PSCustomObject]@{
            OU = $consumerOU
            WindowsVersion = if ($consumerOU -like "*WXI*") { "Windows 11" } else { "Windows 10" }
            TotalComputers = $totalComputers
            Desktops = $desktopCount
            Laptops = $laptopCount
            Status = "Success"
        }
        
        
        Write-Host "  Getting sub-OUs for information..."
        $subOUs = Get-ADOrganizationalUnit -SearchBase $consumerOU -Filter * | Where-Object {$_.DistinguishedName -ne $consumerOU}
        
        foreach ($subOU in $subOUs) {
            try {
                Write-Host "    Processing sub-OU: $($subOU.Name)"
               
                $subOUComputers = Get-ADComputer -SearchBase $subOU.DistinguishedName -Filter * -SearchScope OneLevel
                $subOUCount = $subOUComputers.Count
                
                if ($subOUCount -gt 0) {
                    Write-Host "    Sub-OU $($subOU.Name): $subOUCount computers (DIRECT only, for information)"
                    
                    
                    $results += [PSCustomObject]@{
                        OU = $subOU.DistinguishedName
                        WindowsVersion = if ($consumerOU -like "*WXI*") { "Windows 11" } else { "Windows 10" }
                        TotalComputers = $subOUCount
                        Desktops = "INFO"
                        Laptops = "INFO"
                        Status = "Information Only"
                    }
                }
            }
            catch {
                Write-Host "    Sub-OU $($subOU.Name): Access denied" -ForegroundColor Red
            }
        }
    }
    catch {
        Write-Host "  Error accessing OU: $consumerOU" -ForegroundColor Red
        Write-Host "  Error details: $($_.Exception.Message)" -ForegroundColor Red
        $results += [PSCustomObject]@{
            OU = $consumerOU
            WindowsVersion = if ($consumerOU -like "*WXI*") { "Windows 11" } else { "Windows 10" }
            TotalComputers = 0
            Desktops = 0
            Laptops = 0
            Status = "Error: $($_.Exception.Message)"
        }
    }
}

Write-Host "`nProcessing Consumer USERS DE locations..."
$ConsumerUsersDE = "OU=DE,OU=USERS,OU=CONSUMER,DC=group,DC=pirelli,DC=com"


if (-not (Test-OUExists -OUPath $ConsumerUsersDE)) {
    Write-Host "WARNING: Main USERS DE OU does not exist: $ConsumerUsersDE" -ForegroundColor Yellow
} else {
    foreach ($location in $Locations) {
        try {
            $locationOU = "OU=$location,$ConsumerUsersDE"
            Write-Host "  Processing location: $location"
            $userCount = (Get-ADUser -SearchBase $locationOU -Filter * -ErrorAction SilentlyContinue).Count
            Write-Host "  $location - $userCount users found"
            
            $results += [PSCustomObject]@{
                OU = $locationOU
                DeviceType = "USERS"
                WindowsVersion = "N/A"
                Computers = 0
                Users = $userCount
                Groups = 0
                Status = "Success"
            }
        }
        catch {
            Write-Host "  $location - OU not found or access denied" -ForegroundColor Red
            $results += [PSCustomObject]@{
                OU = "OU=$location,$ConsumerUsersDE"
                DeviceType = "USERS"
                WindowsVersion = "N/A"
                Computers = 0
                Users = 0
                Groups = 0
                Status = "OU not found or access denied"
            }
        }
    }
}


function Get-LocationName($ou) {

    if ($ou -match "OU=(?:DESKTOPS|LAPTOPS),OU=([^,]+),OU=DE,OU=(?:WX|WXI),OU=CONSUMER") {
        return $matches[1]}
    
    elseif ($ou -match "OU=([^,]+),OU=DE,OU=(?:WX|WXI),OU=CONSUMER") {
        return $matches[1] 
    }
 
    elseif ($ou -match "OU=([^,]+),OU=(?:WX|WXI),OU=CONSUMER") {
        return $matches[1]  
    } 
    else {
        return "Unknown"
    }
}


function Get-DeviceType($ou) {
    if ($ou -match "OU=(DESKTOPS|LAPTOPS),") {
        return $matches[1]
    } else {
        return "Mixed"  
    }
}


$finalResults = @()

$ouResults = $results | Where-Object { 
    $_.WindowsVersion -ne 'N/A' -and 
    $_.Status -eq 'Success' -and
    $_.TotalComputers -gt 0
}

Write-Host "`n=== DEBUG: Alle OU-Ergebnisse mit Computern ==="
$ouResults | ForEach-Object { 
    Write-Host "OU: $($_.OU) - Windows: $($_.WindowsVersion) - Computer: $($_.TotalComputers) - Status: $($_.Status)"
}


$userResults = $results | Where-Object { $_.DeviceType -eq 'USERS' }


$locationData = @{}

foreach ($result in $ouResults) {
    $location = Get-LocationName $result.OU
    $deviceType = Get-DeviceType $result.OU
    $count = $result.TotalComputers
    
    Write-Host "DEBUG: OU '$($result.OU)' -> Standort: '$location', Typ: '$deviceType', Count: $count"
    
    if (-not $locationData.ContainsKey($location)) {
        $locationData[$location] = @{
            "User" = 0
            "Desktops W10" = 0
            "Notebooks W10" = 0
            "Desktops W11" = 0
            "Notebooks W11" = 0
        }
    }
    
    if ($result.WindowsVersion -eq "Windows 10") {
        if ($deviceType -eq "DESKTOPS" -or $deviceType -eq "Mixed") {
            $locationData[$location]["Desktops W10"] += $count
        } elseif ($deviceType -eq "LAPTOPS") {
            $locationData[$location]["Notebooks W10"] += $count
        }
    } elseif ($result.WindowsVersion -eq "Windows 11") {
        if ($deviceType -eq "DESKTOPS" -or $deviceType -eq "Mixed") {
            $locationData[$location]["Desktops W11"] += $count
        } elseif ($deviceType -eq "LAPTOPS") {
            $locationData[$location]["Notebooks W11"] += $count
        }
    }
}


foreach ($userResult in $userResults) {
    
    Write-Host "DEBUG: User-OU: $($userResult.OU) - Users: $($userResult.Users)"
    if ($userResult.OU -match "OU=([^,]+),OU=DE,OU=USERS,OU=CONSUMER") {
        $location = $matches[1]
        Write-Host "DEBUG: Extrahierter User-Standort: '$location'"
        if ($locationData.ContainsKey($location)) {
            $locationData[$location]["User"] += $userResult.Users
            Write-Host "DEBUG: User zu bestehendem Standort hinzugefügt: $location"
        } else {
           
            $locationData[$location] = @{
                "User" = $userResult.Users
                "Desktops W10" = 0
                "Notebooks W10" = 0
                "Desktops W11" = 0
                "Notebooks W11" = 0
            }
            Write-Host "DEBUG: Neuer Standort für User erstellt: $location"
        }
    } else {
        Write-Host "DEBUG: User-OU passt nicht zum Pattern: $($userResult.OU)"
    }
}


Write-Host "`n=== DEBUG: Sub-OUs für bessere Standort-Auflösung ==="
foreach ($result in $results) {
    if ($result.Status -eq "Information Only" -and $result.TotalComputers -gt 0) {
        Write-Host "Sub-OU gefunden: $($result.OU) - Computer: $($result.TotalComputers)"
        
        if ($result.OU -match "OU=(?:DESKTOPS|LAPTOPS),OU=([^,]+),OU=DE,OU=(?:WX|WXI),OU=CONSUMER") {
            $location = $matches[1]
            $deviceType = if ($result.OU -match "OU=(DESKTOPS|LAPTOPS),") { $matches[1] } else { "Mixed" }
            
            Write-Host "  -> Standort: '$location', Typ: '$deviceType'"
            
            if (-not $locationData.ContainsKey($location)) {
                $locationData[$location] = @{
                    "User" = 0
                    "Desktops W10" = 0
                    "Notebooks W10" = 0
                    "Desktops W11" = 0
                    "Notebooks W11" = 0
                }
            }
            
            
            $windowsVersion = $result.WindowsVersion
            
            if ($windowsVersion -eq "Windows 10") {
                if ($deviceType -eq "DESKTOPS") {
                    $locationData[$location]["Desktops W10"] += $result.TotalComputers
                } elseif ($deviceType -eq "LAPTOPS") {
                    $locationData[$location]["Notebooks W10"] += $result.TotalComputers
                }
            } elseif ($windowsVersion -eq "Windows 11") {
                if ($deviceType -eq "DESKTOPS") {
                    $locationData[$location]["Desktops W11"] += $result.TotalComputers
                } elseif ($deviceType -eq "LAPTOPS") {
                    $locationData[$location]["Notebooks W11"] += $result.TotalComputers
                }
            }
        }
    }
}


foreach ($location in $locationData.Keys) {
    $data = $locationData[$location]
    $finalResults += [PSCustomObject]@{
        "OU DE" = $location
        "User" = $data["User"]
        "Desktops W10" = $data["Desktops W10"]
        "Notebooks W10" = $data["Notebooks W10"]
        "Desktops W11" = $data["Desktops W11"]
        "Notebooks W11" = $data["Notebooks W11"]
        "Users Presumed" = $data["User"]
    }
}


if ($finalResults.Count -eq 0) {
    Write-Host "WARNUNG: Keine Standorte gefunden, erstelle Gesamt-Übersicht"
    
    $totalW10 = ($ouResults | Where-Object {$_.WindowsVersion -eq "Windows 10"} | Measure-Object -Property TotalComputers -Sum).Sum
    $totalW11 = ($ouResults | Where-Object {$_.WindowsVersion -eq "Windows 11"} | Measure-Object -Property TotalComputers -Sum).Sum
    $totalUsers = ($userResults | Measure-Object -Property Users -Sum).Sum
    
    $finalResults += [PSCustomObject]@{
        "OU DE" = "Gesamt DE"
        "User" = $totalUsers
        "Desktops W10" = $totalW10
        "Notebooks W10" = 0
        "Desktops W11" = $totalW11
        "Notebooks W11" = 0
        "Users Presumed" = $totalUsers
    }
}

$totalUsers = ($finalResults | Measure-Object -Property "User" -Sum).Sum
$totalDesktopsW10 = ($finalResults | Measure-Object -Property "Desktops W10" -Sum).Sum
$totalNotebooksW10 = ($finalResults | Measure-Object -Property "Notebooks W10" -Sum).Sum
$totalDesktopsW11 = ($finalResults | Measure-Object -Property "Desktops W11" -Sum).Sum
$totalNotebooksW11 = ($finalResults | Measure-Object -Property "Notebooks W11" -Sum).Sum

$finalResults += [PSCustomObject]@{
    "OU DE" = "Total DE"
    "User" = $totalUsers
    "Desktops W10" = $totalDesktopsW10
    "Notebooks W10" = $totalNotebooksW10
    "Desktops W11" = $totalDesktopsW11
    "Notebooks W11" = $totalNotebooksW11
    "Users Presumed" = $totalUsers
}

Write-Host "`n=== DEBUG: Finale Ergebnisse für Dashboard ==="
$finalResults | ForEach-Object { 
    Write-Host "Standort: $($_.'OU DE')"
    Write-Host "  User: $($_.'User')"
    Write-Host "  Desktops W10: $($_.'Desktops W10')"
    Write-Host "  Notebooks W10: $($_.'Notebooks W10')"
    Write-Host "  Desktops W11: $($_.'Desktops W11')"
    Write-Host "  Notebooks W11: $($_.'Notebooks W11')"
    Write-Host "  ---"
}


$htmlPath = Join-Path $env:USERPROFILE "Desktop\\Consumer_Dashboard.html"


Write-Host "`n=== DEBUG: Dashboard-Statistiken ==="
Write-Host "totalUsers: $totalUsers"
Write-Host "totalDesktopsW10: $totalDesktopsW10"
Write-Host "totalNotebooksW10: $totalNotebooksW10"
Write-Host "totalDesktopsW11: $totalDesktopsW11"
Write-Host "totalNotebooksW11: $totalNotebooksW11"
Write-Host "Windows 10 Geräte: $($totalDesktopsW10 + $totalNotebooksW10)"
Write-Host "Windows 11 Geräte: $($totalDesktopsW11 + $totalNotebooksW11)"
Write-Host "Gesamt Computer: $(($totalDesktopsW10 + $totalNotebooksW10) + ($totalDesktopsW11 + $totalNotebooksW11))"

$html = @"
<!DOCTYPE html>
<html lang="de">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Consumer IT Asset Dashboard</title>
    <style>
        * {
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }
        
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }
        
        .container {
            max-width: 1400px;
            margin: 0 auto;
            background: white;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }
        
        .header {
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 30px;
            text-align: center;
        }
        
        .header h1 {
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }
        
        .header p {
            font-size: 1.2em;
            opacity: 0.9;
        }
        
        .stats-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(250px, 1fr));
            gap: 20px;
            padding: 30px;
            background: #f8f9fa;
        }
        
        .stat-card {
            background: white;
            padding: 25px;
            border-radius: 10px;
            text-align: center;
            box-shadow: 0 5px 15px rgba(0,0,0,0.08);
            transition: transform 0.3s ease;
        }
        
        .stat-card:hover {
            transform: translateY(-5px);
        }
        
        .stat-number {
            font-size: 2.5em;
            font-weight: bold;
            margin-bottom: 10px;
        }
        
        .stat-label {
            color: #666;
            font-size: 1.1em;
        }
        
        .windows10 { color: #0078d4; }
        .windows11 { color: #0066cc; }
        .users { color: #16a085; }
        .total { color: #e74c3c; }
        
        .table-container {
            padding: 30px;
            overflow-x: auto;
        }
        
        .data-table {
            width: 100%;
            border-collapse: collapse;
            background: white;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 10px 25px rgba(0,0,0,0.1);
        }
        
        .data-table th {
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            padding: 15px 12px;
            text-align: center;
            font-weight: 600;
            font-size: 0.95em;
        }
        
        .data-table td {
            padding: 12px;
            text-align: center;
            border-bottom: 1px solid #eee;
            transition: background-color 0.3s ease;
        }
        
        .data-table tbody tr:hover {
            background-color: #f8f9fa;
        }
        
        .data-table tbody tr:nth-child(even) {
            background-color: #f9f9f9;
        }
        
        .location-cell {
            text-align: left !important;
            font-weight: 600;
            color: #2c3e50;
        }
        
        .total-row {
            background: linear-gradient(135deg, #e74c3c 0%, #c0392b 100%) !important;
            color: white !important;
            font-weight: bold;
        }
        
        .total-row td {
            border-bottom: none !important;
        }
        
        .number-cell {
            font-family: 'Courier New', monospace;
            font-weight: 600;
        }
        
        .footer {
            background: #2c3e50;
            color: white;
            text-align: center;
            padding: 20px;
            font-size: 0.9em;
        }
        
        .badge {
            display: inline-block;
            padding: 4px 8px;
            border-radius: 15px;
            font-size: 0.8em;
            font-weight: bold;
            margin-left: 5px;
        }
        
        .badge-w10 { background: #e3f2fd; color: #1976d2; }
        .badge-w11 { background: #f3e5f5; color: #7b1fa2; }
        
        /* Migration Section Styles */
        .migration-section {
            margin: 0 30px 30px 30px;
            padding: 0;
        }
        
        .migration-card {
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            border-radius: 15px;
            padding: 30px;
            color: white;
            box-shadow: 0 10px 30px rgba(0,0,0,0.2);
        }
        
        .migration-header {
            text-align: center;
            margin-bottom: 30px;
        }
        
        .migration-header h2 {
            font-size: 2.2em;
            font-weight: 300;
            margin-bottom: 10px;
        }
        
        .migration-header p {
            font-size: 1.1em;
            opacity: 0.9;
        }
        
        .migration-stats {
            display: grid;
            grid-template-columns: auto 1fr;
            gap: 40px;
            align-items: center;
        }
        
        .progress-circle {
            width: 180px;
            height: 180px;
            border-radius: 50%;
            background: conic-gradient(
                #4CAF50 0deg,
                #4CAF50 $(
                    $totalComputers = ($totalDesktopsW10 + $totalNotebooksW10) + ($totalDesktopsW11 + $totalNotebooksW11)
                    if ($totalComputers -gt 0) {
                        $percentage = [math]::Round((($totalDesktopsW11 + $totalNotebooksW11) / $totalComputers) * 100, 1)
                        "$percentage%"
                    } else { "0%" }
                ),
                rgba(255,255,255,0.3) $(
                    $totalComputers = ($totalDesktopsW10 + $totalNotebooksW10) + ($totalDesktopsW11 + $totalNotebooksW11)
                    if ($totalComputers -gt 0) {
                        $percentage = [math]::Round((($totalDesktopsW11 + $totalNotebooksW11) / $totalComputers) * 100, 1)
                        "$percentage%"
                    } else { "0%" }
                ),
                rgba(255,255,255,0.3) 100%
            );
            display: flex;
            align-items: center;
            justify-content: center;
            position: relative;
            box-shadow: 0 5px 20px rgba(0,0,0,0.3);
        }
        
        .progress-inner {
            width: 140px;
            height: 140px;
            border-radius: 50%;
            background: rgba(255,255,255,0.15);
            backdrop-filter: blur(10px);
            display: flex;
            flex-direction: column;
            align-items: center;
            justify-content: center;
            text-align: center;
        }
        
        .progress-percentage {
            font-size: 2.5em;
            font-weight: bold;
            color: #4CAF50;
            text-shadow: 0 2px 4px rgba(0,0,0,0.3);
        }
        
        .progress-label {
            font-size: 0.9em;
            margin-top: 5px;
            opacity: 0.9;
        }
        
        .migration-details {
            display: flex;
            flex-direction: column;
            gap: 15px;
        }
        
        .detail-item {
            display: flex;
            align-items: center;
            gap: 15px;
            padding: 15px;
            background: rgba(255,255,255,0.1);
            border-radius: 10px;
            backdrop-filter: blur(5px);
            border: 1px solid rgba(255,255,255,0.2);
        }
        
        .detail-icon {
            font-size: 1.3em;
            width: 30px;
            text-align: center;
        }
        
        .detail-text {
            font-size: 1.1em;
            flex: 1;
        }
        
        .detail-text strong {
            color: #4CAF50;
            font-weight: 600;
        }
        
        @media (max-width: 768px) {
            .header h1 { font-size: 2em; }
            .stats-grid { grid-template-columns: 1fr; }
            .table-container { padding: 15px; }
            .data-table th, .data-table td { padding: 8px 6px; font-size: 0.9em; }
            .migration-stats { grid-template-columns: 1fr; text-align: center; }
            .progress-circle { margin: 0 auto; }
        }
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>WX auf WXI</h1>
        </div>
        
        <div class="stats-grid">
            <div class="stat-card">
                <div class="stat-number total">$($totalUsers)</div>
                <div class="stat-label">Alle Benutzer</div>
            </div>
            <div class="stat-card">
                <div class="stat-number windows10">$($totalDesktopsW10 + $totalNotebooksW10)</div>
                <div class="stat-label">WX Geräte</div>
            </div> 
            <div class="stat-card">
                <div class="stat-number windows11">$($totalDesktopsW11 + $totalNotebooksW11)</div>
                <div class="stat-label">WXI Geräte</div>
            </div>
            <div class="stat-card">
                <div class="stat-number total">$(($totalDesktopsW10 + $totalNotebooksW10) + ($totalDesktopsW11 + $totalNotebooksW11))</div>
                <div class="stat-label">Alle Geräte</div>
            </div>
        </div>
        
        <!-- Migrations-Fortschritt -->
        <div class="migration-section">
            <div class="migration-card">
                <div class="migration-header">
                    <h2>Progress</h2>
                    <p>Das sind die PCs die bereits umgestellt wurden</p>
                </div>
                
                <div class="migration-stats">
                    <div class="progress-circle">
                        <div class="progress-inner">
                            <div class="progress-percentage">$(
                                $totalComputers = ($totalDesktopsW10 + $totalNotebooksW10) + ($totalDesktopsW11 + $totalNotebooksW11)
                                if ($totalComputers -gt 0) {
                                    [math]::Round((($totalDesktopsW11 + $totalNotebooksW11) / $totalComputers) * 100, 1)
                                } else { 0 }
                            )%</div>
                            <div class="progress-label">Migriert</div>
                        </div>
                    </div>
                    
                    <div class="migration-details">
                        <div class="detail-item">
                         
                            <span class="detail-text">Migriert: <strong>$($totalDesktopsW11 + $totalNotebooksW11)</strong> PCs</span>
                        </div>
                        <div class="detail-item">
                      
                            <span class="detail-text">Verbleibend: <strong>$(
                                $remaining = $totalDesktopsW10 + $totalNotebooksW10
                                if ($remaining -eq 0) { "Migration abgeschlossen! 🎉" } else { "$remaining PCs" }
                            )</strong></span>
                        </div>
                    </div>
                </div>
            </div>
        </div>
        
        <div class="table-container">
            <table class="data-table">
                <thead>
                    <tr>
                        <th>Standort</th>
                        <th>User</th>
                        <th>Desktops W10</th>
                        <th>Laptops W10</th>
                        <th>Desktops W11</th>
                        <th>Notebooks W11</th>
                    </tr>
                </thead>
                <tbody>
"@

foreach ($row in $finalResults) {
    $rowClass = if ($row."OU DE" -eq "Total DE") { "total-row" } else { "" }
    $locationClass = if ($row."OU DE" -eq "Total DE") { "" } else { "location-cell" }
    
    $html += @"
                    <tr class="$rowClass">
                        <td class="$locationClass">$($row."OU DE")</td>
                        <td class="number-cell">$($row."User")</td>
                        <td class="number-cell">$($row."Desktops W10")</td>
                        <td class="number-cell">$($row."Notebooks W10")</td>
                        <td class="number-cell">$($row."Desktops W11")</td>
                        <td class="number-cell">$($row."Notebooks W11")</td>
                    </tr>
"@
}

$html += @"
                </tbody>
            </table>
        </div>
    </div>
</body>
</html>
"@


Start-Process $htmlPath


Write-Host "`n=== DEBUG: ==="
$results | ForEach-Object { 
    Write-Host "OU: $($_.OU)"
    Write-Host "  WindowsVersion: $($_.WindowsVersion)"
    Write-Host "  TotalComputers: $($_.TotalComputers)"
    Write-Host "  Status: $($_.Status)"
    Write-Host "  DeviceType: $($_.DeviceType)"
    Write-Host "  ---"
}

