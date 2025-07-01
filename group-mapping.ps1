Import-Module ActiveDirectory
Import-Module ImportExcel

$excelPath = "C:\Pfad\zur\Datei.xlsx"
$data = Import-Excel -Path $excelPath

# Log-Datei im selben Ordner wie Excel-Datei
$excelFolder = Split-Path $excelPath -Parent
$logPath = Join-Path $excelFolder "ADGroupScript_Errors.txt"

if (Test-Path $logPath) { Remove-Item $logPath }
New-Item -Path $logPath -ItemType File -Force | Out-Null

$createdGroupsCount = 0
$prefix = "x"

foreach ($row in $data) {
    $rwName       = $row."Groupname-RW"
    $rwDesc       = $row."Groupname-RW_desc"
    $subName      = $row."Groupname-SUB"
    $subDesc      = $row."Groupname-SUB_desc"
    $memberGrpRaw = $row."Member_Grp"

    $groupsToCreate = @()
    if ($rwName) { $groupsToCreate += @{ Name = $rwName; Desc = $rwDesc } }
    if ($subName) { $groupsToCreate += @{ Name = $subName; Desc = $subDesc } }

    if ($groupsToCreate.Count -eq 0) {
        $msg = "Zeile übersprungen – kein Gruppenname vorhanden."
        Write-Warning $msg
        Add-Content -Path $logPath -Value $msg
        continue
    }

    foreach ($grp in $groupsToCreate) {
        $grpName = $grp.Name
        $grpDesc = $grp.Desc

        if ($grpName -match "x(.+?)_RW$") {
            $grpDesc = "Modify permission on x$($Matches[1]) _RW"
        }
        elseif ($grpName -match "KSPDaten_(R_BUENDL_.+?)_SUB$") {
            $grpDesc = "List permission to subfolders up to $($Matches[1]) _SUB"
        }

        if (-not (Get-ADGroup -Filter "Name -eq '$grpName'" -ErrorAction SilentlyContinue)) {
            try {
                New-ADGroup -Name $grpName `
                            -GroupCategory Security `
                            -GroupScope DomainLocal `
                            -Description $grpDesc `
                            -Path "OU=Gruppen,DC=deinedomain,DC=local" # anpassen
                Write-Host "Gruppe erstellt: $grpName"
                $createdGroupsCount++
            } catch {
                $msg = "Fehler beim Erstellen von '$grpName': $_"
                Write-Warning $msg
                Add-Content -Path $logPath -Value $msg
            }
        } else {
            Write-Host "Gruppe vorhanden: $grpName"
        }
    }

    $memberGroupNames = $memberGrpRaw -split "\s+" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { "$prefix$($_.Trim())" }

    $validMemberGroups = foreach ($gName in $memberGroupNames) {
        if (Get-ADGroup -Identity $gName -ErrorAction SilentlyContinue) {
            $gName
        } else {
            $msg = "Gruppe nicht gefunden: $gName"
            Write-Warning $msg
            Add-Content -Path $logPath -Value $msg
        }
    }

    foreach ($targetGroup in $groupsToCreate) {
        foreach ($memberName in $validMemberGroups) {
            try {
                Add-ADGroupMember -Identity $targetGroup.Name -Members $memberName
                Write-Host "$memberName zu $($targetGroup.Name) hinzugefügt"
            } catch {
                $msg = "Fehler beim Hinzufügen von $memberName zu $($targetGroup.Name): $_"
                Write-Warning $msg
                Add-Content -Path $logPath -Value $msg
            }
        }
    }
}

Write-Host "`nErfolgreich erstellte Gruppen: $createdGroupsCount"
Add-Content -Path $logPath -Value "`nErfolgreich erstellte Gruppen: $createdGroupsCount"
