Import-Module ActiveDirectory
Import-Module ImportExcel

$excelPath = "x"
$data = Import-Excel -Path $excelPath

# Präfix
$prefix = "x"

foreach ($row in $data) {
    $rwName       = $row."Groupname-RW"
    $rwDesc       = $row."Groupname-RW_desc"
    $subName      = $row."Groupname-SUB"
    $subDesc      = $row."Groupname-SUB_desc"
    $memberGrpRaw = $row."Member_Grp"

    $groupsToCreate = @()
    if ($rwName) {
        $groupsToCreate += @{ Name = $rwName; Desc = $rwDesc }
    }
    if ($subName) {
        $groupsToCreate += @{ Name = $subName; Desc = $subDesc }
    }

    if ($groupsToCreate.Count -eq 0) {
        Write-Warning "Zeile übersprungen – kein RW oder SUB Gruppenname vorhanden."
        continue
    }

    foreach ($grp in $groupsToCreate) {
        $grpName = $grp.Name
        $grpDesc = $grp.Desc 

        if ($grpName -match "x(.+?)_RW$") {
            $wildcard = $Matches[1]
            $grpDesc = "Members of this group have Modify permission on x${wildcard} _RW"
        }
        elseif ($grpName -match "KSPDaten_(R_BUENDL_.+?)_SUB$") {
            $subPart = $Matches[1]
            $grpDesc = "Members of this group have List permission to subfolders up to $subPart _SUB"
        }

        if (-not (Get-ADGroup -Filter "Name -eq '$grpName'" -ErrorAction SilentlyContinue)) {
            try {
                New-ADGroup -Name $grpName `
                            -GroupCategory Security `
                            -GroupScope DomainLocal `
                            -Description $grpDesc `
                            -Path "x"
                Write-Host "Gruppe erstellt: $grpName"
            } catch {
                Write-Warning "Fehler beim Erstellen von '$grpName': $_"
            }
        } else {
            Write-Host "Gruppe bereits vorhanden: $grpName"
        }
    }


    #  Grp für Mitgliedschaften 
    Write-Host "Ursprüngliche Gruppenrohwerte: '$memberGrpRaw'"
    $memberGroupNames = $memberGrpRaw -split "\s+" | Where-Object { $_.Trim() -ne "" } | ForEach-Object { "$prefix$($_.Trim())" }
    Write-Host "Erkannte Gruppen für Mitgliedschaft: $($memberGroupNames -join ', ')"

    $validMemberGroups = foreach ($gName in $memberGroupNames) {
        if (Get-ADGroup -Identity $gName -ErrorAction SilentlyContinue) {
            $gName
        } else {
            Write-Warning "Gruppe nicht gefunden: $gName"
        }
    }

    # Mitgliedschaften hinzufügen
    foreach ($targetGroup in $groupsToCreate) {
        foreach ($memberName in $validMemberGroups) {
            try {
                Add-ADGroupMember -Identity $targetGroup.Name -Members $memberName
                Write-Host "$memberName wurde zu $($targetGroup.Name) hinzugefügt"
            } catch {
                Write-Warning "Fehler beim Hinzufügen von $memberName zu $($targetGroup.Name)"
                Write-Error $_
            }
        }
    }
}
