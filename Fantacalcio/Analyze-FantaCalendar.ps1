Clear-Host

#Input Variables

$ExcelFileFullPath =  (Get-Location).Path + "\Calendario_FantaCascina2025-26.xlsx"
$TotGiornate = 38

#Step A: Modify Excel Calendar File to get the desired format to compute data

Write-Host "---------------------------------------" -ForegroundColor Yellow
Write-Host "Step 1: Modifying Excel Calendar File..." -ForegroundColor Yellow
Write-Host "---------------------------------------" -ForegroundColor Yellow

# Load Excel COM object
$excel = New-Object -ComObject Excel.Application
$excel.Visible = $false
$excel.DisplayAlerts = $false

# Open workbook
$workbook = $excel.Workbooks.Open($ExcelFileFullPath)
$sheet = $workbook.Sheets.Item(1)

# 1. Delete first two rows --> OK
$sheet.Rows("1:2").Delete() | out-null

# 2. Cut and paste G2:K96 to A97 --> OK
$sourceRange = $sheet.Range("G2:K96")
$destRange = $sheet.Range("A97:E191")
$sourceRange.Cut() | out-null
$sheet.Paste($destRange)

# 3. Replace "Âª Giornata serie a" with "" in column A --> OK
$usedRange = $sheet.UsedRange
$rowCount = $usedRange.Rows.Count
for ($i = 1; $i -le $rowCount; $i++) {
    $cell = $sheet.Cells.Item($i, 1)
    if ($cell.Value2 -like "*Âª Giornata Lega*") {
        $cell.Value2 = $cell.Value2 -replace "Âª Giornata Lega", ""
    }
}

# 4. Write "Giornata" in F1 --> OK
$sheet.Cells.Item(1, 6).Value2 = "Giornata"

# 5. Write "Risultati" in G1 --> OK
$sheet.Cells.Item(1, 7).Value2 = "Risultati"

# 6. Write formula in F2 and fill down
$formulaF = 'IF(MOD(ROW(A2),5)=2,"N/A",IF(MOD(ROW(A2),5)=3,INDEX(A:A,ROW(A2)-1),IF(MOD(ROW(A2),5)=4,INDEX(A:A,ROW(A2)-2),IF(MOD(ROW(A2),5)=0,INDEX(A:A,ROW(A2)-3),INDEX(A:A,ROW(A2)-4)))))'
$sheet.Cells.Item(2, 6).Formula = "=$formulaF"
$sheet.Range("F2:F191").FillDown() | Out-Null

# 7. Write formula in G2 and fill down
$formulaG = 'CONCATENATE(A2,"|",B2,"|",C2,"|",D2)'
$sheet.Cells.Item(2, 7).Formula = "=$formulaG"
$sheet.Range("G2:G191").FillDown() | Out-Null

# Save and close
$workbook.Save()
$workbook.Close($false)
$excel.Quit()
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($sheet) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($workbook) | Out-Null
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel) | Out-Null
[GC]::Collect()
[GC]::WaitForPendingFinalizers()

$CalendarioTotale = Import-Excel -Path $ExcelFileFullPath | where-object { $_.Giornata -ne "N/A"} | Sort-Object -Property Giornata
Remove-Item -Path $ExcelFileFullPath -Force -ErrorAction SilentlyContinue

Write-Host "---------------------------------------" -ForegroundColor Green
Write-Host "Step 1: Modifying Excel Calendar File - Completed!" -ForegroundColor Green
Write-Host "---------------------------------------" -ForegroundColor Green

#Step B: Compute Rank and Results Data

Write-Host "---------------------------------------" -ForegroundColor Yellow
Write-Host "Step 2: Computing Rank and Results Data..." -ForegroundColor Yellow
Write-Host "---------------------------------------" -ForegroundColor Yellow

#1- Build the Rank to update

$PrimaGiornata = $CalendarioTotale | where-object {$_.Giornata -eq 1} | select-object Risultati
$TeamsList = @()
Foreach($Partita in $PrimaGiornata){
    $Team1 = $Partita.Risultati.Split("|")[0]
    $Team2 = $Partita.Risultati.Split("|")[3]
    $TeamsList += $Team1
    $TeamsList += $Team2
}

$FinalRank = @()
Foreach($Team in $TeamsList){
    $hash = [ordered]@{
        Squadra = $Team
        FantaPunti = 0
        RealScore = 0
        ExpectedScore = 0
    }
    $Item = New-Object PSObject -Property $hash
    $FinalRank = $FinalRank + $Item
}

for ($j = 1; $j -le $TotGiornate; $j++) {
    $Percent = (($j / $TotGiornate) * 100)
    Write-Progress -Status "Percent Complete: $Percent%" -PercentComplete (($j/$TotGiornate)*100) -Activity "Processing giornata $j" -Id 0
    $GiornataDaAnalizzare = $CalendarioTotale | where-object {$_.Giornata -eq $j}

    $PunteggiGiornata = @()
    foreach ($match in $GiornataDaAnalizzare) {
        $Team1 = $match.Risultati.Split("|")[0]
        $Team2 = $match.Risultati.Split("|")[3]
        $ScoreTeam1 = [decimal]$match.Risultati.Split("|")[1]
        $ScoreTeam2 = [decimal]$match.Risultati.Split("|")[2]

        $hash = [ordered]@{
            Squadra = $Team1
            FantaPunti = $ScoreTeam1
        }
        $Item = New-Object PSObject -Property $hash
        $PunteggiGiornata = $PunteggiGiornata + $Item

        $hash2 = [ordered]@{
            Squadra = $Team2
            FantaPunti = $ScoreTeam2
        }
        $Item2 = New-Object PSObject -Property $hash2
        $PunteggiGiornata = $PunteggiGiornata + $Item2
    }

    foreach ($match in $GiornataDaAnalizzare) {
        $Team1 = $match.Risultati.Split("|")[0]
        $Team2 = $match.Risultati.Split("|")[3]
        $ScoreTeam1 = [decimal]$match.Risultati.Split("|")[1]
        $ScoreTeam2 = [decimal]$match.Risultati.Split("|")[2]
        $DeltaScore = $ScoreTeam1 - $ScoreTeam2

        if($ScoreTeam1 -lt 66 -and $ScoreTeam2 -lt 66){
            $RealPointsTeam1 = 1
            $RealPointsTeam2 = 1
        }
        elseif($ScoreTeam1 -gt 65.5 -and $ScoreTeam2 -lt 66){
            $RealPointsTeam1 = 3
            $RealPointsTeam2 = 0
        }
        elseif($ScoreTeam1 -lt 66 -and $ScoreTeam2 -gt 65.5){
            $RealPointsTeam1 = 0
            $RealPointsTeam2 = 3
        }
        elseif ($DeltaScore -gt 2.5){
            $RealPointsTeam1 = 3
            $RealPointsTeam2 = 0
        }
        elseif ($DeltaScore -lt -2.5){
            $RealPointsTeam1 = 0
            $RealPointsTeam2 = 3
        }
        else{
            $RealPointsTeam1 = 1
            $RealPointsTeam2 = 1
        }

        $OtherOpponentsTeam1 = $PunteggiGiornata | Where-Object {$_.Squadra -ne $Team1}
        $OtherOpponentsTeam2 = $PunteggiGiornata | Where-Object {$_.Squadra -ne $Team2}

        [decimal]$TotExpectedScoreTeam1 = 0
        Foreach($OtherOpponent in $OtherOpponentsTeam1){
            $OtherOpponentScore = $OtherOpponent.FantaPunti
            $DeltaScore = $ScoreTeam1 - $OtherOpponentScore
            if($ScoreTeam1 -lt 66 -and $OtherOpponentScore -lt 66){
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 1
            }
            elseif($ScoreTeam1 -gt 65.5 -and $OtherOpponentScore -lt 66){
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 3
            }
            elseif($ScoreTeam1 -lt 66 -and $OtherOpponentScore -gt 65.5){
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 0
            }
            elseif ($DeltaScore -gt 2.5){
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 3
            }
            elseif ($DeltaScore -lt -2.5){
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 0
            }
            else{
                $TotExpectedScoreTeam1 = $TotExpectedScoreTeam1 + 1
            }   
        }
        $ExpectedScoreTeam1 = $TotExpectedScoreTeam1 / ($TeamsList.Count -1)

        [decimal]$TotExpectedScoreTeam2 = 0
        Foreach($OtherOpponent in $OtherOpponentsTeam2){
            $OtherOpponentScore = $OtherOpponent.FantaPunti
            $DeltaScore = $ScoreTeam2 - $OtherOpponentScore
            if($ScoreTeam2 -lt 66 -and $OtherOpponentScore -lt 66){
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 1
            }
            elseif($ScoreTeam2 -gt 65.5 -and $OtherOpponentScore -lt 66){
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 3
            }
            elseif($ScoreTeam2 -lt 66 -and $OtherOpponentScore -gt 65.5){
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 0
            }
            elseif ($DeltaScore -gt 2.5){
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 3
            }
            elseif ($DeltaScore -lt -2.5){
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 0
            }
            else{
                $TotExpectedScoreTeam2 = $TotExpectedScoreTeam2 + 1
            }   
        }
        $ExpectedScoreTeam2 = $TotExpectedScoreTeam2 / ($TeamsList.Count -1)

        Foreach($line in $FinalRank){
            if($line.Squadra -eq $Team1){
                $line.RealScore += $RealPointsTeam1
                $line.FantaPunti += $ScoreTeam1
                $line.ExpectedScore += $ExpectedScoreTeam1
            }
            if($line.Squadra -eq $Team2){
                $line.RealScore += $RealPointsTeam2
                $line.FantaPunti += $ScoreTeam2
                $line.ExpectedScore += $ExpectedScoreTeam2
            }
        }
    }
}

Write-Host "---------------------------------------" -ForegroundColor Green
Write-Host "Step 2: Computing Rank and Results Data - Completed!" -ForegroundColor Green
Write-Host "---------------------------------------" -ForegroundColor Green

$FinalRank = $FinalRank | Sort-Object -Property RealScore, FantaPunti -Descending
$FinalRank
