Clear-Host

import-module importexcel

$Path = ".\Recap_Fanta.csv"
$Report = Import-csv -path $Path -Encoding utf8

#Env Variables
$TotDays = (Get-ChildItem $Path | Select-Object -Property FullName, Name, @{n='PropertyCount'; e={@((Import-Csv -LiteralPath $_.FullName | Select-Object -First 1).psobject.Properties).Count}}).PropertyCount - 1
$FinalRank = @()
$j = 0
$Tot = $Report.count

Foreach ($line in $Report){
    $j ++
    $TotRealPoints = 0
    $TotExpectedPoints = 0
    $TotOwnPoints = 0
    $Team = $line.Team
    Write-Host "Calculating the Expect Score of the team $($Team) - [$($j)/$($Tot)]" -ForegroundColor Yellow
    $i = $Null
    for ($i = 1; $i -lt $TotDays + 1; $i++){
        #Find the opponent of the analyzed team
        $Giornata = 'Giornata ' + $i
        $OwnData = $line.$Giornata
        $OwnData = $OwnData.split("|")
        $OpponentName = $OwnData[0]
        $OwnScore = $OwnData[1]
        $OpponentFullData = $Report | Select-Object Team, $Giornata | Where-Object {$_.Team -eq $OpponentName}
        $OpponentData = $OpponentFullData.$Giornata
        $OpponentData = $OpponentData.split("|")
        $OpponentScore = $OpponentData[1]
        #Compute the real score of the analyzed team
        $DeltaScore = $OwnScore - $OpponentScore
        if($OwnScore -lt 66 -and $OpponentScore -lt 66){
            $RealPoints = 1
        }
        elseif($OwnScore -gt 65.5 -and $OpponentScore -lt 66){
            $RealPoints = 3
        }
        elseif($OwnScore -lt 66 -and $OpponentScore -gt 65.5){
            $RealPoints = 0
        }
        elseif ($DeltaScore -gt 2.5){
            $RealPoints = 3
        }
        elseif ($DeltaScore -lt -2.5){
            $RealPoints = 0
        }
        else{
            $RealPoints = 1
        }
        $TotRealPoints = $TotRealPoints + $RealPoints
        $TotOwnPoints = $TotOwnPoints + $OwnScore
        #Compute the expected score of the analyzed team
        $OtherTeams = $Report | Select-Object Team | Where-Object {$_.Team -ne $Team}
        $ExpectedScoreArray = @()
        foreach ($OtherTeam in $OtherTeams){
            $OtherTeamFullData = $Report | Select-Object Team, $Giornata | Where-Object {$_.Team -eq $OtherTeam.Team}
            $OtherTeamData = $OtherTeamFullData.$Giornata
            $OtherTeamData = $OtherTeamData.split("|")
            $OtherTeamScore = $OtherTeamData[1]
            $DeltaOTScore = $OwnScore - $OtherTeamScore
            if($OwnScore -lt 66 -and $OtherTeamScore -lt 66){
                $ExpectedOTPoints = 1
            }
            elseif($OwnScore -gt 65.5 -and $OtherTeamScore -lt 66){
                $ExpectedOTPoints = 3
            }
            elseif($OwnScore -lt 66 -and $OtherTeamScore -gt 65.5){
                $ExpectedOTPoints = 0
            }
            elseif ($DeltaOTScore -gt 2.5){
                $ExpectedOTPoints = 3
            }
            elseif ($DeltaOTScore -lt -2.5){
                $ExpectedOTPoints = 0
            }
            else{
                $ExpectedOTPoints = 1
            }
            $ExpectedScoreArray = $ExpectedScoreArray + $ExpectedOTPoints
        }
        $ExpectedPoints = ($ExpectedScoreArray | Measure-Object -Average).Average
        $TotExpectedPoints = $TotExpectedPoints + $ExpectedPoints
    }
    $hash = [ordered]@{
        Team = $Team
        RealScore = $TotRealPoints
        ExpectedScore = $TotExpectedPoints
        TotalPoints = $TotOwnPoints
    }
    $Item = New-Object psobject -Property $hash
    $FinalRank = $FinalRank + $Item
}

$FinalRank | Select-Object Team, RealScore, TotalPoints | Sort-Object -Descending -Property RealScore | Export-Excel -Path .\SFC_Ranks.xlsx -WorksheetName "Real Rank" -Title "Real Rank" -TitleBold -TableName "Real_Rank" -TableStyle Medium28
$FinalRank | Select-Object Team, ExpectedScore, TotalPoints | Sort-Object -Descending -Property ExpectedScore | Export-Excel -Path .\SFC_Ranks.xlsx -WorksheetName "Expected Rank" -Title "Expected Rank" -TitleBold -TableName "Expected_Rank" -TableStyle Medium28