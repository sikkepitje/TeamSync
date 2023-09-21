<#
    .SYNOPSIS

    TeamSync script Export-SchoolDataSync.ps1; koppeling tussen Magister en School Data Sync.

    .DESCRIPTION

    TeamSync is een koppeling tussen Magister en School Data Sync.
    TeamSync script Export-SchoolDataSync.ps1 (transformeren en uitvoeren)
    bepaalt actieve teams en genereert CSV-bestanden ten behoeve van 
    School Data Sync.

    Versie 20230921
    Auteur Paul Wiegmans (p.wiegmans@svok.nl)

    naar een voorbeeld door Wim den Ronde, Eric Redegeld, Joppe van Daalen

    .PARAMETER Inifilename

    bepaalt de bestandsnaam van het configuratiebestand, relatief ten opzichte van het pad van dit script.

    .INPUTS
    Diverse; Zie READM.adoc
    .OUTPUTS
    Diverse; Zie READM.adoc
    .LINK

    https://github.com/sikkepitje/teamsync

    .NOTES

    TO DO 
    * situatie voor Magister zonder SSO: gebruik Emailaddress i.p.v. Login

#>
[CmdletBinding()]
param (
    [Parameter(
        HelpMessage="Geef de naam van de te gebruiken INI-file, bij verstek 'Export-SchoolDataSync.ini'"
    )]
    [Alias('Inifile','Inibestandsnaam','Config','Configfile','Configuratiebestand')]
    [String]  $Inifilename = "Export-SchoolDataSync.ini"
)
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
# scriptnaam in venstertitel
$selfpath_base = $MyInvocation.MyCommand.Path.replace(".ps1","")  # compleet pad zonder extensie
$host.ui.RawUI.WindowTitle = Split-Path -Leaf $selfpath_base
$logCountLimit  = 7
$selfpath = $MyInvocation.MyCommand.Path
$selfdir  = Split-Path -Parent $selfpath
$selfname  = Split-Path -Leaf $selfpath 
$selfbasename  = [System.IO.Path]::GetFileNameWithoutExtension($selfpath)
$logBaseFilename = "$selfdir\Log\$selfbasename"
$currentLogFilename = "$logBaseFilename.log"

# variabelen initialisatie
$importdatamap = "ImportData"
$exportfiltermap = "ExportFilter"
$exportkladmap = "ExportKlad"
$exportdatamap = "Exportdata"
$brin = $null
$schoolnaam = $null
$teamid_prefix = ""
$teamnaam_prefix = ""
$teamnaam_suffix = ""
$maakklassenteams = "1"
$logtag = "INIT" 
$toonresultaat = "0"
$bon_match_docentlesgroep_aan_leerlingklas = "0"
$docenten_per_team_limiet = "0"  # maximum toegestaan aantal docenten per team. 0 betekent geen limiet

#region Functies
function PreviousLogFilename($Number) {
    return ("$logBaseFilename.{0:d2}.log" -f $Number)
}

function LogRotate() {
    # Keep 9 logs, delete oldest, rename the rest
    Write-Host "Rotating the logs..." -ForegroundColor Cyan
    New-Item -Path "$selfdir" -ItemType Directory -Name "Log" -ErrorAction SilentlyContinue | Out-Null
    Remove-Item -Path (PreviousLogFilename -Number $logCountLimit) -Force -Confirm:$False -ea:SilentlyContinue 
    ($logCountLimit)..1 | ForEach-Object {
        $oud = PreviousLogFilename -Number $_
        $nieuw = PreviousLogFilename -Number ($_ + 1)
        #Write-Host "  Renaming ($oud) to ($nieuw)" -ForegroundColor cyan
        Rename-Item -Path $oud -NewName $nieuw -ea:SilentlyContinue
    }
    Rename-Item -Path $currentLogFilename -NewName (PreviousLogFilename -Number 1) -ea:SilentlyContinue
}

Function Write-Log {
    Param (
        [Parameter(Position=0)][Alias('Message')][string]$Tekst="`n"
    )

    #Write-Host $Tekst
    $log = "$(Get-Date -f "yyyy-MM-ddTHH:mm:ss.fff") [$logtag] $tekst"
    $log | Out-File -FilePath $currentLogFilename -Append
    Write-Host $log
}

$illegal_characters = "[^\S]|[\~\""\#\%\&\*\:\<\>\?\/\\{\|}\.\[\]]"
$safe_character = "_"
function ConvertTo-SISID([string]$Naam) {
    return $Naam -replace $illegal_characters, $safe_character
    # https://support.microsoft.com/en-us/office/invalid-file-names-and-file-types-in-onedrive-and-sharepoint-64883a5d-228e-48f5-b3d2-eb39e07630fa
}
Function ConvertTo-ASCII([string]$naam) {
    # Verwijder alle tekens anders dan a-z, A-Z, 0-9, underscore en punt
    $naam -replace '[^a-zA-Z0-9_.]', ''
}

#endregion Functies

# Start hoofdprogramma
LogRotate
Write-Log ""
Write-Log ("START " + $MyInvocation.MyCommand.Name)
Try {
    # Lees instellingen uit bestand met key=value
    $filename_settings = $herePath + "\" + $Inifilename
    Write-Log ("Configuratiebestand: " + $filename_settings)
    $settings = Get-Content $filename_settings -Encoding UTF8 | ConvertFrom-StringData
    foreach ($key in $settings.Keys) {
        Set-Variable -Name $key -Value $settings.$key -Scope global
        Write-Log ("Configuratieparameter: " + $key + "=" + $settings.$key)
    }
    <# $teamid_prefix = $settings.teamid_prefix #>
    if (!$brin)  { Throw "Configuratieparameter 'BRIN' is vereist"}
    if (!$schoolnaam)  { Throw "Configuratieparameter 'schoolnaam' is vereist"}
    if (!$teamid_prefix)  { Throw "Configuratieparameter 'teamid_prefix' is vereist"}
    $teamid_prefix = $teamid_prefix.trim() + " "
    if ($teamnaam_prefix) {
        $teamnaam_prefix = $teamnaam_prefix.trim() + " "
    }
    if ($teamnaam_suffix) {
        $teamnaam_suffix = " " + $teamnaam_suffix.trim()
    }
    $toonresultaat = $toonresultaat -ne "0"  # maak boolean
    $bon_match_docentlesgroep_aan_leerlingklas = $bon_match_docentlesgroep_aan_leerlingklas -ne "0" # maak boolean
    [int]$docenten_per_team_limiet = $docenten_per_team_limiet # maak integer

    $logtag = $teamid_prefix
    $host.ui.RawUI.WindowTitle = ((Split-Path -Leaf $MyInvocation.MyCommand.Path) -replace ".ps1") + " " + $logtag
    Write-Log ("Schoolnaam     : " + $schoolnaam)

    # datamappen
    $importPath         = "$herePath\$importdatamap"
    $filterPath         = "$herePath\$exportfiltermap"
    $tempPath           = "$herePath\$exportkladmap"
    $outputPath         = "$herePath\$exportdatamap"
    Write-Log ("ImportDataMap    : " + $importPath)
    Write-Log ("ExportFilterMap  : " + $filterPath)
    Write-Log ("ExportKladMap    : " + $tempPath)
    Write-Log ("ExportDataMap    : " + $outputPath)

    New-Item -path $tempPath -ItemType Directory -ea:Silentlycontinue
    New-Item -path $outputPath -ItemType Directory -ea:Silentlycontinue

    # Import
    $filename_mag_leerling_xml  = $importPath + "\magister_leerling.clixml"
    $filename_mag_docent_xml    = $importPath + "\magister_docent.clixml"
    $filename_mag_vak_xml       = $importPath + "\magister_vak.clixml"
    
    # Filters
    $filename_excl_docent       = $filterPath + "\excl_docent.csv"
    $filename_incl_docent       = $filterPath + "\incl_docent.csv"
    $filename_excl_klas         = $filterPath + "\excl_klas.csv"
    $filename_incl_klas         = $filterPath + "\incl_klas.csv"
    $filename_excl_studie       = $filterPath + "\excl_studie.csv"
    $filename_incl_studie       = $filterPath + "\incl_studie.csv"
    $filename_incl_locatie      = $filterPath + "\incl_locatie.csv"
    $filename_excl_teamnaam     = $filterPath + "\excl_teamnaam.csv"
    $filename_incl_teamnaam     = $filterPath + "\incl_teamnaam.csv"

    # Kladbestanden
    $hteamid                    = $teamid_prefix.trim() -replace(" ","_")
    $filename_t_hteamfull       = $tempPath + "\hteamfull_"   + $hteamid + ".csv"
    $filename_t_hteamactief     = $tempPath + "\hteamactief_" + $hteamid + ".csv"
    $filename_t_hteam0ll        = $tempPath + "\hteam0ll_"    + $hteamid + ".csv"
    $filename_t_hteam0doc       = $tempPath + "\hteam0doc_"   + $hteamid + ".csv"

    # Files OUT
    $filename_School            = $outputPath + "\School.csv"
    $filename_Section           = $outputPath + "\Section.csv"
    $filename_Student           = $outputPath + "\Student.csv"
    $filename_StudentEnrollment = $outputPath + "\StudentEnrollment.csv"
    $filename_Teacher           = $outputPath + "\Teacher.csv"
    $filename_TeacherRoster     = $outputPath + "\TeacherRoster.csv"

    # controleer vereiste bestanden
    if (!(Test-Path -Path $filename_mag_leerling_xml)) {  Throw "Vereist bestand ontbreekt: " + $filename_mag_leerling_xml }
    if (!(Test-Path -Path $filename_mag_docent_xml)) {  Throw "Vereist bestand ontbreekt: " + $filename_mag_docent_xml }
    if (!(Test-Path -Path $filename_mag_vak_xml)) {  Throw "Vereist bestand ontbreekt: " + $filename_mag_vak_xml }


    function ConvertTo-Teamnaam([string]$Naam) {
        return ($teamnaam_prefix + $naam + $teamnaam_suffix)
    }

    ################# LEES DATA van Import-Magister
    $mag_leer = Import-Clixml -Path $filename_mag_leerling_xml
    # velden: Stamnr, Id, Login, Roepnaam, Tussenv, Achternaam, Lesperiode, 
    # Leerjaar, Klas, Studie, Profiel, Groepen, Vakken, Email, Locatie
    $mag_doc = Import-Clixml -Path $filename_mag_docent_xml
    # velden: Stamnr, Id, Login, Roepnaam, Tussenv, Achternaam, Naam, Code, 
    # Functie, Groepvakken, Klasvakken, Docentvakken, Locatie
    # velden van mag_doc[].Groepvakken:  Klas, Vakcode
    $mag_vak = Import-Clixml -Path $filename_mag_vak_xml
    # $mag_vak['Vakcode'] = 'VakOmschrijving'

    # Zet om in hashtabel, kapitaliseer alle woorden in vakomschrijving, behalve en en and
    $vakoms = @{}
    foreach ($kvp in $mag_vak.GetEnumerator()) {
        $vakoms[$kvp.key] = (Get-Culture).TextInfo.ToTitleCase($kvp.value).replace(" En "," en ").replace(" And "," and ")
    }
    $mag_vak = $vakoms

    # sorteer voor de mooi
    foreach ($docent in $mag_doc) {
        $docent.Groepvakken = $docent.Groepvakken | Sort-Object -Property "Klas"
        $docent.Klasvakken = $docent.Klasvakken | Sort-Object
        $docent.Docentvakken = $docent.Docentvakken | Sort-Object
    }

    Write-Log ("Leerlingen     : " + $mag_leer.count)
    Write-Log ("Docenten       : " + $mag_doc.count)
    Write-Log ("Vakken         : " + $mag_vak.count)

    if ($mag_doc.count -eq 0) {
        Throw "Er zijn nul docenten. Er is niets te doen"
    }

    # filters toepassen op leerlingen
    if (Test-Path $filename_excl_studie) {
        $filter_excl_studie = $(Get-Content -Path $filename_excl_studie -Encoding UTF8) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Studie -notmatch $filter_excl_studie}
        Write-Log ("L na uitsluiting studie: " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_studie) {
        $filter_incl_studie = $(Get-Content -Path $filename_incl_studie -Encoding UTF8) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Studie -match $filter_incl_studie}
        Write-Log ("L na insluiting studie : " + $mag_leer.count)
    }
    if (Test-Path $filename_excl_klas) {
        $filter_excl_klas = $(Get-Content -Path $filename_excl_klas -Encoding UTF8) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Klas -notmatch $filter_excl_klas}
        Write-Log ("L na uitsluiting klas  : " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_klas) {
        $filter_incl_klas = $(Get-Content -Path $filename_incl_klas -Encoding UTF8) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Klas -match $filter_incl_klas}
        Write-Log ("L na insluiting klas   : " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_locatie) {
        $filter_incl_locatie = $(Get-Content -Path $filename_incl_locatie -Encoding UTF8) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Locatie -match $filter_incl_locatie}
        Write-Log ("L na insluiting locatie: " + $mag_leer.count)
    }

    # filter toepassen op docent
    if (Test-Path $filename_excl_docent) {
        $filter_excl_docent = $(Get-Content -Path $filename_excl_docent -Encoding UTF8) -join '|'
        $mag_doc = $mag_doc | Where-Object {$_.Id -notmatch $filter_excl_docent}
        Write-Log ("D na uitsluiting docent: " + $mag_doc.count)
    }
    if (Test-Path $filename_incl_docent) {
        $filter_incl_docent = $(Get-Content -Path $filename_incl_docent -Encoding UTF8) -join '|'
        $mag_doc = $mag_doc | Where-Object {$_.Id -match $filter_incl_docent}
        Write-Log ("D na insluiting docent : " + $mag_doc.count)
    }

    ################# Teams bepalen aan de hand van docent groepvakken (a la WootsSyncReadPhase.ps1)
    $team = @{}
    # associatieve array van records:
    #   Id           :
    #   groep        : groep van docent.groepvak[n]
    #   vak          : vakcode van docent.groepvak[n]
    #   VakOms       : omschrijving bij vakcode
    #   Doctal       : aantal docenten
    #   Docent       : lijst van docentid's
    #   Lltal        : aantal leerlingen
    #   TypeL        : type groep : "lgrp" of "klas" oftewel waarom deze leerling hierin zit
    #   Leerling     : lijst van leerlingid's 
    # index is groep

    function New-Team($id, $groep, $vak)
    {
        # maak een nieuw teamrecord met $naam, geindexeerd op Teamid (dit wordt 'Section SIS ID')
        return [PSCustomObject]@{
            Id      = $id
            Naam    = $id
            Groep   = $groep
            Vak     = $vak
            VakOms  = $mag_vak[$vak]
            Doctal  = 0
            Docent  = @()
            Lltal   = 0
            TypeL   = ""
            Leerling = @()
        }
    }

    $activity = "Teams voor docenten maken ..."
    Write-Log ($activity)
    $teller = 0
    $docentprocent = 100 / [Math]::Max($mag_doc.count, 1)
    foreach ($docent in $mag_doc) {
        foreach ($groepvak in $docent.groepvakken) {
            if ($groepvak.vakcode.length -ge 1) {  # skip null-items
                $groep = $groepvak.Klas
                $vak = $groepvak.Vakcode
                $id = "{0}@{1}" -f ($groep, $vak) # tijdelijk identifier uniek voor de combinatie van groep en vak
                if ($team.Keys -notcontains $id) {
                    $tm = New-Team -id $id -groep $groep -vak $vak
                    $team[$id] = $tm
                } else {
                    $tm = $team[$id]
                }
                if ($tm.Docent -notcontains $docent.id) {
                    $tm.Docent += $docent.id
                    $tm.Doctal += 1
                }
            }
        }
        if (!(++$teller % 10)) {
            Write-Progress -PercentComplete ($docentprocent * $teller) `
                -Activity $activity -Status "Docent $teller van $($mag_doc.count)" 
        }
    }
    Write-Progress -Activity $activity -status "Docent" -Completed

    # maak opzoektabel groep->team
    $groepteams =@{} 
    $team.Values | foreach {
        if ($groepteams.Keys -notcontains $_.Groep) {
            $grpteam = [PSCustomObject]@{
                Groep = $_.Groep
                Aantal = 0
                Teams = @()
            }
            $groepteams[$_.Groep] = $grpteam
        } else {
            $grpteam = $groepteams[$_.Groep]
        }
        $grpteam.Teams += $_.naam
        $grpteam.Aantal += 1        
    }
    
    function ToevoegenAan-Team ($Leerling, $Groep, $Label) {
        $teams = $groepteams[$groep].Teams  # zoek bijbehorend teamindex(en) in opzoektabel
        if ($teams) {
            foreach ($samegroup in $teams) {
                $tm = $team[$samegroup] 
                if ($tm) {
                    $tm.Leerling += $leerling.id 
                    $tm.Lltal += 1
                    if (!$tm.TypeL.contains($Label)){
                        $tm.TypeL += "$Label"
                    }
                } else { 
                    Write-Warning "Team niet gevonden voor $groep"
                }
            }
        }
    }

    $activity = "Teams voor leerlingen maken ..."
    Write-Log ($activity)
    $teller = 0
    $leerlingprocent = 100 / [Math]::Max($mag_leer.count, 1)
    foreach ($leerling in $mag_leer) {
        ToevoegenAan-Team -Leerling $leerling -Groep $leerling.klas -Label "klas"
        foreach ($groep in $leerling.groepen) {
            ToevoegenAan-Team -Leerling $leerling -Groep $groep -Label "groep"
        }

        if (!(++$teller % 50)) {
            Write-Progress -PercentComplete ($leerlingprocent * $teller) `
                -Activity $activity -status "Leerling $teller van $($mag_leer.count)" 
        }
    }
    Write-Progress -Activity $activity -status "Leerling" -Completed

    $team = $team.Values | Sort-Object id

    # Teamnaam en Id  bepalen volgens gewenst formaat
    foreach ($t in $team) {
        $t.Naam = "{0}{1} {2} {3}" -f ($teamnaam_prefix, $t.Groep, $t.VakOms, $teamnaam_suffix) 
        $t.Id = ConvertTo-ASCII (("{0}{1} {2}" -f ($teamid_prefix, $t.Groep, $t.Vak)) -replace $illegal_characters, $safe_character)
    }
    
    Write-Log ("Team Totaal: {0} " -f $team.count)

    # Filteren op teamnaam
    if (Test-Path $filename_excl_teamnaam) {
        $filter_excl_teamnaam = $(Get-Content -Path $filename_excl_teamnaam -Encoding UTF8) -join '|'
        $team = $team | Where-Object {$_.Naam -notmatch $filter_excl_teamnaam}
        Write-Log ("Team na uitsluiting teamnaam: " + $team.count)
    }
    if (Test-Path $filename_incl_teamnaam) {
        $filter_incl_teamnaam = $(Get-Content -Path $filename_incl_teamnaam -Encoding UTF8) -join '|'
        $team = $team | Where-Object {$_.Naam -match $filter_incl_teamnaam}
        Write-Log ("Team na insluiting teamnaam : " + $team.count)
    }
    # filter op aantal docenten
    if ($docenten_per_team_limiet -gt 0) {
        $team = $team | Where-Object {$_.doctal -le $docenten_per_team_limiet}
        Write-Log ("Team na toepassen docentenlimiet : " + $team.count)
    }

    # Maak makkelijk leesbare lijsten om te helpen bij foutzoeken en fijnafstelling. 
    $hteam = $team | Select-Object Id, Naam, Groep,Vak,VakOms,
        @{Name = 'Aantal_docenten'; Expression = {$_.Doctal}},
        TypeL, 
        @{Name = 'Aantal_leerlingen'; Expression = {$_.Lltal}},
        @{Name = 'Docenten'; Expression = {($_.docent | Sort-Object) -join ","}},
        @{Name = 'Leerlingen'; Expression = {($_.leerling | Sort-Object) -join ","}}

    # Splits de teams in 3 lijsten: actief, zonder leerlingen, zonder docenten.
    $hteamactief = $hteam | Where-Object {($_.Aantal_leerlingen -gt 0) -and ($_.Aantal_docenten -gt 0)}
    $hteam0ll = $hteam | Where-Object {$_.Aantal_docenten -eq 0}
    $hteam0doc = $hteam | Where-Object {$_.Aantal_leerlingen -eq 0}
    
    Write-Log ("Teams actief          : " + $hteamactief.count )
    Write-Log ("Teams zonder leerling : " + $hteam0ll.count )
    Write-Log ("Teams zonder docent   : " + $hteam0doc.count)

    # Bewaar human readable lijsten in CSV om te helpen bij foutzoeken en fijnafstelling. 
    $hteam | Export-Csv -Path $filename_t_hteamfull -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    $hteamactief | Export-Csv -Path $filename_t_hteamactief -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    $hteam0doc | Export-CSV -Path $filename_t_hteam0doc -NoTypeInformation -Encoding UTF8 -Delimiter ";"
    $hteam0ll | Export-Csv -Path $filename_t_hteam0ll -NoTypeInformation -Encoding UTF8 -Delimiter ";"

    # Bewaar actieve teams ook als clixml
    $hteamactief | Export-Clixml -Path ($filename_t_hteamactief + ".clixml")

    # voor visuele controle
    if ($toonresultaat) {
        $hteamactief | Out-GridView  # dit zijn de actieve teams
    }
    
    ################# UITVOER
    $activity = "School Data Sync CSV v1 lijsten samenstellen ..."
    Write-Log ($activity)
    # Ik maak de uiteindelijke bestanden aan, die naar School Data Sync worden geupload.

    # voorbereiden SDS formaat CSV bestanden
    $school = [System.Collections.ArrayList]@()               # 'SIS ID','Name'    bijv "20MH","Jac P. Thijsse College"
    $section =  [System.Collections.ArrayList]@()             # 'SIS ID','School SIS ID','Section Name'  bijv 'SDS_1920_1A_ak','20MH','SDS 1920 1A ak'
    $student =  [System.Collections.ArrayList]@()             # 'SIS ID','School SIS ID','Username'   bijv '10935','20MH','10935'
    $studentenrollment = [System.Collections.ArrayList]@()    # 'Section SIS ID','SIS ID'   bijv 'SDS_1920_1A','11210'
    $teacher =  [System.Collections.ArrayList]@()             # 'SIS ID','School SIS ID','Username','First Name','Last Name'  bijv "ABl","20MH","ABl","Aaaaaa","Bbbbb"
    $teacherroster =  [System.Collections.ArrayList]@()       # 'Section SIS ID','SIS ID'  bijv "SDS_1920_1A","DZn"

    # actieve leerlingen actieve docenten tabel 
    $teamdoc = [System.Collections.ArrayList]@()
    $teamleer = [System.Collections.ArrayList]@()
    # maak docentopzoektabel
    $hashdoc = @{}
    $mag_doc | ForEach-Object { $hashdoc[$_.Id] = $_}

    $teamactief = $team | Where-Object {($_.lltal -gt 0) -and ($_.doctal -gt 0)}

    $teller = 0
    $teamprocent = 100 / [Math]::Max($teamactief.count, 1)

    foreach ($t in $teamactief) {
        $rec = 1 | Select-Object 'SIS ID','School SIS ID','Section Name'
        $rec.'SIS ID' = $t.id 
        $rec.'School SIS ID' = $brin
        $rec.'Section Name' = $t.naam 
        $null = $section.Add($rec)

        foreach ($leerling in $t.leerling) {
            $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
            $rec.'Section SIS ID' = $t.id
            $rec.'SIS ID' = $leerling
            $null = $studentenrollment.Add($rec)
            if ($teamleer -notcontains $leerling) {
                $null = $teamleer.Add($leerling)
            }
        }

        foreach ($docent in $t.docent) {
            $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
            $rec.'Section SIS ID' = $t.id
            $rec.'SIS ID' = $docent
            $null = $teacherroster.Add($rec)
            if ($teamdoc -notcontains $docent) {
                $null = $teamdoc.Add($docent)
            }
        }
        if (!(++$teller % 10)) {
            Write-Progress -PercentComplete ($teamprocent * $teller) `
                -Activity $activity -Status "Team $teller van $($teamactief.count)" 
        }
    }
    Write-Progress -Activity $activity -Completed

    # actieve docenten opzoeken 
    foreach ($doc in $teamdoc) {
        $rec = 1 | Select-Object 'SIS ID','School SIS ID','Username','First Name','Last Name'
        $rec.'SIS ID' = $hashdoc[$doc].Id
        $rec.'School SIS ID' = $brin
        $rec.'Username' = $hashdoc[$doc].Id
        $rec.'First Name' = $hashdoc[$doc].Roepnaam
        if ($hashdoc[$doc].Tussenv -ne '') {
            $rec.'Last Name' = $hashdoc[$doc].Tussenv + " " + $hashdoc[$doc].Achternaam
        } else {
            $rec.'Last Name' = $hashdoc[$doc].Achternaam
        }
        $null = $teacher.Add($rec)
    }
    foreach ($leer in $teamleer) {
        $rec = 1 | Select-Object 'SIS ID','School SIS ID','Username'
        $rec.'SIS ID' = $leer
        $rec.'School SIS ID' = $brin
        $rec.'Username' = $leer
        $null = $student.Add($rec)
    }

    # Maak een school
    $schoolrec = 1 | Select-Object 'SIS ID',Name
    $schoolrec.'SIS ID' = $brin
    $schoolrec.Name = $schoolnaam
    $null = $school.Add($schoolrec)

    Write-Log ("School               : " + $school.count)
    Write-Log ("Student              : " + $student.count)
    Write-Log ("Studentenrollment    : " + $Studentenrollment.count)
    Write-Log ("Teacher              : " + $teacher.count)
    Write-Log ("Teacherroster        : " + $teacherroster.count)
    Write-Log ("Section              : " + $section.count)

    # Sorteer de teams voor de mooi
    $section = $section | Sort-Object 'SIS ID'
    $student = $student | Sort-Object 'SIS ID'
    $studentenrollment = $studentenrollment | Sort-Object 'Section SIS ID','SIS ID'
    $teacher = $teacher | Sort-Object 'SIS ID'
    $teacherroster = $teacherroster | Sort-Object 'Section SIS ID','SIS ID'

    # Alles opslaan
    Write-Log ("Lijsten voor School Data Sync opslaan ...")
    $school | Export-Csv -Path $filename_School -Encoding UTF8 -NoTypeInformation
    $section | Export-Csv -Path $filename_Section -Encoding UTF8 -NoTypeInformation
    $student | Export-Csv -Path $filename_Student -Encoding UTF8 -NoTypeInformation
    $studentenrollment | Export-Csv -Path $filename_StudentEnrollment -Encoding UTF8 -NoTypeInformation
    $teacher | Export-Csv -Path $filename_Teacher -Encoding UTF8 -NoTypeInformation
    $teacherroster | Export-Csv -Path $filename_TeacherRoster -Encoding UTF8 -NoTypeInformation

    $stopwatch.Stop()
    Write-Log ("Klaar in " + $stopwatch.Elapsed.Hours + " uur " + $stopwatch.Elapsed.Minutes + " minuten " + $stopwatch.Elapsed.Seconds + " seconden ")    
} 
Catch {

    $e = $_.Exception
    $line = $_.InvocationInfo.ScriptLineNumber
    $msg = $e.Message 
 
    "$(Get-Date -f "yyyy-MM-ddTHH:mm:ss:fff") [$logtag] caught exception: $msg at line $line" | Out-File -FilePath $currentLogFilename -Append
    Write-Error "Caught exception: $msg at line $line"    
    exit 1  
}
