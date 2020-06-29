<#
    .SYNOPSIS

    TeamSync script deel 2; koppeling tussen Magister en School Data Sync.

    .DESCRIPTION

    TeamSync is een koppeling tussen Magister en School Data Sync.
    TeamSync script deel 2 (transformeren en uitvoeren)
    bepaalt actieve teams en genereert CSV-bestanden ten behoeve van 
    School Data Sync.

    Versie 20200629
    Auteur Paul Wiegmans (p.wiegmans@svok.nl)

    naar een voorbeeld door Wim den Ronde, Eric Redegeld, Joppe van Daalen

    .PARAMETER Inifilename

    bepaalt de bestandsnaam van het INI-bestand, waarin benodigde parameters 
    worden gelezen, relatief ten opzichte van het pad van dit script.

    .LINK

    https://github.com/sikkepitje/teamsync

    .NOTES

    TO DO 
    * situatie voor Magister zonder SSO : gebruik Emailaddress i.p.v. Login

#>
[CmdletBinding()]
param (
    [Parameter(
        HelpMessage="Geef de naam van de te gebruiken INI-file, bij verstek 'TeamSync.ini'"
    )]
    [Alias('Inibestandsnaam')]
    [String]  $Inifilename = "TeamSync.ini"
)
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
# scriptnaam in venstertitel
$host.ui.RawUI.WindowTitle = (Split-Path -Leaf $MyInvocation.MyCommand.Path).replace(".ps1","")
Start-Transcript -path $MyInvocation.MyCommand.Path.replace(".ps1",".log")

$teamnaam_prefix = ""
$maakklassenteams = "1"
$datainvoermap = "data_in"
$datakladmap = "data_temp"
$datauitvoermap = "data_uit"
$useemail = "0"

# Lees instellingen uit bestand met key=value
$filename_settings = $herePath + "\" + $Inifilename
Write-Host "INI-bestand: " $filename_settings
$settings = Get-Content $filename_settings | ConvertFrom-StringData
foreach ($key in $settings.Keys) {
    Set-Variable -Name $key -Value $settings.$key
}
<# $teamnaam_prefix = $settings.teamnaam_prefix #>
if (!$brin)  { Throw "BRIN is vereist"}
if (!$schoolnaam)  { Throw "schoolnaam is vereist"}
if (!$magisterUser)  { Throw "magisterUser is vereist"}
if (!$magisterPass)  { Throw "magisterPass is vereist"}
if (!$magisterUrl)  { Throw "magisterUrl is vereist"}
if (!$teamnaam_prefix)  { Throw "teamnaam_prefix is vereist"}
$teamnaam_prefix += " "  # teamnaam prefix wordt altijd gevolgd door een spatie
$useemail = $useemail -ne "0"  # maak echte boolean
Write-Host "Schoolnaam:" $schoolnaam

# datamappen
$inputPath = $herePath + "\$datainvoermap"
$tempPath = $herePath + "\$datakladmap"
$outputPath = $herePath + "\$datauitvoermap"
Write-Host "datainvoermap :" $inputPath
Write-Host "datakladmap   :" $tempPath
Write-Host "datauitvoermap:" $outputPath

New-Item -path $inputPath -ItemType Directory -ea:Silentlycontinue
New-Item -path $tempPath -ItemType Directory -ea:Silentlycontinue
New-Item -path $outputPath -ItemType Directory -ea:Silentlycontinue

# Files IN
$filename_excl_docent = $inputPath + "\excl_docent.csv"
$filename_incl_docent = $inputPath + "\incl_docent.csv"
$filename_excl_klas  = $inputPath + "\excl_klas.csv"
$filename_incl_klas  = $inputPath + "\incl_klas.csv"
$filename_excl_studie   = $inputPath + "\excl_studie.csv"
$filename_incl_studie   = $inputPath + "\incl_studie.csv"
$filename_incl_locatie  = $inputPath + "\incl_locatie.csv"

# Files TEMP
$filename_mag_leerling_xml = $tempPath + "\mag_leerling.clixml"
$filename_mag_docent_xml = $tempPath + "\mag_docent.clixml"
$filename_mag_vak_xml = $tempPath + "\mag_vak.clixml"
$filename_t_teamactief = $tempPath + "\teamactief.csv"
$filename_t_team0ll = $tempPath + "\team0ll.csv"
$filename_t_team0doc = $tempPath + "\team0doc.csv"
$filename_log = $tempPath + "\teamsync.log"

# Files OUT
$filename_School = $outputPath + "\School.csv"
$filename_Section = $outputPath + "\Section.csv"
$filename_Student = $outputPath + "\Student.csv"
$filename_StudentEnrollment = $outputPath + "\StudentEnrollment.csv"
$filename_Teacher = $outputPath + "\Teacher.csv"
$filename_TeacherRoster = $outputPath + "\TeacherRoster.csv"

# controleer vereiste bestanden
if (!(Test-Path -Path $filename_mag_leerling_xml)) {  Throw "Vereist bestand ontbreekt:" + $filename_mag_leerling_xml }
if (!(Test-Path -Path $filename_mag_docent_xml)) {  Throw "Vereist bestand ontbreekt:" + $filename_mag_docent_xml }
if (!(Test-Path -Path $filename_mag_vak_xml)) {  Throw "Vereist bestand ontbreekt:" + $filename_mag_vak_xml }


function ConvertTo-SISID([string]$Naam) {
    return $Naam.replace(' ','_')
}

$team = @{}
# associatieve array van records:
#   Naam         : weergavenaam
#   lltal        : aantal leerlingen
#   doctal       : aantal docenten
#   leerling     : lijst van leerlingid's 
#   docent       : lijst van docentid's
# index is teamid

function New-Team($naam, $id)
{
    # maak een nieuw teamrecord met $naam, geindexeerd op Teamid (dit wordt 'Section SIS ID')
    $rec = 1 | Select-Object Id, Naam, lltal, leerling, doctal, docent
    $rec.Naam = $naam       # weergavenaam van team
    $rec.Id = $id
    $rec.lltal = 0          # aantal leerlingen
    $rec.leerling = @()     # lijst met leerlingid
    $rec.doctal = 0         # aantal docenten
    $rec.docent = @()       # lijst met docentid
    return $rec
}

################# LEES TUSSENDATA
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

Write-Host "Leerlingen           :" $mag_leer.count
Write-Host "Docenten             :" $mag_doc.count
Write-Host "Vakken               :" $mag_vak.count

if ($mag_doc.count -eq 0) {
    Throw "Geen docenten!"
}

# filters toepassen op leerlingen
if (Test-Path $filename_excl_studie) {
    $filter_excl_studie = $(Get-Content -Path $filename_excl_studie) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Studie -notmatch $filter_excl_studie}
    Write-Host "L na uitsluiting studie :" $mag_leer.count
}
if (Test-Path $filename_incl_studie) {
    $filter_incl_studie = $(Get-Content -Path $filename_incl_studie) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Studie -match $filter_incl_studie}
    Write-Host "L na insluiting studie  :" $mag_leer.count
}
if (Test-Path $filename_excl_klas) {
    $filter_excl_klas = $(Get-Content -Path $filename_excl_klas) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Klas -notmatch $filter_excl_klas}
    Write-Host "L na uitsluiting klas   :" $mag_leer.count
}
if (Test-Path $filename_incl_klas) {
    $filter_incl_klas = $(Get-Content -Path $filename_incl_klas) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Klas -match $filter_incl_klas}
    Write-Host "L na insluiting klas    :" $mag_leer.count
}
if (Test-Path $filename_incl_locatie) {
    $filter_incl_locatie = $(Get-Content -Path $filename_incl_locatie) -join '|'
    $mag_leerling = $mag_leerling | Where-Object {$_.Locatie -match $filter_incl_locatie}
    Write-Host "L na insluiting locatie:" $mag_leerling.count
}

# filter toepassen op docent
if (Test-Path $filename_excl_docent) {
    $filter_excl_docent = $(Get-Content -Path $filename_excl_docent) -join '|'
    $mag_doc = $mag_doc | Where-Object {$_.Code -notmatch $filter_excl_docent}
    Write-Host "D na uitsluiting docent :" $mag_doc.count
}
if (Test-Path $filename_incl_docent) {
    $filter_incl_docent = $(Get-Content -Path $filename_incl_docent) -join '|'
    $mag_doc = $mag_doc | Where-Object {$_.Code -match $filter_incl_docent}
    Write-Host "D na insluiting docent  :" $mag_doc.count
}

################# LEERLINGEN -> TEAMS BEPALEN
$teller = 0
$leerlingprocent = 100 / $mag_leer.count
foreach ($leerling in $mag_leer) {

    # verzamel de stamklassen
    if ($maakklassenteams -ne "0") {
        $teamnaam = $teamnaam_prefix + $leerling.Klas
        $teamid = ConvertTo-SISID -Naam ($teamnaam)
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }
        $team[$teamid].lltal += 1
        $team[$teamid].leerling += @($leerling.Id)
    }

    # corrigeer lege groepen artefact uit CliXML
    if ($leerling.groepen.ToString() -eq "") {
        $leerling.groepen = $null
    }
    # verzamel de lesgroepen
    # een team voor elke lesgroep
    if ($leerling.groepen) {
        foreach ($groep in $leerling.groepen) {
            $teamnaam = $teamnaam_prefix + $groep
            $teamid = ConvertTo-SISID -Naam $teamnaam
            if ($team.Keys -notcontains $teamid) {
                $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
            }
            $team[$teamid].lltal += 1
            $team[$teamid].leerling += @($leerling.Id)
        }
    }

    # verzamel de vakken
    # een team voor elke vakklas

    foreach ($vak in $leerling.vakken) {
        $teamnaam = $teamnaam_prefix + $leerling.klas + " " + $vak
        $teamid = ConvertTo-SISID -Naam $teamnaam        
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].lltal += 1
        $team[$teamid].leerling += @($leerling.Id)
    }

    if (!($teller++ % 20)) {
        Write-Progress -PercentComplete ($leerlingprocent * $teller) `
            -Activity "Teams bepalen" -status "Leerling $teller van $($mag_leer.count)" 
    }
}
Write-Progress -Activity "Teams bepalen" -status "Leerling" -Completed

################# DOCENTEN TEAMs BEPALEN
$teller = 0
$docentprocent = 100 / $mag_doc.count
foreach ($docent in $mag_doc ) {

    # verzamel groepen per docent
    foreach ($groepvak in $docent.groepvakken) {
        # velden van mag_doc[].Groepvakken:  Klas, Vakcode
        
        # maak team voor de klas
        $teamnaam = $teamnaam_prefix + $groepvak.Klas
        $teamid = ConvertTo-SISID -Naam $teamnaam        
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].doctal += 1
        if ($team[$teamid].docent -notcontains $docent.Id) {
            $team[$teamid].docent += @($docent.Id)
        }
    
        #maak team voor het vak
        $teamnaam = $teamnaam_prefix + $groepvak.Klas + " " + $groepvak.Vakcode
        $teamid = ConvertTo-SISID -Naam $teamnaam
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].doctal += 1
        if ($team[$teamid].docent -notcontains $docent.Id) {
            $team[$teamid].docent += @($docent.Id)
        }
    }

    # corrigeer lege klasvakken artefact uit CliXML
    if ($docent.Klasvakken.Tostring() -eq "") {
        $docent.Klasvakken = $null
    }
    # verzamelen klasvakken
    # LET OP: Normaliter wordt dit niet gedaan. Controleer nut.
    foreach ($klasvak in $docent.klasvakken) {
        #maak team voor klasvak
        $teamnaam = $teamnaam_prefix + $klasvak
        $teamid = ConvertTo-SISID -Naam $teamnaam
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].doctal += 1  
        if ($team[$teamid].docent -notcontains $docent.Id) {
            $team[$teamid].docent += @($docent.Id)
        }      
    }

    # verzamel docentvakken
    # Normaliter wordt dit niet gedaan.
    #foreach ($vak in $docent.docentvakken) { }

    if (!($teller++ % 10)) {
        Write-Progress -PercentComplete ($docentprocent * $teller) `
            -Activity "Teams bepalen" -Status "Docent $teller van $($mag_doc.count)" 
    }
}
Write-Progress -Activity "Teams bepalen" -status "Docent" -Completed

################# ACTIEVE TEAMS BEPALEN
# associatieve array omzetten naar simpele lijst
$team = $team.Values

Write-Host "  Omschrijvingen toevoegen.."

# Voeg vakomschrijving toe aan teamnaam
foreach ($t in $team) {
    # Bepaal vakcode met laatste deel van de naam (klas/vak of cluster)
    # voorbeeld: "Naamprefix 4h.biol3"
    # voorbeeld: "Naamprefix 1b ne"
    $vak = $t.naam.split(". ")[-1]  # split lesgroep op spatie of punt, vak staat achter punt
    # verwijder cijfers op het eind
    While (($vak[-1] -match "^\d+$") -and ($vak.length -gt 2)) {
        $vak = $vak.substring(0, $vak.length - 1)
    }
    $omschrijving = $mag_vak[$vak]
    # indien geen omschrijving gevonden, knip letters van vak totdat omschrijving is gevonden of totdat 2 letters overblijven
    while ((!$omschrijving.length) -and ($vak.length -gt 2)) { 
        $vak = $vak.substring(0, $vak.length - 1)
        $omschrijving = $mag_vak[$vak]
    }
    # indien omschrijving is gevonden, voeg toe aan naam
    if ($omschrijving.length) {
        $t.naam += " " + $omschrijving
    }
}

Write-Host "  Teams totaal          :" $team.count
# Actieve teams bevatten zowel leerlingen als docenten.
# Splits de teams in 3 lijsten: actief, zonder leerlingen, zonder docenten.
$teamactief = $team | Where-Object {($_.lltal -gt 0) -and ($_.doctal -gt 0)}
$team0doc = $team | Where-Object {$_.doctal -eq 0}
$team0ll = $team | Where-Object {$_.lltal -eq 0}

Write-Host "  Teams actief          :" $teamactief.count 
Write-host "  Teams zonder leerling :" $team0ll.count 
Write-Host "  Teams zonder docent   :" $team0doc.count

# door mensen leesbare CSVs uitvoeren ter controle
$hteamactief = $teamactief | Sort-Object Id | Select-Object Id, Naam, Doctal,`
    @{Name = 'Docenten'; Expression = {($_.docent | Sort-Object) -join ","}},`
    Lltal, `
    @{Name = 'Leerlingen'; Expression = {($_.leerling | Sort-Object) -join ","}}
$hteamactief | Export-Csv -Path $filename_t_teamactief -NoTypeInformation -Encoding UTF8

$hteam0ll = $team0ll | Sort-Object Id | Select-Object Id, Naam, Doctal,`
    @{Name = 'Docenten'; Expression = {($_.docent | Sort-Object) -join ","}},`
    Lltal, `
    @{Name = 'Leerlingen'; Expression = {($_.leerling | Sort-Object) -join ","}}
$hteam0ll | Export-Csv -Path $filename_t_team0ll -NoTypeInformation -Encoding UTF8

$hteam0doc = $team0doc | Sort-Object Id | Select-Object Id, Naam, Doctal,`
    @{Name = 'Docenten'; Expression = {($_.docent | Sort-Object) -join ","}},`
    Lltal, `
    @{Name = 'Leerlingen'; Expression = {($_.leerling | Sort-Object) -join ","}}
$hteam0doc | Export-Csv -Path $filename_t_team0doc -NoTypeInformation -Encoding UTF8

# voor visuele controle
#$hteamactief | Out-GridView
#$hteam0ll | Out-GridView
#$hteam0doc | Out-GridView

################# UITVOER
Write-Host "Uitvoer..."
# Ik maak de uiteindelijke bestanden aan, die naar School Data Sync worden geupload.

# voorbereiden SDS formaat CSV bestanden
$school = @()               # 'SIS ID','Name'    bijv "20MH","Jac P. Thijsse College"
$section =  @()             # 'SIS ID','School SIS ID','Section Name'  bijv 'SDS_1920_1A_ak','20MH','SDS 1920 1A ak'
$student =  @()             # 'SIS ID','School SIS ID','Username'   bijv '10935','20MH','10935'
$studentenrollment = @()    # 'Section SIS ID','SIS ID'   bijv 'SDS_1920_1A','11210'
$teacher =  @()             # 'SIS ID','School SIS ID','Username','First Name','Last Name'  bijv "ABl","20MH","ABl","Aaaaaa","Bbbbb"
$teacherroster =  @()       # 'Section SIS ID','SIS ID'  bijv "SDS_1920_1A","DZn"

# actieve leerlingen actieve docenten tabel 
$teamdoc = @()
$teamleer = @()
# maak docentopzoektabel
$hashdoc = @{}
$mag_doc | ForEach-Object { $hashdoc[$_.Id] = $_}

$teller = 0
$teamprocent = 100 / $teamactief.count

foreach ($t in $teamactief) {
    $rec = 1 | Select-Object 'SIS ID','School SIS ID','Section Name'
    $rec.'SIS ID' = $t.naam 
    $rec.'School SIS ID' = $brin
    $rec.'Section Name' = $t.naam 
    $section += $rec

    foreach ($leerling in $t.leerling) {
        $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $rec.'Section SIS ID' = $t.naam
        $rec.'SIS ID' = $leerling
        $studentenrollment += $rec
        if ($teamleer -notcontains $leerling) {
            $teamleer += $leerling
        }
    }

    foreach ($docent in $t.docent) {
        $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $rec.'Section SIS ID' = $t.naam
        $rec.'SIS ID' = $docent
        $teacherroster += $rec
        if ($teamdoc -notcontains $docent) {
            $teamdoc += $docent
        }
    }
    if (!($teller++ % 10)) {
        Write-Progress -PercentComplete ($teamprocent * $teller) `
            -Activity "Lijsten genereren" -Status "Team $teller van $($teamactief.count)" 
    }
}
Write-Progress -Activity "Lijsten genereren" -Completed

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
    $teacher += $rec
}
foreach ($leer in $teamleer) {
    $rec = 1 | Select-Object 'SIS ID','School SIS ID','Username'
    $rec.'SIS ID' = $leer
    $rec.'School SIS ID' = $brin
    $rec.'Username' = $leer
    $student += $rec
}

# Maak een school
$schoolrec = 1 | Select-Object 'SIS ID',Name
$schoolrec.'SIS ID' = $brin
$schoolrec.Name = $schoolnaam
$school += $schoolrec

Write-Host "School               :" $school.count
Write-Host "Student              :" $student.count
Write-Host "Studentenrollment    :" $Studentenrollment.count
Write-Host "Teacher              :" $teacher.count
Write-Host "Teacherroster        :" $teacherroster.count
Write-Host "Section              :" $section.count

# Sorteer de teams voor de mooi
$section = $section | Sort-Object 'SIS ID'
$student = $student | Sort-Object 'SIS ID'
$studentenrollment = $studentenrollment | Sort-Object 'Section SIS ID','SIS ID'
$teacher = $teacher | Sort-Object 'SIS ID'
$teacherroster = $teacherroster | Sort-Object 'Section SIS ID','SIS ID'

# Alles opslaan
$school | Export-Csv -Path $filename_School -Encoding UTF8 -NoTypeInformation
$section | Export-Csv -Path $filename_Section -Encoding UTF8 -NoTypeInformation
$student | Export-Csv -Path $filename_Student -Encoding UTF8 -NoTypeInformation
$studentenrollment | Export-Csv -Path $filename_StudentEnrollment -Encoding UTF8 -NoTypeInformation
$teacher | Export-Csv -Path $filename_Teacher -Encoding UTF8 -NoTypeInformation
$teacherroster | Export-Csv -Path $filename_TeacherRoster -Encoding UTF8 -NoTypeInformation

$stopwatch.Stop()
Write-Host "Klaar (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
Stop-Transcript
