<#
    Transform-MagToSDS.ps1

    20 mei 2020 Paul Wiegmans
    naar een voorbeeld van Wim den Ronde, Eric Redegeld, Joppe van Daalen

    Dit script leest tussenresultaat van de schijf, 
    met de uit Magister opgehaalde gegevens,
    berekent de actieve teams, 
    en genereert bestanden voor School Data Sync.

    Dit is stap 2 in TeamSync v2 proces voor omzetting van gegevens van Magister 
    naar School Data Sync
    * Transformeren
    * Uitvoeren
#>
$stopwatch = [Diagnostics.Stopwatch]::StartNew()
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition

# dot-source een fast object module
. ($herePath + "\ObjectFast.ps1")

$teamnaam_prefix = ""
$maakklassenteams = "1"
$maaklesgroepenteams = "1"
$maakvakkenteams = "1"

Write-Host " "
Write-Host "Start..."
$inputPath = $herePath + "\data_in"
$tempPath = $herePath + "\data_temp"
$outputPath = $herePath + "\data_out"
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
$filename_excl_vak  = $inputPath + "\excl_vak.csv"
$filename_excl_lesgroep = $inputPath + "\excl_lesgroep.csv"

# Files TEMP
$filename_mag_leerling_xml = $tempPath + "\mag_leerling.clixml"
$filename_mag_docent_xml = $tempPath + "\mag_docent.clixml"
$filename_mag_vak_xml = $tempPath + "\mag_vak.clixml"

# Files OUT
$filename_School = $outputPath + "\School.csv"
$filename_Section = $outputPath + "\Section.csv"
$filename_Student = $outputPath + "\Student.csv"
$filename_StudentEnrollment = $outputPath + "\StudentEnrollment.csv"
$filename_Teacher = $outputPath + "\Teacher.csv"
$filename_TeacherRoster = $outputPath + "\TeacherRoster.csv"

# Lees instellingen uit bestand met key=value
$filename_settings = $herePath + "\teamsync.ini"
$settings = Get-Content $filename_settings | ConvertFrom-StringData
foreach ($key in $settings.Keys) {
    Set-Variable -Name $key -Value $settings.$key
}
<# uit Teamsync.ini
$teamnaam_prefix = $settings.teamnaam_prefix
#>
if (!$brin)  { Throw "BRIN is vereist"}
if (!$schoolnaam)  { Throw "schoolnaam is vereist"}
if (!$magisterUser)  { Throw "magisterUser is vereist"}
if (!$magisterPass)  { Throw "magisterPass is vereist"}
if (!$magisterUrl)  { Throw "magisterUrl is vereist"}
if (!$teamnaam_prefix)  { Throw "teamnaam_prefix is vereist"}
$teamnaam_prefix += " "  # teamnaam prefix wordt altijd gevolgd door een spatie

function ConvertTo-SISID([string]$Naam) {
    return $Naam.replace(' ','_')
}

# voorbereiden SDS formaat CSV bestanden
$school = @() # "SIS ID,Name"
$section =  @() # "SIS ID,School SIS ID,Section Name"
$student =  @() # "SIS ID,School SIS ID,Username"
$studentenrollment = @() # "Section SIS ID,SIS ID"
$teacher =  @() # "SIS ID,School SIS ID,Username,First Name,Last Name"
$teacherroster =  @() # "Section SIS ID,SIS ID"

$filter_excl_vak = @()
if (Test-Path $filename_excl_vak) {
    $filter_excl_vak = Get-Content -Path $filename_excl_vak 
}
$filter_excl_lesgroep = @()
if (Test-Path $filename_excl_lesgroep) {
    $filter_excl_lesgroep = Get-Content -Path $filename_excl_lesgroep 
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
# velden: Stamnr, Login, Roepnaam, Tussenv, Achternaam, Lesperiode, 
# Leerjaar, Klas, Studie, Profiel, Groepen, Vakken
$mag_doc = Import-Clixml -Path $filename_mag_docent_xml
# velden: Stamnr, Login, Roepnaam, Tussenv, Achternaam, Naam, Code, 
# Functie, Groepvakken, Klasvakken, Docentvakken
# velden van mag_doc[].Groepvakken:  Klas, Vakcode
$mag_vak = Import-Clixml -Path $filename_mag_vak_xml
# $mag_vak['Vakcode'] = 'VakOmschrijving'

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

# voorfilteren
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

################# LEERLINGEN TEAMS BEPALEN
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
        $team[$teamid].leerling += @($leerling.Login)
    }

    # verzamel de lesgroepen
    # een team voor elke lesgroep
    foreach ($groep in $leerling.groepen) {
        $teamnaam = $teamnaam_prefix + $groep
        $teamid = ConvertTo-SISID -Naam $teamnaam
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }
        $team[$teamid].lltal += 1
        $team[$teamid].leerling += @($leerling.Login)
    }

    # verzamel de vakken
    # een team voor elke vakklas

    foreach ($vak in $leerling.vakken) {
        $teamnaam = $teamnaam_prefix + "vakklas " + $leerling.klas + " " + $vak
        $teamid = ConvertTo-SISID -Naam $teamnaam        
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].lltal += 1
        $team[$teamid].leerling += @($leerling.Login)
    }

    Write-Progress -Activity "Teams bepalen" -status `
        "Leerling $teller van $($mag_leer.count)" -PercentComplete ($leerlingprocent * $teller++)
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
        $team[$teamid].docent += @($docent.Login)    # Login In plaats van Stamnr/code
    
        #maak team voor het vak
        $teamnaam = $teamnaam_prefix + "vakklas " + $groepvak.Klas + " " + $groepvak.Vakcode
        $teamid = ConvertTo-SISID -Naam $teamnaam
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].doctal += 1
        $team[$teamid].docent += @($docent.Login)    # Login In plaats van Stamnr/code
    }

    # verzamelen klasvakken
    # LET OP: Normaliter wordt dit niet gedaan. Controleer nut.
    foreach ($klasvak in $docent.klasvakken) {
        #maak team voor klasvak
        $teamnaam = $teamnaam_prefix + "KV " + $klasvak
        $teamid = ConvertTo-SISID -Naam $teamnaam
        if ($team.Keys -notcontains $teamid) {
            $team[$teamid] = New-Team -Naam $teamnaam -ID $teamid
        }        
        $team[$teamid].doctal += 1
        $team[$teamid].docent += @($docent.Login)    # Login In plaats van Stamnr/code
    }

    # verzamel docentvakken
    # LET OP: Normaliter wordt dit niet gedaan. Controleer nut.
    #foreach ($vak in $docent.docentvakken) { }

    Write-Progress -Activity "Teams bepalen" -status `
        "Docent $teller van $($mag_doc.count)" -PercentComplete ($docentprocent * $teller++)
}
Write-Progress -Activity "Teams bepalen" -status "Docent" -Completed

# Ik maak van de associatieve array een lijst
$team = $team.Values

################# ACTIEVE TEAMS BEPALEN
# We willen alleen teams waarin zowel leerlingen als docent lid van zijn.
# Controleer op geldige leden voor elk team.
Write-Host "  Teams totaal          :" $team.count

$team0doc = $team | Where-Object {$_.doctal -eq 0}
$team0ll = $team | Where-Object {$_.lltal -eq 0}
$teamactief = $team | Where-Object {($_.lltal -gt 0) -and ($_.doctal -gt 0)}


Write-Host "  Teams actief          :" $teamactief.count 
Write-host "  Teams zonder leerling :" $team0ll.count 
Write-Host "  Teams zonder docent   :" $team0doc.count

$teamactief | Out-GridView
$team0ll | Out-GridView
$team0doc | Out-GridView
exit 61

################# AFWERKING EN UITVOER
# Hier gaan we de uiteindelijke bestanden aanmaken die we aan SDS voeren. 



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
$studentenrollment  = $studentenrollment | Sort-Object 'Section SIS ID' 
$teacher = $teacher | Sort-Object 'SIS ID'
$teacherroster = $teacherroster | Sort-Object 'Section SIS ID' 

# Alles opslaan
$school | Export-Csv -Path $filename_School -Encoding UTF8 -NoTypeInformation
$section | Export-Csv -Path $filename_Section -Encoding UTF8 -NoTypeInformation
$student | Export-Csv -Path $filename_Student -Encoding UTF8 -NoTypeInformation
$studentenrollment | Export-Csv -Path $filename_StudentEnrollment -Encoding UTF8 -NoTypeInformation
$teacher | Export-Csv -Path $filename_Teacher -Encoding UTF8 -NoTypeInformation
$teacherroster | Export-Csv -Path $filename_TeacherRoster -Encoding UTF8 -NoTypeInformation



$stopwatch.Stop()
Write-Host "Uitvoer klaar (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
