<#
    TeamSync-JPT.ps1

	2020-05-17 Paul Wiegmans SVOK
	
    Script dat leerling,docent,lesgroep,klas en roostergegevens ophaalt uit Magister
	en CSV-bestanden produceert voor upload naar School Data Sync.
    Naar een voorbeeld van Fons Vitae

#>
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

$brin = $null
$schoolnaam = $null
$magisterUser = $null
$magisterPass = $null
$magisterUrl = $null
$teamnaam_prefix = ""
$maakklassenteams = "1"
$maaklesgroepenteams = "1"
$maakvakkenteams = "1"

$jaarlaag_heeft_lesgroepen = "3", "4", "5", "6"  # er wordt alleen voor deze jaarlagen gezocht naar lesgroepen

Write-Host " "
Write-Host "Start..."
$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
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
$filename_t_leerling = $tempPath + "\leerling.csv"
$filename_t_docent = $tempPath + "\docent.csv"
$filename_t_groep = $tempPath + "\groep.csv"

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
<#
$brin = $settings.brin
$schoolnaam = $settings.schoolnaam
$magisterUser = $settings.magisterUser
$magisterPass = $settings.magisterPass
$magisterUrl = $settings.magisterUrl
$teamnaam_prefix = $settings.teamnaam_prefix
#>
if (!$brin)  { Throw "BRIN is vereist"}
if (!$schoolnaam)  { Throw "schoolnaam is vereist"}
if (!$magisterUser)  { Throw "magisterUser is vereist"}
if (!$magisterPass)  { Throw "magisterPass is vereist"}
if (!$magisterUrl)  { Throw "magisterUrl is vereist"}
if (!$teamnaam_prefix)  { Throw "teamnaam_prefix is vereist"}

function Invoke-Webclient($url) {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    $feed = [xml]$wc.downloadstring($url)
    return $feed.Response.Data    
}

function ADFunction ($Url = $magisterUrl, $Function, $SessionToken, $stamnr = $null) {
    if ($stamnr) {
        return Invoke-Webclient -Url ($Url + "?library=ADFuncties&function=" + 
            $Function + "&SessionToken=" + $SessionToken + "&LesPeriode=&StamNr=" + $stamnr + "&Type=XML")
    } else {
        return Invoke-Webclient -Url ($Url + "?library=ADFuncties&function=" + 
            $Function + "&SessionToken=" + $SessionToken + "&LesPeriode=&Type=XML")
    }
}

# voorbereiden SDS formaat CSV bestanden
$school = @() # "SIS ID,Name"
$section =  @() # "SIS ID,School SIS ID,Section Name"
$student =  @() # "SIS ID,School SIS ID,Username"
$studentenrollment = @() # "Section SIS ID,SIS ID"
$teacher =  @() # "SIS ID,School SIS ID,Username,First Name,Last Name"
$teacherroster =  @() # "Section SIS ID,SIS ID"

$team = @()
$ruwegroepen = @()

$filter_excl_vak = @()
if (Test-Path $filename_excl_vak) {
    $filter_excl_vak = Get-Content -Path $filename_excl_vak 
}
$filter_excl_lesgroep = @()
if (Test-Path $filename_excl_lesgroep) {
    $filter_excl_lesgroep = Get-Content -Path $filename_excl_lesgroep 
}

# haal sessiontoken
$MyToken = ""
$GetToken_URL = $magisterUrl + "?Library=Algemeen&Function=Login&UserName=" + 
$magisterUser + "&Password=" + $magisterPass + "&Type=XML"
$feed = [xml](new-object system.net.webclient).downloadstring($GetToken_URL)
$MyToken = $feed.response.SessionToken

################# VERZAMEL LEERLINGEN
# Ophalen leerlingdata, selecteer attributen, en bewaar hele tabel
Write-Host "Ophalen leerlingen..."
$data = ADFunction -U $magisterUrl -Ses $MyToken -F "GetActiveStudents"
$leerlingen = $data.Leerlingen.Leerling | Select-Object `
    stamnr_str,achternaam,tussenv,roepnaam,'loginaccount.naam',klas,studie,profiel.code
$leerlingen | Export-Csv -Path $filename_t_leerling -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Leerlingen           :" $leerlingen.count

# voorfilteren
if (Test-Path $filename_excl_studie) {
    $filter_excl_studie = $(Get-Content -Path $filename_excl_studie) -join '|'
    $leerlingen = $leerlingen | Where-Object {$_.Studie -notmatch $filter_excl_studie}
    Write-Host "Leerlingen na uitsluitend filteren studie :" $leerlingen.count
}
if (Test-Path $filename_incl_studie) {
    $filter_incl_studie = $(Get-Content -Path $filename_incl_studie) -join '|'
    $leerlingen = $leerlingen | Where-Object {$_.Studie -match $filter_incl_studie}
    Write-Host "Leerlingen na insluitend filteren studie :" $leerlingen.count
}
if (Test-Path $filename_excl_klas) {
    $filter_excl_klas = $(Get-Content -Path $filename_excl_klas) -join '|'
    $leerlingen = $leerlingen | Where-Object {$_.Klas -notmatch $filter_excl_klas}
    Write-Host "Leerlingen na uitsluitend filteren klas  :" $leerlingen.count
}
if (Test-Path $filename_incl_klas) {
    $filter_incl_klas = $(Get-Content -Path $filename_incl_klas) -join '|'
    $leerlingen = $leerlingen | Where-Object {$_.Klas -match $filter_incl_klas}
    Write-Host "Leerlingen na insluitend filteren klas   :" $leerlingen.count
}

$teller = 0
$leerlingprocent = 100 / $leerlingen.count
foreach ($leerling in $leerlingen) {
    $stamnr = $leerling.Stamnr_str
    $nieuwteam = @()

    # verzamel de stamklassen
    # een team voor elke klas
    if ($maakklassenteams -ne "0") {
        $ruwegroepen += $leerling.Klas
        $klas = $teamnaam_prefix + $leerling.Klas
        $nieuwteam += $klas

        $stenrec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $stenrec.'Section SIS ID' = $klas.replace(' ', '')
        $stenrec.'SIS ID' = $stamnr
        $studentenrollment += $stenrec
    }

    # verzamel de lesgroepen
    # een team voor elke lesgroep
    if ($maaklesgroepenteams -ne "0") {
        $leerjaar = $leerling.Klas[0]
        $data = ADFunction -U $magisterUrl -Ses $MyToken -F "GetLeerlingGroepen"-st $stamnr
        foreach ($node in $data.vakken.vak) {
            $ruwegroepen += $node.groep
            if ($filter_excl_lesgroep -notcontains $node.groep) {
                # filtervoorbeeld
                # if ($node.groep -ne "6ventlC") { # filtervoorbeeld 2
                $lesgroep = $teamnaam_prefix + $node.groep
                $nieuwteam += $lesgroep

                $stenrec = 1 | Select-Object 'Section SIS ID','SIS ID'
                $stenrec.'Section SIS ID' = $lesgroep.replace(' ', '')
                $stenrec.'SIS ID' = $stamnr
                $studentenrollment += $stenrec
            }
        }
    }

    # verzamel de vakken
    # een team voor elk vak 
    if ($maakvakkenteams -ne "0") {
        $data = ADFunction -U $magisterUrl -Ses $MyToken -F "GetLeerlingVakken" -st $stamnr
        foreach ($node in $data.vakken.vak) {
            $ruwegroepen += $node.vak
            if ($filter_excl_vak -notcontains $node.vak) {
                $vak = $teamnaam_prefix + $node.vak
                $nieuwteam += $vak

                $stenrec = 1 | Select-Object 'Section SIS ID','SIS ID'
                $stenrec.'Section SIS ID' = $vak.replace(' ', '')
                $stenrec.'SIS ID' = $stamnr
                $studentenrollment += $stenrec
            }
        }
    }

    # verzamel unieke teams
    $compare = compare-object -referenceobject $team -differenceobject $nieuwteam
    $compare | foreach-object {
        if ($_.sideindicator -eq "=>") {
            $team += $_.inputobject
        }
    }

    # Verzamel studenten
    $studrec = 1 | Select-Object 'SIS ID','School SIS ID',Username
    $studrec.'SIS ID' = $stamnr
    $studrec.'School SIS ID' = $brin
    $studrec.Username = $leerling.'loginaccount.naam'
    $student += $studrec

    Write-Progress -Activity "Magister data verwerken" -status `
        "Leerling $teller van $($leerlingen.count)" -PercentComplete ($leerlingprocent * $teller++)
}
Write-Progress -Activity "Magister data verwerken" -status "Leerling" -Completed

$team = $team | Sort-Object -Unique
Write-Host "Aanmeldingen         :" ($studentenrollment.count - 1) # minus kopregel

################# VERZAMEL DOCENTEN
Write-Host "Ophalen docenten..."
$data = ADFunction -U $magisterUrl -Ses $MyToken -F "GetActiveEmpoyees"  
$docenten = $data.Personeelsleden.Personeelslid | Select-Object `
    stamnr_str,achternaam,tussenv,roepnaam,loginaccount.naam,code,Functie.Omschr

# JPT: Om onbekende redenen staan sommige personeelsleden dubbel erin. 
# Met hun voornaam in 'loginaccount.naam' . Filter ze eruit.
$docenten = $docenten | Where-Object {$_.code -eq $_.'loginaccount.naam'}
$docenten | Export-Csv -Path $filename_t_docent -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Docenten ongefilterd :" $docenten.count

if (Test-Path $filename_excl_docent) {
    $filter_excl_docent = $(Get-Content -Path $filename_excl_docent) -join '|'
    $docenten = $docenten | Where-Object {$_.Code -notmatch $filter_excl_docent}
    Write-Host "Docenten na uitsluitend filteren docent :" $docenten.count
}

if (Test-Path $filename_incl_docent) {
    $filter_incl_docent = $(Get-Content -Path $filename_incl_docent) -join '|'
    $docenten = $docenten | Where-Object {$_.Code -match $filter_incl_docent}
    Write-Host "Docenten na insluitend filteren docent :" $docenten.count
}

$teller = 0
$docentprocent = 100 / $docenten.count
foreach ($user in $docenten ) {
    $nieuwteam = @()
    $stamnr = $user.'Code'  # worden codes herbruikt?
    $docentnr = $user.'stamnr_str'
    $voornaam = $user.'Roepnaam'

    if ($user.'Tussenv' -ne '') {
        $achternaam = $user.'Tussenv' + " " + $user.'Achternaam'
    } else {
        $achternaam = $user.'Achternaam'
    }

    # verzamel groepen per docent
    $data = ADFunction -U $magisterUrl -Ses $MyToken -F "GetPersoneelgroepVakken" -St $docentnr
    foreach ($dkv in $data.Lessen.Les) {
        $klasvak = $teamnaam_prefix + $dkv.'klas'
        $nieuwteam += $klasvak

        $terorec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $terorec.'Section SIS ID' = $klasvak.replace(' ', '')
        $terorec.'SIS ID' = $stamnr
        $teacherroster += $terorec
    }

    # verzamel unieke teams
    $compare = compare-object -referenceobject $team -differenceobject $nieuwteam
    $compare | foreach-object {
        if ($_.sideindicator -eq "=>") {
            $team += $_.inputobject
        }
    }

    # Voeg docent toe aan lijst, indien nog niet toegevoegd
    if ($teacher.'SIS ID' -notcontains $stamnr) {
        $tearec = 1 | Select-Object 'SIS ID','School SIS ID',Username,'First Name','Last Name'
        $tearec.'SIS ID' = $stamnr
        $tearec.'School SIS ID' = $brin
        $tearec.Username = $user.'Code'
        $tearec.'First Name' = $voornaam
        $tearec.'Last Name' = $achternaam
        $teacher += $tearec
    }
    Write-Progress -Activity "Magister uitlezen" -status `
        "Docent $teller van $($docenten.count)" -PercentComplete ($docentprocent * $teller++)
}
Write-Progress -Activity "Magister uitlezen" -status "Docent" -Completed

################# TEAMS
# We willen alleen teams waarin zowel leerlingen als docent lid van zijn.
# Controleer op geldige leden voor elk team.
Write-Host "Verzamelen actieve teams ..."
Write-Host "Teams                :" $team.count
$team = $team | Sort-Object -Unique
$team = $team | Where-Object {$_ -in $teacherroster.'Section SIS ID'} | Where-Object {$_ -in $studentenrollment.'Section SIS ID'}
Write-Host "Teams actief         :" $team.count

$vakprocent = 100 / $team.count
$teller = 0
foreach ($tm in $team) {
    $secrec = 1 | Select-Object 'SIS ID','School SIS ID','Section Name'
    $secrec.'SIS ID' = $tm.replace(' ', '')
    $secrec.'School SIS ID' =  $brin
    $secrec.'Section Name' = $tm
    $section += $secrec

    Write-Progress -Activity "Verzamelen teams" -status `
        "Team $teller van $($team.count)" -PercentComplete ($vakprocent * $teller++)
}
Write-Progress -Activity "Verzamelen teams" -status "Vak" -Completed

# we willen alleen de docenten die in een actief team zitten
Write-Host "Docentgroepen        :" $teacherroster.count
$teacherroster = $teacherroster | Where-Object {$_.'Section SIS ID' -in $team}
Write-Host "Docentgroepen actief :" $teacherroster.count
Write-Host "Docenten             :" $teacher.count
$teacher = $teacher | Where-Object {$_.'SIS ID' -in $teacherroster.'SIS ID'}
Write-Host "Docenten actief      :" $teacher.count

################# AFWERKING

$ruwegroepen = $ruwegroepen | Sort-Object -Unique
Write-Host "Groepen uniek        :" $ruwegroepen.count
$ruwegroepen | Out-File -FilePath $filename_t_groep -Encoding UTF8

# Maak een school
$schoolrec = 1 | Select-Object 'SIS ID',Name
$schoolrec.'SIS ID' = $brin
$schoolrec.Name = $schoolnaam
$school += $schoolrec

Write-Host "Oprollen en aftaaien ..."
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
Write-Host "Klaar! (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
