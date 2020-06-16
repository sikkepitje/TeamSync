<#
    TeamSync-JPT.ps1

    20 mei 2020 Paul Wiegmans
    naar een voorbeeld van Eric Redegeld,
    naar een voorbeeld van Wim den Ronde

    Poging ophalen van gegeven via magister Webservices ADFuncties
    naar voorbeeld van Fons Vitae

    Opmerkingen:

    Toon studentengegevens:
    get-content ".\data\studenten.csv" | convertfrom-csv -delimiter ";" | out-gridview

    wijzigingen:
    * aanmaken vakklassen voor klassikale vakken: vakken in brugklas, (netl,entl,me, enz) in bovenbouw
    * bewaart ruwe magistergegevens in tijdelijke bestanden t.b.v. debuggen: llgroep, llvak, teamlid, docvak.

    Dit is een verouderde versie van TeamSync. Gebruik dit niet!
    Zie Ophalen-MagisterData.ps1 voor de verder ontwikkelde versie van dit script. 
    
#>
$stopwatch = [Diagnostics.Stopwatch]::StartNew()

$herePath = Split-Path -parent $MyInvocation.MyCommand.Definition
. ($herePath + "\ObjectFast.ps1")
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
$vakken_bovenbouw_klassikaal = "netl", "entl", "lo", "ckv", "maat", "WVStage", "LV"
$vakken_brugklas_klassikaal = "ne*","ak","aktvt","bi","en","eng","fa","gs","lo","me","mu","ne","ntc","pe","te","tot","up","mat","wi"
$vakken_onderbouw_klassikaal = "me","ne","eng","mus","art","bio","du","fa","geo","his","mat"
# Welke groepen zijn klassikaal voor welk jaar/studie/klas?

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
$filename_t_leerling = $tempPath + "\leerling.csv"
$filename_t_docent = $tempPath + "\docent.csv"
$filename_t_groep = $tempPath + "\groep.csv"
$filename_t_teamlid = $tempPath + "\teamlid.csv"
$filename_t_llgroep = $tempPath + "\llgroep.csv"
$filename_t_llvak = $tempPath + "\llvak.csv"
$filename_t_docvak = $tempPath + "\docvak.csv"
$filename_t_docklasvak = $tempPath + "\docklasvak.csv"
$filename_t_docgroepvak = $tempPath + "\docgroepvak.csv"

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
$teamnaam_prefix += " "  # teamnaam prefix wordt altijd gevolgd door een spatie
function Invoke-Webclient($url) {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    try {
        $feed = [xml]$wc.downloadstring($url)
    } catch {
        Throw "Invoke-Webclient: er ging iets mis"
    }
    return $feed.Response.Data    
}
function ADFunction ($Url = $magisterUrl, $Function, $SessionToken, $Stamnr = $null) {
    if ($stamnr) {
        return Invoke-Webclient -Url ($Url + "?library=ADFuncties&function=" + 
            $Function + "&SessionToken=" + $SessionToken + "&LesPeriode=&StamNr=" + $Stamnr + "&Type=XML")
    } else {
        return Invoke-Webclient -Url ($Url + "?library=ADFuncties&function=" + 
            $Function + "&SessionToken=" + $SessionToken + "&LesPeriode=&Type=XML")
    }
}
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
$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveStudents"
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

$llgroep = @()
$llvak = @()
$llklas = @()   
$teller = 0
$leerlingprocent = 100 / $leerlingen.count
foreach ($leerling in $leerlingen) {
    $stamnr = $leerling.Stamnr_str
    $klas = $leerling.Klas
    $leerjaar = $klas[0]
    $nieuwteam = @()

    # verzamel de stamklassen
    if ($maakklassenteams -ne "0") {
        $ruwegroepen += $leerling.Klas
        $teamnaam = ConvertTo-SISID -Naam ($teamnaam_prefix + $leerling.Klas)
        $nieuwteam += $teamnaam

        # Ik maak GEEN team voor elke klas
        $stenrec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $stenrec.'Section SIS ID' = $teamnaam
        $stenrec.'SIS ID' = $stamnr
        $studentenrollment += $stenrec
    }

    # verzamel de lesgroepen
    # een team voor elke lesgroep
    if (($maaklesgroepenteams -ne "0") -and ($leerjaar -in $jaarlaag_heeft_lesgroepen)) {
        $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetLeerlingGroepen" -Stamnr $stamnr
        foreach ($node in $data.vakken.vak) {
            $ruwegroepen += $node.groep

            $rec = 1 | Select-Object "Stamnr","groep"
            $rec.Stamnr = $node.Stamnr
            $rec.groep = $node.groep
            $llgroep += $rec

            if ($filter_excl_lesgroep -notcontains $node.groep) {
                $teamnaam = ConvertTo-SISID -Naam ($teamnaam_prefix + $node.groep)
                $nieuwteam += $teamnaam

                $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
                $rec.'Section SIS ID' = $teamnaam
                $rec.'SIS ID' = $stamnr
                $studentenrollment += $rec
            }
        }
    }

    # verzamel de vakken
    # een team voor elke vakklas

    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetLeerlingVakken" -Stamnr $stamnr
    foreach ($node in $data.vakken.vak) {
        $ruwegroepen += $node.vak
        $rec = 1 | Select-Object "Stamnr","Vak"
        $rec.Stamnr = $node.Stamnr
        $rec.Vak = $node.Vak
        $llvak += $rec

        if ($filter_excl_vak -notcontains $node.vak) {

            #$is_vakklas = (($leerjaar -eq "1") -and ($vakken_brugklas_klassikaal -contains $node.vak)) `
            #-or (($leerjaar -in "2","3") -and ($vakken_bovenbouw_klassikaal -contains $node.vak)) `
            #-or (($leerjaar -ge "4") -and ($vakken_bovenbouw_klassikaal -contains $node.vak))

            # Ik weet niet voor welk vak dit moet gebeuren. 
            # Laten we dit simpel houden voor nu: Maak een vakklas voor ELKE klas!
            $is_vakklas = $true
            if ($is_vakklas) {
                # als dit een klassikaal vak is, maak een "vakklas"
                $teamnaam = ConvertTo-SISID -Naam ($teamnaam_prefix + $leerling.klas + " " + $node.vak + " vakklas")
                $nieuwteam += $teamnaam
    
                $stenrec = 1 | Select-Object 'Section SIS ID','SIS ID'
                $stenrec.'Section SIS ID' = $teamnaam
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
    # verzamel unieke klassen
    if ($klas -notin $llklas) {
        $llklas += $klas
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

Write-Host "Sorteren..."
$team = $team | Sort-Object -Unique
Write-Host "Aanmeldingen         :" ($studentenrollment.count - 1) # minus kopregel

# tijdelijke gegevens opslaan
$llgroep = $llgroep | Sort-Object "Stamnr","groep"
$llgroep | Export-CSV -Path $filename_t_llgroep -Encoding UTF8 -NoTypeInformation
$llvak = $llvak | Sort-Object "Stamnr","Vak"
$llvak | Export-CSV -Path $filename_t_llvak -Encoding UTF8 -NoTypeInformation

################# VERZAMEL DOCENTEN
Write-Host "Ophalen docenten..."
$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveEmpoyees"  
$docenten = $data.Personeelsleden.Personeelslid | Select-Object `
    stamnr_str,achternaam,tussenv,roepnaam,loginaccount.naam,code,Functie.Omschr

# JPT: Om onbekende redenen staan sommige personeelsleden dubbel erin. 
# Met hun voornaam in 'loginaccount.naam' . Filter ze eruit.
$docenten = $docenten | Where-Object {$_.code -eq $_.'loginaccount.naam'}
$docenten | Export-Csv -Path $filename_t_docent -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Docenten ongefilterd :" $docenten.count
if ($docenten.count -eq 0) {
    Throw "Geen docenten ?? Stopt!"
}
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
$docvak = @()
$docklasvak = @()
$docgroepvak = @()
$docentprocent = 100 / $docenten.count
foreach ($docent in $docenten ) {
    $nieuwteam = @()
    $docentnr = $docent.stamnr_str
    $voornaam = $docent.Roepnaam

    if ($docent.Tussenv -ne '') {
        $achternaam = $docent.Tussenv + " " + $docent.Achternaam
    } else {
        $achternaam = $docent.Achternaam
    }

    # verzamel groepen per docent
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelGroepVakken" -Stamnr $docentnr
    foreach ($elem in $data.Lessen.Les) {
        # maak een klasteam voor deze docent
        $teamnaam = ConvertTo-SISID -Naam ($teamnaam_prefix + $elem.Klas )
        $nieuwteam += $teamnaam

        $rec = 1 | Select-Object "Code","Personeelslid_Stamnr","Klas","Vak_Vakcode","Vak_Omschrijving"
        $rec.Code = $docent.Code
        $rec.Personeelslid_Stamnr = $elem.'Personeelslid.Stamnr'
        $rec.Klas = $elem.Klas
        $rec.Vak_Vakcode = $elem.'Vak.Vakcode'
        $rec.Vak_Omschrijving = $elem.'Vak.Omschrijving'
        $docgroepvak += $rec
        
        $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
        $rec.'Section SIS ID' = $teamnaam
        $rec.'SIS ID' = $docent.Code
        $teacherroster += $rec

        if ($elem.Klas -in $llklas) {
            # maak een vakklas voor deze docent
            $teamnaam = ConvertTo-SISID -Naam ($teamnaam_prefix + $elem.Klas + " " + $elem.'Vak.Vakcode' + " vakklas")
            $nieuwteam += $teamnaam

            $rec = 1 | Select-Object 'Section SIS ID','SIS ID'
            $rec.'Section SIS ID' = $teamnaam
            $rec.'SIS ID' = $docent.Code
            $teacherroster += $rec

            # INFO : docvak
            $rec = 1 | Select-Object "Personeelslid_Stamnr","Klas","Vak_Vakcode","Vak_Omschrijving"
            $rec.Personeelslid_Stamnr = $elem.'Personeelslid.Stamnr'
            $rec.Klas = $elem.Klas
            $rec.Vak_Vakcode = $elem.'Vak.Vakcode'
            $rec.Vak_Omschrijving = $elem.'Vak.Omschrijving'
            $docvak += $rec           
        }
    }

    # verzamelen klasvakken
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelKlasVakken" -Stamnr $docentnr
    if ($data.Lessen.Les) {
        foreach ($elem in $data.Lessen.Les) {
            Write-Host "+PersKlasVak" 
            # INFO : docklasvak
            $rec = 1 | Select-Object "Code","Personeelslid_Stamnr","Klas_Lesgroep","Klas"
            $rec.Code = $docent.Code
            $rec.Personeelslid_Stamnr = $elem.'Personeelslid.Stamnr'
            $rec.Klas = $elem.Klas
            $rec.Klas_Lesgroep = $elem.Klas_Lesgroep
            $docklasvak += $rec
        }
    }

    # verzamel vakken per docent
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelVakken" -Stamnr $docentnr
    if ($data.Lessen.Les) {
        foreach ($elem in $data.Lessen.Les) {
            $rec = 1 | Select-Object "Code","Personeelslid_Stamnr","Vak_Vakcode","Vak_Omschrijving"
            $rec.Code = $docent.Code
            $rec.Personeelslid_Stamnr = $elem.'Personeelslid.Stamnr'
            $rec.Vak_Vakcode = $elem.'Vak.Vakcode'
            $rec.Vak_Omschrijving = $elem.'Vak.Omschrijving'
            $docvak += $rec
        }
    }

    # verzamel unieke teams
    $compare = compare-object -referenceobject $team -differenceobject $nieuwteam
    $compare | foreach-object {
        if ($_.sideindicator -eq "=>") {
            $team += $_.inputobject
        }
    }

    # Voeg docent toe aan lijst, indien nog niet toegevoegd
    if ($teacher.'SIS ID' -notcontains $docent.Code) {
        $tearec = 1 | Select-Object 'SIS ID','School SIS ID',Username,'First Name','Last Name'
        $tearec.'SIS ID' = $docent.Code
        $tearec.'School SIS ID' = $brin
        $tearec.Username = $docent.Code
        $tearec.'First Name' = $voornaam
        $tearec.'Last Name' = $achternaam
        $teacher += $tearec
    }
    Write-Progress -Activity "Magister uitlezen" -status `
        "Docent $teller van $($docenten.count)" -PercentComplete ($docentprocent * $teller++)
}
Write-Progress -Activity "Magister uitlezen" -status "Docent" -Completed

#$docenten | Out-GridView

$docvak = $docvak | Sort-Object "Code","Personeelslid_Stamnr","Klas","Vak_Vakcode"
#$docvak | Out-Gridview
$docvak | Export-Csv -Path $filename_t_docvak -Encoding UTF8 -NoTypeInformation

$docklasvak = $docklasvak | Sort-Object "Code","Personeelslid.Stamnr","Klas_Lesgroep","Klas"
#$docklasvak | Out-Gridview
$docklasvak | Export-Csv -Path $filename_t_docklasvak -Encoding UTF8 -NoTypeInformation

$docgroepvak = $docgroepvak | Sort-Object Code,"Personeelslid_Stamnr",Klas
#$docgroepvak | Out-Gridview
$docgroepvak | Export-Csv -Path $filename_t_docgroepvak -Encoding UTF8 -NoTypeInformation

################# TEAMS
# We willen alleen teams waarin zowel leerlingen als docent lid van zijn.
# Controleer op geldige leden voor elk team.
Write-Host "Verzamelen actieve teams ..."
Write-Host "Teams                :" $team.count
$team = $team | Sort-Object -Unique
#$team = $team | Where-Object {$_ -in $teacherroster.'Section SIS ID'} | Where-Object {$_ -in $studentenrollment.'Section SIS ID'}
Write-Host "Teams actief         :" $team.count

$vakprocent = 100 / $team.count
$teller = 0
foreach ($tm in $team) {
    $secrec = 1 | Select-Object 'SIS ID','School SIS ID','Section Name'
    $secrec.'SIS ID' = $tm
    $secrec.'School SIS ID' =  $brin
    $secrec.'Section Name' = $tm
    $section += $secrec

    Write-Progress -Activity "Verzamelen teams" -status `
        "Team $teller van $($team.count)" -PercentComplete ($vakprocent * $teller++)
}
Write-Progress -Activity "Verzamelen teams" -status "Vak" -Completed

# we willen alleen de docenten die in een actief team zitten WERKT NIET?
Write-Host "Docentgroepen        :" $teacherroster.count
#$teacherroster = $teacherroster | Where-Object {$_.'Section SIS ID' -in $team.'SIS ID'}
#Write-Host "Docentgroepen actief :" $teacherroster.count
Write-Host "Docenten             :" $teacher.count
#$teacher = $teacher | Where-Object {$_.'SIS ID' -in $teacherroster.'SIS ID'}
#Write-Host "Docenten actief      :" $teacher.count

################# AFWERKING EN UITVOER

$ruwegroepen = $ruwegroepen | Sort-Object -Unique
Write-Host "Groepen uniek        :" $ruwegroepen.count
$ruwegroepen | Out-File -FilePath $filename_t_groep -Encoding UTF8

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


################# CONTROLEWEERGAVE
# Ter controle: Verzamel een lijst van teams waar voor ieder team staat hoeveel leerlingen 
# en hoeveel docenten er lid van zijn.
# Voor eenvoud, gebruik 3 associatieve arrays, één voor bijhouden van leerlingen 
# die lid zijn van een team, één voor docentenaantallen, één voor docentcodes. Gebruik teamnaam als index. 

Write-Host "Weergave ter controle"
$teamleer = @{}
$teamdoc = @{}
$teamdoccode = @{}

Write-Host "  Teams tellen..."
foreach ($tm in $team) {
    $teamleer.Add($tm, 0)
    $teamdoc.Add($tm, 0)
    $teamdoccode.Add($tm, "")
}
Write-Host "  Leerlingen tellen..."
foreach ($sten in $studentenrollment) {
    $teamleer[$sten.'Section SIS ID'] += 1
}
Write-Host "  Docenten tellen..."
foreach ($tero in $teacherroster) {
    $teamdoc[$tero.'Section SIS ID'] += 1
    $teamdoccode[$tero.'Section SIS ID'] += " " + $tero.'SIS ID'
}

Write-Host "  Tabel opbouwen..."
$teamlid = @()
foreach ($tm in $team) {
    $tl = 1 | Select-Object Team,Leerlingen,Docenten,Code 
    $tl.Team = $tm
    $tl.Leerlingen = $teamleer[$tm]
    $tl.Docenten = $teamdoc[$tm] 
    $tl.Code = $teamdoccode[$tm]
    $teamlid += $tl
}

Write-Host "  Weergeven..."
$teamlid | Out-GridView
$teamlid | Export-Csv -Path $filename_t_teamlid -Encoding UTF8 -NoTypeInformation

$stopwatch.Stop()
Write-Host "Uitvoer klaar (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
