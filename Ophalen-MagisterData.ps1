<#
    Get-JPTMagisterData.ps1

    16 juni 2020 Paul Wiegmans
    naar een voorbeeld van Wim den Ronde, Eric Redegeld, Joppe van Daalen

    Dit script haalt gegevens over leerlingen, docenten en vakken op uit Magister
    webservices, en bewaart dit als tussenresultaat op schijf, 
    ten behoeve van het aanmaken van gegevensbestanden voor School Data Sync.

    Dit is stap 1 in een Teamsync v2 conversie van JPT Magister naar School Data Sync.
    * Ophalen
    Hierna volgt Transformeren-Naar-SchoolDataSync.ps1
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

$jaarlaag_heeft_lesgroepen = "3", "4", "5", "6"  # er wordt alleen voor deze jaarlagen gezocht naar lesgroepen

Write-Host " "
Write-Host "Start..."
$inputPath = $herePath + "\data_in"
$tempPath = $herePath + "\data_temp"
$outputPath = $herePath + "\data_out"
New-Item -path $inputPath -ItemType Directory -ea:Silentlycontinue
New-Item -path $tempPath -ItemType Directory -ea:Silentlycontinue
New-Item -path $outputPath -ItemType Directory -ea:Silentlycontinue

# Bestanden in map Data_in
$filename_excl_docent = $inputPath + "\excl_docent.csv"
$filename_incl_docent = $inputPath + "\incl_docent.csv"
$filename_excl_klas   = $inputPath + "\excl_klas.csv"
$filename_incl_klas   = $inputPath + "\incl_klas.csv"
$filename_excl_studie = $inputPath + "\excl_studie.csv"
$filename_incl_studie = $inputPath + "\incl_studie.csv"

# Bestanden in map Data_TEMP
$filename_t_leerling  = $tempPath + "\leerling.csv"
$filename_t_docent    = $tempPath + "\docent.csv"
$filename_mag_leerling_xml = $tempPath + "\mag_leerling.clixml"
$filename_mag_docent_xml   = $tempPath + "\mag_docent.clixml"
$filename_mag_vak_xml      = $tempPath + "\mag_vak.clixml"

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

# voor data minimalisatie houden we lijstje met vakken bij
$mag_vak = @{}   # associatieve array van vakomschrijvingen geindexeerd op vakcodes

# haal sessiontoken
$MyToken = ""
$GetToken_URL = $magisterUrl + "?Library=Algemeen&Function=Login&UserName=" + 
$magisterUser + "&Password=" + $magisterPass + "&Type=XML"
$feed = [xml](new-object system.net.webclient).downloadstring($GetToken_URL)
$MyToken = $feed.response.SessionToken

################# VERZAMEL LEERLINGEN
# Ophalen leerlingdata, selecteer attributen, en bewaar hele tabel
Write-Host "Ophalen leerlingen..."
#$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveStudents"
$mag_leerling = (ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveStudents").Leerlingen.Leerling | Select-Object `
    @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},`
    @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},`
    Roepnaam,Tussenv,Achternaam,`
    @{Name = 'Lesperiode'; Expression = {$_.'Lesperiode.Korteomschrijving'}},`
    @{Name = 'Leerjaar'; Expression = {$_.'Leerfase.leerjaar'}},`
    Klas,`
    Studie,`
    @{Name = 'Profiel'; Expression = {$_.'profiel.code'}},`
    @{Name = 'Groepen'; Expression = { @() }},`
    @{Name = 'Vakken'; Expression = { @() }}
# velden: Stamnr, Login, Roepnaam, Tussenv, Achternaam, Lesperiode, 
# Leerjaar, Klas, Studie, Profiel, Groepen, Vakken

# tussentijds opslaan
$mag_leerling | Export-Csv -Path $filename_t_leerling -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Leerlingen           :" $mag_leerling.count
$data = $null 

# voorfilteren
if (Test-Path $filename_excl_studie) {
    $filter_excl_studie = $(Get-Content -Path $filename_excl_studie) -join '|'
    $mag_leerling = $mag_leerling | Where-Object {$_.Studie -notmatch $filter_excl_studie}
    Write-Host "Leerlingen na uitsluitend filteren studie :" $mag_leerling.count
}
if (Test-Path $filename_incl_studie) {
    $filter_incl_studie = $(Get-Content -Path $filename_incl_studie) -join '|'
    $mag_leerling = $mag_leerling | Where-Object {$_.Studie -match $filter_incl_studie}
    Write-Host "Leerlingen na insluitend filteren studie :" $mag_leerling.count
}
if (Test-Path $filename_excl_klas) {
    $filter_excl_klas = $(Get-Content -Path $filename_excl_klas) -join '|'
    $mag_leerling = $mag_leerling | Where-Object {$_.Klas -notmatch $filter_excl_klas}
    Write-Host "Leerlingen na uitsluitend filteren klas  :" $mag_leerling.count
}
if (Test-Path $filename_incl_klas) {
    $filter_incl_klas = $(Get-Content -Path $filename_incl_klas) -join '|'
    $mag_leerling = $mag_leerling | Where-Object {$_.Klas -match $filter_incl_klas}
    Write-Host "Leerlingen na insluitend filteren klas   :" $mag_leerling.count
}

$teller = 0
$leerlingprocent = 100 / $mag_leerling.count
foreach ($leerling in $mag_leerling) {

    # verzamel de lesgroepen
    # een team voor elke lesgroep
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetLeerlingGroepen" -Stamnr $leerling.Stamnr
    foreach ($groepnode in $data.vakken.vak) {
        <#
        Stamnr Lesgroep groep
        ------ -------- -----
        9479   11286    4h.maatA
        #>
        $leerling.groepen += @($groepnode.groep)
    }

    # verzamel de vakken
    # een team voor elke vakklas
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetLeerlingVakken" -Stamnr $leerling.Stamnr
    foreach ($vaknode in $data.vakken.vak) {
        <#
        Stamnr Vak
        ------ ---
        11300  wi
        #>
        $leerling.Vakken += @($vaknode.Vak)
    }

    Write-Progress -Activity "Magister data verwerken" -status `
        "Leerling $teller van $($mag_leerling.count)" -PercentComplete ($leerlingprocent * $teller++)
}
Write-Progress -Activity "Magister data verwerken" -status "Leerling" -Completed

$mag_leerling | Export-Clixml -Path $filename_mag_leerling_xml -Encoding UTF8

#$mag_leerling | Out-GridView

################# VERZAMEL DOCENTEN
Write-Host "Ophalen docenten..."
$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveEmpoyees"  
$mag_docent = $data.Personeelsleden.Personeelslid | Select-Object `
    @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},`
    @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},`
    Roepnaam,Tussenv,Achternaam,`
    @{Name = 'Naam'; Expression = {$_.'Loginaccount.Volledige_naam'}},`
    Code,`
    @{Name = 'Functie'; Expression = { $_.'Functie.Omschr' }},`
    @{Name = 'Groepvakken'; Expression = { $null }},`
    @{Name = 'Klasvakken'; Expression = { @() }},`
    @{Name = 'Docentvakken'; Expression = { @() }}
# velden: Stamnr, Login, Roepnaam, Tussenv, Achternaam, Naam, Code, 
# Functie, Groepvakken, Klasvakken, Docentvakken

# JPT: Om onbekende redenen staan sommige personeelsleden dubbel erin. 
# De docenten met voornaam in login eruit filteren.
$mag_docent = $mag_docent | Where-Object {$_.code -eq $_.login}

# tussentijds opslaan
$mag_docent | Export-Csv -Path $filename_t_docent -Delimiter ";" -NoTypeInformation -Encoding UTF8

Write-Host "Docenten ongefilterd :" $mag_docent.count
if ($mag_docent.count -eq 0) {
    Throw "Geen docenten ?? Stopt!"
}
if (Test-Path $filename_excl_docent) {
    $filter_excl_docent = $(Get-Content -Path $filename_excl_docent) -join '|'
    $mag_docent = $mag_docent | Where-Object {$_.Code -notmatch $filter_excl_docent}
    Write-Host "Docenten na uitsluitend filteren docent :" $mag_docent.count
}

if (Test-Path $filename_incl_docent) {
    $filter_incl_docent = $(Get-Content -Path $filename_incl_docent) -join '|'
    $mag_docent = $mag_docent | Where-Object {$_.Code -match $filter_incl_docent}
    Write-Host "Docenten na insluitend filteren docent :" $mag_docent.count
}

$teller = 0
$docentprocent = 100 / $mag_docent.count
foreach ($docent in $mag_docent ) {

    # verzamel Groepvakken
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelGroepVakken" -Stamnr $docent.stamnr
    foreach ($gvnode in $data.Lessen.Les) {
        <# velden: 
        Personeelslid.Stamnr Klas     Vak.Vakcode Vak.Omschrijving
        -------------------- ----     ----------- ----------------
        11                   4v.dutl1 dutl        Duitse taal en literatuur 
        #>
        $rec = 1 | Select-Object Klas, Vakcode
        $rec.Klas = $gvnode.Klas
        $rec.Vakcode = $gvnode.'Vak.Vakcode'
        $docent.Groepvakken += @($rec) 

        if ($mag_vak.keys -notcontains $gvnode.'Vak.Vakcode') {
            $mag_vak[$gvnode.'Vak.Vakcode'] = $gvnode.'Vak.Omschrijving'
        }
    }

    # verzamelen Klasvakken
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelKlasVakken" -Stamnr $docent.stamnr
    foreach ($kvnode in $data.Lessen.Les) {
        <# velden:
        Personeelslid.Stamnr Klas_Lesgroep Klas
        -------------------- ------------- ----
        11                   11182         5vD 
        #>
        $docent.Klasvakken += @($kvnode.Klas)
    }

    # verzamel Docentvakken
    $data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetPersoneelVakken" -Stamnr $docent.stamnr
    foreach ($dvnode in $data.Lessen.Les) {
        <# velden:
        Personeelslid.Stamnr Vak.Vakcode Vak.Omschrijving
        -------------------- ----------- ----------------
        11                   dutl        Duitse taal en literatuur 
        #>
        $docent.Docentvakken += @($dvnode.'Vak.Vakcode')

        if ($mag_vak.keys -notcontains $dvnode.'Vak.Vakcode') {
            $mag_vak[$dvnode.'Vak.Vakcode'] = $dvnode.'Vak.Omschrijving'
        }
    }

    Write-Progress -Activity "Magister uitlezen" -status `
        "Docent $teller van $($mag_docent.count)" -PercentComplete ($docentprocent * $teller++)
}
Write-Progress -Activity "Magister uitlezen" -status "Docent" -Completed

$mag_docent | Export-Clixml -Path $filename_mag_docent_xml -Encoding UTF8
$mag_vak | Export-Clixml -Path $filename_mag_vak_xml -Encoding UTF8

#$mag_docent |  Out-GridView
#$mag_vak | Out-GridView

################# EINDE

$stopwatch.Stop()
Write-Host "Uitvoer klaar (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
