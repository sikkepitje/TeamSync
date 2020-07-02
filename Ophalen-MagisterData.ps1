<#
    .SYNOPSIS

    TeamSync script deel 1; koppeling tussen Magister en School Data Sync.

    .DESCRIPTION

    TeamSync is een koppeling tussen Magister en School Data Sync.
    TeamSync script deel 1 (ophalen) haalt gegevens op uit Medius (Magister)
    Webservice.

    Versie 20200702
    Auteur Paul Wiegmans (p.wiegmans@svok.nl)

    naar een voorbeeld door Wim den Ronde, Eric Redegeld, Joppe van Daalen

    .PARAMETER Inifilename

    bepaalt de bestandsnaam van het INI-bestand, waarin benodigde parameters 
    worden gelezen, relatief ten opzichte van het pad van dit script.

    .LINK

    https://github.com/sikkepitje/teamsync

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
$ADSearchBase = ""
$ADServer = "" 

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
if ($useemail) {
    if (!$ADSearchBase)  { Throw "ADSearchBase is vereist"}
    if (!$ADServer)  { Throw "ADServer is vereist"}    
}
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
$filename_t_leerling = $tempPath + "\leerling.csv"
$filename_t_docent = $tempPath + "\docent.csv"
$filename_mag_leerling_xml = $tempPath + "\mag_leerling.clixml"
$filename_mag_docent_xml = $tempPath + "\mag_docent.clixml"
$filename_mag_vak_xml = $tempPath + "\mag_vak.clixml"
$filename_persemail_xml = $tempPath + "\personeelemail.clixml"

if ($useemail) {
    Write-Host "Ophalen personeel uit AD"
    Import-Module activedirectory
   
    $users = Get-ADUser -Filter * -Server $ADserver -SearchBase $ADsearchbase -Properties employeeid
    
    # Extraheer uit employeeid (bijvoorbeeld "bc435") een stamnr
    $medew = $users | Select-Object UserPrincipalName,employeeid,
        @{Name = 'Stamnr'; Expression = {$_.employeeid.replace("bc","")}}
    $medew = $medew | Where-Object {$_.Stamnr -ne $null} | Where-Object {$_.Stamnr -gt 0}
    # Velden: UserPrincipalName, employeeid, stamnr
    Write-Host "Aantal:" $medew.count 
    # maak hashtable
    $email = @{}
    foreach ($user in $medew) {
        $email[$user.stamnr] = $user.UserPrincipalName
    }
     # hashtable $email["$stamnr"] geeft $UserPrincipalName
    $email | Export-Clixml -Path $filename_persemail_xml
}

function Invoke-Webclient($url) {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    try {
        $feed = [xml]$wc.downloadstring($url)
    } catch {
        Throw "Invoke-Webclient: er ging iets mis"
    }
    if ($feed.Response.Exception) {
        Write-Warning ("Invoke-Webclient:" + $feed.Response.Exception + ":" + $feed.Response.ExceptionMsg)
        return $feed.Response
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

# voor dataminimalisatie houd ik een lijstje met vakken bij
$mag_vak = @{}   # associatieve array van vakomschrijvingen geindexeerd op vakcodes

# haal sessiontoken
$MyToken = ""
$GetToken_URL = $magisterUrl + "?Library=Algemeen&Function=Login&UserName=" + 
$magisterUser + "&Password=" + $magisterPass + "&Type=XML"
$feed = [xml](new-object system.net.webclient).downloadstring($GetToken_URL)
if ($feed.Response.Result -ne "True") {
    Throw $feed.Response.ResultMessage
}
$MyToken = $feed.response.SessionToken

################# VERZAMEL LEERLINGEN

# Ophalen leerlingdata, selecteer attributen, en bewaar hele tabel
Write-Host "Ophalen leerlingen..."
$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveStudents"
<#
Achternaam                              Property              string Achternaam {get;set;}
Administratieve_eenheid.Omschrijving    Property              string Administratieve_eenheid.Omschrijving {get;set;}
Administratieve_eenheid.Plaats          Property              string Administratieve_eenheid.Plaats {get;set;}
Adres                                   Property              string Adres {get;set;}
c_vrij1                                 Property              string c_vrij1 {get;set;}
c_vrij2                                 Property              string c_vrij2 {get;set;}
Email                                   Property              string Email {get;set;}
geb_datum_str                           Property              string geb_datum_str {get;set;}
Geslacht                                Property              string Geslacht {get;set;}
Klas                                    Property              string Klas {get;set;}
Land___Nationaliteit.Land               Property              string Land___Nationaliteit.Land {get;set;}
Leerfase.Leerjaar                       Property              string Leerfase.Leerjaar {get;set;}
Lesperiode.Korte_omschrijving           Property              string Lesperiode.Korte_omschrijving {get;set;}
Loginaccount.Naam                       Property              string Loginaccount.Naam {get;set;}
Nr                                      Property              string Nr {get;set;}
Nr.tv                                   Property              string Nr.tv {get;set;}
Personeelslid.Volledige_naam            Property              string Personeelslid.Volledige_naam {get;set;}
Plaats.Woonplaats                       Property              string Plaats.Woonplaats {get;set;}
Postcode                                Property              string Postcode {get;set;}
Profiel.Code                            Property              string Profiel.Code {get;set;}
Profiel.Omschrijving                    Property              string Profiel.Omschrijving {get;set;}
Roepnaam                                Property              string Roepnaam {get;set;}
sis_pers0.sis_pers0.sis_pers0__naam_vol Property              string sis_pers0.sis_pers0.sis_pers0__naam_vol {get;set;}
stamnr_str                              Property              string stamnr_str {get;set;}
Straat                                  Property              string Straat {get;set;}
Studie                                  Property              string Studie {get;set;}
Tel._1_geheim                           Property              string Tel._1_geheim {get;set;}
Telefoon                                Property              string Telefoon {get;set;}
Telefoon_2                              Property              string Telefoon_2 {get;set;}
Tussenv                                 Property              string Tussenv {get;set;}
Volledige_naam                          Property              string Volledige_naam {get;set;}
Voorletters                             Property              string Voorletters {get;set;}
#>
#$data.Leerlingen.Leerling | ogv
#exit 71

$mag_leer = $data.Leerlingen.Leerling | Select-Object `
    @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},`
    @{Name = 'Id'; Expression = { if ($useemail) {$_.Email} Else {$_.'loginaccount.naam'}}}, `
    @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},`
    Roepnaam,Tussenv,Achternaam,`
    @{Name = 'Lesperiode'; Expression = {$_.'Lesperiode.Korteomschrijving'}},`
    @{Name = 'Leerjaar'; Expression = {$_.'Leerfase.leerjaar'}},`
    Klas,`
    Studie,`
    @{Name = 'Profiel'; Expression = {$_.'profiel.code'}},`
    @{Name = 'Groepen'; Expression = { @() }},`
    @{Name = 'Vakken'; Expression = { @() }},
    Email, `
    @{Name = 'Locatie'; Expression = { $_.'Administratieve_eenheid.Omschrijving' }}

# velden: Stamnr, Id, Login, Roepnaam, Tussenv, Achternaam, Lesperiode, 
# Leerjaar, Klas, Studie, Profiel, Groepen, Vakken, Email, Locatie

# tussentijds opslaan
$mag_leer | Export-Csv -Path $filename_t_leerling -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Leerlingen           :" $mag_leer.count

# voorfilteren
if (Test-Path $filename_excl_studie) {
    $filter_excl_studie = $(Get-Content -Path $filename_excl_studie) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Studie -notmatch $filter_excl_studie}
    Write-Host "Leerlingen na uitsluitend filteren studie :" $mag_leer.count
}
if (Test-Path $filename_incl_studie) {
    $filter_incl_studie = $(Get-Content -Path $filename_incl_studie) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Studie -match $filter_incl_studie}
    Write-Host "Leerlingen na insluitend filteren studie :" $mag_leer.count
}
if (Test-Path $filename_excl_klas) {
    $filter_excl_klas = $(Get-Content -Path $filename_excl_klas) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Klas -notmatch $filter_excl_klas}
    Write-Host "Leerlingen na uitsluitend filteren klas  :" $mag_leer.count
}
if (Test-Path $filename_incl_klas) {
    $filter_incl_klas = $(Get-Content -Path $filename_incl_klas) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Klas -match $filter_incl_klas}
    Write-Host "Leerlingen na insluitend filteren klas   :" $mag_leer.count
}
if (Test-Path $filename_incl_locatie) {
    $filter_incl_locatie = $(Get-Content -Path $filename_incl_locatie) -join '|'
    $mag_leer = $mag_leer | Where-Object {$_.Locatie -match $filter_incl_locatie}
    Write-Host "Leerlingen na insluitend filteren locatie:" $mag_leer.count
}

if ($mag_leer.count -lt 1) {
    Throw "Geen leerlingen... Niets te doen"
}
$teller = 0
$leerlingprocent = 100 / $mag_leer.count
foreach ($leerling in $mag_leer) {

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
        "Leerling $teller van $($mag_leer.count)" -PercentComplete ($leerlingprocent * $teller++)
}
Write-Progress -Activity "Magister data verwerken" -status "Leerling" -Completed

$mag_leer | Export-Clixml -Path $filename_mag_leerling_xml -Encoding UTF8

#$mag_leer | Out-GridView

################# VERZAMEL DOCENTEN
Write-Host "Ophalen docenten..."
$data = ADFunction -Url $magisterUrl -Sessiontoken $MyToken -Function "GetActiveEmpoyees"  
<#
Achternaam                           Property              string Achternaam {get;set;}
Administratieve_eenheid.Omschrijving Property              string Administratieve_eenheid.Omschrijving {get;set;}
Adres                                Property              string Adres {get;set;}
Code                                 Property              string Code {get;set;}
c_vrij1                              Property              string c_vrij1 {get;set;}
c_vrij2                              Property              string c_vrij2 {get;set;}
c_vrij3                              Property              string c_vrij3 {get;set;}
c_vrij4                              Property              string c_vrij4 {get;set;}
datum_in_str                         Property              string datum_in_str {get;set;}
dVertrek_str                         Property              string dVertrek_str {get;set;}
Functie.Omschr                       Property              string Functie.Omschr {get;set;}
Geheim                               Property              string Geheim {get;set;}
Huisnr                               Property              string Huisnr {get;set;}
Huisnr._toevoeging                   Property              string Huisnr._toevoeging {get;set;}
Loginaccount.Naam                    Property              string Loginaccount.Naam {get;set;}
Loginaccount.Volledige_naam          Property              string Loginaccount.Volledige_naam {get;set;}
M_V                                  Property              string M_V {get;set;}
Off._voornamen                       Property              string Off._voornamen {get;set;}
Oude_personeelscode                  Property              string Oude_personeelscode {get;set;}
Plaats                               Property              string Plaats {get;set;}
Postcode                             Property              string Postcode {get;set;}
Roepnaam                             Property              string Roepnaam {get;set;}
stamnr_str                           Property              string stamnr_str {get;set;}
Straat                               Property              string Straat {get;set;}
Telefoon                             Property              string Telefoon {get;set;}
Telefoon_2                           Property              string Telefoon_2 {get;set;}
Telefoon_3                           Property              string Telefoon_3 {get;set;}
Telefoon_4                           Property              string Telefoon_4 {get;set;}
Tussenv                              Property              string Tussenv {get;set;}
Voorletters                          Property              string Voorletters {get;set;}
#>
#$data.Personeelsleden.Personeelslid | ogv
#exit 45

$mag_doc = $data.Personeelsleden.Personeelslid | Select-Object `
    @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},`
    @{Name = 'Id'; Expression = { if ($useemail) {$email[$_.stamnr_str]} Else {$_.'loginaccount.naam'}}}, `
    @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},`
    Roepnaam,Tussenv,Achternaam,`
    @{Name = 'Naam'; Expression = {$_.'Loginaccount.Volledige_naam'}},`
    Code,`
    @{Name = 'Functie'; Expression = { $_.'Functie.Omschr' }},`
    @{Name = 'Groepvakken'; Expression = { $null }},`
    @{Name = 'Klasvakken'; Expression = { @() }},`
    @{Name = 'Docentvakken'; Expression = { @() }}, `
    @{Name = 'Locatie'; Expression = { $_.'Administratieve_eenheid.Omschrijving' }}
# velden: Stamnr, Id, Login, Roepnaam, Tussenv, Achternaam, Naam, Code, 
# Functie, Groepvakken, Klasvakken, Docentvakken, Locatie

# JPT: Om onbekende redenen staan sommige personeelsleden dubbel erin. 
# Filter de docenten met voornaam als login eruit.
$mag_doc = $mag_doc | Where-Object {$_.code -eq $_.login}

# tussentijds opslaan
$mag_doc | Export-Csv -Path $filename_t_docent -Delimiter ";" -NoTypeInformation -Encoding UTF8
Write-Host "Docenten ongefilterd :" $mag_doc.count

# Wanneer id is gebaseerd op email, filter de medewerkers eruit
# waarvan email niet kon worden opgezocht in AD
if ($useemail) {
    $mag_doc = $mag_doc | Where-Object {$_.Id -ne $null}
    Write-Host "D na uitfilteren van lege Ids:" $mag_doc.count
}

# voorfilteren
if ($mag_doc.count -eq 0) {
    Throw "Geen docenten ?? Stopt!"
}

if (Test-Path $filename_excl_docent) {
    $filter_excl_docent = $(Get-Content -Path $filename_excl_docent) -join '|'
    $mag_doc = $mag_doc | Where-Object {$_.Code -notmatch $filter_excl_docent}
    Write-Host "Docenten na uitsluitend filteren docent :" $mag_doc.count
}

if (Test-Path $filename_incl_docent) {
    $filter_incl_docent = $(Get-Content -Path $filename_incl_docent) -join '|'
    $mag_doc = $mag_doc | Where-Object {$_.Code -match $filter_incl_docent}
    Write-Host "Docenten na insluitend filteren docent :" $mag_doc.count
}

$teller = 0
$docentprocent = 100 / $mag_doc.count
foreach ($docent in $mag_doc ) {

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
        "Docent $teller van $($mag_doc.count)" -PercentComplete ($docentprocent * $teller++)
}
Write-Progress -Activity "Magister uitlezen" -status "Docent" -Completed

$mag_doc | Export-Clixml -Path $filename_mag_docent_xml -Encoding UTF8
$mag_vak | Export-Clixml -Path $filename_mag_vak_xml -Encoding UTF8

################# EINDE

$stopwatch.Stop()
Write-Host "Uitvoer klaar (uu:mm.ss)" $stopwatch.Elapsed.ToString("hh\:mm\.ss")
Stop-Transcript -ea SilentlyContinue
