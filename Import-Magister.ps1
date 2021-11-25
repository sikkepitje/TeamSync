<#
    .SYNOPSIS

    TeamSync script Import-Magister.ps1; koppeling tussen Magister en School Data Sync.

    .DESCRIPTION

    TeamSync is een koppeling tussen Magister en School Data Sync.
    TeamSync script Import-Magister.ps1 (ophalen) haalt gegevens op uit Medius (Magister)
    Webservice.

    Versie 20211125
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

#>
[CmdletBinding()]
param (
    [Parameter(
        HelpMessage="Geef de naam van de te gebruiken configuratiebestand, bij verstek 'Import-Magister.ini'"
    )]
    [Alias('Inifile','Inibestandsnaam','Config','Configfile','Configuratiebestand')]
    [String]  $Inifilename = "Import-Magister.ini"
)
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

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

# initialisatie constanten 
function Constante ($name, $value) { Set-Variable -Name $Name -Value $Value -Option Constant -Scope Global -Erroraction:SilentlyContinue }
# constanten voor koppelmethode configuratievariabelen
Constante KOPPEL_MEDEWERKERID_AAN_CODE      'code'
Constante KOPPEL_MEDEWERKERID_AAN_LOGIN     'loginaccount'
Constante KOPPEL_MEDEWERKERID_AAN_CSVUPN    'csv_upn'
Constante KOPPEL_LEERLINGID_AAN_LOGIN       'loginaccount'
Constante KOPPEL_LEERLINGID_AAN_EMAIL       'email'

# initialisatie variabelen 
$importfiltermap = "importfilter"
$importkladmap = "importklad"
$importdatamap = "importdata"
$handhaafJPTMedewerkerCodeIsLogin = "0"
$logtag = "IMPORT" 
$medewerker_id = "NIETBESCHIKBAAR"
$leerling_id = "NIETBESCHIKBAAR"
$toondata = "0"

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
    Param ([Parameter(Position=0)][Alias('Message')][string]$Tekst="`n")

    $log = "$(Get-Date -f "yyyy-MM-ddTHH:mm:ss.fff") [$logtag] $tekst"
    $log | Out-File -FilePath $currentLogFilename -Append
    Write-Host $log
}

function Invoke-Webclient($url) {
    $wc = New-Object System.Net.WebClient
    $wc.Encoding = [System.Text.Encoding]::UTF8
    try {
        $feed = [xml]$wc.downloadstring($url)
    } catch {
        $e = $_.Exception
        $line = $_.InvocationInfo.ScriptLineNumber
        $msg = $e.Message 
        Write-Log "Invoke-Webclient: caught exception: $e at line $line : $msg"
        Throw "Invoke-Webclient: caught exception: $e at line $line : $msg"
    }
    if ($feed.Response.Exception) {
        Write-Log  ("Invoke-Webclient: " + $feed.Response.Exception + ": " + $feed.Response.ExceptionMsg)
        Write-Warning ("Invoke-Webclient: " + $feed.Response.Exception + ": " + $feed.Response.ExceptionMsg)
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

################# VERZAMEL LEERLINGEN
function Verzamel_leerlingen() 
{
    # Ophalen leerlingdata, selecteer attributen, en bewaar hele tabel
    Write-Log "Ophalen leerlingen..."
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

    $script:mag_leer = $data.Leerlingen.Leerling | Select-Object `
        @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},
        @{Name = 'Id'; Expression = { ""}},
        @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},
        Email,
        Roepnaam,Tussenv,Achternaam,
        @{Name = 'Lesperiode'; Expression = {$_.'Lesperiode.Korteomschrijving'}},
        @{Name = 'Leerjaar'; Expression = {$_.'Leerfase.leerjaar'}},
        Klas,
        Studie,
        @{Name = 'Profiel'; Expression = {$_.'profiel.code'}},
        @{Name = 'Groepen'; Expression = { @() }},
        @{Name = 'Vakken'; Expression = { @() }},
        @{Name = 'Locatie'; Expression = { $_.'Administratieve_eenheid.Omschrijving' }}

    # velden: Stamnr, Id, Login, Roepnaam, Tussenv, Achternaam, Lesperiode, 
    # Leerjaar, Klas, Studie, Profiel, Groepen, Vakken, Email, Locatie
    # oude code: @{Name = 'Id'; Expression = { if ($useemail) {$_.Email} Else {$_.'loginaccount.naam'}}},

    if ($KOPPEL_LEERLINGID_AAN_EMAIL -eq $leerling_id) {
        foreach ($l in $mag_leer) {
            $l.Id = $l.Email
        }
    } 
    elseif ($KOPPEL_LEERLINGID_AAN_LOGIN -eq $leerling_id) {
        foreach ($l in $mag_leer) {
            $l.Id = $l.Login
        }
    }

    # tussentijds opslaan
    $mag_leer | Export-Csv -Path $filename_t_leerling -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Log ("Leerlingen: " + $mag_leer.count)
    # ID moet gevuld zijn; skip leerlingen zonder e-mail
    $mag_leer = $mag_leer | Where-Object {$_.id.length -gt 0}
    Write-Log ("Leerlingen met geldige ID: " + $mag_leer.count)

    # ID omzetten naar kleine letters 
    foreach ($l in $mag_leer) {
        $l.Id = $l.Id.ToLower()
    }

    # voorfilteren
    if (Test-Path $filename_excl_studie) {
        $filter_excl_studie = $(Get-Content -Path $filename_excl_studie) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Studie -notmatch $filter_excl_studie}
        Write-Log ("Leerlingen na uitsluitend filteren studie: " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_studie) {
        $filter_incl_studie = $(Get-Content -Path $filename_incl_studie) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Studie -match $filter_incl_studie}
        Write-Log ("Leerlingen na insluitend filteren studie: " + $mag_leer.count)
    }
    if (Test-Path $filename_excl_klas) {
        $filter_excl_klas = $(Get-Content -Path $filename_excl_klas) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Klas -notmatch $filter_excl_klas}
        Write-Log ("Leerlingen na uitsluitend filteren klas: " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_klas) {
        $filter_incl_klas = $(Get-Content -Path $filename_incl_klas) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Klas -match $filter_incl_klas}
        Write-Log ("Leerlingen na insluitend filteren klas: " + $mag_leer.count)
    }
    if (Test-Path $filename_incl_locatie) {
        $filter_incl_locatie = $(Get-Content -Path $filename_incl_locatie) -join '|'
        $mag_leer = $mag_leer | Where-Object {$_.Locatie -match $filter_incl_locatie}
        Write-Log ("Leerlingen na insluitend filteren locatie: " + $mag_leer.count)
    }

    if ($mag_leer.count -lt 1) {
        Throw "Er zijn nul leerlingen. Uitvoering stopt"
    }
    $teller = 0
    $leerlingprocent = 100 / [Math]::Max($mag_leer.count, 1)
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

        Write-Progress -Activity "Import Magister leerlinggroepen" -status `
            "Leerling $teller van $($mag_leer.count)" -PercentComplete ($leerlingprocent * $teller++)
    }
    Write-Progress -Activity "Import Magister leerlinggroepen" -status "Leerling" -Completed
    $mag_leer | Export-Clixml -Path $filename_mag_leerling_xml -Encoding UTF8
    if ($toondata) {
        $mag_leer | Out-GridView   # Magister leerlingenlijst met gekoppelde ID
    }
}

################# VERZAMEL DOCENTEN

function Verzamel_docenten() 
{
    Write-Log "Ophalen docenten..."
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

    # Selecteer de belangrijke attributen; paar Id aan AD->Email of Magister->accountnaam.
    $script:mag_doc = $data.Personeelsleden.Personeelslid | Select-Object `
        @{Name = 'Stamnr'; Expression = {$_.stamnr_str}},
        @{Name = 'Id'; Expression = { $null }},
        @{Name = 'Login'; Expression = {$_.'loginaccount.naam'}},
        Code, Roepnaam, Tussenv, Achternaam,
        @{Name = 'Naam'; Expression = {$_.'Loginaccount.Volledige_naam'}},
        @{Name = 'Functie'; Expression = { $_.'Functie.Omschr' }},
        @{Name = 'Groepvakken'; Expression = { $null }},
        @{Name = 'Klasvakken'; Expression = { @() }},
        @{Name = 'Docentvakken'; Expression = { @() }},
        @{Name = 'Locatie'; Expression = { $_.'Administratieve_eenheid.Omschrijving' }}
    # velden: Stamnr, Id, Login, Code, Roepnaam, Tussenv, Achternaam, Naam,  
    # Functie, Groepvakken, Klasvakken, Docentvakken, Locatie
    # oude code: @{Name = 'Id'; Expression = { if ($useemail) {$upntabel[$_.stamnr_str]} Else {$_.'loginaccount.naam'}}},

    if ($KOPPEL_MEDEWERKERID_AAN_CODE -eq $medewerker_id) {
        foreach ($mw in $mag_doc) {
            $mw.Id = $mw.Code
        }
    } 
    elseif ($KOPPEL_MEDEWERKERID_AAN_LOGIN -eq $medewerker_id) {
        foreach ($mw in $mag_doc) {
            $mw.Id = $mw.Login
        }
    }
    elseif ($KOPPEL_MEDEWERKERID_AAN_CSVUPN -eq $medewerker_id) {
        foreach ($mw in $mag_doc) {
            $mw.Id = $upntabel[$mw.stamnr]
        }
    }

    # tussentijds opslaan
    $mag_doc | Export-Csv -Path $filename_t_docent -Delimiter ";" -NoTypeInformation -Encoding UTF8
    Write-Log ("Docenten : " + $mag_doc.count)

    # Speciaal geval JPT: Om onbekende redenen staan sommige personeelsleden dubbel erin. 
    # Docenten met voornaam als login zijn overtollig. 
    # Filter alle medewerker eruit waarvan Magister:code ongelijk is aan Magister:loginaccount.name.
    if ($handhaafJPTMedewerkerCodeIsLogin) {
        $script:mag_doc = $mag_doc | Where-Object {$_.code -eq $_.login}
        Write-Log ("handhaafJPTMedewerkerCodeIsLogin: Docenten na uitfilteren van dubbele Ids: " + $mag_doc.count)
    }

    # Filter docenten met meer dan één rol 
    if ($True) {
        $script:mag_doc = $mag_doc | Sort-Object id -Unique
        Write-Log ("MaakIdUniek: Docenten na uniek maken Ids: " + $mag_doc.count )
    }

    # Algemeen: filter de medewerkers eruit zonder Id
    if ($True) {
        $script:mag_doc = $mag_doc | Where-Object {$null -ne $_.Id}
        Write-Log ("IdNotNull: Docenten na uitfilteren van lege Ids: " + $mag_doc.count)
    }
    Write-Log ("Docenten met geldige Id: " + $mag_doc.count)

    # omzetten naar kleine letters
    foreach ($mw in $mag_doc) {
        $mw.ID = $mw.ID.ToLower()
    }

    # voorfilteren
    if ($mag_doc.count -eq 0) {
        Throw "Er zijn nul docenten. Er is niets te doen"
    }

    if (Test-Path $filename_excl_docent) {
        $filter_excl_docent = $(Get-Content -Path $filename_excl_docent) -join '|'
        $mag_doc = $mag_doc | Where-Object {$_.Id -notmatch $filter_excl_docent}
        Write-Log ("Docenten na uitsluitend filteren docent: " + $mag_doc.count)
    }

    if (Test-Path $filename_incl_docent) {
        $filter_incl_docent = $(Get-Content -Path $filename_incl_docent) -join '|'
        $mag_doc = $mag_doc | Where-Object {$_.Id -match $filter_incl_docent}
        Write-Log ("Docenten na insluitend filteren docent: " + $mag_doc.count)
    }

    $teller = 0
    $docentprocent = 100 / [Math]::Max($mag_doc.count, 1)
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

        Write-Progress -Activity "Import Magister docentgroepen" -status `
            "Docent $teller van $($mag_doc.count)" -PercentComplete ($docentprocent * $teller++)
    }
    Write-Progress -Activity "Import Magister docentgroepen" -status "Docent" -Completed
    
    $mag_doc | Export-Clixml -Path $filename_mag_docent_xml -Encoding UTF8
    $mag_vak | Export-Clixml -Path $filename_mag_vak_xml -Encoding UTF8

    if ($toondata) {
        $mag_doc | Out-GridView   # Magister docentenlijst met gekoppelde ID
        $mag_vak | Out-GridView   # Magister vakcodes en omschrijvingen
    }

}
#endregion Functies

LogRotate
Write-Log ""
Write-Log ("START " + $MyInvocation.MyCommand.Name)
Try {

    # Configuratieparameter inlezen
    $filename_settings = $herePath + "\" + $Inifilename
    Write-Log ("Configuratiebestand: " + $filename_settings)
    $settings = Get-Content $filename_settings | ConvertFrom-StringData
    foreach ($key in $settings.Keys) {
        Set-Variable -Name $key -Value $settings.$key -Scope global
    }
    # Configuratieparameter validatie
    if (!$magisterUser)  { Throw "Configuratieparameter 'magisterUser' is vereist"}
    if (!$magisterPass)  { Throw "Configuratieparameter 'magisterPass' is vereist"}
    if (!$magisterUrl)  { Throw "Configuratieparameter 'magisterUrl' is vereist"}
    if ($medewerker_id -eq "NIETBESCHIKBAAR") { Throw "Configuratieparameter 'medewerker_id' is vereist"}
    if ($leerling_id -eq "NIETBESCHIKBAAR") { Throw "Configuratieparameter 'leerling_id is' vereist"}
    $handhaafJPTMedewerkerCodeIsLogin = $handhaafJPTMedewerkerCodeIsLogin -ne "0"  # maak boolean
    $toondata = $toondata -ne "0"  # maak boolean
    if ($medewerker_id -notin $KOPPEL_MEDEWERKERID_AAN_CODE, $KOPPEL_MEDEWERKERID_AAN_LOGIN, $KOPPEL_MEDEWERKERID_AAN_CSVUPN) {
        Throw "Geen geldige koppelmethode voor medewerkers: $medewerker_id "
    }
    if ($leerling_id -notin $KOPPEL_LEERLINGID_AAN_EMAIL, $KOPPEL_LEERLINGID_AAN_LOGIN) {
        Throw "Geen geldige koppelmethode voor leerling: $leerling_id "
    }

    # datamappen
    $filterPath     = "$herePath\$importfiltermap"
    $tempPath       = "$herePath\$importkladmap"
    $dataPath       = "$herePath\$importdatamap"
    Write-Log ("ImportFilterMap : " + $filterPath)
    Write-Log ("ImportKladMap   : " + $tempPath)
    Write-Log ("ImportDataMap   : " + $dataPath)

    New-Item -path $tempPath -ItemType Directory -ea:Silentlycontinue
    New-Item -path $dataPath -ItemType Directory -ea:Silentlycontinue

    # Filterbestanden (R)
    $filename_excl_docent       = $filterPath + "\excl_docent.csv"
    $filename_incl_docent       = $filterPath + "\incl_docent.csv"
    $filename_excl_klas         = $filterPath + "\excl_klas.csv"
    $filename_incl_klas         = $filterPath + "\incl_klas.csv"
    $filename_excl_studie       = $filterPath + "\excl_studie.csv"
    $filename_incl_studie       = $filterPath + "\incl_studie.csv"
    $filename_incl_locatie      = $filterPath + "\incl_locatie.csv"
    $filename_mwupncsv          = $filterPath + "\Medewerker_UPN.csv"

    # Kladbestanden (W)
    $filename_t_leerling        = $tempPath + "\leerling.csv"
    $filename_t_docent          = $tempPath + "\docent.csv"
    $filename_persemail_xml     = $tempPath + "\personeelemail.clixml"

    # importdata geproduceerd (W)
    $filename_mag_leerling_xml  = $dataPath + "\magister_leerling.clixml"
    $filename_mag_docent_xml    = $dataPath + "\magister_docent.clixml"
    $filename_mag_vak_xml       = $dataPath + "\magister_vak.clixml"

    if ($KOPPEL_MEDEWERKERID_AAN_CSVUPN -eq $medewerker_id) {
        $users = Import-CSV  -Path $filename_mwupncsv
        if (!$users)  {
            Throw "Geen records in de medewerker_upn tabel"
        }
        # maak een hashtable
        $upntabel = @{}
        foreach ($user in $users) {
            $upntabel[$user.employeeId] = $user.UserPrincipalName
        }
        # hashtable $upntabel["$stamnr"] geeft $UserPrincipalName
    }

    # voor dataminimalisatie houd ik een lijstje met vakken bij
    $mag_vak = @{}   # associatieve array van vakomschrijvingen geindexeerd op vakcodes

    # haal sessiontoken
    $MyToken = ""
    $GetToken_URL = $magisterUrl + "?Library=Algemeen&Function=Login&UserName=" + 
    $magisterUser + "&Password=" + $magisterPass + "&Type=XML"
    $feed = [xml](new-object system.net.webclient).downloadstring($GetToken_URL)
    if ($feed.Response.Result -ne "True") {
        Throw "Fatale fout in GetToken: " + $feed.Response.ResultMessage
    }
    $MyToken = $feed.response.SessionToken
    
    Verzamel_docenten
    Verzamel_leerlingen

    ################# EINDE

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
