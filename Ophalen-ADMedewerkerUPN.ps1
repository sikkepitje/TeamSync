<#
.SYNOPSIS
    VOORBEELD; Haalt lijst van employeeId en UserPrincipalNames op uit Active Directory
    voor gebruik in TeamSync met configvariabele 'medewerker=csv_upn'.
.DESCRIPTION
    (c) 2020 Paul Wiegmans. All rights reserved. 
    Script provided as-is without any warranty of any kind. 
    Use it freely at your own risks.
.INPUTS
    None
.OUTPUTS
    CSV bestand in "./data_in/Medewerker_UPN.csv"
.NOTES
    Versie:     1.0
    Auteur:     Paul Wiegmans
    GitHub:     https://github.com/sikkepitje/TeamSync
    Datum:      25-11-2020
.EXAMPLE
    .\Ophalen-ADMedewerkerUPN.ps1
#>    

# v--AANPASSEN!--v
$ADsearchbase="OU=Medewerkers,OU=Users,OU=MyBusiness,DC=domain,DC=your,DC=cloud"
$ADserver="dc1.domain.your.cloud"
$filename_medewerkerUPN = "./data_in/Medewerker_UPN.csv"

Write-Host ("Ophalen UserPrincipalNames van personeel uit AD")
Import-Module activedirectory

$users = Get-ADUser -Filter * -Server $ADserver `
    -SearchBase $ADsearchbase -Properties employeeid
# Maak een tabel van employeeId,UserPrincipalName
# van alle medewerkers met een employeeId groter dan 0.
$medew = $users | Select-Object `
    @{Name = 'employeeid'; Expression = {$_.employeeid -replace "[A-Za-z]"}}, 
    UserPrincipalName | `
    Where-Object {$_.employeeid -ne $null} | `
    Where-Object {$_.employeeid -gt 0}
Write-Host "Aantal medewerkers met geldige employeeId:" $medew.count 

# Bewaar tabel in CSV-bestand
$medew | Export-CSV -Path $filename_medewerkerUPN -NoTypeInformation
