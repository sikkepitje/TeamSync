<#
.SYNOPSIS
    Laat zien hoe je gebruikers uit Azure AD leest.
.DESCRIPTION
    Dit script dient als voorbeeld om te tonen hoe we een lijst met
    UserPrincipalNames en employeeIds van medewerkers kunnen ophalen uit Azure
    Active Directory.

    In dit geval worden medewerkers van een school geselecteerd aan de hand van
    extensionAttribute1="BON" en extensionAttribute2="Medewerker"

.INPUTS
    None
.OUTPUTS
    None
.NOTES
    Versie:     1.0
    Auteur:     Paul Wiegmans
    GitHub:     https://github.com/sikkepitje/TeamSync
    Datum:      7-7-2021

    Het script produceert uitvoer in het CSV bestand "Medewerker_UPN.csv"
.EXAMPLE
    .\Ophalen-AzureADMedewerkersUPN.ps1
#>    

Install-Module AzureAD
Connect-AzureAD

$filename_medewerkerUPN = "./Medewerker_UPN.csv"
$ea1="extension_92c806d48fcc4740a5b5166f298e334e_extensionAttribute1"
$ea2="extension_92c806d48fcc4740a5b5166f298e334e_extensionAttribute2"

Write-Host "Alle gebruikers ophalen (dit kan even duren)..."
$medewerkers = Get-AzureADUser -All $true `
| Where-Object {$_.ExtensionProperty.$ea1 -eq "BON"  -and $_.ExtensionProperty.$ea2 -eq "Medewerker"}

$mijnlijst = $medewerkers | Select-Object `
    @{Name = 'employeeId'; Expression = {$_.ExtensionProperty.employeeId -replace "[A-Za-z]"}}, 
    UserPrincipalName | `
    Where-Object {$_.employeeId -ne $null} | `
    Where-Object {$_.employeeId -gt 0}
Write-Host "Aantal medewerkers met geldige employeeId:" $mijnlijst.count 

$mijnlijst | Export-CSV -Path $filename_medewerkerUPN -NoTypeInformation
