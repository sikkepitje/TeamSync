<#
    Ophalen-EntraMedewerkerUPN-voorbeeld.ps1
    20-02-2024 p.wiegmans@svok.nl 

    Voorbeeldscript: haalt een lijst met gebruikers (upn, employeeid, mail) op uit Entra ID 
    en bewaart in een CSV.
#>

Import-Module Microsoft.Graph.Authentication
Import-Module Microsoft.Graph.Users

$selfdir = Split-Path -Parent $MyInvocation.MyCommand.Path

Connect-MgGraph -ClientId "11112222-3333-4444-5555-666677778888" `
    -TenantId "11112222-3333-4444-5555-666677778888" `
    -CertificateName "CN=My Certificate" -NoWelcome
    
$users = Get-MgUser -All -Property EmployeeId, UserPrincipalName, Mail

$users | Export-Csv -Path "$selfdir\Medewerker_UPN.csv" -Delimiter "," -Encoding UTF8 -NoTypeInformation
