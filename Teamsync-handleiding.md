# Teamsync Handleiding

Paul Wiegmans (p.wiegmans@svok.nl) 

## Bestandsformaten

### Instellingen

Het bestand 'Teamsync.ini' bevat een aantal naam-waarde-paren, die parameters definieert die nodig zijn voor Teamsync. Deze parameters worden gespeficeerd als een naam-waarde-paar. Deze heeft de volgende vorm:

    <naam>=<waarde>

Aanhalingstekens zijn toegestaan maar niet nodig. Spaties in de waarde-deel zijn toegestaan. 

De volgende parametersnamen zijn vereist en moeten een waarde hebben: 
* brin
* schoolnaam
* magisterUser
* magisterPassword
* magisterUrl
* teamnaam-prefix


### Filters 

* excl_studie.csv    , filtert leerlingen op studie door weglating.
* incl_studie.csv    , filtert leerlingen op studie door insluiting. 
* incl_klas.csv      , filtert leerlingen op klas door opname.
* incl_docent.csv    , filtert docenten op code door opname.

In de map 'Data_in' kunnen één of meer van bovenstaande filterbestanden worden aangemaakt. Deze bestanden bevatten filters, die selectief records uit de invoer filteren. Ze kunnen **exclusief** filteren, dat wil zeggen dat overeenkomende records worden weggegooid en uitgesloten van verwerking, of ze kunnen **inclusief** filteren, dat wil zeggen dat uitsluitend de overeenkomende records worden verwerkt.

Het gebruik van deze filterbestanden is optioneel. Als ze bestaan, worden ze ingelezen en toegepast op de invoer. Als ze niet bestaan, wordt er niet gefilterd.

Indien gebruikt, dan kan elk van deze bestand een of meer filters bevatten, elk op een eigen regel, die worden toegepast met behulp van de "match"-operator voor het filteren van de leerlingen of docenten. Elke filter match een deel van de invoer. Wildcards zijn niet nodig. Alle tekens met een speciale betekenis voor de "match"-operator zijn hierbij toegelaten.

Speciale betekenis hebben :
* '^' matcht het begin van een zoekterm 
* '$' matcht het eind van een zoekterm

### Uitvoer 

De voor School Data Sync geschikte uitvoer worden aangemaakt in de map 'Data_out'. Het script maakt volgens de specificaties van SDS de volgende bestanden aan. 

* School.csv
* Section.csv
* Student.csv
* StudentEnrollment.csv
* Teacher.csv
* TeacherRoster.csv

### Tussenresultaten

In de map 'Data_temp' worden de ongefilterde verzameling van ingelezen leerlingen en docenten opgeslagen in een bestand, elk met een deelverzameling van de attributen zoals die uit Magister worden gelezen. 
* docent.csv
* leerling.csv
