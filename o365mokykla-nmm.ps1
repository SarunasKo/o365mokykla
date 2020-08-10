#------------------------------------------------------------------------------------------------------------------
#
# MIT License
#
# Copyright (c) 2020 Sarunas Koncius
#
#
# Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated
# documentation files (the "Software"), to deal in the Software without restriction, including without limitation
# the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and
# to permit persons to whom the Software is furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all copies or substantial portions of
# the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO
# THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF
# CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER
# DEALINGS IN THE SOFTWARE.
#
#------------------------------------------------------------------------------------------------------------------
#
# PowerShell Source Code
#
# NAME:
#    o365mokykla-nmm.ps1
#
# AUTHOR:
#    Sarunas Koncius
#
# VERSION:
# 	 0.1.0 20200810
#
# MODIFIED:
#	 2020-08-10
#
#------------------------------------------------------------------------------------------------------------------


<#
	.SYNOPSIS
        PowerShell skriptas, skirtas atnaujinti mokyklos Office 365 aplinką naujiems mokslo metams.

	.DESCRIPTION
        PowerShell skriptas leidžia eksportuoti informaciją apie mokytojų, mokinių bei grupių paskyras į CSV
        failus, koreguoti paskyrų informaciją CSV failuose naudojant Excel programą ir įrašyti padarytus
        pakeitimus atgal į Office 365 aplinką. 

	.NOTES
        DĖMESIO! Šį PowerShell skriptą reikia atsidaryti Windows PowerShell ISE aplinkoje ir vykdyti ne visą
        skriptą, o tik pasirinktinai pažymėtas kodo eilutes!

#>


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 0: parengti PowerShell aplinką Office 365 paslaugų valdymui.
#
# Šiuos veiksmus kompiuteryje reikia atlikti vieną kartą prieš naudojant skriptą. Jeigu jūsų kompiuteryje 
# PowerShell jau buvo naudojamas Office 365 aplinkos valdymui anksčiau, tikėtina, kad šie veiksmai jau buvo
# atlikti. Jeigu dar neesate šių veiksmų atlikę savo kompiuteryje, prieš pradedant dirbti su šiuo skriptu, 
# įdiekite MSOnline modulį Azure AD valdymui ir leiskite vykdyti skriptus Exchange Online valdymui.
#
#------------------------------------------------------------------------------------------------------------------

# Startuokite Windows PowerShell ISE aplinką ar Windows PowerShell komandinės eilutės aplinką administratoriaus 
# teisėmis ir panaudokite šiame skripto žingsnyje esančias komandas. Pažymėkite kodo eilutę su pirmąją komanda
# "Install-Module -Name MSOnline" ir ją įvykdykite naudodami "Run Selection" mygtuką įrankių juostoje arba mygtuką
# "F8" klaviatūroje. Tuomet pažymėkite antrąją kodo eilutę "Set-ExecutionPolicy RemoteSigned" ir ją įvykdykite
# tokiu pačiu būdu.

Install-Module -Name MSOnline
Set-ExecutionPolicy RemoteSigned


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 1: PowerShell aplinkoje prisijungti prie Office 365 paslaugų naudojant visuotinio administratoriaus
# teises turinčią paskyrą.
#
# Šiuos veiksmus reikia atlikti kiekvieną kartą, kai atidarote PowerShell skriptą ir norite vykdyti komandas,
# kurios valdo Office 365 aplinką (Azure AD ir Exchange Online paslaugas).
#
#------------------------------------------------------------------------------------------------------------------

# Dialogo lange įvesti visuotinio admnistratoriaus teises turinčio vartotojo prisijungimo vardą ir slaptažodį
$UserCredential = Get-Credential
# Prisijungti prie Azure AD
connect-msolservice -credential $UserCredential
# Suformuoti Exchange Online valdymo sesiją
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
#Importuoti Exchange Online valdymo sesiją
Import-PSSession $Session -DisableNameChecking
# Aktyviu katalogu nustatyti darbastalį, kad jame būtų galima patogiai rasti ir saugoti CSV failus
Set-Location -Path $Env:USERPROFILE\OneDrive\Desktop


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 2: peržiūrėti turimas licencijas, rasti savo Office 365 aplinkos identifikatorių ir jį įrašyti šio
# skripto kode eilutės xxx, xxx, xxx ir xxx.
#
# Šį veiksmą reikia atlikti vieną kartą, pritaikant skriptą savo mokyklos Office 365 aplinkai.
#
#------------------------------------------------------------------------------------------------------------------

# Pažymėjus ir įvykdžius žemiau esančią komandą, PowerShell konsolės lange bus matomi maždaug tokie rezultatai:
#
# AccountSkuId                                 ActiveUnits WarningUnits ConsumedUnits
# ------------                                 ----------- ------------ -------------
# o365mokykla:STANDARDWOFFPACK_STUDENT         1000000     0            0         
# o365mokykla:STANDARDWOFFPACK_FACULTY         500000      0            1          
#
# Stulpelyje AccountSkuId prieš licencijos pavadinimą iki dvitaškio yra rodomas mokyklos Office 365 aplinkos
# identifikatorius. Šiame pavyzdyje matomas Office 365 aplinkos identifikatorius yra "o365mokykla" (be kabučių).

# Priskirkite reišmes kintamiesiems, kurie bus naudojami šiame skripte. Po to pažymėkite kodo bloką su 
# kintamaisiais ir įvykdykite šias kodo eilutes naudodami "Run Selection" mygtuką įrankių juostoje arba 
# mygtuką "F8" klaviatūroje.

Get-MsolAccountSku

# !!! SVARBU!
# !!! Savo mokyklos Office 365 aplinkos identifikatorių reikia įrašyti šio skripto eilutėse: 181 ir 253.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 3: pritaikyti skriptą naujiems mokslo metams.
#
# Kintamųjų reikšmes reikės modifikuoti kiekvieną kartą ruošiant Office 365 aplinką naujiems mokslo metams.
# Šio žingsnio kodą reikės įvykdyti kiekvieną kartą, kai norėsite naudotis skriptu. 
#
#------------------------------------------------------------------------------------------------------------------

# Priskirkite reišmes kintamiesiems, kurie bus naudojami šiame skripte. Po to pažymėkite kodo bloką su 
# kintamaisiais ir įvykdykite šias kodo eilutes naudodami "Run Selection" mygtuką įrankių juostoje arba 
# mygtuką "F8" klaviatūroje.

# Mokyklos naudojamas interneto domeno vardas
$Domeno_vardas = "eportfelis.net"

# Visuotinio administratoriaus teises turinčios paskyros e. pašto adresas
$VisuotinioAdministratoriausSmtpAdresas = "o365.administratorius@eportfelis.net"

# Ankstesnieji mokslo metai
$Ankstesnieji_mokslo_metai = "2019-2020"

# Naujieji mokslo metai
$Naujieji_mokslo_metai = "2020-2021"

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus mokytojų paskyrų informaciją
$Pradinis_mokytoju_paskyru_failas    = ".\o365mokykla_2020-2021_mokytojai_pradinis.csv"

# CSV failas su atnaujinta mokytojų paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_mokytoju_paskyru_failas = ".\o365mokykla_2020-2021_mokytojai_atnaujintas.csv" 

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus mokinių paskyrų informaciją
$Pradinis_mokiniu_paskyru_failas     = ".\o365mokykla_2020-2021_mokiniai_pradinis.csv" 

# CSV failas su atnaujinta mokytojų paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_mokiniu_paskyru_failas  = ".\o365mokykla_2020-2021_mokiniai_atnaujintas.csv" 

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus mokinių paskyrų informaciją
$Pradinis_grupiu_saraso_failas       = ".\o365mokykla_2020-2021_grupes_pradinis.csv"

# CSV failas su atnaujinta mokytojų paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_grupiu_saraso_failas    = ".\o365mokykla_2020-2021_grupes_atnaujintas.csv"


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 4: eksportuoti informaciją apie dabartines mokytojų paskyras iš Office 365 aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Atrinkti mokytojų paskyras pagal paskyroms priskirtą licenciją iš Office 365 aplinkos
$DabartinisMokytojuSarasas = Get-MsolUser -All |
    Where-Object { $_.Licenses.AccountSKUid -eq "o365mokykla:STANDARDWOFFPACK_FACULTY" } # <<< !!! Vietoje o365mokykla įrašykite savo mokyklos Office 365 aplinkos ID, kurį parodo Get-MsolAccountSku komanda !!!
# Eksportuoti atrinktų mokytojų paskyrų informaciją į CSV failą
$DabartinisMokytojuSarasas |
     Select UserPrincipalName, DisplayName, FirstName, LastName, Title, Department, City, Office |
     Export-CSV $Pradinis_mokytoju_paskyru_failas -Encoding UTF8


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 5: eksportuotą mokytojų paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite mokytojų paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Naudodami Excel programą atidarykite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas
#      skripto kintamajame $Pradinis_mokytoju_paskyru_failas (155 eilutė).
#
#   2. Išsaugokite CSV failą nauju vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#      $Atnaujintas_mokytoju_paskyru_failas (158 eilutėje).
#
#   3. Ištrinkite pirmąją duomenų eilutę, prasidedančią simboliais "#TYPE", kad anglų kalba nurodyti stulpelių
#      pavadinimai atsirastų pirmoje eilutėje.
#
#   4. Surikiuokite mokytojų paskyrų duomenis jums patogia tvarka, pavyzdžiui, pagal pavardę.
#
#   5. Peržiūrėkite mokytojų paskyrų informaciją, padarykite reikiamus pakeitimus bet kuriame stulpelyje, išskyrus
#      stulpelį "UserPrincipalName".
#
#   6. Visoms mokytojų paskyroms stulpelyje "Office" įrašykite naujuosius mokslo metus, kurie yra įrašyti skripto
#      kintamajame $Naujieji_mokslo_metai (152 eilutėje).
#
#   7. Paskyroms tų mokytojų, kurie nebedirbs mokykloje naujaisiais mokslo metais, "Office" stulpelyje mokslo metus
#      pakeiskite į praėjusius, kurie įrašyti skripto kintamajame $Ankstesnieji_mokslo_metai (149 eilutė).
#
#   8. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 6: importuoti atnaujintų mokytojų paskyrų informaciją iš CVS failo į Office 365 aplinką.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokytojų paskyrų informaciją iš CVS failo
$AtnaujintasMokytojuSarasas = Import-Csv $Atnaujintas_mokytoju_paskyru_failas -Encoding UTF8
# Atnaujintą informaciją įrašyti į mokytojų paskyras, esančias Office 365 aplinkoje
$AtnaujintasMokytojuSarasas |
    foreach { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.Title -Department $_.Department -City $_.City -Office $_.Office }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 7: Office 365 aplinkoje blokuoti prisijungimą tų mokytojų paskyroms, kurioms "Office" stulpelyje nėra
# nurodyti naujieji mokslo metai.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokytojų paskyrų informaciją iš CVS failo
$AtnaujintasMokytojuSarasas = Import-Csv $Atnaujintas_mokytoju_paskyru_failas -Encoding UTF8
# Blokuoti prisijungimą tų mokytojų paskyroms, kurioms "Office" stulpelyje nėra nurodyti naujieji mokslo metai
$AtnaujintasMokytojuSarasas | 
    foreach {if ($_.Office -ne $Naujieji_mokslo_metai -and (Get-MsolUserRole -UserPrincipalName $_.UserPrincipalName) -eq $Null) { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -BlockCredential $true } }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 8: eksportuoti informaciją apie dabartines mokinių paskyras iš Office 365 aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Atrinkti mokinių paskyras pagal paskyroms priskirtą licenciją iš Office 365 aplinkos
$DabartinisMokiniuSarasas = Get-MsolUser -All | Where-Object {$_.Licenses.AccountSKUid -eq "o365mokykla:STANDARDWOFFPACK_STUDENT"} # <<< !!! Vietoje o365mokykla įrašykite savo mokyklos Office 365 aplinkos ID, kurį parodo Get-MsolAccountSku komanda !!!
# Eksportuoti atrinktų mokinių paskyrų informaciją į CSV failą
$DabartinisMokiniuSarasas | 
    Select UserPrincipalName, DisplayName, FirstName, LastName, Title, Department, City, Office | Export-CSV $Pradinis_mokiniu_paskyru_failas -Encoding UTF8


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 9: eksportuotą mokinių paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite mokinių paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Naudodami Excel programą atidarykite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas
#      skripto kintamajame $Pradinis_mokiniu_paskyru_failas (161 eilutėje).
#
#   2. Išsaugokite CSV failą nauju vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#      $Atnaujintas_mokiniu_paskyru_failas (164 eilutėje).
#
#   3. Ištrinkite pirmąją duomenų eilutę, parasidedančią simboliais "#TYPE", kad anglų kalba nurodyti stulpelių
#      pavadinimai atsirastų pirmoje eilutėje.
#
#   4. Surikiuokite paskyrų duomenis, kad vienos klasės mokiniai būtų vienoje grupėje, o klasės būtų surikiuotos
#      mažejimo tvarka.
#
#   5. Search/Replace vaiksmais pervadinkite alumnais vyriausių klasių mokinių pareigas ir jų klasės informaciją.
#
#   6. Eilės tvarka (!) nuo vyriausių iki jauniausių klasių Search/Replace veiksmais atnaujinkite mokinių pareigas
#      ir klasės informaciją.
#
#   7. Peržiūrėkite mokinių paskyrų informaciją, padarykite reikiamus pakeitimus bet kuriame stulpelyje, išskyrus
#      stulpelį "UserPrincipalName".
#
#   8. Visoms mokinių paskyroms stulpelyje "Office" įrašykite naujuosius mokslo metus, kurie yra įrašyti skripto
#      kintamajame $Naujieji_mokslo_metai (152 eilutėje).
#
#   9. Paskyroms tų mokinių, kurie nebesimokys mokykloje naujaisiais mokslo metais, "Office" stulpelyje mokslo
#      metus pakeiskite į praėjusius, kurie įrašyti skripto kintamajame $Ankstesnieji_mokslo_metai (149 eilutėje).
#
#   10. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 10: importuoti atnaujintų mokinių paskyrų informaciją iš CVS failo į Office 365 aplinką.
# 
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintasMokiniuSarasas = Import-Csv $Atnaujintas_mokiniu_paskyru_failas -Encoding UTF8
# Atnaujintą informaciją įrašyti į mokinių paskyras, esančias Office 365 aplinkoje
$AtnaujintasMokiniuSarasas | 
    foreach { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.Title -Department $_.Department -City $_.City -Office $_.Office }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 11: Office 365 aplinkoje blokuoti prisijungimą tų mokinių paskyroms, kurioms "Office" stulpelyje nėra
# nurodyti naujieji mokslo metai.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintasMokytojuSarasas = Import-Csv $Atnaujintas_mokiniu_paskyru_failas -Encoding UTF8
# Blokuoti prisijungimą tų mokinių paskyroms, kurioms "Office" stulpelyje nėra nurodyti naujieji mokslo metai
$AtnaujintasMokytojuSarasas | 
    foreach {if ($_.Office -ne $Naujieji_mokslo_metai -and (Get-MsolUserRole -UserPrincipalName $_.UserPrincipalName) -eq $Null) { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -BlockCredential $true } }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 12: eksportuoti informaciją apie dabartines saugos grupių su įgalintu e. paštu paskyras iš Office 365
# aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Išrinkti saugos grupių su įgalintu e. paštu paskyras ir eksportuoti informaciją apie jas į CSV failą
Get-DistributionGroup -ResultSize unlimited -Filter "RecipientTypeDetails -eq 'MailUniversalSecurityGroup'" |
    Select-Object Guid, Identity, Id, Name, DisplayName, Alias, EmailAddresses, PrimarySmtpAddress, WindowsEmailAddress |
    Export-Csv $Pradinis_grupiu_saraso_failas -Encoding UTF8


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 13: eksportuotą grupių paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite mokinių paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Naudodami Excel programą atidarykite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas
#      skripto kintamajame $Pradinis_grupiu_saraso_failas (167 eilutėje).
#
#   2. Išsaugokite CSV failą nauju vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#      $Atnaujintas_grupiu_saraso_failas (170 eilutėje).
#
#   3. Ištrinkite pirmąją duomenų eilutę, parasidedančią simboliais "#TYPE", kad anglų kalba nurodyti stulpelių
#      pavadinimai atsirastų pirmoje eilutėje.
#
#   4. Ištrinkite eilutes su grupių paskyromis, kurio nėra susijusios su klasėmis.
#
#   5. Surikiuokite grupių paskyrų duomenis klasių mažejimo tvarka.
#
#   6. Search/Replace vaiksmais pervadinkite alumnais vyriausių klasių grupes.
#
#   7. Eilės tvarka (!) nuo vyriausių iki jauniausių klasių Search/Replace veiksmais arba tiesiog koreguodami
#      duomenis, atnaujinkite grupių informaciją, pakeisdami *ir* jų pavadinimus, *ir* e. pašo adresą. 
#
#   8. Peržiūrėkite grupių paskyrų informaciją, padarykite reikiamus pakeitimus bet kuriame stulpelyje, išskyrus
#      stulpelį "Guid".
#
#   9. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 14: importuoti atnaujintų grupių paskyrų informaciją iš CVS failo į Office 365 aplinką.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintosGrupes = Import-Csv $Atnaujintas_grupiu_saraso_failas -Encoding UTF8
# Atnaujintą informaciją įrašyti į grupių paskyras, esančias Office 365 aplinkoje
$AtnaujintosGrupes |
    foreach { Set-DistributionGroup -Identity $_.Guid -Name $_.Name -DisplayName $_.DisplayName -Alias $_.Alias -EmailAddresses $_.PrimarySmtpAddress -IgnoreNamingPolicy }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 15: prieš uždarant šį PowerShell skripto failą arba Windows PowerShell ISE programą, uždaryti Exchange
# Online paslaugos valdymo sesiją.
#
# Šį veiksmą rekomenduojama atlikti kiekvieną kartą, kai atidarote uždarote šį PowerShell skriptą arba uždarote
# Windows PowerShell ISE programą.
#
#------------------------------------------------------------------------------------------------------------------

# Uždaryti Exchange Online paslaugos valdymo sesiją
Remove-PSSession $Session

#------------------------------------------------------------------------------------------------------------------
#
# PowerShell skripto pabaiga.
#
#------------------------------------------------------------------------------------------------------------------
