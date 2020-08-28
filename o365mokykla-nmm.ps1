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
# 	 0.9.5 20200828
#
# MODIFIED:
#	 2020-08-28
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
# Žingsnis 00: parengti PowerShell aplinką Office 365 paslaugų valdymui.
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
# Žingsnis 01: PowerShell aplinkoje prisijungti prie Office 365 paslaugų naudojant visuotinio administratoriaus
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

# Aktyviu katalogu nustatyti darbalaukį, kad jame būtų galima patogiai rasti ir saugoti CSV failus
Set-Location -Path $Env:USERPROFILE\Desktop


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 02: peržiūrėti turimas licencijas, rasti savo Office 365 aplinkos identifikatorių ir jį įrašyti šio
# skripto kode eilutėse 313 ir 494.
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
# !!! Savo mokyklos Office 365 aplinkos identifikatorių reikia įrašyti šio skripto eilutėse: 313 ir 494.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 03: pritaikyti skriptą naujiems mokslo metams.
#
# Kintamųjų reikšmes reikės modifikuoti kiekvieną kartą ruošiant Office 365 aplinką naujiems mokslo metams.
# Šio žingsnio kodą reikės įvykdyti kiekvieną kartą, kai norėsite naudotis skriptu. 
#
#------------------------------------------------------------------------------------------------------------------

# Priskirkite reišmes kintamiesiems, kurie bus naudojami šiame skripte. Po to pažymėkite kodo bloką su 
# kintamaisiais ir įvykdykite šias kodo eilutes naudodami "Run Selection" mygtuką įrankių juostoje arba 
# mygtuką "F8" klaviatūroje.

# Mokyklos naudojamas interneto domeno vardas
$Domeno_vardas = "o365mokykla.lt"

# Visuotinio administratoriaus teises turinčios paskyros e. pašto adresas
$VisuotinioAdministratoriausSmtpAdresas = "o365.administratorius@o365mokykla.lt"

# Mokytojų saugos grupės e. pašto adresas
$GrupesVisiMokytojaiSmtpAdresas = "visi.mokytojai@o365mokykla.lt"

# Ankstesnieji mokslo metai
$Ankstesnieji_mokslo_metai = "2019-2020"

# Naujieji mokslo metai
$Naujieji_mokslo_metai = "2020-2021"

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus mokytojų paskyrų informaciją
$Pradinis_mokytoju_paskyru_failas    = ".\o365mokykla_2020-2021_mokytojai_pradinis.csv"

# CSV failas su atnaujinta mokytojų paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_mokytoju_paskyru_failas = ".\o365mokykla_2020-2021_mokytojai_atnaujintas.csv" 

# Jūsų parengtas CSV failas, kuriame yra surašyta informacija apie naujus mokytojus
$Nauju_mokytoju_saraso_failas        = ".\o365mokykla_2020-2021_mokytojai_nauji.csv"

# CSV failas su naujų mokytojų paskyrų informacija, kurią sukurs skriptas
$Nauju_mokytoju_paskyru_failas       = ".\o365mokykla_2020-2021_mokytojai_nauji_paskyros.csv" 

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus mokinių paskyrų informaciją
$Pradinis_mokiniu_paskyru_failas     = ".\o365mokykla_2020-2021_mokiniai_pradinis.csv" 

# CSV failas su atnaujinta mokinių paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_mokiniu_paskyru_failas  = ".\o365mokykla_2020-2021_mokiniai_atnaujintas.csv" 

# Jūsų parengtas CSV failas, kuriame yra surašyta informacija apie naujus mokinius
$Nauju_mokiniu_saraso_failas        = ".\o365mokykla_2020-2021_mokiniai_nauji.csv"

# CSV failas su naujų mokinių paskyrų informacija, kurią sukurs skriptas
$Nauju_mokiniu_paskyru_failas       = ".\o365mokykla_2020-2021_mokiniai_nauji_paskyros.csv" 

# CSV failas, kurį sukurs PowerShell skriptas, eksportavus grupių paskyrų informaciją
$Pradinis_grupiu_saraso_failas       = ".\o365mokykla_2020-2021_grupes_pradinis.csv"

# CSV failas su atnaujinta grupių paskyrų informacija, kurį parengsite Excel programoje
$Atnaujintas_grupiu_saraso_failas    = ".\o365mokykla_2020-2021_grupes_atnaujintas.csv"


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 04: eksportuoti informaciją apie dabartines mokytojų paskyras iš Office 365 aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Atrinkti mokytojų paskyras pagal paskyroms priskirtą licenciją iš Office 365 aplinkos
$DabartinisMokytojuSarasas = Get-MsolUser -All | Where-Object { $_.Licenses.AccountSKUid -like "*STANDARDWOFFPACK_FACULTY*" }

# Eksportuoti atrinktų mokytojų paskyrų informaciją į CSV failą
$DabartinisMokytojuSarasas | Select UserPrincipalName, DisplayName, FirstName, LastName, Title, Department, City, Office | Export-CSV $Pradinis_mokytoju_paskyru_failas -Encoding UTF8 -Delimiter ";"


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 05: eksportuotą mokytojų paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite mokytojų paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Naudodami Excel programą atidarykite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas
#      skripto kintamajame $Pradinis_mokytoju_paskyru_failas (162 skripto eilutėje).
#
#   2. Išsaugokite CSV failą nauju vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#      $Atnaujintas_mokytoju_paskyru_failas (165 skripto eilutėje).
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
#      kintamajame $Naujieji_mokslo_metai (159 skripto eilutėje).
#
#   7. Paskyroms tų mokytojų, kurie nebedirbs mokykloje naujaisiais mokslo metais, "Office" stulpelyje mokslo metus
#      pakeiskite į praėjusius, kurie įrašyti skripto kintamajame $Ankstesnieji_mokslo_metai (156 skripto eilutėje).
#
#   8. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 06: importuoti atnaujintų mokytojų paskyrų informaciją iš CVS failo į Office 365 aplinką.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokytojų paskyrų informaciją iš CVS failo
$AtnaujintasMokytojuSarasas = Import-Csv $Atnaujintas_mokytoju_paskyru_failas -Encoding UTF8 -Delimiter ";"

# Atnaujintą informaciją įrašyti į mokytojų paskyras, esančias Office 365 aplinkoje
$AtnaujintasMokytojuSarasas | foreach { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.Title -Department $_.Department -City $_.City -Office $_.Office }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 07: Office 365 aplinkoje blokuoti prisijungimą tų mokytojų paskyroms, kurioms "Office" stulpelyje nėra
# nurodyti naujieji mokslo metai.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokytojų paskyrų informaciją iš CVS failo
$AtnaujintasMokytojuSarasas = Import-Csv $Atnaujintas_mokytoju_paskyru_failas -Encoding UTF8 -Delimiter ";"

# Blokuoti prisijungimą tų mokytojų paskyroms, kurioms "Office" stulpelyje nėra nurodyti naujieji mokslo metai ir
# vartotojo paskyra neturi jai priskirtos administratoriaus rolės.
$AtnaujintasMokytojuSarasas | foreach {if ($_.Office -ne $Naujieji_mokslo_metai -and (Get-MsolUserRole -UserPrincipalName $_.UserPrincipalName) -eq $Null) { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -BlockCredential $true } }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 08: Sukurti paskyras naujiems mokytojams
#
#------------------------------------------------------------------------------------------------------------------

# Patikrinti, ar CSV failas su mokytojų sąrašu yra tinkamas.
# Įvykdžius kodą turi būti rodomi mokytojų sąrašo duomenys trijuose stulpeliuose. Jeigu duomenys matosi viename 
# stulpelyje, CSV faile skyrybos ženklą kablelį pakeiskite kabliataškiu arba atvirkščiai. Stulpelių pavadinimai
# turi būti "Pavardė", "Vardas" ir "Pareigos", bet jų eilės tvarka nėra svarbi. Pataisykite CSV failą, jeigu
# stulpelių pavadinimai yra kiti. Pakoregavę CSV failą, grįžkite prie 275 skripto eilutės CSV failui patikrinti.
$NaujuMokytojuSarasas = Import-Csv $Nauju_mokytoju_saraso_failas -Encoding UTF8 -Delimiter ";"
$NaujuMokytojuSarasas | ft

# Sukurti naujas mokytojų paskyras, naudojant informaciją iš naujų mokytojo sąrašo ir formuojant CSV failą su
# naujų mokytojų paskyrų duomenimis - prisijungimo vardu ir slaptažodžiu.
function Remove-StringNonLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
Out-File -FilePath $Nauju_mokytoju_paskyru_failas -InputObject "Mokytojas;VartotojoID;Slaptažodis" -Encoding UTF8
$NaujuMokytojuSarasas = Import-Csv $Nauju_mokytoju_saraso_failas -Encoding UTF8 -Delimiter ";"
$Licencijos = Get-MsolAccountSku
foreach ($Licencija in $Licencijos) {
    if ($Licencija.AccountSkuId -like "*STANDARDWOFFPACK_FACULTY*") { $YraO365Licenciju = $Licencija.ActiveUnits - $Licencija.ConsumedUnits }
}
if ( $YraO365Licenciju -lt $NaujuMokytojuSarasas.Count ) { throw "Nepakanka licencijų." }
foreach ($NaujasMokytojas in $NaujuMokytojuSarasas) {
    $NewFirstName = (Get-Culture).textinfo.totitlecase($NaujasMokytojas.Vardas.ToLower())
    $NewLastName = (Get-Culture).textinfo.totitlecase($NaujasMokytojas.Pavardė.ToLower())
    $NewTitle = $NaujasMokytojas.Pareigos
    $NewOffice = $Naujieji_mokslo_metai
    If ($NewFirstName.Contains(" ")) {
        $OnlyFirstName = $NewFirstName.Substring(0, $NewFirstName.IndexOf(" ")) 
    } else { 
        $OnlyFirstName = $NewFirstName 
    }
    If ($NewLastName.Contains(" ")) { 
        $OnlyLastName = $NewLastName.Substring($NewLastName.LastIndexOf(" ")+1,$NewLastName.Length-$NewLastName.LastIndexOf(" ")-1) 
    } else { 
        $OnlyLastName = $NewLastName 
    }
    $NewDisplayName = $OnlyFirstName + " " + $OnlyLastName
    $NewUserPrincipalName = (Remove-StringNonLatinCharacters $OnlyFirstName.ToLower()) + "." + (Remove-StringNonLatinCharacters $OnlyLastName.ToLower()) + "@" + $Domeno_vardas
    Echo $NewUserPrincipalName
    $EsamasVartotojas = Get-MsolUser -UserPrincipalName $NewUserPrincipalName -ErrorAction SilentlyContinue
    If ($EsamasVartotojas -eq $Null) {
		New-MsolUser -UserPrincipalName $NewUserPrincipalName -DisplayName $NewDisplayName -FirstName $NewFirstName -LastName $NewLastName -Title $NewTitle -Office $NewOffice -PreferredLanguage "lt-LT" -UsageLocation "LT" -ForceChangePassword:$true
        $Slaptazodis = Set-MsolUserPassword -UserPrincipalName $NewUserPrincipalName -ForceChangePassword:$true
        Set-MsolUserLicense -UserPrincipalName $NewUserPrincipalName -AddLicenses "o365mokykla:STANDARDWOFFPACK_FACULTY"
# !!! -----------------------------------------------------------------------------^^^^^^^^^^^------------------ !!!
# !!! Vietoje o365mokykla įrašykite savo mokyklos Office 365 aplinkos ID, kurį parodo Get-MsolAccountSku komanda !!!

    } else {
        $Slaptazodis = "!!!PasikartojantisVartotojoID!!!"
    }
    $Mokytojas = $NewFirstName + " " + $NewLastName
    $VartotojoID = $NewUserPrincipalName
	$PrisijungimoInformacija = "$Mokytojas;$VartotojoID;$Slaptazodis"
    Out-File -FilePath $Nauju_mokytoju_paskyru_failas -InputObject $PrisijungimoInformacija -Encoding UTF8 -append
}

# Nustatyti lietuviškus Office 365 aplinkos ir e. pašto dėžutės parametrus naujoms mokytojų paskyroms. 
# Prieš vykdant šį kodo bloką, Office 365 administratoriaus portale įsitikinkite, kad paskutinėms sukurtoms naujų
# mokytojų paskyroms jau yra sukurtos e.pašto dėžutės.
$NaujosMokytojuPaskyros = Import-Csv $Nauju_mokytoju_paskyru_failas -Encoding UTF8 -Delimiter ";"
$Skaitliukas = 1
foreach ($NaujaMokytojoPaskyra In $NaujosMokytojuPaskyros) {
	$Upn = $NaujaMokytojoPaskyra.VartotojoID
    Echo $Upn
	$ActivityMessage = "Prašome palaukti..."
	$StatusMessage = ("Nustatomi parametrai vartotojui {0} ({1} iš {2})" -f $Upn, $Skaitliukas, @($NaujosMokytojuPaskyros).count)
	$PercentComplete = ($Skaitliukas / @($NaujosMokytojuPaskyros).count * 100)
	Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
	set-MailboxRegionalConfiguration -Identity $NaujaMokytojoPaskyra.VartotojoID -TimeZone "FLE Standard Time" -Language lt-LT –LocalizeDefaultFolderName
    If ($Skaitliukas % 50 -eq 0) { Start-Sleep -m 750 }
	$Skaitliukas++
}

# Įtraukti naujų mokytojų paskyras į saugos grupę "Visi mokytojai"
$NaujosMokytojuPaskyros = Import-Csv $Nauju_mokytoju_paskyru_failas -Encoding UTF8 -Delimiter ";"
$NaujosMokytojuPaskyros | foreach { Add-DistributionGroupMember -Identity $GrupesVisiMokytojaiSmtpAdresas -Member $_.VartotojoID -Confirm:$false -BypassSecurityGroupManagerCheck }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 9: eksportuoti informaciją apie dabartines mokinių paskyras iš Office 365 aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Atrinkti mokinių paskyras pagal paskyroms priskirtą licenciją iš Office 365 aplinkos
$DabartinisMokiniuSarasas = Get-MsolUser -All | Where-Object {$_.Licenses.AccountSKUid -like "*STANDARDWOFFPACK_STUDENT*"} 

# Eksportuoti atrinktų mokinių paskyrų informaciją į CSV failą
$DabartinisMokiniuSarasas | Select UserPrincipalName, DisplayName, FirstName, LastName, Title, Department, City, Office | Export-CSV $Pradinis_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";"


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 10: eksportuotą mokinių paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite mokinių paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Startuokite "Excel" programą ir pradėkite darbą nuo naujo tuščio dokumento.
#
#   2. Kortelėje "Duomenys" ("Data") pasirinkite komandą "Iš teksto/CSV" ("From Text/CSV").
#
#   3. Parinkite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas skripto kintamajame 
#      $Pradinis_mokiniu_paskyru_failas (174 skripto eilutėje), ir nuspauskite mygtuką "Importuoti" ("Import").
#
#   4. Atsidariusiame dialogo lange nuspauskite mygtuką "Transformuoti duomenis" ("Transform data").
#
#   5. Atsidariusiame "Power Query" lange nuspauskite mygtuką "Išplėstinė rengyklė" ("Advanced Editor").
#
#   6. Jeigu "Iplėstinė rengyklė" ("Advanced Editor") lange rodomas tekstas neturi eilučių numerių, įjunkite eilučių numerių rodymą 
#      pasirinkę meniu "Rodymo parinkty" ("Display Options") ir komandą "Rodyti eilučių numerius" ("Display line numbers").
#
#   7. Pažymėkite "Iplėstinė rengyklė" ("Advanced Editor") lange esantį tekstą nuo trečios iki paskutinės eilutės imtinai ir jį ištrinkite.  
#
#   8. Atidarykite darbalaukyje esantį failą o365mokykla_PowerQueryLT_mokiniai.txt arba o365mokykla_PowerQueryEN_mokiniai.txt,
#      pažymėkite jame esantį tekstą nuo trečios iki paskutinės eilutės imtinai ir jį nukopijuokite.
#
#   9. Sugrįžkite į "Iplėstinė rengyklė" ("Advanced Editor") langą, pastatykite kursorių trečioje eilutėje,
#      įkelkite nukopijuotą tekstą ir nuspauskite mygtuką "Atlikta" ("Done"). 
#
#   10. "Power Query" lange nuspauskite mygtuką "Uždaryti ir įkelti" ("Close & Load").
#
#   11. Išsaugokite failą darbalaukyje vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#       $Atnaujintas_mokiniu_paskyru_failas (177 skripto eilutėje), parinkdami failo tipą "CSV UTF-8 (Comma delimited)".
#       Atsidariusiame dialogo lange nuspauskite mygtuką "OK", kad būtų išsaugotas tik aktyvus lapas.
#
#   12. Search/Replace vaiksmais pervadinkite alumnais vyriausių klasių mokinių pareigas ir jų klasės informaciją.
#
#   13. Peržiūrėkite mokinių paskyrų informaciją, padarykite reikiamus pakeitimus bet kuriame stulpelyje, išskyrus
#       stulpelį "UserPrincipalName".
#
#   14. Visoms mokinių paskyroms stulpelyje "Office" įrašykite naujuosius mokslo metus, kurie yra įrašyti skripto
#       kintamajame $Naujieji_mokslo_metai (159 skripto eilutėje).
#
#   15. Paskyroms tų mokinių, kurie nebesimokys mokykloje naujaisiais mokslo metais, "Office" stulpelyje mokslo
#       metus pakeiskite į praėjusius, kurie įrašyti skripto kintamajame $Ankstesnieji_mokslo_metai (156 skripto
#       eilutėje).
#
#   16. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 11: importuoti atnaujintų mokinių paskyrų informaciją iš CVS failo į Office 365 aplinką.
# 
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintasMokiniuSarasas = Import-Csv $Atnaujintas_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";"

# Atnaujintą informaciją įrašyti į mokinių paskyras, esančias Office 365 aplinkoje
$AtnaujintasMokiniuSarasas | foreach { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -DisplayName $_.DisplayName -FirstName $_.FirstName -LastName $_.LastName -Title $_.Title -Department $_.Department -City $_.City -Office $_.Office }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 12: Office 365 aplinkoje blokuoti prisijungimą tų mokinių paskyroms, kurioms "Office" stulpelyje nėra
# nurodyti naujieji mokslo metai.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintasMokiniuSarasas = Import-Csv $Atnaujintas_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";"

# Blokuoti prisijungimą tų mokinių paskyroms, kurioms "Office" stulpelyje nėra nurodyti naujieji mokslo metai ir
# vartotojo paskyra neturi jai priskirtos administratoriaus rolės.
$AtnaujintasMokiniuSarasas | foreach {if ($_.Office -ne $Naujieji_mokslo_metai -and (Get-MsolUserRole -UserPrincipalName $_.UserPrincipalName) -eq $Null) { Set-MsolUser -UserPrincipalName $_.UserPrincipalName -BlockCredential $true } }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 13: sukurti paskyras naujiems mokiniams
#
#------------------------------------------------------------------------------------------------------------------

# Patikrinti, ar CSV failas su mokinių sąrašu yra tinkamas.
# Įvykdžius kodą turi būti rodomi mokinių sąrašo duomenys trijuose stulpeliuose. Jeigu duomenys matosi viename 
# stulpelyje, CSV faile skyrybos ženklą kablelį pakeiskite kabliataškiu arba atvirkščiai. Stulpelių pavadinimai
# turi būti "Pavardė", "Vardas" ir "Klasės/grupės pavadinimas", bet jų eilės tvarka nėra svarbi. Pataisykite CSV
# failą, jeigu stulpelių pavadinimai yra kiti. Pakoregavę CSV failą, grįžkite prie skripto 452-453 eilučių CSV 
# failui patikrinti.
$NaujuMokiniuSarasas = Import-Csv $Nauju_mokiniu_saraso_failas -Encoding UTF8 -Delimiter ";"
$NaujuMokiniuSarasas | ft

# Sukurti naujas mokiniu paskyras, naudojant informaciją iš naujų mokiniu sąrašo ir formuojant CSV failą su
# naujų mokinių paskyrų duomenimis - prisijungimo vardu ir slaptažodžiu.
function Remove-StringNonLatinCharacters {
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
Out-File -FilePath $Nauju_mokiniu_paskyru_failas -InputObject "Klasė;Mokinys;VartotojoID;Slaptažodis" -Encoding UTF8
$NaujuMokiniuSarasas = Import-Csv $Nauju_mokiniu_saraso_failas -Encoding UTF8 -Delimiter ";"
$Licencijos = Get-MsolAccountSku
foreach ($Licencija in $Licencijos) {
    if ($Licencija.AccountSkuId -like "*STANDARDWOFFPACK_STUDENT*") { $YraO365Licenciju = $Licencija.ActiveUnits - $Licencija.ConsumedUnits }
}
if ( $YraO365Licenciju -lt $NaujuMokiniuSarasas.Count ) { throw "Nepakanka licencijų." }
foreach ($NaujasMokinys in $NaujuMokiniuSarasas) {
    $NewFirstName = (Get-Culture).textinfo.totitlecase($NaujasMokinys.Vardas.ToLower())
    $NewLastName = (Get-Culture).textinfo.totitlecase($NaujasMokinys.Pavardė.ToLower())
    $NewDepartment = $NaujasMokinys."Klasės/grupės pavadinimas".ToLower() + " klasė"
    $NewOffice = $Naujieji_mokslo_metai
    If ($NewFirstName.Contains(" ")) {
        $OnlyFirstName = $NewFirstName.Substring(0, $NewFirstName.IndexOf(" ")) 
    } else { 
        $OnlyFirstName = $NewFirstName 
    }
    If ($NewLastName.Contains(" ")) { 
        $OnlyLastName = $NewLastName.Substring($NewLastName.LastIndexOf(" ")+1, $NewLastName.Length-$NewLastName.LastIndexOf(" ")-1) 
    } else { 
        $OnlyLastName = $NewLastName 
    }
    if ($OnlyFirstName.EndsWith("s")) { 
        $NewTitle = $NewDepartment + "s mokinys" 
    } else { 
        $NewTitle = $NewDepartment + "s mokinė" 
    }
    $NewDisplayName = $OnlyFirstName + " " + $OnlyLastName
    $NewUserPrincipalName = (Remove-StringNonLatinCharacters $OnlyFirstName.ToLower()) + "." + (Remove-StringNonLatinCharacters $OnlyLastName.ToLower()) + "@" + $Domeno_vardas
    Echo $NewUserPrincipalName
    $EsamasVartotojas = Get-MsolUser -UserPrincipalName $NewUserPrincipalName -ErrorAction SilentlyContinue
    If ($EsamasVartotojas -eq $Null) {
		New-MsolUser -UserPrincipalName $NewUserPrincipalName -DisplayName $NewDisplayName -FirstName $NewFirstName -LastName $NewLastName -Title $NewTitle -Department $NewDepartment -Office $NewOffice -PreferredLanguage "lt-LT" -UsageLocation "LT" -ForceChangePassword:$true
        $Slaptazodis = Set-MsolUserPassword -UserPrincipalName $NewUserPrincipalName -ForceChangePassword:$true
        Set-MsolUserLicense -UserPrincipalName $NewUserPrincipalName -AddLicenses "o365mokykla:STANDARDWOFFPACK_STUDENT"
# !!! -----------------------------------------------------------------------------^^^^^^^^^^^------------------ !!!
# !!! Vietoje o365mokykla įrašykite savo mokyklos Office 365 aplinkos ID, kurį parodo Get-MsolAccountSku komanda !!!

    } else {
        $Slaptazodis = "!!!PasikartojantisVartotojoID!!!"
    }
    $Klase = $NaujasMokinys."Klasės/grupės pavadinimas".ToLower()
    $Mokinys = $NewFirstName + " " + $NewLastName
    $VartotojoID = $NewUserPrincipalName
	$PrisijungimoInformacija = "$Klase;$Mokinys;$VartotojoID;$Slaptazodis"
    Out-File -FilePath $Nauju_mokiniu_paskyru_failas -InputObject $PrisijungimoInformacija -Encoding UTF8 -append
}

# Nustatyti lietuviškus Office 365 aplinkos ir e. pašto dėžutės parametrus naujoms mokinių paskyroms. 
# Prieš vykdant šį kodo bloką, Office 365 administratoriaus portale įsitikinkite, kad paskutinėms sukurtoms naujų
# mokinių paskyroms jau yra sukurtos e.pašto dėžutės.
$NaujosMokiniuPaskyros = Import-Csv $Nauju_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";"
$Skaitliukas = 1
foreach ($NaujaMokinioPaskyra In $NaujosMokiniuPaskyros) {
	$Upn = $NaujaMokinioPaskyra.VartotojoID
    Echo $Upn
	$ActivityMessage = "Prašome palaukti..."
	$StatusMessage = ("Nustatomi parametrai vartotojui {0} ({1} iš {2})" -f $Upn, $Skaitliukas, @($NaujosMokiniuPaskyros).count)
	$PercentComplete = ($Skaitliukas / @($NaujosMokiniuPaskyros).count * 100)
	Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
	set-MailboxRegionalConfiguration -Identity $NaujaMokinioPaskyra.VartotojoID -TimeZone "FLE Standard Time" -Language lt-LT –LocalizeDefaultFolderName
    If ($Skaitliukas % 50 -eq 0) { Start-Sleep -m 750 }
	$Skaitliukas++
}


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 14: eksportuoti informaciją apie dabartines saugos grupių su įgalintu e. paštu paskyras iš Office 365
# aplinkos į CSV failą.
#
#------------------------------------------------------------------------------------------------------------------

# Išrinkti saugos grupių su įgalintu e. paštu paskyras ir eksportuoti informaciją apie jas į CSV failą
Get-DistributionGroup -ResultSize unlimited -Filter "RecipientTypeDetails -eq 'MailUniversalSecurityGroup'" |
    Select-Object Guid, Identity, Id, Name, DisplayName, Alias, EmailAddresses, PrimarySmtpAddress, WindowsEmailAddress |
    Export-Csv $Pradinis_grupiu_saraso_failas -Encoding UTF8 -Delimiter ";"


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 15: eksportuotą grupių paskyrų informaciją tvarkyti ir atnaujinti Excel programoje.
#
#------------------------------------------------------------------------------------------------------------------

# Atnaujinkite grupių paskyrų duomenis, atlikdami šiuos žingsnius:
#
#   1. Startuokite "Excel" programą ir pradėkite darbą nuo naujo tuščio dokumento.
#
#   2. Kortelėje "Data" pasirinkite komandą "From Text/CSV".
#
#   3. Parinkite PowerShell skripto sukurtą CSV failą, kurio pavadinimas yra nurodytas skripto kintamajame 
#      $Pradinis_grupių_paskyru_failas (186 skripto eilutėje), ir nuspauskite mygtuką "Import".
#
#   4. Atsidariusiame dialogo lange nuspauskite mygtuką "Transform data".
#
#   5. Atsidariusiame "Power Query Editor" lange nuspauskite mygtuką "Advanced Editor".
#
#   6. Jeigu "Advanced Editor" lange rodomas tekstas neturi eilučių numerių, įjunkite eilučių numerių rodymą 
#      pasirinkę meniu "Display Options" ir komandą "Display line numbers".
#
#   7. Pažymėkite "Advanced Editor" lange esantį tekstą nuo trečios iki paskutinės eilutės imtinai ir jį ištrinkite.  
#
#   8. Atidarykite darbalaukyje esantį failą o365mokykla_PowerQuery_grupes.txt, pažymėkite jame esantį tekstą
#      nuo trečios iki paskutinės eilutės imtinai ir jį nukopijuokite.
#
#   9. Sugrįžkite į "Advanced Editor" langą, pastatykite kursorių trečioje eilutėje, įkelkite nukopijuotą tekstą 
#      ir nuspauskite mygtuką "Done". 
#
#   10. "Power Query Editor" lange nuspauskite mygtuką "Close & Load".
#
#   11. Išsaugokite failą darbalaukyje vardu, kurį nurodėte atnaujintų paskyrų failo kintamajame
#       $Atnaujintas_grupiu_paskyru_failas (189 skripto eilutėje), parinkdami failo tipą "CSV UTF-8 (Comma delimited)".
#       Atsidariusiame dialogo lange nuspauskite mygtuką "OK", kad būtų išsaugotas tik aktyvus lapas.
#
#   12. Search/Replace vaiksmais pervadinkite alumnais vyriausių klasių grupes, pavyzdžiui, 13a klasę.
#
#   13. Peržiūrėkite grupių paskyrų informaciją, padarykite reikiamus pakeitimus bet kuriame stulpelyje, išskyrus
#       stulpelį "Guid".
#
#   14. Ištrinkite eilutes su grupių paskyromis, kurio nėra susijusios su klasių saugos grupėmis.
#
#   15. Išsaugokite CSV faile atliktus pakeitimus.


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 16: importuoti atnaujintų grupių paskyrų informaciją iš CVS failo į Office 365 aplinką.
#
#------------------------------------------------------------------------------------------------------------------

# Nuskaityti atnaujintų mokinių paskyrų informaciją iš CVS failo
$AtnaujintosGrupes = Import-Csv $Atnaujintas_grupiu_saraso_failas -Encoding UTF8 -Delimiter ";"

# Atnaujintą informaciją įrašyti į grupių paskyras, esančias Office 365 aplinkoje
$AtnaujintosGrupes |
    foreach { Set-DistributionGroup -Identity $_.Guid -Name $_.Name -DisplayName $_.DisplayName -Alias $_.Alias -EmailAddresses $_.PrimarySmtpAddress -IgnoreNamingPolicy }


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 17: sukurti saugos grupes naujų mokinių klasėms
#
#------------------------------------------------------------------------------------------------------------------

# Sukurti saugos grupes naujoms klasėms
$NaujuKlasiuSarasas = Import-Csv $Nauju_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";" | select Klasė | Where-Object { $_.Klasė -NotLike "*grupė" -and $_.Klasė.Length -gt 0 } | Sort-Object Klasė -Unique
foreach ($NaujaKlase in $NaujuKlasiuSarasas) { 
    $KlasesPilnasPavadinimas = "Visa " + $NaujaKlase.Klasė + " klasė"
    if ($NaujaKlase.Klasė.IndexOf(" ") -ne -1) { $KlasesTrumpasPavadinimas = "visa." + $NaujaKlase.Klasė.Substring(0, $NaujaKlase.Klasė.IndexOf(" ")) } else { $KlasesTrumpasPavadinimas = "visa." + $NaujaKlase.Klasė }
    $KlasesSmtpAdresas = $KlasesTrumpasPavadinimas + "@" + $Domeno_vardas
    New-DistributionGroup -Name $KlasesPilnasPavadinimas -Type Security -DisplayName $KlasesPilnasPavadinimas -Alias $KlasesTrumpasPavadinimas -PrimarySmtpAddress $KlasesSmtpAdresas -MemberJoinRestriction ApprovalRequired -Notes $KlasesPilnasPavadinimas
    Set-DistributionGroup -Identity $KlasesSmtpAdresas -AcceptMessagesOnlyFrom $VisuotinioAdministratoriausSmtpAdresas -RequireSenderAuthenticationEnabled $false
    Set-DistributionGroup -Identity $KlasesSmtpAdresas -AcceptMessagesOnlyFromDLMembers $KlasesSmtpAdresas, $GrupesVisiMokytojaiSmtpAdresas
    Set-DistributionGroup -Identity $KlasesSmtpAdresas -AcceptMessagesOnlyFromSendersOrMembers $KlasesSmtpAdresas, $VisuotinioAdministratoriausSmtpAdresas, $GrupesVisiMokytojaiSmtpAdresas
}


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 18: įtraukti naujų mokinių paskyras į klasių saugos grupes
#
#------------------------------------------------------------------------------------------------------------------

# Įtraukti mokinių paskyras į klasių saugos grupes (CSV)
$NaujosMokiniuPaskyros = Import-Csv $Nauju_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";"
$NaujuKlasiuSarasas = Import-Csv $Nauju_mokiniu_paskyru_failas -Encoding UTF8 -Delimiter ";" | select Klasė | Where-Object { $_.Klasė -match '\d{1}\w' -or $_.Klasė -match '\d{1}\w' } | Sort-Object Klasė -Unique
foreach ($NaujaKlase in $NaujuKlasiuSarasas) {
    $KlasesPilnasPavadinimas = $NaujaKlase.Klasė + " klasė"
    $KlasesTrumpasPavadinimas = "visa." + $KlasesPilnasPavadinimas.Substring(0, $KlasesPilnasPavadinimas.IndexOf(" "))
    $KlasesSmtpAdresas = $KlasesTrumpasPavadinimas + "@" + $Domeno_vardas
    echo $KlasesSmtpAdresas
    $KlasesMokiniai = $NaujosMokiniuPaskyros | Where-Object { $_.Klasė -eq $NaujaKlase.Klasė } | Select VartotojoID
    $KlasesMokiniai | foreach { Add-DistributionGroupMember -Identity $KlasesSmtpAdresas -Member $_.VartotojoID -Confirm:$false -BypassSecurityGroupManagerCheck }
}


#------------------------------------------------------------------------------------------------------------------
#
# Žingsnis 19: prieš uždarant šį PowerShell skripto failą arba Windows PowerShell ISE programą, uždaryti Exchange
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
