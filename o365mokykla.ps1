#------------------------------------------------------------------------------
#
# MIT License
#
# Copyright (c) 2020 Sarunas Koncius
#
#
# Permission is hereby granted, free of charge, to any person obtaining a copy
# of this software and associated documentation files (the "Software"), to deal
# in the Software without restriction, including without limitation the rights
# to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
# copies of the Software, and to permit persons to whom the Software is
# furnished to do so, subject to the following conditions:
#
# The above copyright notice and this permission notice shall be included in all
# copies or substantial portions of the Software.
#
# THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
# IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
# FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
# AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
# LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
# OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
# SOFTWARE.
#
#------------------------------------------------------------------------------
#
# PowerShell Source Code
#
# NAME:
#    o365mokykla.ps1
#
# AUTHOR:
#    Sarunas Koncius
#
# VERSION:
# 	 0.8.0 20200322 
#
# MODIFIED:
#	 2020-03-22
#
#
#------------------------------------------------------------------------------


#
# DĖMESIO! Skriptą reikia atsidaryti Windows PowerShell ISE aplinkoje ir vykdyti tik pažymėtas kodo eilutes!
# 


# Prieš pradedant dirbti su skriptu, įdiekite MSOnline modulį Azure AD valdymui ir leiskite vykdyti skriptus Exchange Online valdymui.
# Startuokite Windows PowerShell ISE aplinką ar Windows PowerShell komandinės eilutės aplinką administratoriaus teisėmis ir naudokite šias komandas:
Install-Module -Name MSOnline
Set-ExecutionPolicy RemoteSigned


# Prisijungti prie Azure AD ir Exchange Online paslaugų valdymo
$UserCredential = Get-Credential
connect-msolservice -credential $UserCredential
$Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection
Import-PSSession $Session -DisableNameChecking
Set-Location -Path $Env:USERPROFILE\Desktop # Darbastalis yra nustatomas aktyviu katalogas - ten galite laikyti CSV failus


# Paeržiūrėti informaciją apie turimas licencijas ir jų kiekius
#
# AccountSkuId                                 ActiveUnits WarningUnits ConsumedUnits
# ------------                                 ----------- ------------ -------------
# o365mokykla:STANDARDWOFFPACK_STUDENT         1000000     0            0         
# o365mokykla:STANDARDWOFFPACK_FACULTY         500000      0            1          
#
# !!! SVARBU!
# !!! Prieš licencijos pavadinimą iki dvitaškio yra rodomas mokyklos Office 365 aplinkos identifikatorius (pavyzdyje - o365mokykla).
# !!! Savo mokyklos Office 365 aplinkos identifikatorių reikia įrašyti 123 ir 190 eilutėse.
#
Get-MsolAccountSku


# Įveskite informaciją apie savo mokyklą, pažymėkite ir įvykdykite kodo eilutes
$Domeno_vardas = "o365mokykla.lt"
$Mokslo_metai = "2019-2020"
$Mokytoju_saraso_failas = ".\o365mokykla_2019-2020_mokytojai.csv" # CSV failas su besimokančių mokinių sąrašu iš mokinių registro 
$Mokytoju_paskyru_failas = ".\o365mokykla_2019-2020_mokytojai_paskyros.csv" # Paskyrų failas bus sukurtas su laikinaisiais slaptažodžiais pirmajam vartotojų prisijungimui
$Mokiniu_saraso_failas = ".\o365mokykla_2019-2020_mokiniai.csv" # CSV failas su besimokančių mokinių sąrašu iš mokinių registro 
$Mokiniu_paskyru_failas = ".\o365mokykla_2019-2020_mokiniai_paskyros.csv" # Paskyrų failas bus sukurtas su laikinaisiais slaptažodžiais pirmajam vartotojų prisijungimui


# Patikrinti, ar CSV failas su mokytojų sąrašu yra tinkamas
# Turi būti rodomi mokytojų sąrašo duomenys trijuose stulpeliuose. Jeigu duomenys matosi viename stulpelyje,
# CSV faile skyrybos ženklą kablelį pakeiskite kabliataškiu arba atvirkščiai.
# Stulpelių pavadinimai turi būti "Pavardė", "Vardas" ir "Pareigos", bet jų eilės tvarka nėra svarbi. 
# Pataisykite CSV failą, jeigu stulpelių pavadinimai yra kiti.
# Pakoregavę CSV failą, grįžkite prie 94-95 eilutės CSV failui patikrinti
$NaujuMokytojuSarasas = Import-Csv $Mokytoju_saraso_failas -Encoding UTF8
$NaujuMokytojuSarasas | ft


# Sukurti naujas mokytojų paskyras
function Remove-StringNonLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
Out-File -FilePath $Mokytoju_paskyru_failas -InputObject "Mokytojas,VartotojoID,Slaptažodis" -Encoding UTF8
$NaujuMokytojuSarasas = Import-Csv $Mokytoju_saraso_failas -Encoding UTF8
foreach ($NaujasMokytojas in $NaujuMokytojuSarasas) {}
foreach ($NaujasMokytojas in $NaujuMokytojuSarasas)
	{
        $NewFirstName = (Get-Culture).textinfo.totitlecase($NaujasMokytojas.Vardas.ToLower())
        $NewLastName = (Get-Culture).textinfo.totitlecase($NaujasMokytojas.Pavardė.ToLower())
        $NewTitle = $NaujasMokytojas.Pareigos
        $NewOffice = $Mokslo_metai
        If ($NewFirstName.Contains(" ")) { $OnlyFirstName = $NewFirstName.Substring(0, $NewFirstName.IndexOf(" ")) } else { $OnlyFirstName = $NewFirstName }
        If ($NewLastName.Contains(" ")) { $OnlyLastName = $NewLastName.Substring($NewLastName.LastIndexOf(" ")+1,$NewLastName.Length-$NewLastName.LastIndexOf(" ")-1) } else { $OnlyLastName = $NewLastName }
        $NewDisplayName = $OnlyFirstName + " " + $OnlyLastName
        $NewUserPrincipalName = (Remove-StringNonLatinCharacters $OnlyFirstName.ToLower()) + "." + (Remove-StringNonLatinCharacters $OnlyLastName.ToLower()) + "@" + $Domeno_vardas
        Echo $NewUserPrincipalName
        $EsamasVartotojas = Get-MsolUser -UserPrincipalName $NewUserPrincipalName -ErrorAction SilentlyContinue
        If ($EsamasVartotojas -eq $Null)
            {
		        New-MsolUser -UserPrincipalName $NewUserPrincipalName -DisplayName $NewDisplayName -FirstName $NewFirstName -LastName $NewLastName -Title $NewTitle -Office $NewOffice -PreferredLanguage "lt-LT" -UsageLocation "LT" -ForceChangePassword:$true
                $Slaptazodis = Set-MsolUserPassword -UserPrincipalName $NewUserPrincipalName -ForceChangePassword:$true
                Set-MsolUserLicense -UserPrincipalName $NewUserPrincipalName -AddLicenses "o365mokykla:STANDARDWOFFPACK_STUDENT" # <<< !!! Prieš dvitaškį įrašykite savo mokyklos Office 365 aplinkos ID !!!
            }
        else
            {
                $Slaptazodis = "!!!PasikartojantisVartotojoID!!!"
            }
        $Mokytojas = $NewFirstName + " " + $NewLastName
        $VartotojoID = $NewUserPrincipalName
		$PrisijungimoInformacija = "$Mokytojas,$VartotojoID,$Slaptazodis"
        Out-File -FilePath $Mokytoju_paskyru_failas -InputObject $PrisijungimoInformacija -Encoding UTF8 -append
	}


# Nustatyti lietuviškus Office 365 aplinkos ir e. pašto dėžutės parametrus naujoms mokytojų paskyroms
# Prieš vykdant šį kodo bloką, Office 365 administratoriaus portale įsitikinkite, 
# kad paskutinėms sukurtoms mokytojų paskyroms jau yra sukurtos e.pašto dėžutės.
$NaujuMokytojuSarasas = Import-Csv $Mokytoju_paskyru_failas -Encoding UTF8
$Skaitliukas = 1
foreach ($NaujasMokytojas In $NaujuMokytojuSarasas)
	{
		$Upn = $NaujasMokytojas.VartotojoID
        Echo $Upn
		$ActivityMessage = "Prašome palaukti..."
		$StatusMessage = ("Nustatomi parametrai vartotojui {0} ({1} iš {2})" -f $Upn, $Skaitliukas, @($NaujuMokytojuSarasas).count)
		$PercentComplete = ($Skaitliukas / @($NaujuMokytojuSarasas).count * 100)
		Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
		set-MailboxRegionalConfiguration -Identity $NaujasMokytojas.VartotojoID -TimeZone "FLE Standard Time" -Language lt-LT –LocalizeDefaultFolderName
		Start-Sleep -m 500
		$Skaitliukas++
	}


# Patikrinimas, ar CSV failas su mokinių sąrašu yra tinkamas
# Turi būti rodomi mokinių sąrašo duomenys trijuose stulpeliuose. Jeigu duomenys matosi viename stulpelyje,
# CSV faile skyrybos ženklą kablelį pakeiskite kabliataškiu arba atvirkščiai.
# Stulpelių pavadinimai turi būti "Pavardė", "Vardas" ir "Klasės/grupės pavadinimas", bet jų eilės tvarka nėra svarbi. 
# Pataisykite CSV failą, jeigu stulpelių pavadinimai yra kiti.
# Pakoregavę CSV failą, grįžkite prie 161-162 eilutės CSV failui patikrinti
$NaujuMokiniuSarasas = Import-Csv $Mokiniu_saraso_failas -Encoding UTF8
$NaujuMokiniuSarasas | ft


# Sukurti naujas mokinių paskyras
function Remove-StringNonLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
Out-File -FilePath $Mokiniu_paskyru_failas -InputObject "Klasė,Mokinys,VartotojoID,Slaptažodis" -Encoding UTF8
$NaujuMokiniuSarasas = Import-Csv $Mokiniu_saraso_failas -Encoding UTF8
foreach ($NaujasMokinys in $NaujuMokiniuSarasas)
	{
        $NewFirstName = (Get-Culture).textinfo.totitlecase($NaujasMokinys.Vardas.ToLower())
        $NewLastName = (Get-Culture).textinfo.totitlecase($NaujasMokinys.Pavardė.ToLower())
        $NewDepartment = $NaujasMokinys."Klasės/grupės pavadinimas".ToLower() + " klasė"
        $NewOffice = $Mokslo_metai
        If ($NewFirstName.Contains(" ")) { $OnlyFirstName = $NewFirstName.Substring(0, $NewFirstName.IndexOf(" ")) } else { $OnlyFirstName = $NewFirstName }
        If ($NewLastName.Contains(" ")) { $OnlyLastName = $NewLastName.Substring($NewLastName.LastIndexOf(" ")+1,$NewLastName.Length-$NewLastName.LastIndexOf(" ")-1) } else { $OnlyLastName = $NewLastName }
        if ($OnlyFirstName.EndsWith("s")) { $NewTitle = $NewDepartment + "s mokinys" } else { $NewTitle = $NewDepartment + "s mokinė" }
        $NewDisplayName = $OnlyFirstName + " " + $OnlyLastName
        $NewUserPrincipalName = (Remove-StringNonLatinCharacters $OnlyFirstName.ToLower()) + "." + (Remove-StringNonLatinCharacters $OnlyLastName.ToLower()) + "@" + $Domeno_vardas
        Echo $NewUserPrincipalName
        $EsamasVartotojas = Get-MsolUser -UserPrincipalName $NewUserPrincipalName -ErrorAction SilentlyContinue
        If ($EsamasVartotojas -eq $Null)
            {
		        New-MsolUser -UserPrincipalName $NewUserPrincipalName -DisplayName $NewDisplayName -FirstName $NewFirstName -LastName $NewLastName -Title $NewTitle -Department $NewDepartment -Office $NewOffice -PreferredLanguage "lt-LT" -UsageLocation "LT" -ForceChangePassword:$true
                $Slaptazodis = Set-MsolUserPassword -UserPrincipalName $NewUserPrincipalName -ForceChangePassword:$true
                Set-MsolUserLicense -UserPrincipalName $NewUserPrincipalName -AddLicenses "o365mokykla:STANDARDWOFFPACK_STUDENT" # <<< !!! Prieš dvitaškį įrašykite savo mokyklos Office 365 aplinkos ID !!!
            }
        else
            {
                $Slaptazodis = "!!!PasikartojantisVartotojoID!!!"
            }
        $Klase = $NaujasMokinys."Klasės/grupės pavadinimas".ToLower()
        $Mokinys = $NewFirstName + " " + $NewLastName
        $VartotojoID = $NewUserPrincipalName
		$PrisijungimoInformacija = "$Klase,$Mokinys,$VartotojoID,$Slaptazodis"
        Out-File -FilePath $Mokiniu_paskyru_failas -InputObject $PrisijungimoInformacija -Encoding UTF8 -append
	}


# Nustatyti lietuviškus Office 365 aplinkos ir e. pašto dėžutės parametrus naujoms mokinių paskyroms
# Prieš vykdant šį kodo bloką, Office 365 administratoriaus portale įsitikinkite, 
# kad paskutinėms sukurtoms mokinių paskyroms jau yra sukurtos e.pašto dėžutės.
$NaujuMokiniuSarasas = Import-Csv $Mokiniu_paskyru_failas -Encoding UTF8
$Skaitliukas = 1
foreach ($NaujasMokinys In $NaujuMokiniuSarasas)
	{
		$Upn = $NaujasMokinys.VartotojoID
        Echo $Upn
		$ActivityMessage = "Prašome palaukti..."
		$StatusMessage = ("Nustatomi parametrai vartotojui {0} ({1} iš {2})" -f $Upn, $Skaitliukas, @($NaujuMokiniuSarasas).count)
		$PercentComplete = ($Skaitliukas / @($NaujuMokiniuSarasas).count * 100)
		Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
		set-MailboxRegionalConfiguration -Identity $NaujasMokinys.VartotojoID -TimeZone "FLE Standard Time" -Language lt-LT –LocalizeDefaultFolderName
		Start-Sleep -m 500
		$Skaitliukas++
	}


# -----------------------------------------------------------------------------
#
# End of Script.
#
# -----------------------------------------------------------------------------
