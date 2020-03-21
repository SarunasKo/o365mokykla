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
# 	 0.7.0 20200321 
#
# MODIFIED:
#	 2020-03-21
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
Set-Location -Path $Env:USERPROFILE\Desktop # Aktyvus katalogas - darbastalio, CSV failą laikykite ten


# Pasižiūrėkite informaciją apie turimas licencijas ir jų kiekius
# Prieš licencijos pavadinimą iki dvitašio yra rodomas mokyklos Office 365 aplinkos identifikatorius - jį reikės nurodyti 114 eilutėje
Get-MsolAccountSku


# Įveskite informaciją apie savo mokyklą, pažymėkite ir įvykdykite kodo eilutes
$Domeno_vardas = "eportfelis.net"
$Mokslo_metai = "2019-2020"
$Mokiniu_saraso_failas = ".\o365mokykla_2019-2020_mokiniai.csv" # CSV failas su besimokančių mokinių sąrašu iš mokinių registro 
$Paskyru_saraso_failas = ".\o365mokykla_2019-2020_paskyros.csv" # Paskyrų failas bus sukurtas su laikinaisiais slaptažodžiais pirmajam vartotojų prisijungimui


# Patikrinimas, ar CSV failas su mokinių sąrašu yra tinkamas
$NaujuMokiniuSarasas = Import-Csv $Mokiniu_saraso_failas -Encoding UTF8
$NaujuMokiniuSarasas | ft


# Turi būti rodomi mokinių sąrašo duomenys trijuose stulpeliuose. Jeigu duomenys matosi viename stulpelyje,
# CSV faile skyrybos ženklą kablelį pakeiskite kabliataškiu arba atvirkščiai.
# Stulpelių pavadinimai turi būti "Pavardė", "Vardas" ir "Klasės/grupės pavadinimas" (be kabučių). 
# Pataisykite CSV failą, jeigu stulpelių pavadinimai yra kiti.
# Pakoregavę CSV failą, grįžkite prie 78-79 eilutės CSV failui patikrinti


# Mokinių paskyrų sukūrimas
function Remove-StringNonLatinCharacters
{
    PARAM ([string]$String)
    [Text.Encoding]::ASCII.GetString([Text.Encoding]::GetEncoding("Cyrillic").GetBytes($String))
}
Out-File -FilePath $Paskyru_saraso_failas -InputObject "Klasė,Mokinys,VartotojoID,Slaptažodis" -Encoding UTF8
$NaujuMokiniuSarasas = Import-Csv $Mokiniu_saraso_failas -Encoding UTF8
foreach ($NaujasMokinys in $NaujuMokiniuSarasas) {}
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
                Set-MsolUserLicense -UserPrincipalName $NewUserPrincipalName -AddLicenses "eportfelis:STANDARDWOFFPACK_STUDENT" # <<< !!! Prieš dvitaškį įrašykite savo mokyklos Office 365 aplinkos ID !!!
            }
        else
            {
                $Slaptazodis = "!!!PasikartojantisUPN!!!"
            }
        $Klase = $NaujasMokinys."Klasės/grupės pavadinimas".ToLower()
        $Mokinys = $NewFirstName + " " + $NewLastName
        $VartotojoID = $NewUserPrincipalName
		$PrisijungimoInformacija = "$Klase,$Mokinys,$VartotojoID,$Slaptazodis"
        Out-File -FilePath $Paskyru_saraso_failas -InputObject $PrisijungimoInformacija -Encoding UTF8 -append
	}


# Lietuviški Office 365 aplinkos ir pašto dėžutės nustatymai naujiems mokiniams
$NaujuMokiniuSarasas = Import-Csv $Paskyru_saraso_failas -Encoding UTF8
$Skaitliukas = 1
foreach ($Mokinys In $NaujuMokiniuSarasas)
	{
		$Upn = $Mokinys.VartotojoID
        Echo $Upn
		$ActivityMessage = "Prašome palaukti..."
		$StatusMessage = ("Nustatomi parametrai vartotojui {0} ({1} iš {2})" -f $Upn, $Skaitliukas, @($NaujuMokiniuSarasas).count)
		$PercentComplete = ($Skaitliukas / @($NaujuMokiniuSarasas).count * 100)
		Write-Progress -Activity $ActivityMessage -Status $StatusMessage -PercentComplete $PercentComplete
		set-MailboxRegionalConfiguration -Identity $Mokinys.VartotojoID -TimeZone "FLE Standard Time" -Language lt-LT –LocalizeDefaultFolderName
		Start-Sleep -m 500
		$Skaitliukas++
	}


# -----------------------------------------------------------------------------
#
# End of Script.
#
# -----------------------------------------------------------------------------