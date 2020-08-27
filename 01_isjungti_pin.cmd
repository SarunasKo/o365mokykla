rem
rem Komandinis failas, skirtas išjungti PIN reikalavimą į kompiuterį jungiantis Azure AD vartotojams
rem Failas: 01_isjungti_pin.cmd
rem Versija: v0.9.4
rem 
@echo off
reg.exe add HKEY_LOCAL_MACHINE\SOFTWARE\Policies\Microsoft\PassportForWork /v Enabled /t REG_DWORD /d 0
