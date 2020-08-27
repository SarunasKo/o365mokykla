rem
rem Komandinis failas, skirtas vietinei administratoriaus paskyrai sukurti
rem Failas: 02_sukurti_vietini_admina.cmd
rem Versija: v0.9.4
rem 
rem Pavyzdyje vietinio administratoriaus vardas: localadmin
rem Jeigu keisite administratoriaus vardą, neužmirškite jį įrašyti abejose komandose!
rem 
rem Pvyzdyje nurodytas administratoriaus slaptažodis: password
rem Būtinai pakeiskite pavyzdyje nurodytą administratoriaus slaptažodį stipriu.
rem Jeigu trūksta idėjų stipriam ir įsimenamam slaptažodžiui, naudokite https://xkpasswd.net/s/
rem
@echo off
net user localadmin password /add /y
net localgroup Administrators localadmin /add
