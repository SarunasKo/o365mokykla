# o365mokykla
 - LT: PowerShell skriptas, skirtas sukurti pradinę Office 365 aplinką mokyklai, kuri pradeda naudotis šiomis paslaugomis
 - EN: PowerShell script to setup initial environment on Office 365 tenant for secondary school in Lithuania

PowerShell skriptas:
 - sukuria mokytojų paskyras iš CSV failo, kuris gali būti sukuriamas eksportavus duomenis iš Excel failo;
 - sukuria mokytojų paskyrų CSV failą su prisijungimo vardais ir laikinaisiais slaptažodžiais;
 - sukuria mokinių paskyras iš CSV failo, kuris yra gaunamas eksportavus duomenis iš mokinių registro;
 - sukuria mokinių paskyrų CSV failą su prisijungimo vardais ir laikinaisiais slaptažodžiais;
 - nustato lietuviškus regiono bei laiko juostos parametrus vartotojų paskyroms ir pašto dėžutėms;
 - išjungia standartinį Focused Inbox nustatymą vartotojų pašto dėžutėms;
 - sukuria saugos grupę su įgalintu e. paštu "Visi mokytojai" ir į narius įtraukia visų mokytojų paskyras;
 - sukuria saugos grupę su įgalintu e. paštu "Visi mokiniai";
 - sukuria saugos grupę su įgalintu e. paštu "Visa mokyka" ir į narius įtraukia grupes "Visi mokytojai" bei "Visi mokiniai";
 - sukuria saugos grupę su įgalinti e. paštu kiekvienai klasei ir į narius įtraukia tos klasės mokinius;
 - suteikia teises mokytojams siųsti laiškus visoms grupėms, o mokiniams - tik savo klasės grupei.
 
 Papildomi failai:
  - o365mokykla_2019-2020_mokytojai.csv - pavyzdinis mokytojų sąrašo CSV failas;
  - o365mokykla_2019-2020_mokytojai_paskyros.csv - pavyzdinis sukurtų mokytojų paskyrų CSV failas;
  - o365mokykla_2019-2020_mokiniai.csv - pavyzdinis mokinių sąrašo CSV failas;
  - o365mokykla_2019-2020_mokiniai_paskyros.csv - pavyzdinis sukurtų mokinių paskyrų CSV failas;
  - o365mokykla_00_Kaip-parengti-Office-365-naudojimui-mokykloje.pptx - PowerPoint pateiktis, aprašanti Office 365 aplinkos parengimo eigą;
  - o365mokykla_05_Kaip-pirmaji-karta-prisijungti-prie-Office.pptx - PowerPoint pateiktis sukurti filmukui apie pirmąjį prisijungimą.
