 # o365mokykla-nmm.ps1
 PowerShell skriptas, skirtas atnaujinti Office 365 mokytojų, mokinių ir grupių paskyras naujiems mokslo metams:
  - eksportuoja esamas mokytojų paskyras į CSV failą, kad jas būtų galima atnaujinti naudojant Excel programą;
  - importuoja atnaujintą mokytojų paskyrų informaciją iš CSV failo į Office 365 aplinką;
  - eksportuoja esamas mokinių paskyras į CSV failą, kad jas būtų galima atnaujinti naudojant Excel programą;
  - importuoja atnaujintą mokinių paskyrų informaciją iš CSV failo į Office 365 aplinką;
  - eksportuoja esamas saugos grupių su įgalintu e. paštu paskyras į CSV failą, kad jas būtų galima atnaujinti naudojant Excel programą;
  - importuoja atnaujintą saugos grupių paskyrų informaciją iš CSV failo į Office 365 aplinką.
 
# o365mokykla.ps1
PowerShell skriptas, skirtas sukurti pradinę Office 365 aplinką mokyklai, kuri pradeda naudotis šiomis paslaugomis:
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
 
 # Papildomi failai:
  - o365mokykla_2019-2020_mokytojai.csv - pavyzdinis mokytojų sąrašo CSV failas;
  - o365mokykla_2019-2020_mokytojai_paskyros.csv - pavyzdinis sukurtų mokytojų paskyrų CSV failas;
  - o365mokykla_2019-2020_mokiniai.csv - pavyzdinis mokinių sąrašo CSV failas;
  - o365mokykla_2019-2020_mokiniai_paskyros.csv - pavyzdinis sukurtų mokinių paskyrų CSV failas;
  - o365mokykla_2020-2021_mokytojai_nauji.csv - pavyzdinis naujų mokytojų sąrašo CSV failas, naudojamas atnaujinant Office 365 aplinkos duoemenis naujiems mokslo metams;
  - o365mokykla_2020-2021_mokiniai_nauji.csv - pavyzdinis naujų mokinių sąrašo CSV failas, naudojamas atnaujinant Office 365 aplinkos duoemenis naujiems mokslo metams;
  - o365mokykla_PowerQueryEN_mokiniai.txt - Power Query užklausa mokinių paskyrų duomenų atnaujinimui angliškoje Excel programos versijoje;
  - o365mokykla_PowerQueryLT_mokiniai.txt - Power Query užklausa mokinių paskyrų duomenų atnaujinimui lietuviškoje Excel programos versijoje;
  - o365mokykla_PowerQueryEN_grupes.txt - Power Query užklausa grupių paskyrų duomenų atnaujinimui angliškoje Excel programos versijoje;
  - o365mokykla_PowerQueryLT_grupes.txt - Power Query užklausa grupių paskyrų duomenų atnaujinimui lietuviškoje Excel programos versijoje;
  - o365mokykla_00_Kaip-parengti-Office-365-naudojimui-mokykloje.pptx - PowerPoint pateiktis apie Office 365 aplinkos parengimo eigą;
  - o365mokykla_05_Kaip-pirmaji-karta-prisijungti-prie-Office.pptx - PowerPoint pateiktis sukurti filmukui apie pirmąjį prisijungimą.
