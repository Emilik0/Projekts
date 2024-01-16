# Treniņu grafiks

Programmas galvenais mērķis ir nokopēt jau pieejamo treniņa grafiku failā Workout.xlsx jaunā Excel lapā un saglabāt to ar šodienas datuma nosaukumu. Tabula sastāv no četrām ailēm. Vingrinājuma nosaukums, atkārtojumu skaits, pieeju skaits un laiks (ja vingrinājums tiek veikts noteiktā laikā). Pēc tam, kad lietotājs ir ievadījis visus treniņa datus un uzsācis programmu, galvenā tabula ar kolonnu nosaukumiem un vingrinājumu nosaukumiem tiks nokopēta uz jaunu lapu un treniņa dati tiks izdzēsti. Tas ļauj lietotājam katru reizi ievadīt jaunākos datus par treniņa gaitu, papildināt vingrinājumu sarakstu un sekot līdzi savam progresam. Programma papildus saskaita treniņā pavadīto laiku, ņemot vērā atpūtu starp piegājieniem, un aprēķina treniņa produktivitāti, aprēķinot, cik no kopējā vingrinājumu saraksta tika izpildīts.

## Izmantotas bibliotēkas 

**openpyxl**
Python modulis openpyxl tiek izmantots darbam ar Excel failiem, gan lasīt gan rakstīt uz tam.

Workbook: Šī klase tiek izmantota, lai izveidotu jaunu Excel workbook. To var izmantot, lai pievienotu lapas, manipulētu ar datiem un saglabātu workbook failā.
load_workbook: Šī funkcija tiek izmantota, lai ielādētu esošo Excel workbook no faila. Tā atgriež workbook objektu, ar kuru pēc tam var strādāt.

**datetime**
To izmanto darbam ar datumiem un laiku Python programmā. Kodā tiek izmantots datetime.now(), lai iegūtu pašreizējo datumu un saglābāt to kā lapas nosaukumu.

**openpyxl.styles**
Klase NamedStyle ļauj definēt un piemērot nosauktos stilus Excel workbook šūnām. Nosaukts stils ir formatēšanas iespēju kopums (piemēram, fonts, aizpildījums, apmales u. c.), kam var piešķirt nosaukumu un pēc tam viegli piemērot workbook šūnām.