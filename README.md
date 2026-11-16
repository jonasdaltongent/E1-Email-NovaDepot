# Handleiding voor de leraar — Les 15: E1 — de e-mailopdracht

<!-- Beginvragen (stap 1 van CONTEXT.md), beantwoord door Jonas op 16-11-2026 (telkens het voorstel gekozen):
     klassen 3ORLO, 4ORLOa, 4ORLOb · les 15 van het jaarplan (module 3), evaluatie E1, uur 2 van W11: 3ORLO wo 18-11
     (2de uur), 4ORLOb wo 18-11 (3de uur), 4ORLOa do 19-11 (7de uur) · vorige les: les 14 (Uitnodigen met de agenda),
     een dag eerder in elke klas → lesstart = kort ophalen uit les 11–14 · lokaal 18, Windows 11 · individuele
     evaluatie: één NovaDepot-casus, een zakelijke e-mail met bijlage (juist gebruik van Aan/CC/BCC) én een
     agenda-uitnodiging, verstuurd naar Jonas · werkdocument CO5_E1-e-mailopdracht met screenshots en een eigen
     nakijklijst · puntenlijst met criteria, 20 punten · de lespagina geeft de opdracht en de criteria, niet het
     antwoord · deadline: het einde van het lesuur · context NovaDepot. 
     Getest op 16-11-2026 met demoleerling3 (Jonas: "Test het?"): Google Agenda werkt voor leerlingen (afspraak maken,
     gasten uitnodigen, Sturen); Gmail niet. Geen ruimtes: het lokaal komt bij Locatie. Dus: de uitnodiging in de echte
     Google Agenda, met Jonas en de duopartner als gasten; de e-mail (E1) in een nagebootst venster, zoals les 12. -->

**Vak:** Toegepaste Informatica
**Doelgroep:** 3ORLO, 4ORLOa en 4ORLOb — 2de graad Organisatie en logistiek, arbeidsmarktgerichte finaliteit
**Lesduur:** 1 × 50 minuten: 8 minuten instructie + 42 minuten zelfstandig werk (evaluatie, geen demo)
**Context:** NovaDepot, *Tom mailt de jobstudenten over hun eerste werkdag en zet de infosessie in de agenda* — module 3 van het jaarplan, les 15 (**E1**, reeks *Communicatie 5*)
**Toestel:** **lokaal 18, Windows 11** — alles in Chrome; de e-mail in een nagebootst venster, de afspraak in de echte Google Agenda
**Lesdag en deadline (het einde van het lesuur):** 3ORLO wo 18 november, 2de uur, **10.05 uur** · 4ORLOb wo 18 november, 3de uur, **11.15 uur** · 4ORLOa do 19 november, 7de uur, **14.35 uur**
**Kernleerplandoelen:** `BV2_04.01`, `BK2_02.01.02`, `BK2_02.10.03` · ondersteunend `BV2_04.04` · **summatief, 20 punten**
**Pages:** https://jonasdaltongent.github.io/E1-Email-NovaDepot/ · dia's: …/presentatie.html

De leerling is Tom. Uit zijn door elkaar geschreven notities kiest ze wat in de **e-mail** aan de vijf jobstudenten hoort (Aan, Cc, Bcc, onderwerp, bijlage, de zes bouwstenen) en wat in de **afspraak** in Google Agenda (titel, dag en uur, locatie, de leraar als gast, beschrijving, **Sturen**). Daarna overtypt ze de afspraak, plakt een schermafbeelding, kijkt haar werk na met de nakijklijst en levert in. De lespagina geeft de opdracht, de notities, de puntenlijst en de klikpaden; **geen voorbeeldzinnen**. Opbouw en bronnen: `lesvoorbereiding.md`.

---

## 1. Inhoud van het pakket

```text
W11 - Les 15 - ORLO - E1 de e-mailopdracht/
├── index.html                   # de lespagina: route · één stap · checklist
├── presentatie.html             # 5 dia's: 1 Zo start je · 2 lesstart · 3 de opdracht · 4 de puntenlijst · 5 deadline en zo werk je verder (kloktijden)
├── css/style.css, css/slides.css
├── js/script.js, js/slides.js
├── assets/                      # logo's, lettertypes, mascotte, het filmpje Inleveren; geen schermafbeeldingen
├── werkdocument/
│   ├── CO5_E1-e-mailopdracht.docx   # het werkdocument dat de leerling INLEVERT (kopie per leerling)
│   ├── maak_werkdocumenten.py       # maakt het opnieuw (python-docx)
│   └── huisstijl.py, go-dalton-gent.png
├── lesvoorbereiding.md
├── dalton-lesfiche.html         # openen, Kopieer de fiche, plakken in je planner
├── classroom.json               # de opdracht voor _tools/zet_opdracht_klaar.py
├── lesdoelen.json               # leerplandoelen voor je jaaroverzicht
└── README.md                    # deze handleiding
```

Geen zip en geen bronbestanden. De leerlingen hebben alleen hun **pdf uit les 04** nodig (die staat in hun Drive).

### 1b. Hoe de lespagina werkt

Links de **route** (*De opdracht*: stap 1 · *De e-mail*: stap 2–3 · *De afspraak*: stap 4 · *Afwerken*: stap 5–6), midden **één stap** met *Toon stapjes* bij de handelingen in de software (de les begint in *Kort*, zonder knop), rechts de **checklist** met 18 taken. Geen *Hulp nodig?*.

- **Startpagina** met vier beelden (de opdracht in Classroom → deze lespagina → je werkdocument CO5 → Google Agenda) en *Dit maak je vandaag*: het skelet van de e-mail en van de afspraak, als balkjes zonder tekst.
- **Stap 1**: twee infokaders: *De opdracht en de notities van Tom* (de notities door elkaar) en *Zo krijg je je punten* (de 13 criteria, algemeen verwoord, samen 20). De leerling typt bij elke notitie in het werkdocument *mail*, *agenda* of *allebei*.
- **Stap 2 en 3**: korte geheugensteunen (*Aan, Cc of Bcc?*; de zes bouwstenen en de toon), **zonder voorbeeldzinnen**, zinsbegin of *Mag je overnemen*. De bijlage met `@` heeft drie stapjes met een tekening van de lijst met suggesties.
- **Stap 4**: Agenda in zes handelingen met *Toon stapjes* (Meer opties, gast, schermafbeelding met **Windows + Shift + S** en **Ctrl + V**, **Opslaan** › het venster met de drie knoppen) en het skelet *Ziet het bij jou er zo uit?*. De tekening van het venster toont de drie knoppen zonder rood kader: de leerling kiest zelf.
- **Stap 6** (*Inleveren*): eerst een zelftest van drie vragen over de **werkwijze**, daarna het blok *Afwerken en inleveren*; de groene regel toont het codewoord **INFOSESSIE**, dat de leerling bovenaan *Deel 3* typt.
- `localStorage` bewaart alleen vinkjes en de laatste stap (voorvoegsel `novadepot_e1_v1_`), met een wisknop.

---

## 2. Klaarzetten

### Stap 1 — Publiceren via GitHub Pages

Repository `jonasdaltongent/E1-Email-NovaDepot`, Pages op `main`, map `/ (root)`; adres hierboven. Dat adres staat op dia 5, in `classroom.json` en in `lesdoelen.json`.

### Stap 2 — De opdracht in Classroom, met de koppeling

`classroom.json`: **Communicatie 5 — E1: de e-mailopdracht**, onderwerp *Digitale competenties*, 3ORLO, 4ORLOa en 4ORLOb, 20 punten. **Eén opdracht per klas, met de deadline van die klas** (het script kent één deadline per run: de concepten zijn per klas gemaakt):

| Klas | Deadline | Inplannen (een half uur vóór de les) |
|---|---|---|
| 3ORLO | wo 18-11 om 10.05 uur | wo 18-11 om 8.45 |
| 4ORLOb | wo 18-11 om 11.15 uur | wo 18-11 om 9.55 |
| 4ORLOa | do 19-11 om 14.35 uur | do 19-11 om 13.15 |

Bijlagen: de link naar de lespagina en `CO5_E1-e-mailopdracht` (**Een kopie maken voor elke leerling**). Geen zip. Status van de concepten: zie de regel in `_index/logboek.md` (16-11-2026); de lesbouwer maakt alleen het **concept**, inplannen doet de hoofdsessie.

Instructietekst:

```text
1. Open de lespagina (link) en je werkdocument CO5_E1-e-mailopdracht.
2. Dit is een evaluatie: je werkt alleen. Je schrijft de e-mail in je werkdocument en maakt de afspraak in Google Agenda.
3. Klik op Inleveren, ten laatste <deadline van de klas>.
```

### Stap 3 — Met de hand, als de koppeling niet werkt

1. Upload `werkdocument/CO5_E1-e-mailopdracht.docx` **in Drive zelf** (**Nieuw** › **Bestand uploaden**, met *Uploads converteren* aan) en voeg het toe met **Bijvoegen** › **Drive** › **Een kopie maken voor elke leerling**.
2. Voeg de lespagina toe met **Link**. Stel per klas de deadline in.

### Afvinklijst vóór de les

- [ ] De lespagina staat online en het adres op dia 5 klopt.
- [ ] De opdracht staat in de drie klassen, met twee bijlagen en de deadline van die klas.
- [ ] **Je e-mailadres staat op het bord.** Het staat nergens anders (lespagina, dia's, werkdocument).
- [ ] Een eigen testafspraak naar jezelf komt aan in je agenda; je weet hoe je de uitnodigingen van maandag 21 december daarna opruimt (± één per leerling).
- [ ] Les 14 is in die klas gegeven (een dag eerder); kijk in de app *Na de les* of er over de agenda iets bij te sturen is.
- [ ] De Dalton-lesfiche staat in je planner (open `dalton-lesfiche.html`, klik op **Kopieer de fiche**, plak).
- [ ] `presentatie.html` opent op de beamer; `N` toont je notities, `F` is volledig scherm.

### 2b. Nagelezen klikpaden

| Handeling | Klikpad / naam | Bron |
|---|---|---|
| Werkdocument openen | in de opdracht rechts bij **Jouw werk** op de naam klikken | `_afspraken/klikpaden.md` |
| Een afspraak maken | **Maken** › **Afspraak** › **Meer opties**: **Titel toevoegen**, datum en uur **tot** uur, **Locatie toevoegen**, **Beschrijving toevoegen**; rechts het tabblad **Gasten** met **Gasten toevoegen**; de eerste keer eerst twee vensters met **OK** | `_afspraken/klikpaden.md` (demoleerling3, 16-11-2026) · [Agenda 72143](https://support.google.com/calendar/answer/72143?hl=nl) |
| De uitnodiging versturen | **Opslaan** › venster *Wil je uitnodigingsmails sturen naar Google Agenda-gasten?* › **Terug naar bewerken** · **Niet sturen** · **Sturen** | `_afspraken/klikpaden.md` (demoleerling3, 16-11-2026) |
| Een bijlage als chip | in het vak **Bijlage** `@` typen, een stuk van de naam (*TV3*), het **rode pdf-icoon** onder *Bestanden* kiezen | `_afspraken/klikpaden.md` (ORLO les 12, demoleerling3, 21-10-2026) |
| Schermafbeelding | **Windows-toets + Shift + S**; het knipsel staat op het klembord; plakken met **Ctrl + V** in het vak | `_afspraken/klikpaden.md` · [Knipprogramma](https://support.microsoft.com/nl-nl/windows/apps/use-snipping-tool-to-capture-screenshots) |
| Inleveren | **Inleveren** › venster **Je werk inleveren?** › **Inleveren** › status **Ingeleverd** | `_afspraken/klikpaden.md` |

**Niet nagelezen, dus zo opgeschreven:** *Open Google Agenda* (de Agenda-help noemt geen adres: de pagina zegt *zoals in les 14*); *druk op Enter na het e-mailadres van de gast* (de help zegt alleen dat je een e-mailadres kan typen); wat er na **Sturen** in de agenda van de leerling staat.

---

## 3. Het verloop van de les

| Fase | Tijd | Wat |
|---|---|---|
| **Instructie** | **8'** | Dia 1 terwijl de computers opstarten (zet **je adres op het bord**). Dia 2 lesstart (2'): *"Vandaag toon je wat je in les 11 tot 14 leerde. Je werkt alleen."* Dia 3 de opdracht (2'): de twee dingen die Tom moet doen; de notities staan door elkaar. Dia 4 de puntenlijst (2'). Dia 5 deadline en inleveren (2'): het einde van het lesuur, eerst de nakijklijst, dan *Inleveren*. |
| **Zelfstandig werk** | **42'** | Dia 5 blijft staan, met **kloktijden** (`_tools/zet_lesmomenten.py`); stap 1 tot 6, inleveren inbegrepen. |

Stap 1–6: 4 + 6 + 12 + 12 + 5 + 3 = **42 minuten**. Geen demo, geen controlevraag, geen exitvraag.

**Rondgang:** alleen de **software** (*Waar is Meer opties? Hoe maak ik een chip? Welke toetsen voor een schermafbeelding?*). Op een vraag over de inhoud: *"Kijk naar de notities en je nakijklijst."* Eerste rondgang: heeft iedereen het **eigen** werkdocument open? Daarna: staat de afspraak op de juiste dag, en zit er een schermafbeelding in? Zeg aan het einde mondeling dat wie niet klaar is, toch inlevert. Minimumroute: eerst de e-mail (12 punten), dan de afspraak (6), dan inleveren.

---

## 4. Verbetersleutel

- **Codewoord** bovenaan *Deel 3 — Nakijken en inleveren*: **INFOSESSIE** (een tikfout telt als juist). Ontbreekt het of is het fout: −2 (`_afspraken/evaluatie-en-privacy.md`).
- De leerling levert alleen het werkdocument in (20 punten: 12 · 6 · 2). De volledige verdeling en de feiten die juist moeten zijn: `lesvoorbereiding.md` § 22. De uitnodiging staat ook in jouw agenda (ma 21-12-2026, 8.00).

### Deel 1 — De e-mail (12)

| Vak | Verwacht | Punten |
|---|---|---|
| Aan, Cc, Bcc | Aan leeg (of Toms eigen adres) · Cc `sara.claes@novadepot.example` · Bcc de vijf adressen van de jobstudenten, letter voor letter | 3 |
| Onderwerp | waarover + de datum (*Eerste werkdag maandag 21 december*) | 1 |
| Aanspreking, Inleiding | een zakelijke aanspreking (½) · wie Tom is en waarom hij mailt (½) | 1 |
| De boodschap | ma 21-12-2026 · 7.30 uur · onthaal · identiteitskaart · veiligheidsschoenen · infosessie 8.00 in de kantine | 3 |
| Slot | wie geen veiligheidsschoenen heeft, antwoordt (½) vóór vrijdag 11 december (½); verwijzing naar de brochure | 1 |
| Slotgroet, Handtekening | *Met vriendelijke groeten* (½) · Tom Peeters, magazijnverantwoordelijke, NovaDepot, toestel 214 (½) | 1 |
| Bijlage | de brochure als chip, of de naam van de pdf | 1 |
| Toon | je- of u-vorm volgehouden, geen emoji, afkortingen of hoofdletterwoorden | 1 |

### Deel 2 — De infosessie in de agenda (6)

| Vak | Verwacht | Punten |
|---|---|---|
| Titel, Datum, Van … tot … | titel met *infosessie* (½) · ma 21 december 2026 (½) · 8.00–9.00 (1) | 2 |
| Locatie, Gasten | *NovaDepot – kantine* (½) · *mijn leraar* (½) | 1 |
| Beschrijving | wie leidt (Lotte Maes) · het doel · de vier punten · wie niet kan: Nee (elk ½) | 2 |
| De knop na Opslaan | **Sturen** | 1 |

### Deel 3 — Nakijken en inleveren (2)

De schermafbeelding in deel 2 (1) en de nakijklijst met de zin over wat de leerling verbeterde (1). De schermafbeelding kan de AI die nakijkt niet lezen zonder namen en adressen te zien: kijk er zelf naar (of: aanwezig = 1).

### Essentiële fouten — geef hier altijd feedback op

- De adressen van de jobstudenten zichtbaar voor elkaar (in Aan of Cc).
- Een verkeerde datum of een verkeerd uur voor de eerste werkdag of de infosessie.
- Geen bijlage.
- De afspraak opgeslagen met **Niet sturen**.

E1 is summatief: de leerling levert niet opnieuw in om punten te verbeteren (voorstel; jij beslist). Wie niets inleverde: 0, inhalen mag (−5).

---

## 5. Schermafbeeldingen (optioneel)

Geen. Alle beelden zijn tekeningen in HTML en CSS (ook de lijst met suggesties, het venster met de drie knoppen en de afspraak): er is geen schermafbeelding van een leerlingaccount nodig. Zie `assets/screenshots/LEESMIJ.md`.

---

## 6. Het materiaal opnieuw maken

```bash
python3 "werkdocument/maak_werkdocumenten.py"
```

Maakt `CO5_E1-e-mailopdracht.docx` opnieuw. Vereist `python-docx` en `Pillow`. `werkdocument/huisstijl.py` is de versie van het sjabloon, met de parameter `velden` bij `mailvenster` (hier: Aan, Cc, Bcc, Onderwerp, Bijlage).

---

## 7. Leerplandoelen in je jaaroverzicht

`lesdoelen.json`: `BV2_04.01`, `BK2_02.01.02` en `BK2_02.10.03` (kern), `BV2_04.04` (ondersteunend). De repository krijgt de `pre-push` hook uit `_sjablonen/pre-push`.

---

## 8. Wat nog moet blijken in de klas

- **Haalbaarheid in 42 minuten.** Een e-mail met zes tekstvakken, een afspraak in Agenda met een schermafbeelding en een nakijklijst, alleen, zonder schrijfhulp: stap 3 (12') en stap 4 (12') zijn krap. De minimumroute is de uitweg.
- **De notities door elkaar.** Zetten de leerlingen de infosessie (uur en plaats) in de mail én het doel en de punten in de afspraak, of mengen ze alles? Hoeveel leerlingen typen echt *mail*, *agenda* of *allebei* bij elke notitie (4 minuten)?
- **De schermafbeelding** met **Windows + Shift + S** en **Ctrl + V** in een vak van een Google-document: nieuw voor de meeste leerlingen, nooit met een leerlingaccount in lokaal 18 geprobeerd. Lukt plakken in een tabelcel? Hoeveel leerlingen typen *lukt niet*?
- **Agenda openen en een gast toevoegen:** de pagina zegt *zoals in les 14* en *druk op Enter*; beide zijn niet nagelezen (§ 2b). Vragen de leerlingen waar ze Agenda openen?
- **De uitnodigingen in jouw agenda:** komen ze allemaal aan, ook van de leerlingen die per ongeluk **Niet sturen** kozen (dan niet; de knop in deel 2 vertelt het)? Hoe ruim je ± één uitnodiging per leerling op?
- **De chip (`@`)** met de pdf uit les 04: zoals les 12 (getest met demoleerling3). Typen leerlingen de naam als ze de pdf niet vinden?
- **Het venster *Nieuw bericht*** met vijf velden: groeit het vak *Bijlage* mee met een chip met een lange naam? Alleen nagekeken in de opbouw.
- **De nakijklijst**: nakijken zonder de antwoorden te kennen; verbeteren de leerlingen echt wat op *nee* staat, of typen ze overal *ja*?
- De lespagina en dia's zijn gemeten met scripts en in headless Chrome bekeken, niet door een mens op een Windows-toestel.
