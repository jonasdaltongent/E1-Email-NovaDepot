#!/usr/bin/env python3
"""
maak_werkdocumenten.py
Genereert het werkdocument voor les 15 "E1 — de e-mailopdracht" (NovaDepot, ORLO, Communicatie 5).

  CO5_E1-e-mailopdracht.docx   het werkdocument dat de leerling INLEVERT (een kopie per leerling)

Geen bronbestanden en geen zip: de leerling werkt in dit document en in Google Agenda. Geen Gmail: de leerlingen
hebben er met hun schoolaccount geen toegang toe (Jonas, 09-10-2026). De e-mail komt in dit document, in een venster
dat eruitziet als "Nieuw bericht" (zoals les 12). De afspraak maakt de leerling in de echte Google Agenda, met de
leraar als enige gast; in deel 2 typt ze over wat ze invulde en plakt ze een schermafbeelding.

Gebruik:  python3 maak_werkdocumenten.py
Vereist:  python-docx, Pillow (huisstijl.py en go-dalton-gent.png staan ernaast)

LET OP bij aanpassen:
 - Alle namen, nummers en adressen zijn verzonnen; de adressen eindigen op .example (RFC 2606).
 - E1 is een evaluatie: GEEN voorbeeldzinnen, GEEN zinsbegin, GEEN ingevulde rij. Alleen de feiten (de notities van Tom,
   bewust door elkaar) en lege vakken. De leerling kiest wat in de mail hoort en wat in de afspraak.
 - Het e-mailadres van de leraar staat NERGENS in dit document; het staat op het bord.
 - De punten staan per deel (12 · 6 · 2) en per criterium in de nakijklijst (lesvoorbereiding § 22). Verander je ze,
   doe het op beide plaatsen. Het punt voor de schermafbeelding (1) staat in § 22 bij *Nakijken*, het vak staat in deel 2.
 - Eén vak per ding dat nagekeken wordt (_afspraken/verbeteren.md). De verwachte antwoorden staan in lesvoorbereiding
   § 14 en § 22.
 - Het deel 'Nakijken en inleveren' begint met het vak voor het codewoord (INFOSESSIE); de leerling krijgt het in stap 6
   van de lespagina na de zelftest.
"""
import os

import huisstijl as hs   # de huisstijl van het werkdocument (uit _sjablonen/lespakket/werkdocument/)

HERE = os.path.dirname(os.path.abspath(__file__))
UIT = os.path.join(HERE, "CO5_E1-e-mailopdracht.docx")

# De notities van Tom (lesvoorbereiding § 4), bewust door elkaar. Elk adres staat op een eigen regel: om te kopiëren.
NOTITIES = [
    "eerste werkdag: maandag 21 december 2026, om 7.30 uur aan het onthaal",
    "Sara Claes, preventiedienst — sara.claes@novadepot.example",
    "infosessie, punten: welkom en rondleiding · veiligheid in het magazijn (Sara) · de planning van de week · vragen",
    "bijlage: de onthaalbrochure (de pdf van les 04)",
    "nieuw van Sara: veiligheidsschoenen meebrengen; wie er geen heeft, antwoordt op de mail vóór vrijdag 11 december: "
    "dan krijgt hij een paar van NovaDepot",
    "doel infosessie: de jobstudenten welkom heten en uitleggen hoe de week verloopt",
    "wie niet kan komen naar de infosessie: antwoordt met Nee",
    "identiteitskaart meebrengen",
    "infosessie: 8.00 tot 9.00 uur, in de kantine — Lotte Maes leidt ze",
    "Tom Peeters, magazijnverantwoordelijke, NovaDepot, toestel 214 — tom.peeters@novadepot.example",
    "de jobstudenten kennen elkaar niet; hun adressen zijn privé:",
    "lina.haddad@mail.example",
    "emre.yilmaz@mail.example",
    "noor.verbeke@mail.example",
    "kobe.desmet@mail.example",
    "ayla.janssens@mail.example",
]

# Deel 2: het veld in Google Agenda, en een vak. Eén vak per ding dat nagekeken wordt.
VELDEN_AGENDA = [
    ("Titel", ""),
    ("Datum", ""),
    ("Van … tot …", ""),
    ("Locatie", ""),
    ("Gasten (zonder e-mailadres)", ""),
    ("Beschrijving", ""),
]

# De nakijklijst: (vraag in leerlingtaal, punten van het criterium in § 22). Samen 19, plus 1 voor de nakijklijst en de zin.
NAKIJKLIJST = [
    ("Staat elke ontvanger in het juiste veld: Aan, Cc of Bcc? Kan een jobstudent het adres van een andere jobstudent zien?", 3),
    ("Staat in het onderwerp waarover de mail gaat, met de datum?", 1),
    ("Begint de mail met een zakelijke aanspreking, en zegt de inleiding wie Tom is en waarom hij mailt?", 1),
    ("Staan dag, uur, onthaal, identiteitskaart, veiligheidsschoenen en de infosessie (uur en plaats) in de mail?", 3),
    ("Vraagt de mail wie geen veiligheidsschoenen heeft om te antwoorden, met de datum?", 1),
    ("Staat er een zakelijke slotgroet, en een handtekening met naam, functie en NovaDepot?", 1),
    ("Zit de onthaalbrochure als bijlage bij de mail?", 1),
    ("Gebruik ik overal je of overal u, zonder emoji, afkortingen of hoofdletterwoorden?", 1),
    ("Heeft de afspraak een titel met infosessie, de juiste dag en het juiste uur?", 2),
    ("Staat de kantine bij Locatie en de leraar bij Gasten?", 1),
    ("Staat in de beschrijving wie leidt, het doel, de vier punten en wie niet kan komen?", 2),
    ("Klikte ik op Sturen?", 1),
    ("Staat de schermafbeelding van mijn afspraak in deel 2?", 1),
]


def nieuwe_bladzijde(doc):
    """Een bladzijdeskeer vóór een deel met puntenvak: page_break_before in een tabelcel (huisstijl.deel met punten)
    negeert Google Documenten. Een kleine lege alinea die zelf op een nieuwe bladzijde begint, doet het wel."""
    p = doc.add_paragraph()
    p.paragraph_format.page_break_before = True
    hs.zonder_ruimte(p)
    p.paragraph_format.line_spacing = hs.Pt(2)
    r = p.add_run("")
    r.font.size = hs.Pt(1)
    return p


def werkdocument():
    doc = hs.nieuw_document()
    doc.core_properties.title = "CO5 E1 de e-mailopdracht - werkdocument"
    doc.core_properties.author = "Toegepaste Informatica"
    hs.kopblok(doc, "E1 — DE E-MAILOPDRACHT", punten=20)

    hs.kader(doc, "De opdracht en de notities van Tom", [
        "Tom Peeters begeleidt de vijf nieuwe jobstudenten van NovaDepot. Hij vraagt jou:",
        ("1.", "Stuur de jobstudenten een e-mail over hun eerste werkdag, met de onthaalbrochure als bijlage. Sara Claes van "
               "de preventiedienst moet weten wat er gestuurd werd. Ze moet niets doen."),
        ("2.", "Zet de infosessie in de agenda. Nodig in de les alleen je leraar uit."),
        "Zijn notities staan door elkaar.",
    ])
    nt = hs.tabel(doc, ["Notitie", "Waar? (mail, agenda of allebei)"], [(n, "") for n in NOTITIES],
                  [hs.BREEDTE - 4.6, 4.6], invul=(1,), rijhoogte=0.7)
    for rij in nt.rows[NOTITIES.index("lina.haddad@mail.example") + 1:]:   # de adressen: geen vak, ze horen bij de notitie erboven
        hs.schaduw(rij.cells[1], "FFFFFF")

    nieuwe_bladzijde(doc)
    hs.deel(doc, 1, "De e-mail aan de jobstudenten", punten=12)
    venster = hs.mailvenster(doc, [("Aanspreking", 1.0), ("Inleiding", 2.0), ("De boodschap", 5.0), ("Slot", 2.0),
                                   ("Slotgroet", 1.0), ("Handtekening", 2.2)],
                             velden=("Aan", "Cc", "Bcc", "Onderwerp", "Bijlage"))
    hs.hoogte(venster.rows[3], 2.3)   # Bcc: vijf adressen
    hs.hoogte(venster.rows[5], 1.2)   # Bijlage: een chip met een lange bestandsnaam

    nieuwe_bladzijde(doc)
    hs.deel(doc, 2, "De infosessie in de agenda", punten=6)
    t = hs.tabel(doc, ["Veld in Google Agenda", "Wat je invulde"], VELDEN_AGENDA, [4.8, hs.BREEDTE - 4.8],
                 invul=(1,), rijhoogte=1.0, vet=(0,))
    hs.hoogte(t.rows[6], 5.2)   # Beschrijving
    hs.vraag(doc, "Op welke knop klikte je na Opslaan?")
    hs.antwoordvak(doc, 0.9)
    hs.vraag(doc, "Schermafbeelding van je afspraak")
    hs.antwoordvak(doc, 8.0)
    hs.tekst(doc, "Lukt de schermafbeelding niet? Typ dan in het vak: lukt niet.", klein=True, cursief=True)

    nieuwe_bladzijde(doc)
    hs.deel(doc, 3, "Nakijken en inleveren", punten=2)
    # eerst het codewoord van de zelftest (15-10-2026): de lespagina toont het pas als de drie vragen van stap 6
    # juist zijn; zo ziet Jonas wie de zelftest echt maakte (_afspraken/lespagina.md, De laatste stap)
    hs.vraag(doc, "Het codewoord van de zelftest (stap 6 op de lespagina):")
    hs.antwoordvak(doc, 0.9)
    rijen = [(v, "", f"/ {p}") for v, p in NAKIJKLIJST]
    rijen.append(("Ik vulde de nakijklijst in en schreef wat ik verbeterde (hieronder).", "", "/ 1"))
    nk = hs.tabel(doc, ["Vraag", "Ja of nee", "Punten"], rijen, [11.2, 2.8, hs.BREEDTE - 14.0], invul=(1,),
                  rijhoogte=0.9)
    laatste = nk.rows[len(rijen)]
    laatste.cells[1].text = ""            # geen ja/nee bij de laatste rij: ze is zelf het bewijs
    hs.schaduw(laatste.cells[1], "FFFFFF")
    hs.vraag(doc, "Wat verbeterde je na het nakijken? Schrijf één zin. Was alles ja? Schrijf dan niets.")
    hs.antwoordvak(doc, 1.6)

    doc.save(UIT)


if __name__ == "__main__":
    werkdocument()
    print("Klaar.")
    print("  werkdocument om in te leveren : " + os.path.basename(UIT))
