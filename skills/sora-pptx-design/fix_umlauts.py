#!/usr/bin/env python3
"""
Ersetzt ASCII-transliterierte deutsche Wörter durch echte Umlaute.
Arbeitet direkt auf pptx-Textframes UND Speaker-Notes.
"""
import sys
import shutil
import re
import zipfile
from lxml import etree

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
}
A = '{%s}' % NS['a']

# Word-by-word Ersetzungen (Reihenfolge wichtig: längere zuerst!)
REPLACEMENTS = [
    # Spezielle grosse → große (ß)
    (r'\bgrosse\b', 'große'),
    (r'\bGrosse\b', 'Große'),
    (r'\bgrossen\b', 'großen'),
    (r'\bgrossem\b', 'großem'),
    (r'\bgrosser\b', 'großer'),
    (r'\bgrosses\b', 'großes'),
    (r'\bgroesste\b', 'größte'),
    (r'\bgroesstem\b', 'größtem'),
    (r'\bgroessten\b', 'größten'),
    (r'\bgroesstes\b', 'größtes'),
    (r'\bgroesser\b', 'größer'),
    (r'\bStrasse\b', 'Straße'),
    (r'\bStrassen\b', 'Straßen'),
    (r'\bausloeschen\b', 'auslöschen'),

    # Politik / Wirtschaft Wortschatz
    (r'\bBuerokratie\b', 'Bürokratie'),
    (r'\bBuerger\b', 'Bürger'),
    (r'\bBuergerinnen\b', 'Bürgerinnen'),
    (r'\bstaerkste\b', 'stärkste'),
    (r'\bStaerke\b', 'Stärke'),
    (r'\bstaerken\b', 'stärken'),
    (r'\bSondervermoegen\b', 'Sondervermögen'),
    (r'\bRueckfuehrung\b', 'Rückführung'),
    (r'\bRueckfuehrungsoffensive\b', 'Rückführungsoffensive'),
    (r'\bRealitaets\b', 'Realitäts'),
    (r'\bRealitaet\b', 'Realität'),
    (r'\bschuetzte\b', 'schützte'),
    (r'\bschuetzt\b', 'schützt'),
    (r'\bSchuetzen\b', 'Schützen'),
    (r'\bVerlaengert\b', 'Verlängert'),
    (r'\bverlaengert\b', 'verlängert'),
    (r'\bverlaengern\b', 'verlängern'),
    (r'\bSprengkoepfe\b', 'Sprengköpfe'),
    (r'\bSprengkopf\b', 'Sprengkopf'),
    (r'\bAUFRUESTUNG\b', 'AUFRÜSTUNG'),
    (r'\bAufruestung\b', 'Aufrüstung'),
    (r'\baufruesten\b', 'aufrüsten'),
    (r'\bmoeglich\b', 'möglich'),
    (r'\bunmoeglich\b', 'unmöglich'),
    (r'\bVertraege\b', 'Verträge'),
    (r'\bWettruesten\b', 'Wettrüsten'),
    (r'\bWETTRUESTEN\b', 'WETTRÜSTEN'),
    (r'\bfuer\b', 'für'),
    (r'\bFuer\b', 'Für'),
    (r'\bueber\b', 'über'),
    (r'\bUeber\b', 'Über'),
    (r'\bUEBERLEITUNG\b', 'ÜBERLEITUNG'),
    (r'\bueberfuellt\b', 'übererfüllt'),
    (r'\buebererfuellt\b', 'übererfüllt'),
    (r'\bgruen\b', 'grün'),
    (r'\bGruen\b', 'Grün'),
    (r'\bgrun\b', 'grün'),  # typo in build: "grun" ohne e
    (r'\bGrun\b', 'Grün'),
    (r'\bHaertere\b', 'Härtere'),
    (r'\bhaertere\b', 'härtere'),
    (r'\bHaerte\b', 'Härte'),
    (r'\bBegruendung\b', 'Begründung'),
    (r'\bbegruendet\b', 'begründet'),
    (r'\blaeuft\b', 'läuft'),
    (r'\blaufen\b', 'laufen'),  # no change but safe
    (r'\bzaeh\b', 'zäh'),
    (r'\bSchaetzt\b', 'Schätzt'),
    (r'\bschaetzt\b', 'schätzt'),
    (r'\bSchaetzung\b', 'Schätzung'),
    (r'\bZuhoeren\b', 'Zuhören'),
    (r'\bzuhoeren\b', 'zuhören'),
    (r'\bAnsaetze\b', 'Ansätze'),
    (r'\bAtommaechten\b', 'Atommächten'),
    (r'\bStaedte\b', 'Städte'),
    (r'\bSTADT\b', 'STADT'),
    (r'\bausloesen\b', 'auslösen'),
    (r'\beingefroren\b', 'eingefroren'),  # kein Umlaut
    (r'\bgeschaetzt\b', 'geschätzt'),
    (r'\bAenderung\b', 'Änderung'),
    (r'\bAenderungen\b', 'Änderungen'),
    (r'\bMaerz\b', 'März'),
    (r'\bwaehrend\b', 'während'),
    (r'\bJahrzehnte\b', 'Jahrzehnte'),  # kein Umlaut
    (r'\bjaehrlich\b', 'jährlich'),
    (r'\bnaehert\b', 'nähert'),
    (r'\bDaumen\b', 'Daumen'),  # kein Umlaut
    (r'\bdurchwachsen\b', 'durchwachsen'),  # kein Umlaut
    (r'\bBruttoinlandsprodukt\b', 'Bruttoinlandsprodukt'),  # kein Umlaut
    (r'\bInlandsprodukt\b', 'Inlandsprodukt'),  # kein Umlaut
    (r'\bsonderweg\b', 'sonderweg'),  # kein Umlaut
    (r'\bgroessten\b', 'größten'),
    (r'\bgroesseren\b', 'größeren'),
    (r'\bnaechsten\b', 'nächsten'),
    (r'\bNachbarn\b', 'Nachbarn'),  # kein Umlaut
    (r'\bOekonomenpanel\b', 'Ökonomenpanel'),
    (r'\bOEKONOMENPANEL\b', 'ÖKONOMENPANEL'),
]


def apply_replacements(text):
    if not text:
        return text
    for pat, repl in REPLACEMENTS:
        text = re.sub(pat, repl, text)
    return text


def process_xml(data_bytes):
    """Walk text elements in slide or notes XML and replace."""
    parser = etree.XMLParser(remove_blank_text=False)
    root = etree.fromstring(data_bytes, parser)
    changed = 0
    # a:t elements hold the visible text
    for t in root.iter(A + 't'):
        if t.text:
            new = apply_replacements(t.text)
            if new != t.text:
                t.text = new
                changed += 1
    return etree.tostring(root, xml_declaration=True,
                          encoding='UTF-8', standalone=True), changed


def main():
    if len(sys.argv) != 3:
        print("Usage: fix_umlauts.py <input.pptx> <output.pptx>")
        sys.exit(1)
    src, dst = sys.argv[1], sys.argv[2]

    total = 0
    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(dst, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if re.match(r'ppt/slides/slide\d+\.xml$', item.filename) or \
                   re.match(r'ppt/notesSlides/notesSlide\d+\.xml$', item.filename):
                    new_data, n = process_xml(data)
                    if n > 0:
                        print(f"  {item.filename}: {n} Ersetzungen")
                        data = new_data
                        total += n
                zout.writestr(item, data)
    print(f"\nGesamt: {total} Ersetzungen. Gespeichert: {dst}")


if __name__ == '__main__':
    main()
