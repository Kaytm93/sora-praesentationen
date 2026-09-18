#!/usr/bin/env python3
"""
Click-to-Reveal Animation Injector v3 (PowerPoint-konform)

Grund für v3: PowerPoint markierte die v2-Ausgabe als reparaturbedürftig, weil
  a) <p:seq> muss <p:prevCondLst>/<p:nextCondLst> enthalten
  b) Jede Klick-Gruppe braucht einen eigenen <p:par>-Container mit delay=indefinite
     (v2 bündelte alle in einen Container → "Repair" entfernte das gesamte Timing)
  c) Effect-IDs müssen global eindeutig sein und streng aufsteigend

Nutzung:
  python3 inject_clicks.py input.pptx output.pptx [--skip 1,4,5,9,10]

Shapes werden per objectName-Präfix "g<N>_" gruppiert. Alle Shapes mit
demselben Präfix faden mit einem Klick gemeinsam ein.

--skip   Folien-Indices (1-basiert), die nicht animiert werden sollen.
         Folie 1 wird IMMER übersprungen (Intro-Charakter).
"""

import sys
import re
import shutil
import zipfile
from lxml import etree
from collections import OrderedDict

NS = {
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
    'p': 'http://schemas.openxmlformats.org/presentationml/2006/main',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
}
P = '{%s}' % NS['p']
A = '{%s}' % NS['a']


def extract_groups(slide_xml):
    """Gib OrderedDict {g_index: [spid, ...]} zurück. g_index startet bei 1."""
    root = etree.fromstring(slide_xml)
    groups = OrderedDict()
    # Alle sp / pic / graphicFrame mit objectName im nvSpPr/nvPicPr/nvGraphicFramePr
    for sp in root.iter():
        tag = etree.QName(sp).localname
        if tag not in ('sp', 'pic', 'graphicFrame', 'cxnSp', 'grpSp'):
            continue
        # nvXxxPr/cNvPr
        nv = None
        for child in sp:
            ltag = etree.QName(child).localname
            if ltag in ('nvSpPr', 'nvPicPr', 'nvGraphicFramePr', 'nvCxnSpPr', 'nvGrpSpPr'):
                nv = child
                break
        if nv is None:
            continue
        cnvpr = nv.find(P + 'cNvPr')
        if cnvpr is None:
            continue
        name = cnvpr.get('name', '') or ''
        spid = cnvpr.get('id')
        if not spid:
            continue
        m = re.match(r'^g(\d+)_', name)
        if not m:
            continue
        gidx = int(m.group(1))
        groups.setdefault(gidx, []).append(spid)
    # Sortiere nach Gruppenindex
    return OrderedDict(sorted(groups.items()))


def build_timing_xml(groups):
    """Baut <p:timing>-Block mit einer Klick-Gruppe pro Index, jede mit eigenem par/indefinite."""
    if not groups:
        return None

    # Effect-IDs: start bei 3, streng aufsteigend
    next_id = [3]

    def nid():
        v = next_id[0]
        next_id[0] += 1
        return str(v)

    # Root timing
    timing = etree.Element(P + 'timing', nsmap={'p': NS['p']})
    tnLst = etree.SubElement(timing, P + 'tnLst')
    par_root = etree.SubElement(tnLst, P + 'par')
    cTn_root = etree.SubElement(par_root, P + 'cTn', {
        'id': '1', 'dur': 'indefinite',
        'restart': 'whenNotActive', 'nodeType': 'tmRoot'
    })
    childTnLst_root = etree.SubElement(cTn_root, P + 'childTnLst')
    seq = etree.SubElement(childTnLst_root, P + 'seq', {
        'concurrent': '1', 'nextAc': 'seek'
    })
    cTn_main = etree.SubElement(seq, P + 'cTn', {
        'id': '2', 'dur': 'indefinite', 'nodeType': 'mainSeq'
    })
    main_children = etree.SubElement(cTn_main, P + 'childTnLst')

    # Pro Klick-Gruppe: ein <p:par> mit delay=indefinite
    for gidx, spids in groups.items():
        # Outer par (wartet auf Klick)
        click_par = etree.SubElement(main_children, P + 'par')
        click_cTn = etree.SubElement(click_par, P + 'cTn', {
            'id': nid(), 'fill': 'hold'
        })
        stc = etree.SubElement(click_cTn, P + 'stCondLst')
        etree.SubElement(stc, P + 'cond', {'delay': 'indefinite'})
        click_children = etree.SubElement(click_cTn, P + 'childTnLst')

        # Mid par (delay=0, startet direkt nach dem Klick)
        mid_par = etree.SubElement(click_children, P + 'par')
        mid_cTn = etree.SubElement(mid_par, P + 'cTn', {
            'id': nid(), 'fill': 'hold'
        })
        mstc = etree.SubElement(mid_cTn, P + 'stCondLst')
        etree.SubElement(mstc, P + 'cond', {'delay': '0'})
        mid_children = etree.SubElement(mid_cTn, P + 'childTnLst')

        # Pro Shape: erster = clickEffect, Rest = withEffect
        for i, spid in enumerate(spids):
            effect_par = etree.SubElement(mid_children, P + 'par')
            effect_cTn = etree.SubElement(effect_par, P + 'cTn', {
                'id': nid(),
                'presetID': '10',
                'presetClass': 'entr',
                'presetSubtype': '0',
                'fill': 'hold',
                'grpId': '0',
                'nodeType': 'clickEffect' if i == 0 else 'withEffect'
            })
            estc = etree.SubElement(effect_cTn, P + 'stCondLst')
            etree.SubElement(estc, P + 'cond', {'delay': '0'})
            echildren = etree.SubElement(effect_cTn, P + 'childTnLst')

            # <p:set> — setze style.visibility auf visible
            set_el = etree.SubElement(echildren, P + 'set')
            cBhvr_set = etree.SubElement(set_el, P + 'cBhvr')
            etree.SubElement(cBhvr_set, P + 'cTn', {
                'id': nid(), 'dur': '1', 'fill': 'hold'
            })
            tgt_set = etree.SubElement(cBhvr_set, P + 'tgtEl')
            etree.SubElement(tgt_set, P + 'spTgt', {'spid': spid})
            attrlst = etree.SubElement(cBhvr_set, P + 'attrNameLst')
            attr = etree.SubElement(attrlst, P + 'attrName')
            attr.text = 'style.visibility'
            to_el = etree.SubElement(set_el, P + 'to')
            etree.SubElement(to_el, P + 'strVal', {'val': 'visible'})

            # <p:animEffect> — fade in
            eff = etree.SubElement(echildren, P + 'animEffect', {
                'transition': 'in', 'filter': 'fade'
            })
            cBhvr_eff = etree.SubElement(eff, P + 'cBhvr')
            etree.SubElement(cBhvr_eff, P + 'cTn', {
                'id': nid(), 'dur': '500'
            })
            tgt_eff = etree.SubElement(cBhvr_eff, P + 'tgtEl')
            etree.SubElement(tgt_eff, P + 'spTgt', {'spid': spid})

    # prevCondLst / nextCondLst im <p:seq> (PFLICHT für PowerPoint!)
    prev = etree.SubElement(seq, P + 'prevCondLst')
    prev_cond = etree.SubElement(prev, P + 'cond', {
        'evt': 'onPrev', 'delay': '0'
    })
    prev_tgt = etree.SubElement(prev_cond, P + 'tgtEl')
    etree.SubElement(prev_tgt, P + 'sldTgt')

    nxt = etree.SubElement(seq, P + 'nextCondLst')
    nxt_cond = etree.SubElement(nxt, P + 'cond', {
        'evt': 'onNext', 'delay': '0'
    })
    nxt_tgt = etree.SubElement(nxt_cond, P + 'tgtEl')
    etree.SubElement(nxt_tgt, P + 'sldTgt')

    # bldLst — für jede Shape eine bldP-Zeile (optional aber sauberer)
    bld = etree.SubElement(timing, P + 'bldLst')
    for gidx, spids in groups.items():
        for spid in spids:
            etree.SubElement(bld, P + 'bldP', {
                'spid': spid, 'grpId': '0'
            })

    return timing


def inject_slide(slide_xml_bytes):
    """Nimm Slide-XML, entferne evtl. altes timing, füge neues ein."""
    root = etree.fromstring(slide_xml_bytes)
    groups = extract_groups(slide_xml_bytes)

    # Entferne bestehendes <p:timing>
    for t in root.findall(P + 'timing'):
        root.remove(t)

    if not groups:
        return etree.tostring(root, xml_declaration=True,
                              encoding='UTF-8', standalone=True), 0

    timing = build_timing_xml(groups)
    if timing is not None:
        root.append(timing)

    return (etree.tostring(root, xml_declaration=True,
                           encoding='UTF-8', standalone=True),
            len(groups))


def main():
    argv = sys.argv[1:]
    skip_slides = {1}  # Folie 1 immer skip
    if '--skip' in argv:
        idx = argv.index('--skip')
        skip_list = argv[idx + 1]
        for s in skip_list.split(','):
            skip_slides.add(int(s.strip()))
        del argv[idx:idx + 2]
    if len(argv) != 2:
        print("Usage: inject_clicks.py <input.pptx> <output.pptx> [--skip 4,5,9,10]")
        sys.exit(1)
    src, dst = argv[0], argv[1]
    print(f"Skip-Liste: Folien {sorted(skip_slides)}")

    # Kopiere zuerst, dann modifiziere in-place im Zip
    shutil.copy(src, dst)

    # Öffne Zip, lese und schreibe Slide-XMLs
    # zipfile kann nicht in-place modifizieren → neuer Zip + ersetzen
    tmp = dst + '.tmp'

    with zipfile.ZipFile(src, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                m = re.match(r'ppt/slides/slide(\d+)\.xml$', item.filename)
                if m:
                    slide_num = int(m.group(1))
                    if slide_num in skip_slides:
                        print(f"  Folie {slide_num}: übersprungen (keine Animation)")
                    else:
                        new_data, n_groups = inject_slide(data)
                        if n_groups > 0:
                            print(f"  Folie {slide_num}: {n_groups} Klick-Gruppen")
                            data = new_data
                        else:
                            print(f"  Folie {slide_num}: keine g*_ Shapes gefunden")
                zout.writestr(item, data)

    shutil.move(tmp, dst)
    print(f"\nGespeichert: {dst}")


if __name__ == '__main__':
    main()
