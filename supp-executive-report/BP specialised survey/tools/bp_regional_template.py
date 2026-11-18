#!/usr/bin/env python3
"""Turn the BP template's U.S. (MTF) comparisons into Asia and Europe comparisons.

    python3 bp_regional_template.py IN.docx OUT.docx

- 8 measure tables (alcohol / cannabis / vaping): per year group Asia | Europe | school
- 6 substance tables per grade: Asia | Europe | school
- 3 gender tables: Males and Females each Asia | Europe | school
- "U.S. norms" wording, table titles and the MTF footnote replaced; MTF note rows removed.
"""
import copy, re, sys
import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

LOG = []
T = lambda e: ''.join(x.text or '' for x in e.iter(qn('w:t')))


def set_tc_text(tc, text):
    ts = list(tc.iter(qn('w:t')))
    if not ts:
        p = tc.find(qn('w:p')) or OxmlElement('w:p')
        r = OxmlElement('w:r'); t = OxmlElement('w:t'); r.append(t); p.append(r)
        ts = [t]
    ts[0].text = text
    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for x in ts[1:]:
        x.text = ''


def set_span(tc, n):
    pr = tc.find(qn('w:tcPr'))
    if pr is None:
        pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
    gs = pr.find(qn('w:gridSpan'))
    if n == 1:
        if gs is not None:
            pr.remove(gs)
        return
    if gs is None:
        gs = OxmlElement('w:gridSpan')
        w = pr.find(qn('w:tcW'))
        (w.addnext(gs) if w is not None else pr.insert(0, gs))
    gs.set(qn('w:val'), str(n))


def set_tc_w(tc, tw):
    pr = tc.find(qn('w:tcPr'))
    if pr is None:
        pr = OxmlElement('w:tcPr'); tc.insert(0, pr)
    w = pr.find(qn('w:tcW'))
    if w is None:
        w = OxmlElement('w:tcW'); pr.insert(0, w)
    w.set(qn('w:w'), str(int(tw))); w.set(qn('w:type'), 'dxa')


def rebuild(tbl, groups, label_tw, total_tw):
    """groups: number of header groups (year groups or sexes). Each becomes Asia | Europe | school."""
    trs = tbl.findall(qn('w:tr'))
    ncol = 1 + 3 * groups
    rest = (total_tw - label_tw) / (ncol - 1)
    widths = [label_tw] + [rest] * (ncol - 1)
    grid = tbl.find(qn('w:tblGrid'))
    for g in list(grid):
        grid.remove(g)
    for w in widths:
        gc = OxmlElement('w:gridCol'); gc.set(qn('w:w'), str(int(w))); grid.append(gc)
    tpr = tbl.find(qn('w:tblPr'))
    tw = tpr.find(qn('w:tblW'))
    if tw is not None:
        tw.set(qn('w:w'), str(int(total_tw))); tw.set(qn('w:type'), 'dxa')
    # row 0: label + group headers spanning 3
    tcs = trs[0].findall(qn('w:tc'))
    for tc in tcs[1:]:
        set_span(tc, 3); set_tc_w(tc, rest * 3)
    set_tc_w(tcs[0], label_tw)
    # row 1: label + (Asia, Europe, school) per group
    tcs = trs[1].findall(qn('w:tc'))
    src_bench, src_school = tcs[1], tcs[2]
    for tc in tcs[1:]:
        trs[1].remove(tc)
    for _ in range(groups):
        for txt, src in [('Asia', src_bench), ('Europe', src_bench), (None, src_school)]:
            c = copy.deepcopy(src)
            if txt:
                set_tc_text(c, txt)
            set_span(c, 1); set_tc_w(c, rest)
            trs[1].append(c)
    set_tc_w(tcs[0], label_tw)
    # loop rows span everything; data row gets {{row[i]}} cells in order
    for tr in trs[2:]:
        tcs = tr.findall(qn('w:tc'))
        if len(tcs) == 1:
            set_span(tcs[0], ncol); set_tc_w(tcs[0], total_tw)
            continue
        sample = tcs[1]
        for tc in tcs[1:]:
            tr.remove(tc)
        for i in range(1, ncol):
            c = copy.deepcopy(sample); set_tc_text(c, '{{row[%d]}}' % i); set_span(c, 1); set_tc_w(c, rest)
            tr.append(c)
        set_tc_text(tcs[0], '{{row[0]}}'); set_tc_w(tcs[0], label_tw)


def main(src, dst):
    doc = docx.Document(src)
    body = doc.element.body
    sect = doc.sections[0]
    total_tw = int((sect.page_width - sect.left_margin - sect.right_margin) / 635)
    measure_vars = ['alcohollifetumetable', 'alcoholl12monthtstable', 'alcoholl30daysstable', 'pot_lifetime_MTF',
                    'pot_12_MTF', 'pot_30_MTF', 'VAPE_12_MTF', 'VAPE_30_MTF']
    n = {'measure': 0, 'substance': 0, 'gender': 0}
    for tbl in list(body.iter(qn('w:tbl'))):
        txt = T(tbl)
        if any(f'in {v}.iterrows' in txt for v in measure_vars):
            rebuild(tbl, 3, int(0.95 * 1440), total_tw); n['measure'] += 1
        elif re.search(r'in ALLSUBSTANCE\d+th_(12|30)_MTF\.iterrows', txt):
            rebuild(tbl, 1, int(1.85 * 1440), total_tw); n['substance'] += 1
        elif re.search(r'in sub12_by_sex_MTFcombo\d+th\.iterrows', txt):
            rebuild(tbl, 2, int(1.85 * 1440), total_tw); n['gender'] += 1
    LOG.append(f'tables rebuilt: {n}')

    # wording
    REPL = [
        ('Substance Use Comparisons to U.S. Norms', 'Substance Use Comparisons to Asia and Europe'),
        ('and U.S. norms for use of various substances', 'and the Asia and Europe international school averages for use of various substances'),
        ('with U.S. norms', 'with the Asia and Europe international school averages'),
        ('Compared to U.S. Norms', 'Compared to Asia and Europe'),
    ]
    below_re = re.compile(r'use less than the U\.S\. average on \{\{(\w+)\}\} of the substances compared')
    removed = 0
    for p in list(body.iter(qn('w:p'))):
        t = T(p)
        if re.fullmatch(r'\s*\*{1,2}\s*(Annual use not measured in MTF Survey|30 Day use not measured in MTF Survey|Identified as .Any Vaping. on MTF Survey)\s*', t):
            p.getparent().remove(p); removed += 1
            continue
        new = t
        for a, b in REPL:
            new = new.replace(a, b)
        new = below_re.sub(lambda m: (f'use less than the Asia average on {{{{{m.group(1)}}}}} and less than the Europe '
                                      f'average on {{{{{m.group(1)}_eu}}}} of the substances compared'), new)
        if new != t:
            ts = list(p.iter(qn('w:t')))
            ts[0].text = new
            ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            for x in ts[1:]:
                x.text = ''
            LOG.append(f'text: {new[:90]}')
    LOG.append(f'MTF note rows removed: {removed}')

    # footnote about Monitoring the Future -> the regional benchmarks
    for part in doc.part.package.parts:
        if str(part.partname) == '/word/footnotes.xml':
            from lxml import etree
            root = etree.fromstring(part.blob)
            for fn in root.findall(qn('w:footnote')):
                if 'Monitoring the Future' in T(fn):
                    p = fn.find(qn('w:p'))
                    for hl in list(p.iter(qn('w:hyperlink'))):
                        hl.getparent().remove(hl)
                    for fld in list(p.iter(qn('w:fldSimple'))):
                        fld.getparent().remove(fld)
                    ts = [t for t in p.iter(qn('w:t'))]
                    ts[0].text = (' The Asia and Europe averages in this report come from Prevention Ed SUPP surveys of '
                                  'international schools in each region, combined across schools and weighted by the number '
                                  'of students. They are included to provide benchmarks for contextualizing {{School}} results.')
                    ts[0].set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                    for x in ts[1:]:
                        x.text = ''
                    for r in list(p.iter(qn('w:r'))):
                        if r.find(qn('w:instrText')) is not None or r.find(qn('w:fldChar')) is not None:
                            r.getparent().remove(r)
                    LOG.append('footnote 1 rewritten (Asia and Europe benchmarks)')
            part._blob = etree.tostring(root, xml_declaration=True, encoding='UTF-8', standalone=True)
    doc.save(dst)
    return LOG


if __name__ == '__main__':
    for l in main(sys.argv[1], sys.argv[2]):
        print(l)
