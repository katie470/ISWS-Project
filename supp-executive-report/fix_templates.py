#!/usr/bin/env python3
"""Apply the Nov 2026 template fixes (T1-T20, Part 3 sizes, and the template side
of C2, C5, C7, C8, C10, C12, C13, C15, C17, C19, C20, C24, C25) to a SUPP
Executive Report template.

    python3 fix_templates.py IN.docx OUT.docx

Safe to run on a template that has already been fixed (every step checks
first). Prints what it changed. Both templates (8th-12th and 7th grade) are
fixed with this same script so they stay identical.
"""
import copy
import re
import sys

import docx
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.shared import Pt, Inches, Twips
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.text.paragraph import Paragraph

LOG = []


def log(msg):
    LOG.append(msg)


# ---------------------------------------------------------------- helpers
def ptext(el):
    return ''.join(t.text or '' for t in el.iter(qn('w:t')))


def is_p(el):
    return el is not None and el.tag == qn('w:p')


def is_tbl(el):
    return el is not None and el.tag == qn('w:tbl')


def style_id(p_el):
    st = p_el.find(qn('w:pPr') + '/' + qn('w:pStyle'))
    return st.get(qn('w:val')) if st is not None else ''


def ppr(p_el):
    pp = p_el.find(qn('w:pPr'))
    if pp is None:
        pp = OxmlElement('w:pPr')
        p_el.insert(0, pp)
    return pp


# w:pPr children must keep schema order or Word reports the file as damaged.
_PPR_ORDER = ['pStyle', 'keepNext', 'keepLines', 'pageBreakBefore', 'framePr', 'widowControl', 'numPr',
              'suppressLineNumbers', 'pBdr', 'shd', 'tabs', 'suppressAutoHyphens', 'kinsoku', 'wordWrap',
              'overflowPunct', 'topLinePunct', 'autoSpaceDE', 'autoSpaceDN', 'bidi', 'adjustRightInd',
              'snapToGrid', 'spacing', 'ind', 'contextualSpacing', 'mirrorIndents', 'suppressOverlap', 'jc',
              'textDirection', 'textAlignment', 'textboxTightWrap', 'outlineLvl', 'divId', 'cnfStyle', 'rPr',
              'sectPr', 'pPrChange']


def set_ppr_child(pp, tag, attrs=None, remove=False):
    """Set (or remove) one pPr child, keeping schema order."""
    name = tag.split(':')[1]
    old = pp.find(qn(tag))
    if old is not None:
        pp.remove(old)
    if remove:
        return None
    el = OxmlElement(tag)
    for k, v in (attrs or {}).items():
        el.set(qn(k), str(v))
    idx = _PPR_ORDER.index(name)
    for i, child in enumerate(list(pp)):
        cname = child.tag.split('}')[1]
        if cname in _PPR_ORDER and _PPR_ORDER.index(cname) > idx:
            child.addprevious(el)
            return el
    pp.append(el)
    return el


def set_flag(p_el, tag, on=True):
    if on:
        set_ppr_child(ppr(p_el), tag)
    else:
        set_ppr_child(ppr(p_el), tag, {'w:val': '0'})


def has_drawing(el):
    return el.find('.//' + qn('w:drawing')) is not None or el.find('.//' + qn('w:pict')) is not None


def set_text(p_el, new):
    """Replace a paragraph's text, keeping the first run's formatting."""
    runs = p_el.findall(qn('w:r'))
    if not runs:
        r = OxmlElement('w:r'); t = OxmlElement('w:t'); r.append(t); p_el.append(r)
        runs = [r]
    first = runs[0]
    for r in runs[1:]:
        p_el.remove(r)
    for child in list(first):
        if child.tag in (qn('w:t'), qn('w:br'), qn('w:tab')):
            first.remove(child)
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = new
    first.append(t)


def replace_in_p(p_el, old, new):
    """Replace text inside a paragraph; run-by-run when possible, so mixed
    formatting (bold words, italics) survives."""
    if old not in ptext(p_el):
        return False
    for t in p_el.iter(qn('w:t')):
        if t.text and old in t.text:
            t.text = t.text.replace(old, new)
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            return True
    # spans several runs: merge text into the run where the match starts
    ts = [t for t in p_el.iter(qn('w:t'))]
    full = ''.join(t.text or '' for t in ts)
    i = full.index(old)
    pos = 0
    start_t = None
    for t in ts:
        L = len(t.text or '')
        if start_t is None and pos + L > i:
            start_t = t
            offset = i - pos
            break
        pos += L
    # put the whole remainder of the paragraph text from this run on in start_t
    rest_ts = ts[ts.index(start_t):]
    rest = ''.join(t.text or '' for t in rest_ts)
    start_t.text = (start_t.text or '')[:offset] + rest[offset:].replace(old, new, 1)
    start_t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    for t in rest_ts[1:]:
        t.text = ''
    return True


def new_p_after(ref_el, text, like=None):
    """Insert a paragraph after ref_el, copying pPr/rPr from `like` (a w:p)."""
    p = OxmlElement('w:p')
    src = like if like is not None else (ref_el if is_p(ref_el) else None)
    if src is not None and src.find(qn('w:pPr')) is not None:
        pp = copy.deepcopy(src.find(qn('w:pPr')))
        for tag in ('w:pageBreakBefore', 'w:numPr', 'w:sectPr'):
            x = pp.find(qn(tag))
            if x is not None:
                pp.remove(x)
        p.append(pp)
    r = OxmlElement('w:r')
    if src is not None:
        r0 = src.find(qn('w:r'))
        if r0 is not None and r0.find(qn('w:rPr')) is not None:
            r.append(copy.deepcopy(r0.find(qn('w:rPr'))))
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = text
    r.append(t)
    p.append(r)
    ref_el.addnext(p)
    return p


def body_children(body):
    return [el for el in body if el.tag in (qn('w:p'), qn('w:tbl'))]


def find_p(body, pred, start=None):
    """First body-level paragraph whose text satisfies pred (after `start`)."""
    seen = start is None
    for el in body:
        if not seen:
            seen = el is start
            continue
        if is_p(el) and pred(ptext(el)):
            return el
    return None


def next_sig(el):
    """Next sibling that is a table or a non-empty / drawing paragraph."""
    n = el.getnext()
    while n is not None and is_p(n) and not ptext(n).strip() and not has_drawing(n):
        n = n.getnext()
    return n


def prev_sig(el):
    n = el.getprevious()
    while n is not None and is_p(n) and not ptext(n).strip() and not has_drawing(n):
        n = n.getprevious()
    return n


def is_spacer(el):
    if not is_p(el) or ptext(el).strip() or has_drawing(el):
        return False
    for tag in ('w:sectPr', 'w:fldChar', 'w:instrText', 'w:bookmarkStart', 'w:object'):
        if el.find('.//' + qn(tag)) is not None:
            return False
    return True


# ---------------------------------------------------------------- styles
def ensure_style(doc, style_id_, name, based_on, ppr_xml=None, rpr_xml=None, outline_none=False):
    styles = doc.styles.element
    for s in styles.findall(qn('w:style')):
        if s.get(qn('w:styleId')) == style_id_:
            return s
    base = None
    for s in styles.findall(qn('w:style')):
        if s.get(qn('w:styleId')) == based_on:
            base = s
    st = OxmlElement('w:style')
    st.set(qn('w:type'), 'paragraph')
    st.set(qn('w:customStyle'), '1')
    st.set(qn('w:styleId'), style_id_)
    n = OxmlElement('w:name'); n.set(qn('w:val'), name); st.append(n)
    b = OxmlElement('w:basedOn'); b.set(qn('w:val'), based_on); st.append(b)
    q = OxmlElement('w:qFormat'); st.append(q)
    if base is not None and base.find(qn('w:pPr')) is not None:
        pp = copy.deepcopy(base.find(qn('w:pPr')))
    else:
        pp = OxmlElement('w:pPr')
    if outline_none:
        ol = pp.find(qn('w:outlineLvl'))
        if ol is not None:
            pp.remove(ol)
        set_ppr_child(pp, 'w:outlineLvl', {'w:val': 9})     # 9 = body text, never in the TOC
        for tag in ('w:pageBreakBefore', 'w:numPr'):
            x = pp.find(qn(tag))
            if x is not None:
                pp.remove(x)
    st.append(pp)
    if base is not None and base.find(qn('w:rPr')) is not None:
        st.append(copy.deepcopy(base.find(qn('w:rPr'))))
    styles.append(st)
    log(f'style added: {name}')
    return st


def style_el(doc, sid):
    for s in doc.styles.element.findall(qn('w:style')):
        if s.get(qn('w:styleId')) == sid:
            return s
    return None


def style_ppr(st):
    pp = st.find(qn('w:pPr'))
    if pp is None:
        pp = OxmlElement('w:pPr')
        rpr = st.find(qn('w:rPr'))
        (rpr.addprevious(pp) if rpr is not None else st.append(pp))
    return pp


def style_id_by_name(doc, name):
    for s in doc.styles.element.findall(qn('w:style')):
        n = s.find(qn('w:name'))
        if n is not None and n.get(qn('w:val')).lower() == name.lower():
            return s.get(qn('w:styleId'))
    return None


def set_style(p_el, sid):
    pp = ppr(p_el)
    set_ppr_child(pp, 'w:pStyle', {'w:val': sid})


# ---------------------------------------------------------------- main fix
def fix(path_in, path_out, seventh=False, protect=()):
    doc = docx.Document(path_in)
    body = doc.element.body
    # Protected sections (e.g. a school-specific section that must stay exactly
    # as written): from the paragraph whose text is the title up to the next
    # top-level heading. They are lifted out, the fixes run, then they go back
    # unchanged in the same place.
    _protected = []
    for _title in protect:
        _start = next((el for el in body if el.tag == qn('w:p') and ptext(el).strip() == _title), None)
        if _start is None:
            LOG.append(f'WARNING: protected section "{_title}" not found')
            continue
        _sec = [_start]
        _n = _start.getnext()
        while _n is not None and _n.tag in (qn('w:p'), qn('w:tbl')):
            _sid = style_id(_n) if _n.tag == qn('w:p') else ''
            if _sid.lower().startswith('heading') and ptext(_n).strip() and _sid == style_id(_start) or \
                    (_sid.lower() in ('heading', 'heading1') and ptext(_n).strip()):
                break
            _sec.append(_n)
            _n = _n.getnext()
        _hold = OxmlElement('w:p')
        _r = OxmlElement('w:r'); _t = OxmlElement('w:t'); _t.text = '@@PROTECTED SECTION@@'
        _r.append(_t); _hold.append(_r)
        _start.addprevious(_hold)
        for _el in _sec:
            body.remove(_el)
        _protected.append((_hold, _sec))
        LOG.append(f'protected (left unchanged): "{_title}" ({len(_sec)} paragraphs/tables)')
    sect = doc.sections[0]
    text_w = int(sect.page_width - sect.left_margin - sect.right_margin)   # EMU
    text_w_tw = int(text_w / 635)                                            # twips
    H1 = style_id_by_name(doc, 'heading 1')
    H2 = style_id_by_name(doc, 'heading 2')
    H3 = style_id_by_name(doc, 'heading 3')
    H4 = style_id_by_name(doc, 'heading 4')
    BODY = style_id_by_name(doc, 'Body') or style_id_by_name(doc, 'Normal')
    heading_ids = {x for x in (H1, H2, H3, H4) if x}

    def all_ps():
        return list(body.iter(qn('w:p')))

    # ============ text fixes (T11, T13, T17-T19, C5, C10, C13, C15, C19) ============
    REPL = [
        # C5: the value now carries "% (5 of 11)" itself
        *[(f'{{{{{v}}}}}%', f'{{{{{v}}}}}') for v in
          ['eighthbelowmeasure12', 'tenthbelowmeasure', 'twelvebelowmeasure', 'eighthbelowmeasure30',
           'tenthbelowmeasure_30', 'twelvebelowmeasure_30']],
        # C10: grade lists from grades with respondents
        ('8th, 10th, and 12th graders', '{{cmp_grades_text}} graders'),
        ('8th, 10th and 12th graders', '{{cmp_grades_text}} graders'),
        ('8th, 10th and 12th grade students', '{{cmp_grades_text}} grade students'),
        ('8th, 10th, and 12th grade students', '{{cmp_grades_text}} grade students'),
        # C12: the "cluster" claim is now part of the data-driven sentence
        ('These findings show how risk factors can “cluster.” {{clustering_sentence}}', '{{clustering_sentence}}'),
        # C13
        ('Of students who have had alcohol within the past 30 days:', '{{circumstances_lead}}'),
        # C15
        ('We cannot conclude from this data that a causal relationship exists between alcohol or cannabis use and grades. '
         'The findings are, however, in keeping with other research that posits a negative association between '
         'drinking/drug use and academic performance.', '{{grades_closing_sentence}}'),
        # C19: quoted answer wording comes from the survey
        ('believe that teenage drinking “is never a good thing to do,” or that “drinking is alright, but a person should not get drunk.”',
         'believe that {{att_never_good_quote}}'),
        # T17
        ('health of individual student’s and the wider community', 'health of individual students and the wider community'),
        ("health of individual student's and the wider community", 'health of individual students and the wider community'),
        # T18
        ('Our initial recommendations based on the SUPP results are found in the Executive Summary of this report.  In sum, '
         'Prevention Ed recommends', 'Based on the SUPP results, Prevention Ed recommends'),
        ('Our initial recommendations based on the SUPP results are found in the Executive Summary of this report. In sum, '
         'Prevention Ed recommends', 'Based on the SUPP results, Prevention Ed recommends'),
        ('We would like to restate here, however, our belief that the school community', 'We also believe that the school community'),
        # T19
        ('Prevention Education Solutions', 'Prevention Ed'),
        ('FCD', 'Prevention Ed'),
        # T13
        ('High-risk drinkers', 'Higher-risk drinkers'),
        ('Low-risk drinkers', 'Lower-risk drinkers'),
        ('“Higher-risk”drinkers', '“Higher-risk” drinkers'),
        # T11 (references to section names in running text)
        ('in the Frequency of cannabis use section', 'in the Frequency of Cannabis Use section'),
        ('Alcohol and cannabis use in Past 12 Months', 'Alcohol and Cannabis Use in Past 12 Months'),
    ]
    for p in all_ps():
        for old, new in REPL:
            while old in ptext(p):
                if not replace_in_p(p, old, new):
                    break
                log(f'text: "{old[:50]}" -> "{new[:50]}"')
    # T20: "good kid" / "bad kid" style wording
    for p in all_ps():
        if re.search(r'(?i)\b(good|bad) kids?\b', ptext(p)):
            log(f'WARNING: "good/bad kid" wording needs a manual rewrite: {ptext(p)[:80]}')
    # T11 headings / labels: one capitalisation
    CAPS = {'Cannabis use': 'Cannabis Use', 'Frequency of cannabis use': 'Frequency of Cannabis Use',
            'Cannabis use Lifetime Prevalence': 'Cannabis Use Lifetime Prevalence',
            'Cannabis use Past 12 Months': 'Cannabis Use Past 12 Months',
            'Cannabis use Past 30 Days': 'Cannabis Use Past 30 Days',
            'Younger Student’s Perceptions of 12th Grade Alcohol Use': 'Younger Students’ Perceptions of 12th Grade Alcohol Use',
            "Younger Student's Perceptions of 12th Grade Alcohol Use": "Younger Students' Perceptions of 12th Grade Alcohol Use"}
    for p in all_ps():
        t = ptext(p).strip()
        if t in CAPS:
            set_text(p, CAPS[t])
            log(f'T11: "{t}" -> "{CAPS[t]}"')

    # ============ C25: survey date and report date, labelled ============
    first_cover = True
    for p in all_ps():
        if ptext(p).strip() == '{{date_today}}':
            if first_cover:
                # main cover page: survey date only (no report date)
                set_text(p, 'Survey date: {{date_survey}}')
                first_cover = False
                log('C25: cover page shows the survey date')
                continue
            set_text(p, 'Report date: {{report_date}}')
            sp = new_p_after(p, 'Survey date: {{date_survey}}', like=p)
            p.addprevious(sp)          # survey date first, then report date
            log('C25: labelled survey date + report date')

    # NOTE: bright green text shading marks paragraphs that need a manual review
    # in every report. It is left exactly as it is (the pipeline lists those
    # paragraphs in _QA_lint.txt).

    # ============ T1: spacer paragraphs ============
    # Cover page: its vertical position came from 14 empty paragraphs; use
    # space-before on the school name instead.
    first_text = next(el for el in body if is_p(el) and ptext(el).strip())
    # Empty paragraphs that only hold bookmarks (e.g. empty Heading 1 lines used
    # as spacers): move the bookmarks to the next real paragraph, then they are
    # ordinary spacers. Left in place they become blank pages once Heading 1
    # starts a new page.
    moved = 0
    for el in list(body):
        if not is_p(el) or ptext(el).strip() or has_drawing(el):
            continue
        if el.find('.//' + qn('w:sectPr')) is not None or el.find('.//' + qn('w:fldChar')) is not None:
            continue
        bms = el.findall(qn('w:bookmarkStart')) + el.findall(qn('w:bookmarkEnd'))
        if not bms:
            continue
        nxt = el.getnext()
        while nxt is not None and not (is_p(nxt) and ptext(nxt).strip()):
            nxt = nxt.getnext()
        if nxt is None:
            continue
        anchor = nxt.find(qn('w:pPr'))
        for bm in bms:
            el.remove(bm)
            if anchor is not None:
                anchor.addnext(bm); anchor = bm
            else:
                nxt.insert(0, bm); anchor = bm
        moved += 1
    log(f'T1: moved bookmarks out of {moved} empty paragraphs')
    removed = 0
    for el in list(body):
        if is_spacer(el):
            # a hard page break inside an empty paragraph is handled below (T5)
            if el.find('.//' + qn('w:br')) is not None:
                continue
            body.remove(el)
            removed += 1
    set_ppr_child(ppr(first_text), 'w:spacing', {'w:before': int(2.6 * 1440)})
    log(f'T1: removed {removed} empty spacer paragraphs from the body')
    # empty paragraphs inside table cells next to real text (T12: one paragraph per header cell)
    removed = 0
    for tc in body.iter(qn('w:tc')):
        ps = tc.findall(qn('w:p'))
        if len(ps) > 1 and any(ptext(p).strip() or has_drawing(p) for p in ps):
            for p in ps:
                if is_spacer(p) and p.find('.//' + qn('w:br')) is None:
                    tc.remove(p)
                    removed += 1
    log(f'T12: removed {removed} empty paragraphs inside table cells')

    # ============ T5: one pagination mechanism (no hard page breaks) ============
    n_br = 0
    for br in list(body.iter(qn('w:br'))):
        if br.get(qn('w:type')) != 'page':
            continue
        p = br.getparent().getparent()
        r = br.getparent()
        r.remove(br)
        nxt = p.getnext() if not ptext(p).strip() else p.getnext()
        while nxt is not None and is_p(nxt) and is_spacer(nxt):
            nxt = nxt.getnext()
        if is_p(nxt) and not nxt.find('.//' + qn('w:sectPr')) is not None:
            set_flag(nxt, 'w:pageBreakBefore')
        if is_p(p) and is_spacer(p) and p.getparent() is body:
            body.remove(p)
        n_br += 1
    log(f'T5: replaced {n_br} hard page breaks with "page break before"')

    # charts placed in a heading style (e.g. {{ heavy_table_chart }} in Heading 2) -> Body
    for p in body.findall(qn('w:p')):
        if style_id(p) in heading_ids and re.fullmatch(r'\s*\{\{\s*\w+\s*\}\}\s*', ptext(p)) and 'chart' in ptext(p).lower():
            set_style(p, BODY)
            log(f'chart paragraph taken out of a heading style: {ptext(p).strip()}')
    # empty headings (spacers in a heading style) are gone with T1; headings with
    # text keep with the next paragraph
    # ============ T2: page break before every section title ============
    h1 = style_el(doc, H1)
    set_ppr_child(style_ppr(h1), 'w:keepNext')
    set_ppr_child(style_ppr(h1), 'w:pageBreakBefore')
    log('T2: Heading 1 style = page break before + keep with next')
    for p in body.findall(qn('w:p')):
        if style_id(p) == H1:
            # a pageBreakBefore=0 left in the paragraph would cancel the style
            x = p.find(qn('w:pPr') + '/' + qn('w:pageBreakBefore'))
            if x is not None:
                x.getparent().remove(x)
            prv = prev_sig(p)
            if is_p(prv) and style_id(prv) in heading_ids | {'Title'}:
                set_flag(p, 'w:pageBreakBefore', on=False)   # e.g. KEY DATA ANALYSIS -> Alcohol Use
        if style_id(p) in heading_ids:
            set_flag(p, 'w:keepNext')
    for p in body.findall(qn('w:p')):
        if ptext(p).strip() == 'Vaping Nicotine' and style_id(p) in heading_ids:
            set_flag(p, 'w:pageBreakBefore')
            log('T2: page break before "Vaping Nicotine"')

    # ============ T15: sub-headings that are not section titles: non-TOC style ============
    SUB = ensure_style(doc, 'SUPPSubheading', 'SUPP Subheading', H4 or H3, outline_none=True)
    LABEL = ensure_style(doc, 'SUPPTableLabel', 'SUPP Table Label', BODY, outline_none=True)
    lp = style_ppr(LABEL)
    set_ppr_child(lp, 'w:keepNext')
    set_ppr_child(lp, 'w:spacing', {'w:before': 120, 'w:after': 60, 'w:line': 240, 'w:lineRule': 'auto'})
    set_ppr_child(lp, 'w:jc', {'w:val': 'left'})
    CMP = ensure_style(doc, 'SUPPComparisonTitle', 'SUPP Comparison Title', H4 or H3, outline_none=True)
    cp = style_ppr(CMP)
    set_ppr_child(cp, 'w:keepNext')
    set_ppr_child(cp, 'w:jc', {'w:val': 'center'})
    set_ppr_child(cp, 'w:spacing', {'w:before': 0, 'w:after': 120})
    for p in body.findall(qn('w:p')):
        if style_id(p) in {H3, H4} and ptext(p).strip():
            set_style(p, 'SUPPSubheading')
    log('T15: Heading 3/4 sub-headings moved to "SUPP Subheading" (not in the TOC)')

    # ============ T10: one comparison title for alcohol, cannabis and vaping ============
    n_t10 = 0
    for p in list(body.findall(qn('w:p'))):
        t = ptext(p).strip()
        if re.fullmatch(r'Comparison [Bb]etween U\.S\. Norms', t):
            nxt = p.getnext()
            set_text(p, '{{cmp_grades_title}}')
            set_style(p, 'SUPPComparisonTitle')
            for tag in ('w:jc', 'w:ind'):
                set_ppr_child(ppr(p), tag, remove=True)
            set_flag(p, 'w:pageBreakBefore')     # 2.2: page break before the comparison title
            if is_p(nxt) and re.match(r'\s*and \{\{School\}\}', ptext(nxt)):
                body.remove(nxt)
            n_t10 += 1
    log(f'T10: {n_t10} comparison titles made identical (centered, one paragraph)')

    # ============ C12: cluster examples printed only where the gap is real ============
    cl = find_p(body, lambda t: '{{cigarettes_12_risk10_12[0]}}' in t)
    if cl is not None:
        ex = [cl]
        n = cl.getnext()
        while is_p(n) and re.search(r'risk10_12\[|risk_12_bullets\[', ptext(n)):
            ex.append(n)
            n = n.getnext()
        start = new_p_after(ex[-1], '{%p for b in cluster_examples %}', like=ex[0])
        mid = new_p_after(start, '{{ b }}', like=ex[0])
        # keep the bullet numbering on the repeated line
        if ex[0].find(qn('w:pPr') + '/' + qn('w:numPr')) is not None:
            set_ppr_child(ppr(mid), 'w:numPr')
            ppr(mid).replace(ppr(mid).find(qn('w:numPr')), copy.deepcopy(ex[0].find(qn('w:pPr') + '/' + qn('w:numPr'))))
        new_p_after(mid, '{%p endfor %}', like=ex[0])
        for e in ex:
            body.remove(e)
        log(f'C12: {len(ex)} fixed cluster bullets replaced by a data-driven list')

    # ============ C24: 8th-grade-only blocks shown only when 8th graders took the survey ============
    def wrap_if(first_el, last_el, cond='has_8th'):
        p_if = OxmlElement('w:p'); r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = '{%p if ' + cond + ' %}'
        r.append(t); p_if.append(r); first_el.addprevious(p_if)
        p_end = OxmlElement('w:p'); r = OxmlElement('w:r'); t = OxmlElement('w:t'); t.text = '{%p endif %}'
        r.append(t); p_end.append(r); last_el.addnext(p_end)

    def notes_after(tbl):
        last = tbl
        n = tbl.getnext()
        while is_p(n) and ptext(n).strip().startswith('*'):
            last = n
            n = n.getnext()
        return last

    n_c24 = 0
    for var in ['ALLSUBSTANCE8th_12_MTF', 'ALLSUBSTANCE8th_30_MTF']:
        tbl = next((t for t in body.findall(qn('w:tbl')) if var in ptext(t)), None)
        _pp = prev_sig(prev_sig(tbl)) if tbl is not None else None
        if tbl is None or (_pp is not None and 'has_8th' in ptext(_pp)):
            continue
        label = prev_sig(tbl)
        intro = prev_sig(label) if is_p(label) else None
        first = intro if (is_p(intro) and '8th graders' in ptext(intro)) else label
        wrap_if(first, notes_after(tbl))
        n_c24 += 1
    tbl8 = next((t for t in body.findall(qn('w:tbl')) if 'sub12_by_sex_MTFcombo8th' in ptext(t)), None)
    if tbl8 is not None:
        last = notes_after(tbl8)
        n = last.getnext()
        while is_p(n) and 'outputstringssexbullets[' in ptext(n):
            last = n
            n = n.getnext()
        wrap_if(tbl8, last)
        n_c24 += 1
    log(f'C24: {n_c24} 8th-grade-only blocks made conditional on has_8th')

    # gender tables: one grade per page, each with its own label (2.2)
    gen = find_p(body, lambda t: t.strip() == 'Substance Use Over Past 12 Months')
    if gen is not None and is_tbl(next_sig(gen)) is False:
        gen = None
    first_done = False
    for g in ['8th', '10th', '12th']:
        tbl = next((t for t in body.findall(qn('w:tbl')) if f'sub12_by_sex_MTFcombo{g}' in ptext(t)), None)
        if tbl is None:
            continue
        prv = tbl.getprevious()
        if is_p(prv) and 'by Gender Compared to U.S. Norms' in ptext(prv):
            continue                      # already fixed
        lab = OxmlElement('w:p')
        tbl.addprevious(lab)              # inside the has_8th block for 8th
        set_text(lab, f'{g}-Grade Substance Use Over the Past 12 Months by Gender Compared to U.S. Norms')
        set_style(lab, 'SUPPTableLabel')
        if first_done:
            set_flag(lab, 'w:pageBreakBefore')
        first_done = True
    gen = find_p(body, lambda t: t.strip() == 'Substance Use Over Past 12 Months')
    if gen is not None and 'by Gender' in ptext(gen.getprevious() if is_p(gen.getprevious()) else gen) or \
            (gen is not None and is_p(next_sig(gen)) and '{%p if has_8th %}' in ptext(next_sig(gen))):
        body.remove(gen)                  # replaced by the per-grade labels
    for _t in ('Past-12-Months Substance Use by Gender', 'Past-12-Months Substance Use Comparisons',
               'Past-30-Days Substance Use Comparisons'):
        _gp = find_p(body, lambda t, _t=_t: t.strip() == _t)
        if _gp is not None:
            set_style(_gp, 'SUPPSubheading')
            set_flag(_gp, 'w:keepNext')
    log('2.2: gender tables labelled per grade, one grade per page')

    # C10/C24: statements about 7th-8th graders only when the school has them;
    # grade-range labels come from the grades present
    n_low = 0
    for p in list(body.findall(qn('w:p'))):
        t = ptext(p)
        if '{{seventhand8th' in t and 'has_lower_grades' not in ptext(p.getprevious() if p.getprevious() is not None else p):
            replace_in_p(p, '7th-8th graders', '{{lower_grades_label}} graders')
            wrap_if(p, p, 'has_lower_grades')
            n_low += 1
        elif '9th-12th graders' in t and '{{ninto12th' in t:
            replace_in_p(p, '9th-12th graders', '{{upper_grades_label}} graders')
    log(f'C10: {n_low} 7th-8th statements made conditional; grade-range labels from the data')

    # ============ C17: footnote under the "Greatest Number" table ============
    gl = find_p(body, lambda t: t.strip().startswith('Greatest Number of Whole Drinks Consumed in a Row'))
    if gl is not None and '{{greatest_footnote}}' not in ptext(body):
        tbl = next_sig(gl)
        if is_tbl(tbl):
            fp = new_p_after(tbl, '{{greatest_footnote}}', like=next_sig(tbl) if is_p(next_sig(tbl)) else None)
            for r in fp.iter(qn('w:rPr')):
                sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '18'); r.append(sz)
            log('C17: footnote placeholder added under the Greatest Number table')

    # ============ C2/C3/C7/C8: methodology text ============
    ip = find_p(body, lambda t: t.startswith('We identified {{removed_N}} students'))
    if ip is not None and '{{exclusion_summary}}' not in ptext(body):
        new_p_after(ip, '{{exclusion_summary}}', like=ip)
        log('C2: exclusion summary added after "We identified..."')
    rp = find_p(body, lambda t: t.strip().startswith('Yes. The survey was administered to all students'))
    if rp is not None and '{{time_rule_text}}' not in ptext(body):
        _has_k10 = 'Kessler' in ptext(body)
        q = new_p_after(rp, 'How are past-12-month use and mental health scores counted?' if _has_k10 else
                        'How is past-12-month use counted?',
                        like=find_p(body, lambda t: t.strip().startswith('Were survey respondents representative')))
        a1 = new_p_after(q, '{{time_rule_text}}', like=rp)
        if _has_k10:
            new_p_after(a1, '{{k10_method_text}}', like=rp)
        log('C7/C8: counting rules and K10 scoring added to the methodology')

    # ============ C20 / T14: bullets and statistic + "Yet," pairs ============
    ps = body.findall(qn('w:p'))
    n_c20 = 0
    for i in range(1, len(ps) - 1):
        a, b, c = ps[i - 1], ps[i], ps[i + 1]
        na = a.find(qn('w:pPr') + '/' + qn('w:numPr'))
        nc = c.find(qn('w:pPr') + '/' + qn('w:numPr'))
        nb = b.find(qn('w:pPr') + '/' + qn('w:numPr'))
        if na is not None and nc is not None and nb is None and ptext(b).strip() and \
                re.match(r'\s*(\{\{|[\d.]+%|Yet,)', ptext(b)) and style_id(a) == style_id(b) == style_id(c):
            ppr(b)
            set_ppr_child(ppr(b), 'w:numPr')
            ppr(b).replace(ppr(b).find(qn('w:numPr')), copy.deepcopy(na))
            n_c20 += 1
            log(f'C20: bullet added to "{ptext(b)[:60]}"')
    n_t14 = 0
    for p in body.findall(qn('w:p')):
        if ptext(p).strip().startswith('Yet,'):
            prv = p.getprevious()
            if is_p(prv) and ptext(prv).strip():
                set_flag(prv, 'w:keepNext')
                set_ppr_child(ppr(prv), 'w:spacing', {'w:before': 120, 'w:after': 0})
                set_ppr_child(ppr(p), 'w:spacing', {'w:before': 0, 'w:after': 160})
                n_t14 += 1
    log(f'T14: {n_t14} statistic / "Yet," pairs kept together')

    # ============ T9: one label style for question text + table titles ============
    n_lab = 0
    for tbl in body.findall(qn('w:tbl')):
        rows = tbl.findall(qn('w:tr'))
        if len(rows) == 1 and len(rows[0].findall(qn('w:tc'))) == 1:
            continue            # highlight / definition boxes
        prv = prev_sig(tbl)
        for _ in range(2):
            _t = ptext(prv).strip() if is_p(prv) else ''
            if _t and len(_t) < 220 and style_id(prv) not in heading_ids \
                    and not _t.startswith(('{%', '*')) and style_id(prv) not in ('SUPPComparisonTitle',) \
                    and not re.fullmatch(r'\{\{[^}]*\}\}', _t) and not _t.endswith(('.', ':', '.”', '."')) \
                    and prv.find(qn('w:pPr') + '/' + qn('w:numPr')) is None:
                set_style(prv, 'SUPPTableLabel')
                for tag in ('w:jc', 'w:spacing', 'w:ind'):
                    set_ppr_child(ppr(prv), tag, remove=True)
                n_lab += 1
                prv = prev_sig(prv)
            else:
                break
    log(f'T9: {n_lab} question/table labels use "SUPP Table Label" (left aligned, single spacing)')

    # ============ T12: header cells: one paragraph, identical spacing ============
    for tbl in body.iter(qn('w:tbl')):
        rows = tbl.findall(qn('w:tr'))
        for tr in rows[:2]:
            for tc in tr.findall(qn('w:tc')):
                for p in tc.findall(qn('w:p')):
                    set_ppr_child(ppr(p), 'w:spacing', {'w:before': 0, 'w:after': 0, 'w:line': 240, 'w:lineRule': 'auto'})

    # ============ T8 + Part 3: one table width, set column widths ============
    def set_widths(tbl, widths_in):
        grid = tbl.find(qn('w:tblGrid'))
        cols = grid.findall(qn('w:gridCol'))
        if len(cols) != len(widths_in):
            return False
        tw = [int(w * 1440) for w in widths_in]
        for gc, w in zip(cols, tw):
            gc.set(qn('w:w'), str(w))
        tpr = tbl.find(qn('w:tblPr'))
        for tag, attrs in [('w:tblW', {'w:w': sum(tw), 'w:type': 'dxa'}), ('w:tblInd', {'w:w': 0, 'w:type': 'dxa'})]:
            x = tpr.find(qn(tag))
            if x is None:
                x = OxmlElement(tag)
                tpr.append(x)
            for k, v in attrs.items():
                x.set(qn(k), str(v))
        lay = tpr.find(qn('w:tblLayout'))
        if lay is None:
            lay = OxmlElement('w:tblLayout'); tpr.append(lay)
        lay.set(qn('w:type'), 'fixed')
        for tr in tbl.findall(qn('w:tr')):
            pos = 0
            for tc in tr.findall(qn('w:tc')):
                gs = tc.find(qn('w:tcPr') + '/' + qn('w:gridSpan'))
                span = int(gs.get(qn('w:val'))) if gs is not None else 1
                tcpr = tc.find(qn('w:tcPr'))
                if tcpr is None:
                    tcpr = OxmlElement('w:tcPr'); tc.insert(0, tcpr)
                w = tcpr.find(qn('w:tcW'))
                if w is None:
                    w = OxmlElement('w:tcW'); tcpr.insert(0, w)
                w.set(qn('w:w'), str(sum(tw[pos:pos + span]))); w.set(qn('w:type'), 'dxa')
                pos += span
        return True

    full = text_w_tw / 1440.0
    counts = {}
    for tbl in body.findall(qn('w:tbl')):
        rows = tbl.findall(qn('w:tr'))
        ncol = len(tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol')))
        txt = ptext(tbl)
        lab = ptext(prev_sig(tbl)) if is_p(prev_sig(tbl)) else ''
        if len(rows) == 1 and ncol == 1:
            kind, w = 'box', [full]
        elif 'highlowrisk_consequence' in txt and ncol == 3:
            kind, w = 'consequences', [4.7] + [(full - 4.7) / 2] * 2
        elif re.search(r'_k10|K10', txt + lab) and ncol == 7:
            kind, w = 'k10', [1.35] + [(full - 1.35) / 6] * 6
        elif re.search(r'ALLSUBSTANCE|sub12_by_sex|all_sub_risks', txt):
            kind, w = 'substance', [1.85] + [(full - 1.85) / (ncol - 1)] * (ncol - 1)
        elif lab.startswith('During the Past 30 Days') and ncol in (6, 7):
            kind, w = 'freq30', [2.0] + [(full - 2.0) / (ncol - 1)] * (ncol - 1)
        elif ncol in (6, 7) and 'Grade 12' in txt and re.search(r'(?i)parent|family|guardian', lab):
            kind, w = 'parental', [1.5] + [(full - 1.5) / (ncol - 1)] * (ncol - 1)
        elif ncol in (6, 7) and 'Grade 12' in txt:
            kind, w = 'agree', [1.25] + [(full - 1.25) / (ncol - 1)] * (ncol - 1)
        else:
            # every other table: same total width, label column kept, the rest equal
            grid = [int(gc.get(qn('w:w')) or 0) for gc in tbl.find(qn('w:tblGrid')).findall(qn('w:gridCol'))]
            first = min(grid[0] / 1440.0, full * 0.45) if grid and grid[0] else full / ncol
            kind, w = 'other', [first] + [(full - first) / max(1, ncol - 1)] * (ncol - 1)
        if set_widths(tbl, w):
            counts[kind] = counts.get(kind, 0) + 1
    log(f'T8: table widths set to the text width ({full:.2f} in): {counts}')

    # ============ 2.2 approved page layouts ============
    def pbb(p, why):
        if is_p(p):
            set_flag(p, 'w:pageBreakBefore')
            log(f'2.2 page break before "{ptext(p).strip()[:50]}" ({why})')
    # U.S. norm comparisons: one page per measure
    for t in ['Alcohol Use Past 12 Months', 'Alcohol Use Past 30 Days', 'Cannabis Use Past 12 Months',
              'Cannabis Use Past 30 Days', 'Vaping Nicotine Past 30 Days']:
        pbb(find_p(body, lambda x, t=t: x.strip() == t), 'one measure per page')
    # Kessler: three pages
    pbb(find_p(body, lambda x: x.strip() == 'Kessler Psychological Distress Scale'), 'K10 page 1')
    pbb(find_p(body, lambda x: x.strip().startswith('K10 Anxiety Scores')), 'K10 page 2')
    pbb(find_p(body, lambda x: x.strip().startswith('Full K10 Scores')), 'K10 page 3')
    # Higher-risk subgroups p.B / p.C are Heading 1 (break from the style)
    # Culture of Use + school attitudes: two table+chart pairs per page; one when
    # the page also has intro text or bullets
    for sec_title in ['Culture of Use: Attitudes and Perceptions', 'Attitudes and Feelings About School']:
        h = find_p(body, lambda x, s=sec_title: x.strip() == s)
        if h is None:
            continue
        el = h.getnext()
        pairs = []
        cur = {'pre': [], 'label': None, 'extra': False}
        while el is not None and not (is_p(el) and style_id(el) == H1):
            if is_tbl(el):
                rows = el.findall(qn('w:tr'))
                if len(rows) == 1 and len(rows[0].findall(qn('w:tc'))) == 1:
                    el = el.getnext(); continue
                lab = prev_sig(el)
                cur['label'] = lab
            elif is_p(el) and re.search(r'\{\{\s*\w*chart\w*\s*\}\}', ptext(el), re.I):
                pairs.append(cur)
                cur = {'pre': [], 'label': None, 'extra': False}
            elif is_p(el) and ptext(el).strip():
                if cur['label'] is None:
                    cur['pre'].append(el)
                else:
                    cur['extra'] = True       # bullets between table and chart
            el = el.getnext()
        on_page, page_has_extra = 0, True     # page 1 carries the section intro
        for i, pr in enumerate(pairs):
            # the label paragraphs are part of 'pre' until the table is reached
            pre_text = [p for p in pr['pre'] if p is not pr['label'] and style_id(p) != 'SUPPTableLabel']
            extra = pr['extra'] or bool(pre_text and i > 0)
            if i == 0:
                on_page, page_has_extra = 1, True
                continue
            if page_has_extra or extra or on_page >= 2:
                first = pr['pre'][0] if pr['pre'] else pr['label']
                pbb(first, f'{sec_title}: pair {i + 1}')
                on_page, page_has_extra = 1, extra
            else:
                on_page += 1
    # Student stressors: same gaps everywhere (0.4 in before a question, 0.25 in to its table)
    h = find_p(body, lambda x: x.strip() == 'Student Stressors and Mental Health')
    if h is not None:
        el = h.getnext()
        while el is not None and not (is_p(el) and ptext(el).strip() == 'Kessler Psychological Distress Scale'):
            if is_p(el) and ptext(el).startswith('One of my greatest stressors'):
                set_ppr_child(ppr(el), 'w:spacing', {'w:before': int(0.4 * 1440), 'w:after': int(0.25 * 1440)})
            el = el.getnext()
        log('2.2: stressor questions spaced 0.4 in before / 0.25 in to the table')

    # ============ T15/T16: table of contents ============
    for it in body.iter(qn('w:instrText')):
        if it.text and it.text.strip().startswith('TOC'):
            it.text = ' TOC \\o "1-2" \\h \\z \\u '
            log('T15: TOC field limited to heading levels 1-2')
    for name in ('toc 1', 'toc 2'):
        sid = style_id_by_name(doc, name)
        st = style_el(doc, sid) if sid else None
        if st is None:
            continue
        rpr = st.find(qn('w:rPr'))
        if rpr is None:
            rpr = OxmlElement('w:rPr'); st.append(rpr)
        for tag in ('w:b', 'w:bCs', 'w:rFonts', 'w:sz', 'w:szCs', 'w:caps'):
            x = rpr.find(qn(tag))
            if x is not None:
                rpr.remove(x)
        f = OxmlElement('w:rFonts')
        for a in ('w:ascii', 'w:hAnsi', 'w:cs', 'w:eastAsia'):
            f.set(qn(a), 'Calibri')
        rpr.insert(0, f)
        b = OxmlElement('w:b'); b.set(qn('w:val'), '0'); rpr.append(b)
        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), '24'); rpr.append(sz)
        set_ppr_child(style_ppr(st), 'w:spacing', {'w:before': 0, 'w:after': 60, 'w:line': 276, 'w:lineRule': 'auto'})
    # TOC entry paragraphs carry their own bold/size: clear them so the style applies
    for p in body.findall(qn('w:p')):
        if style_id(p).lower().startswith('toc'):
            for rpr in p.iter(qn('w:rPr')):
                for tag in ('w:b', 'w:bCs', 'w:sz', 'w:szCs', 'w:rFonts', 'w:caps'):
                    x = rpr.find(qn(tag))
                    if x is not None:
                        rpr.remove(x)
            set_ppr_child(ppr(p), 'w:spacing', remove=True)
    log('T16: TOC styles Calibri 12, regular, 1.15 line spacing (Word refreshes it when the report opens)')

    for _hold, _sec in _protected:
        for _el in _sec:
            _hold.addprevious(_el)
        _hold.getparent().remove(_hold)
    doc.save(path_out)
    return LOG


if __name__ == '__main__':
    # python3 fix_templates.py IN.docx OUT.docx [--protect "Section title" ...]
    args = sys.argv[1:]
    prot = [args[i + 1] for i, a in enumerate(args) if a == '--protect']
    files = [a for i, a in enumerate(args) if a != '--protect' and (i == 0 or args[i - 1] != '--protect')]
    src, dst = files[0], files[1]
    for line in fix(src, dst, seventh='7th' in src, protect=prot):
        print(line)
