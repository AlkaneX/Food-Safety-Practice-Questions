"""Build the ServSafe question-bank spreadsheet (and the website's data) from questions.json.

    python3 servsafe/build.py

Writes servsafe/servsafe-question-bank.xlsx and refreshes the question data inside
servsafe/index.html (between the QUESTION-DATA markers). Needs openpyxl.
"""
import json
import re
from pathlib import Path

from openpyxl import Workbook
from openpyxl.comments import Comment
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

HERE = Path(__file__).resolve().parent
DATA = json.loads((HERE / 'questions.json').read_text(encoding='utf8'))
CHAPTERS = {int(k): v for k, v in DATA['chapters'].items()}
TOPICS = {t[0]: t for t in DATA['topics']}
SOURCES = DATA['sources']                # display order
ITEMS = sorted(DATA['items'], key=lambda x: (SOURCES.index(x['src']), x['id']))

TYPE_NAME = {'tf': 'True / false', 'mc': 'Multiple choice', 'multi': 'Select all that apply', 'order': 'Put in order',
             'list': 'List', 'fill': 'Fill in the blank', 'short': 'Short answer'}
FONT = 'Arial'
HEAD_FILL = PatternFill('solid', fgColor='1D2B36')
ASSIGNED_FILL = PatternFill('solid', fgColor='FFE98A')     # answer not in the source: assigned
PICTURE_FILL = PatternFill('solid', fgColor='F4D6CC')      # can't be answered without the missing picture
THIN = Side(style='thin', color='C9CFD4')
GRID = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)
TOP = Alignment(wrap_text=True, vertical='top')
LETTERS = 'ABCDEFGHIJKLMNOP'


def choices(x):
    if x['type'] in ('tf', 'mc', 'multi', 'order'):
        return '\n'.join(f'{LETTERS[i]}. {o}' for i, o in enumerate(x['opts']))
    return ''


def answer(x):
    if x.get('ansText'):
        return x['ansText']
    t = x['type']
    if t in ('tf', 'mc'):
        return f'{LETTERS[x["ans"]]}. {x["opts"][x["ans"]]}'
    if t == 'multi':
        return '\n'.join(f'{LETTERS[i]}. {x["opts"][i]}' for i in x['ans'])
    if t == 'order':
        return '\n'.join(f'{n}. {x["opts"][i]}' for n, i in enumerate(x['ans'], 1))
    if t == 'list':
        return '\n'.join('• ' + s for s in x['ansList'])
    return ''


def answer_from(x):
    if not x.get('assigned'):
        return 'Source'
    if x.get('img') and 'determined' in (x.get('ansText') or ''):
        return 'Not answerable (picture missing)'
    return f'Assigned ({x.get("conf", "medium")} confidence)'


def on_site(x):
    if x.get('web') is False:
        return 'Not on the website (picture missing)'
    if x.get('mcq'):
        return 'Multiple choice (rewritten)'
    return 'Multiple choice' if x['type'] == 'mc' else 'Knowledge point'


COLS = [('ID', 10), ('Source', 20), ('Where in source', 18), ('Chapter', 9), ('Chapter name', 26), ('Topic', 28),
        ('Question type', 15), ('Question', 55), ('Answer choices', 45), ('Correct answer', 40), ('Answer from', 17),
        ('Explanation', 48), ('What it tests', 48), ('Note', 32), ('On the website as', 20)]


def header(ws, row, labels):
    for c, label in enumerate(labels, 1):
        cell = ws.cell(row=row, column=c, value=label)
        cell.font = Font(name=FONT, bold=True, color='FFFFFF')
        cell.fill = HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical='center')
        cell.border = GRID


def question_rows(ws, items):
    header(ws, 1, [c for c, _ in COLS])
    for r, x in enumerate(items, 2):
        values = [x['id'], x['src'], x.get('no', ''), x['ch'], CHAPTERS[x['ch']], TOPICS[x['tp']][2],
                  TYPE_NAME[x['type']], x['q'], choices(x), answer(x), answer_from(x), x.get('why', ''), x.get('kn', ''),
                  x.get('note', ''), on_site(x)]
        fill = PICTURE_FILL if answer_from(x).startswith('Not') else ASSIGNED_FILL if x.get('assigned') else None
        for c, v in enumerate(values, 1):
            cell = ws.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10, bold=(c == 10))
            cell.alignment = TOP
            cell.border = GRID
            if fill:
                cell.fill = fill
    for c, (_, w) in enumerate(COLS, 1):
        ws.column_dimensions[get_column_letter(c)].width = w
    ws.freeze_panes = 'B2'
    ws.auto_filter.ref = f'A1:{get_column_letter(len(COLS))}{len(items) + 1}'
    ws.print_title_rows = '1:1'
    ws.page_setup.orientation = 'landscape'
    ws.sheet_properties.pageSetUpPr.fitToPage = True
    ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
    return len(items) + 1


def build_xlsx(out):
    wb = Workbook()
    wb.calculation.fullCalcOnLoad = True
    ws = wb.active
    ws.title = 'Questions'
    last = question_rows(ws, ITEMS)
    q = f'Questions!$%s$2:$%s${last}'
    src_rng, ch_rng, from_rng = q % ('B', 'B'), q % ('D', 'D'), q % ('K', 'K')

    sm = wb.create_sheet('Summary', 0)
    sm['A1'] = 'ServSafe Manager (9th edition) practice question bank'
    sm['A1'].font = Font(name=FONT, size=14, bold=True)
    # Short single-line notes (no merged cells) so every spreadsheet app shows them in full.
    notes = [
        (None, 'Every question from the class materials, one row each on the Questions sheet: question, all answer choices, correct answer,'),
        (None, 'chapter (1–10) and what it tests. Use the filter arrows on row 1 to pick a chapter or source.'),
        (ASSIGNED_FILL, 'Yellow rows: the source gives no answer, so one was assigned from ServSafe Manager 9th ed. / FDA Food Code rules'),
        (None, '(the Explanation column gives the rule). They are also listed on the Assigned answers sheet.'),
        (PICTURE_FILL, 'Pink rows: the question depends on a picture that isn’t in the file, so it can’t be answered. Kept for completeness only.'),
        (None, 'To quiz yourself, hide columns J–M (Correct answer, Answer from, Explanation, What it tests).'),
        (None, 'The website is multiple choice only: column O shows how each question appears there, and the Website rewrites sheet'),
        (None, 'shows true/false, select-all, put-in-order and fill-in questions next to their multiple-choice versions.'),
    ]
    for i, (fill, n) in enumerate(notes, 2):
        sm.cell(row=i, column=2, value=n).font = Font(name=FONT, size=10)
        if fill:
            sm.cell(row=i, column=1).fill = fill

    hrow = len(notes) + 3
    header(sm, hrow, ['Chapter', 'Chapter name'] + SOURCES + ['Total', 'Assigned answers'])
    tcol = get_column_letter(3 + len(SOURCES))
    for r, ch in enumerate(range(1, 11), hrow + 1):
        sm.cell(row=r, column=1, value=ch)
        sm.cell(row=r, column=2, value=CHAPTERS[ch])
        for c, s in enumerate(SOURCES, 3):
            col = get_column_letter(c)
            sm.cell(row=r, column=c, value=f'=COUNTIFS({ch_rng},$A{r},{src_rng},{col}${hrow})')
        sm.cell(row=r, column=3 + len(SOURCES), value=f'=SUM(C{r}:{get_column_letter(2 + len(SOURCES))}{r})')
        sm.cell(row=r, column=4 + len(SOURCES), value=f'=COUNTIFS({ch_rng},$A{r},{from_rng},"Assigned*")')
    trow = hrow + 11
    sm.cell(row=trow, column=1, value='Total')
    for c in range(3, 5 + len(SOURCES)):
        col = get_column_letter(c)
        sm.cell(row=trow, column=c, value=f'=SUM({col}{hrow + 1}:{col}{trow - 1})')
    for row in sm.iter_rows(min_row=hrow + 1, max_row=trow, max_col=4 + len(SOURCES)):
        for cell in row:
            cell.font = Font(name=FONT, size=10, bold=cell.row == trow or cell.column >= 3 + len(SOURCES))
            cell.border = GRID
            cell.alignment = Alignment(vertical='top', wrap_text=cell.column == 2)
    sm.row_dimensions[hrow].height = 44
    for c, w in enumerate([9, 34] + [13] * len(SOURCES) + [9, 11], 1):
        sm.column_dimensions[get_column_letter(c)].width = w
    sm.page_setup.orientation = 'landscape'
    sm.sheet_properties.pageSetUpPr.fitToPage = True
    sm.page_setup.fitToWidth, sm.page_setup.fitToHeight = 1, 0

    asg = wb.create_sheet('Assigned answers')
    question_rows(asg, [x for x in ITEMS if x.get('assigned')])

    # The website is multiple choice only: show each rewritten question next to its original.
    rw = wb.create_sheet('Website rewrites')
    rcols = [('ID', 10), ('Source', 20), ('Chapter', 9), ('Original type', 16), ('Original question', 50),
             ('Original answer', 32), ('Website question (multiple choice)', 50), ('Website choices', 45), ('Website answer', 32)]
    header(rw, 1, [c for c, _ in rcols])
    for r, x in enumerate([x for x in ITEMS if x.get('mcq')], 2):
        m = x['mcq']
        values = [x['id'], x['src'], x['ch'], TYPE_NAME[x['type']], x['q'], answer(x), m['q'],
                  '\n'.join(f'{LETTERS[i]}. {o}' for i, o in enumerate(m['opts'])), f'{LETTERS[m["ans"]]}. {m["opts"][m["ans"]]}']
        for c, v in enumerate(values, 1):
            cell = rw.cell(row=r, column=c, value=v)
            cell.font = Font(name=FONT, size=10, bold=(c == 9))
            cell.alignment = TOP
            cell.border = GRID
    for c, (_, w) in enumerate(rcols, 1):
        rw.column_dimensions[get_column_letter(c)].width = w
    rw.freeze_panes = 'B2'
    rw.auto_filter.ref = f'A1:I{rw.max_row}'

    tp = wb.create_sheet('Topics')
    header(tp, 1, ['Chapter', 'Topic', 'Questions'])
    for r, (tid, ch, name) in enumerate(DATA['topics'], 2):
        tp.cell(row=r, column=1, value=ch)
        tp.cell(row=r, column=2, value=name)
        tp.cell(row=r, column=3, value=f'=COUNTIF(Questions!$F$2:$F${last},B{r})')
    tr = len(DATA['topics']) + 2
    tp.cell(row=tr, column=2, value='Total')
    tp.cell(row=tr, column=3, value=f'=SUM(C2:C{tr - 1})')
    for row in tp.iter_rows(min_row=2, max_row=tr, max_col=3):
        for cell in row:
            cell.font = Font(name=FONT, size=10, bold=cell.row == tr)
            cell.border = GRID
    for c, w in enumerate([9, 46, 11], 1):
        tp.column_dimensions[get_column_letter(c)].width = w
    tp.freeze_panes = 'A2'
    wb.save(out)


def build_site():
    page = HERE / 'index.html'
    if not page.exists():
        return False
    html = page.read_text(encoding='utf8')
    payload = json.dumps({k: DATA[k] for k in ('chapters', 'topics', 'sources')} | {'items': ITEMS},
                         ensure_ascii=False, separators=(',', ':')).replace('</', '<\\/')
    new, n = re.subn(r'(<script id="question-data" type="application/json">).*?(</script>)',
                     lambda m: m.group(1) + payload + m.group(2), html, flags=re.S)
    if n != 1:
        raise SystemExit('index.html: question-data script tag not found')
    page.write_text(new, encoding='utf8')
    return True


if __name__ == '__main__':
    build_xlsx(HERE / 'servsafe-question-bank.xlsx')
    print(f'servsafe-question-bank.xlsx: {len(ITEMS)} questions, {sum(1 for x in ITEMS if x.get("assigned"))} assigned')
    if build_site():
        print('index.html: question data refreshed')
