from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os

app = Flask(__name__)
CORS(app)

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'EXTERIOR_MASTER_TEMPLATE.docx')

def set_cell_text(cell, new_text, bold=None, italic=None):
    for para in cell.paragraphs:
        if not para.runs:
            continue
        first = para.runs[0]
        for run in para.runs[1:]:
            run.text = ""
        first.text = new_text
        if bold is not None: first.bold = bold
        if italic is not None: first.italic = italic
        return

def set_row_cell_text(row_el, col_idx, text, bold=None):
    cells = row_el.findall(qn('w:tc'))
    if col_idx >= len(cells): return
    cell = cells[col_idx]
    for p in cell.findall(qn('w:p')):
        runs = p.findall(qn('w:r'))
        if runs:
            r = runs[0]
            t = r.find(qn('w:t'))
            if t is None:
                t = OxmlElement('w:t')
                r.append(t)
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            rpr = r.find(qn('w:rPr'))
            if rpr is not None:
                if bold is not None:
                    b = rpr.find(qn('w:b'))
                    if bold and b is None:
                        b = OxmlElement('w:b'); rpr.insert(0, b)
                    elif not bold and b is not None:
                        rpr.remove(b)
                i_el = rpr.find(qn('w:i'))
                if i_el is not None: rpr.remove(i_el)
            for extra_r in runs[1:]:
                t2 = extra_r.find(qn('w:t'))
                if t2 is not None: t2.text = ''
            break

def get_para_text(el, doc):
    try:
        from docx.text.paragraph import Paragraph
        return Paragraph(el, doc).text.strip()
    except:
        return ''

def get_table_first_cell(el, doc):
    try:
        from docx.table import Table
        return Table(el, doc).rows[0].cells[0].text.strip()
    except:
        return ''

def remove_side_block(doc, side_label):
    search = side_label + ' of House'
    body = doc.element.body
    children = list(body)
    for i, child in enumerate(children):
        tag = child.tag.split('}')[-1]
        if tag == 'p' and search in get_para_text(child, doc):
            # Remove heading + next 3 elements (paint brand para, table, empty para)
            to_remove = [child]
            j = i + 1
            while j < len(children) and len(to_remove) <= 3:
                nc = children[j]
                nc_tag = nc.tag.split('}')[-1]
                nc_text = get_para_text(nc, doc) if nc_tag == 'p' else get_table_first_cell(nc, doc)
                # Stop if we hit another major section
                if nc_text and any(x in nc_text for x in ['of House','PROJECT PHOTOS','WARRANTY','PAYMENT','COST','NEXT','ESTIMATE','Inspection','Materials']):
                    break
                to_remove.append(nc)
                j += 1
            for el in to_remove:
                try: body.remove(el)
                except: pass
            return

def rebuild_paint_table(tbl, surfaces):
    rows = list(tbl.rows)
    for row in rows[1:]:
        row._element.getparent().remove(row._element)
    hdr_el = tbl.rows[0]._element
    for idx, sf in enumerate(surfaces):
        qty = sf.get('qty')
        nm = f"{sf['name']} × {qty}" if qty and int(str(qty)) > 1 else sf['name']
        new_row = copy.deepcopy(hdr_el)
        fill = 'FFFFFF' if idx % 2 == 0 else 'FAF5F5'
        for tc in new_row.findall(qn('w:tc')):
            shd = tc.find(f'.//{qn("w:shd")}')
            if shd is not None:
                shd.set(qn('w:fill'), fill)
                shd.set(qn('w:color'), 'auto')
        tbl._element.append(new_row)
        vals = [nm, sf.get('paint',''), sf.get('sheen',''), sf.get('color','TBD'), f"{sf.get('pc',2)} / {sf.get('prc',0)}"]
        for ci, val in enumerate(vals):
            set_row_cell_text(new_row, ci, val, bold=(ci == 0))

def generate_proposal(E):
    doc = Document(TEMPLATE_PATH)
    sides_data = E.get('sides', [])
    active_labels = [s['label'] for s in sides_data]
    all_sides = ['Front', 'Left', 'Right', 'Back']

    # 1. Proposal # / License / Date
    for para in doc.paragraphs:
        if 'Proposal #:' in para.text and 'License:' in para.text:
            for run in para.runs:
                run.text = run.text.replace('0135', E.get('proposalNum', '____'))
                run.text = run.text.replace('05/14/2026', E.get('dateIssued', ''))
            break

    # 2. Client info
    ct = doc.tables[0]
    addr = f"{E['client']['street']}, {E['client']['city']}, {E['client']['state']} {E['client']['zip']}"
    set_cell_text(ct.rows[0].cells[1], E.get('subject',''), italic=True)
    set_cell_text(ct.rows[1].cells[1], E['client']['name'], italic=True)
    set_cell_text(ct.rows[2].cells[1], addr, italic=True)
    set_cell_text(ct.rows[3].cells[1], E['client']['phone'], italic=True)
    set_cell_text(ct.rows[4].cells[1], E['client']['email'], italic=True)

    # 3. Power wash bullets
    pw_cell = doc.tables[1].rows[0].cells[0]
    bps = [p for p in pw_cell.paragraphs if p.text.strip() and 'Power Washing' not in p.text]
    for i, item in enumerate(E.get('powerWash', [])):
        if i < len(bps) and bps[i].runs:
            for r in bps[i].runs: r.text = ''
            bps[i].runs[0].text = item

    # 4. Surface prep bullets
    sp_cell = doc.tables[2].rows[0].cells[0]
    sps = [p for p in sp_cell.paragraphs if p.text.strip() and 'Surface Preparation' not in p.text]
    for i, item in enumerate(E.get('surfacePrep', [])):
        if i < len(sps) and sps[i].runs:
            for r in sps[i].runs: r.text = ''
            sps[i].runs[0].text = item

    # 5. Build map of side label -> paint table BEFORE removing anything
    side_table_map = {}
    body_children = list(doc.element.body)
    for i, child in enumerate(body_children):
        if child.tag.split('}')[-1] == 'tbl':
            from docx.table import Table
            t = Table(child, doc)
            if t.rows[0].cells[0].text.strip() == 'Surface':
                # Look back for the side heading
                for prev in reversed(body_children[:i]):
                    if prev.tag.split('}')[-1] == 'p':
                        txt = get_para_text(prev, doc)
                        for sl in all_sides:
                            if sl + ' of House' in txt:
                                side_table_map[sl] = t
                                break
                        if txt: break

    # 6. Update paint tables for active sides
    for i, side in enumerate(sides_data):
        label = side['label']
        if label in side_table_map:
            rebuild_paint_table(side_table_map[label], side.get('surfaces', []))

    # 7. Remove unused side blocks
    unused = [s for s in all_sides if s not in active_labels]
    for sl in unused:
        remove_side_block(doc, sl)

    # 8. Renumber remaining side headings
    for i, side in enumerate(sides_data):
        label = side['label']
        num = i + 3
        for para in doc.paragraphs:
            if label + ' of House' in para.text and para.runs:
                para.runs[0].text = f"{num}.  {label} of House"
                for r in para.runs[1:]: r.text = ''
                break

    # 9. Carpentry
    carp_enabled = E.get('carpentry', {}).get('enabled', False)
    if not carp_enabled:
        body = doc.element.body
        to_remove = []
        for child in list(body):
            tag = child.tag.split('}')[-1]
            if tag == 'p' and 'Inspection & Carpentry' in get_para_text(child, doc):
                to_remove.append(child)
            elif tag == 'tbl' and 'IMPORTANT' in get_table_first_cell(child, doc):
                to_remove.append(child)
        for el in to_remove:
            try: body.remove(el)
            except: pass
    else:
        num = len(active_labels) + 3
        for para in doc.paragraphs:
            if 'Inspection & Carpentry' in para.text and para.runs:
                para.runs[0].text = f"{num}.  Inspection & Carpentry (Work Change Order)"
                for r in para.runs[1:]: r.text = ''
                break

    # 10. Duration
    for para in doc.paragraphs:
        if '(X–X)' in para.text:
            for run in para.runs:
                run.text = run.text.replace('(X–X)', E.get('duration', 'X–X'))

    # 11. Porta Potty
    if not E.get('portaPotty', False):
        for tbl in doc.tables:
            for row in list(tbl.rows):
                if 'Porta Potty' in row.cells[0].text:
                    row._element.getparent().remove(row._element)
                    break

    buf = io.BytesIO()
    doc.save(buf)
    buf.seek(0)
    return buf

@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if request.method == 'OPTIONS':
        return '', 200
    try:
        E = request.get_json()
        if not E:
            return jsonify({'error': 'No data provided'}), 400
        buf = generate_proposal(E)
        client_name = E.get('client', {}).get('name', 'Client').replace(' ', '_')
        date = E.get('dateIssued', '').replace('/', '-')
        filename = f"LUCAZProposal_{client_name}_{date}.docx"
        return send_file(buf, mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document', as_attachment=True, download_name=filename)
    except Exception as e:
        import traceback
        return jsonify({'error': str(e), 'trace': traceback.format_exc()}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
