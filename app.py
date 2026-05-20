from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os, json

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

def set_row_cell_text(row_el, col_idx, text, bold=None, italic=False):
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
                        b = OxmlElement('w:b')
                        rpr.insert(0, b)
                    elif not bold and b is not None:
                        rpr.remove(b)
                i_el = rpr.find(qn('w:i'))
                if i_el is not None:
                    rpr.remove(i_el)
            for extra_r in runs[1:]:
                t2 = extra_r.find(qn('w:t'))
                if t2 is not None: t2.text = ''
            break

def generate_proposal(E):
    doc = Document(TEMPLATE_PATH)

    # 1. Proposal # / License / Date
    for para in doc.paragraphs:
        if 'Proposal #:' in para.text and 'License:' in para.text:
            for run in para.runs:
                run.text = run.text.replace('0135', E['proposalNum'])
                run.text = run.text.replace('05/14/2026', E['dateIssued'])
            break

    # 2. Client info table
    ct = doc.tables[0]
    addr = f"{E['client']['street']}, {E['client']['city']}, {E['client']['state']} {E['client']['zip']}"
    set_cell_text(ct.rows[0].cells[1], E['subject'],              italic=True)
    set_cell_text(ct.rows[1].cells[1], E['client']['name'],       italic=True)
    set_cell_text(ct.rows[2].cells[1], addr,                      italic=True)
    set_cell_text(ct.rows[3].cells[1], E['client']['phone'],      italic=True)
    set_cell_text(ct.rows[4].cells[1], E['client']['email'],      italic=True)

    # 3. Power wash bullets (table 1)
    pw_cell = doc.tables[1].rows[0].cells[0]
    bullet_paras = [p for p in pw_cell.paragraphs if p.text.strip() and 'Power Washing' not in p.text]
    for i, item in enumerate(E.get('powerWash', [])):
        if i < len(bullet_paras) and bullet_paras[i].runs:
            for run in bullet_paras[i].runs: run.text = ''
            bullet_paras[i].runs[0].text = item

    # 4. Surface prep bullets (table 2)
    sp_cell = doc.tables[2].rows[0].cells[0]
    sp_paras = [p for p in sp_cell.paragraphs if p.text.strip() and 'Surface Preparation' not in p.text]
    for i, item in enumerate(E.get('surfacePrep', [])):
        if i < len(sp_paras) and sp_paras[i].runs:
            for run in sp_paras[i].runs: run.text = ''
            sp_paras[i].runs[0].text = item

    # 5. Paint spec tables (tables 3-6) — one per side
    side_tables = [doc.tables[3], doc.tables[4], doc.tables[5], doc.tables[6]]
    for si, side in enumerate(E.get('sides', [])):
        if si >= len(side_tables): break
        tbl = side_tables[si]
        # Remove all data rows (keep header row 0)
        data_rows = list(tbl.rows)[1:]
        for dr in data_rows:
            dr._element.getparent().remove(dr._element)
        # Add new rows cloned from header
        hdr_el = tbl.rows[0]._element
        for idx, sf in enumerate(side.get('surfaces', [])):
            nm = f"{sf['name']} × {sf['qty']}" if sf.get('qty') and sf['qty'] > 1 else sf['name']
            new_row = copy.deepcopy(hdr_el)
            fill = 'FFFFFF' if idx % 2 == 0 else 'FAF5F5'
            for tc in new_row.findall(qn('w:tc')):
                shd = tc.find(f'.//{qn("w:shd")}')
                if shd is not None:
                    shd.set(qn('w:fill'), fill)
                    shd.set(qn('w:color'), 'auto')
            tbl._element.append(new_row)
            vals = [nm, sf.get('paint',''), sf.get('sheen',''), sf.get('color',''), f"{sf.get('pc',2)} / {sf.get('prc',0)}"]
            for ci, val in enumerate(vals):
                set_row_cell_text(new_row, ci, val, bold=(ci == 0))

    # 6. Side subheadings in body text
    side_num_map = {0: 3, 1: 4, 2: 5, 3: 6}
    side_label_keys = ['Front of House', 'Left of House', 'Right of House', 'Back of House']
    for para in doc.paragraphs:
        for i, key in enumerate(side_label_keys):
            if key in para.text and para.runs:
                sides = E.get('sides', [])
                if i < len(sides):
                    new_label = f"{side_num_map[i]}.  {sides[i]['label']} of House"
                    para.runs[0].text = new_label
                    for r in para.runs[1:]: r.text = ''

    # 7. Duration
    for para in doc.paragraphs:
        if '(X–X)' in para.text:
            for run in para.runs:
                run.text = run.text.replace('(X–X)', E.get('duration', 'X–X'))

    # 8. Porta Potty toggle
    if not E.get('portaPotty', True):
        cost_tbl = doc.tables[10]
        for row in list(cost_tbl.rows):
            if 'Porta Potty' in row.cells[0].text:
                row._element.getparent().remove(row._element)
                break

    # 9. Carpentry toggle — hide the box if off
    if not E.get('carpentry', {}).get('enabled', True):
        carp_tbl = doc.tables[7]
        carp_tbl._element.getparent().remove(carp_tbl._element)

    # Save to bytes
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
        return send_file(
            buf,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
