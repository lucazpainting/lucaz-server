from flask import Flask, request as flask_request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import copy, io, os, json, time
import requests as http_requests
from google.oauth2 import service_account

app = Flask(__name__)
CORS(app)

# ── GOOGLE DRIVE SETUP ──
SERVICE_ACCOUNT_FILE = os.path.join(os.path.dirname(__file__), 'service_account.json')
SCOPES = ['https://www.googleapis.com/auth/drive']
PROPOSALS_FOLDER_NAME = 'Proposals'
PROPOSALS_FOLDER_ID = '1MQA2U0eaEBM0w_T75-pHaYyO55T2d5nY'

def get_drive_token():
    """Get a fresh access token for the service account"""
    import google.auth.transport.requests
    from credentials import SERVICE_ACCOUNT_INFO
    creds = service_account.Credentials.from_service_account_info(SERVICE_ACCOUNT_INFO, scopes=SCOPES)
    creds.refresh(google.auth.transport.requests.Request())
    return creds.token

def drive_request(method, url, token, **kwargs):
    """Make an authenticated Drive API request"""
    headers = kwargs.pop('headers', {})
    headers['Authorization'] = f'Bearer {token}'
    return http_requests.request(method, url, headers=headers, **kwargs)

def get_or_create_folder(token, name, parent_id=None):
    """Find folder by name, create if not exists"""
    q = f"name='{name}' and mimeType='application/vnd.google-apps.folder' and trashed=false"
    if parent_id:
        q += f" and '{parent_id}' in parents"
    res = drive_request('GET', 'https://www.googleapis.com/drive/v3/files',
        token, params={'q': q, 'fields': 'files(id,name)'})
    data = res.json()
    files = data.get('files', [])
    if files:
        return files[0]['id']
    # Create folder
    meta = {'name': name, 'mimeType': 'application/vnd.google-apps.folder'}
    if parent_id:
        meta['parents'] = [parent_id]
    res = drive_request('POST', 'https://www.googleapis.com/drive/v3/files',
        token, json=meta, params={'fields': 'id'})
    return res.json().get('id')

def save_to_drive(doc_bytes, file_name, client_name, status='Active', existing_file_id=None):
    """Save or update proposal in Drive"""
    try:
        token = get_drive_token()
        # Use hardcoded Proposals folder ID directly
        proposals_id = PROPOSALS_FOLDER_ID
        status_id = get_or_create_folder(token, status, proposals_id)
        client_id = get_or_create_folder(token, client_name, status_id)

        mimetype = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'

        if existing_file_id:
            # Update in place
            res = drive_request('PATCH',
                f'https://www.googleapis.com/upload/drive/v3/files/{existing_file_id}',
                token,
                params={'uploadType': 'media', 'fields': 'id'},
                headers={'Content-Type': mimetype},
                data=doc_bytes
            )
            return res.json().get('id', existing_file_id)
        else:
            # Multipart upload
            boundary = f'lucaz_{int(time.time())}'
            meta = json.dumps({'name': file_name, 'parents': [client_id]})
            body = (
                f'--{boundary}\r\n'
                f'Content-Type: application/json; charset=UTF-8\r\n\r\n'
                f'{meta}\r\n'
                f'--{boundary}\r\n'
                f'Content-Type: {mimetype}\r\n\r\n'
            ).encode() + doc_bytes + f'\r\n--{boundary}--'.encode()

            res = drive_request('POST',
                'https://www.googleapis.com/upload/drive/v3/files',
                token,
                params={'uploadType': 'multipart', 'fields': 'id'},
                headers={'Content-Type': f'multipart/related; boundary={boundary}'},
                data=body
            )
            return res.json().get('id')
    except Exception as e:
        print(f'Drive save error: {e}')
        import traceback; traceback.print_exc()
        return None

def move_drive_file(file_id, old_status, new_status):
    """Move file between status folders"""
    try:
        token = get_drive_token()
        proposals_id = PROPOSALS_FOLDER_ID
        old_folder_id = get_or_create_folder(token, old_status, proposals_id)
        new_folder_id = get_or_create_folder(token, new_status, proposals_id)
        res = drive_request('PATCH',
            f'https://www.googleapis.com/drive/v3/files/{file_id}',
            token,
            params={'addParents': new_folder_id, 'removeParents': old_folder_id, 'fields': 'id'}
        )
        return 'id' in res.json()
    except Exception as e:
        print(f'Drive move error: {e}')
        return False

TEMPLATE_PATH = os.path.join(os.path.dirname(__file__), 'EXTERIOR_MASTER_TEMPLATE.docx')

# Fixed table indices in master template — never changes
SIDE_TABLE_IDX = {'Front': 3, 'Left': 4, 'Right': 5, 'Back': 6}

def set_cell_text(cell, new_text, bold=None, italic=None):
    for para in cell.paragraphs:
        if not para.runs: continue
        first = para.runs[0]
        for run in para.runs[1:]: run.text = ""
        first.text = new_text
        if bold is not None: first.bold = bold
        if italic is not None: first.italic = italic
        return

def set_run_text(row_el, col_idx, text, bold=None):
    """Set text in a cell, preserving formatting but fixing color to black (auto)"""
    cells = row_el.findall(qn('w:tc'))
    if col_idx >= len(cells): return
    for p in cells[col_idx].findall(qn('w:p')):
        runs = p.findall(qn('w:r'))
        if runs:
            r = runs[0]
            # Set text
            t = r.find(qn('w:t'))
            if t is None:
                t = OxmlElement('w:t'); r.append(t)
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            # Fix run properties
            rpr = r.find(qn('w:rPr'))
            if rpr is None:
                rpr = OxmlElement('w:rPr'); r.insert(0, rpr)
            # Remove white color — set to auto/black
            color_el = rpr.find(qn('w:color'))
            if color_el is not None:
                rpr.remove(color_el)
            # Set bold
            if bold is not None:
                b = rpr.find(qn('w:b'))
                if bold and b is None:
                    b = OxmlElement('w:b'); rpr.insert(0, b)
                elif not bold and b is not None:
                    rpr.remove(b)
            # Remove italic
            i_el = rpr.find(qn('w:i'))
            if i_el is not None: rpr.remove(i_el)
            # Clear extra runs
            for extra_r in runs[1:]:
                t2 = extra_r.find(qn('w:t'))
                if t2 is not None: t2.text = ''
            break

def get_para_text(el, doc):
    try:
        from docx.text.paragraph import Paragraph
        return Paragraph(el, doc).text.strip()
    except: return ''

def rebuild_paint_table(tbl, surfaces):
    """Completely rebuild table with only the surfaces from the estimate"""
    # Clone a DATA row (row 1) as template — not header row
    # This ensures text color is black not white
    if len(tbl.rows) < 2:
        data_row_template = copy.deepcopy(tbl.rows[0]._element)
    else:
        data_row_template = copy.deepcopy(tbl.rows[1]._element)
    
    hdr_el = copy.deepcopy(tbl.rows[0]._element)

    # Remove ALL rows
    for row in list(tbl.rows):
        tbl._element.remove(row._element)

    # Re-add header row
    tbl._element.append(copy.deepcopy(hdr_el))

    # Add one data row per surface using data row template
    for idx, sf in enumerate(surfaces):
        qty = sf.get('qty')
        try: qty_int = int(str(qty)) if qty else 0
        except: qty_int = 0
        nm = f"{sf['name']} × {qty_int}" if qty_int > 1 else sf['name']

        new_row = copy.deepcopy(data_row_template)

        # Set alternating background
        fill = 'FFFFFF' if idx % 2 == 0 else 'FAF5F5'
        for tc in new_row.findall(qn('w:tc')):
            shd = tc.find(f'.//{qn("w:shd")}')
            if shd is not None:
                shd.set(qn('w:fill'), fill)
                shd.set(qn('w:color'), 'auto')

        tbl._element.append(new_row)

        # Set values — text color will be black since cloned from data row
        vals = [nm, sf.get('paint',''), sf.get('sheen',''), sf.get('color','TBD'), f"{sf.get('pc',2)} / {sf.get('prc',0)}"]
        for ci, val in enumerate(vals):
            set_run_text(new_row, ci, val, bold=(ci == 0))

def remove_side_block(doc, side_label):
    """Remove heading + paint brand line + paint table + spacer for a given side"""
    search = side_label + ' of House'
    body = doc.element.body
    children = list(body)
    for i, child in enumerate(children):
        if child.tag.split('}')[-1] == 'p':
            txt = get_para_text(child, doc)
            if search in txt:
                # Collect: this heading + up to 3 more elements
                to_remove = [child]
                j = i + 1
                while j < len(children) and len(to_remove) <= 3:
                    nc = children[j]
                    nc_tag = nc.tag.split('}')[-1]
                    # Stop if we hit another side or major section
                    if nc_tag == 'p':
                        nc_txt = get_para_text(nc, doc)
                        if nc_txt and any(x in nc_txt for x in [
                            'of House','PROJECT PHOTOS','WARRANTY','PAYMENT',
                            'COST','NEXT','ESTIMATE','Inspection','PROJECT INFORMATION'
                        ]):
                            break
                    to_remove.append(nc)
                    j += 1
                for el in to_remove:
                    try: body.remove(el)
                    except: pass
                return

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
    pw_items = E.get('powerWash', [])
    for i, item in enumerate(pw_items):
        if i < len(bps) and bps[i].runs:
            for r in bps[i].runs: r.text = ''
            bps[i].runs[0].text = item
    for i in range(len(pw_items), len(bps)):
        if bps[i].runs:
            for r in bps[i].runs: r.text = ''

    # 4. Surface prep bullets
    sp_cell = doc.tables[2].rows[0].cells[0]
    sps = [p for p in sp_cell.paragraphs if p.text.strip() and 'Surface Preparation' not in p.text]
    sp_items = E.get('surfacePrep', [])
    for i, item in enumerate(sp_items):
        if i < len(sps) and sps[i].runs:
            for r in sps[i].runs: r.text = ''
            sps[i].runs[0].text = item
    for i in range(len(sp_items), len(sps)):
        if sps[i].runs:
            for r in sps[i].runs: r.text = ''

    # 5. Rebuild paint tables using FIXED indices BEFORE removing anything
    for side in sides_data:
        label = side['label']
        tbl_idx = SIDE_TABLE_IDX.get(label)
        if tbl_idx is not None and tbl_idx < len(doc.tables):
            rebuild_paint_table(doc.tables[tbl_idx], side.get('surfaces', []))

    # 6. Remove ALL unused side blocks
    unused = [s for s in all_sides if s not in active_labels]
    for sl in unused:
        remove_side_block(doc, sl)

    # 7. Renumber remaining side headings
    for i, side in enumerate(sides_data):
        num = i + 3
        for para in doc.paragraphs:
            if side['label'] + ' of House' in para.text and para.runs:
                para.runs[0].text = f"{num}.  {side['label']} of House"
                for r in para.runs[1:]: r.text = ''
                break

    # 8. Carpentry
    carp_enabled = E.get('carpentry', {}).get('enabled', False)
    if not carp_enabled:
        body = doc.element.body
        to_remove = []
        for child in list(body):
            tag = child.tag.split('}')[-1]
            if tag == 'p' and 'Inspection & Carpentry' in get_para_text(child, doc):
                to_remove.append(child)
            elif tag == 'tbl':
                try:
                    from docx.table import Table
                    if 'IMPORTANT' in Table(child, doc).rows[0].cells[0].text:
                        to_remove.append(child)
                except: pass
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

    # 9. Duration — handle split runs "(X" + "–" + "X)" across multiple runs
    duration = E.get('duration', '')
    if duration:
        for para in doc.paragraphs:
            if 'anticipated to take approximately' in para.text and '(X' in para.text:
                # Rebuild the full paragraph text replacing placeholder
                full_text = para.text
                if '(X' in full_text and 'X)' in full_text:
                    new_text = full_text.replace('(X–X)', duration).replace('(X-X)', duration)
                    # Set first run to full text, clear the rest
                    if para.runs:
                        para.runs[0].text = new_text
                        for run in para.runs[1:]:
                            run.text = ''
                break

    # 9b. Cost table — update with real numbers from dashboard
    cost_map = {
        'Subtotal': E.get('subtotal',''),
        'Sales Tax': E.get('salesTaxAmt',''),
        'Total Cost for Project': E.get('total',''),
        'Deposit Due': E.get('deposit',''),
        'Balance Due': E.get('balance',''),
        'Porta Potty': E.get('portaPottyAmt','$200.00'),
    }
    for tbl in doc.tables:
        for row in tbl.rows:
            if len(row.cells) < 2: continue
            cell0 = row.cells[0].text.strip()
            for key, val in cost_map.items():
                if val and key in cell0:
                    set_cell_text(row.cells[1], val)
                    break

    # 9c. Photos — embed images, remove empty cells individually
    photos = E.get('photos', {})
    import base64
    from docx.shared import Inches

    # Build label to key map — include custom labels from dashboard
    label_to_key = {
        'Back': 'Back', 'Left': 'Left', 'Garage': 'add1',
        'Front': 'Front', 'Right': 'Right',
        'Additional 1': 'add1', 'Additional 2': 'add2', 'Additional 3': 'add3'
    }
    # Add custom labels — if user typed "Deck" for add1, map 'Deck' -> 'add1'
    photo_labels = E.get('photoLabels', {})
    for key, custom_label in photo_labels.items():
        if custom_label:
            label_to_key[custom_label] = key
    # Also update the template cell labels to show custom text
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if '[ Insert Photo Here ]' not in cell.text: continue
                for para in cell.paragraphs:
                    txt = para.text.strip()
                    # Check if this label matches a custom key
                    for slot_key, custom_label in photo_labels.items():
                        default_labels = {'add1':'Additional 1','add2':'Additional 2','add3':'Additional 3'}
                        if txt == default_labels.get(slot_key) and custom_label:
                            if para.runs:
                                para.runs[0].text = custom_label
                                for r in para.runs[1:]: r.text = ''

    # Find photo tables and process them
    photo_tables_to_remove = []
    for tbl in doc.tables:
        is_photo_table = any(
            '[ Insert Photo Here ]' in cell.text
            for row in tbl.rows
            for cell in row.cells
        )
        if not is_photo_table:
            continue

        has_any_photo_in_table = False
        for row in tbl.rows:
            for cell in row.cells:
                if '[ Insert Photo Here ]' not in cell.text:
                    continue
                label = ''
                for para in cell.paragraphs:
                    if '[ Insert Photo Here ]' not in para.text and para.text.strip():
                        label = para.text.strip()
                        break
                photo_key = label_to_key.get(label) or label
                photo_data = photos.get(photo_key) if photo_key else None
                if photo_data and photo_data.startswith('data:image'):
                    has_any_photo_in_table = True
                    try:
                        header, b64 = photo_data.split(',', 1)
                        img_bytes = base64.b64decode(b64)
                        img_buf = io.BytesIO(img_bytes)
                        if cell.paragraphs:
                            first_para = cell.paragraphs[0]
                            for run in first_para.runs:
                                run.text = ''
                            run = first_para.add_run()
                            run.add_picture(img_buf, width=Inches(1.9))
                    except Exception as photo_err:
                        print(f'Photo embed error: {photo_err}')

        if not has_any_photo_in_table:
            photo_tables_to_remove.append(tbl)

    # Remove entire photo tables that have no photos
    for tbl in photo_tables_to_remove:
        try: tbl._element.getparent().remove(tbl._element)
        except: pass

    # 11. Fix signature section — line first, label underneath
    sig_tbl = None
    for tbl in doc.tables:
        if len(tbl.rows) > 0 and len(tbl.rows[0].cells) >= 3:
            if 'Client' in tbl.rows[0].cells[0].text or 'Signature' in tbl.rows[0].cells[0].text:
                sig_tbl = tbl
                break

    if sig_tbl:
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        def add_line_para(cell, space_before=120):
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            pBdr = OxmlElement('w:pBdr')
            bottom = OxmlElement('w:bottom')
            bottom.set(qn('w:val'), 'single')
            bottom.set(qn('w:sz'), '4')
            bottom.set(qn('w:space'), '1')
            bottom.set(qn('w:color'), '333333')
            pBdr.append(bottom)
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), str(space_before))
            spacing.set(qn('w:after'), '40')
            pPr.append(pBdr)
            pPr.append(spacing)
            p.append(pPr)
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = ' '
            r.append(t)
            p.append(r)
            cell._element.append(p)

        def add_label_para(cell, text, bold=False):
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            spacing = OxmlElement('w:spacing')
            spacing.set(qn('w:before'), '20')
            spacing.set(qn('w:after'), '80')
            pPr.append(spacing)
            p.append(pPr)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            clr = OxmlElement('w:color')
            clr.set(qn('w:val'), '555555')
            rPr.append(clr)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '18')
            rPr.append(sz)
            fn = OxmlElement('w:rFonts')
            fn.set(qn('w:ascii'), 'Calibri')
            rPr.append(fn)
            if bold:
                b = OxmlElement('w:b')
                rPr.append(b)
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            p.append(r)
            cell._element.append(p)

        sig_configs = [
            (0, 'Client Name', 'Client Signature'),
            (2, 'Contractor', 'Authorized Signature')
        ]
        for col_idx, name_label, sig_label in sig_configs:
            if col_idx >= len(sig_tbl.rows[0].cells):
                continue
            cell = sig_tbl.rows[0].cells[col_idx]
            # Clear text from existing paragraphs, keep structure
            for para in cell.paragraphs:
                for run in para.runs:
                    run.text = ''
            # Just update the label text in existing paragraphs
            paras = [p for p in cell.paragraphs]
            labels = [name_label, sig_label, 'Date']
            label_idx = 0
            for para in paras:
                txt = para.text.strip()
                # If paragraph has a border (it's a line), skip
                pb = para._element.find(qn('w:pBdr'))
                if pb is not None:
                    continue
                # It's a label paragraph — set the text
                if label_idx < len(labels):
                    if not para.runs:
                        run = para.add_run()
                    para.runs[0].text = labels[label_idx]
                    label_idx += 1
    # 10. Porta Potty
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

@app.route('/test-generate', methods=['GET'])
def test_generate():
    try:
        E = {
            'proposalNum': '0001',
            'dateIssued': '05/26/2026',
            'subject': 'Test',
            'client': {'name':'Test Client','street':'123 Main St','city':'Ossining','state':'NY','zip':'10562','phone':'(914) 555-0000','email':'test@test.com'},
            'powerWash': ['Full house exterior power wash — all sides'],
            'surfacePrep': ['Scraping and sanding of all peeling or flaking paint'],
            'sides': [{'label':'Front','surfaces':[{'name':'Siding — Clapboard','qty':None,'paint':'Regal Select 100% Acrylic','sheen':'Flat','color':'Color match','pc':2,'prc':0}]}],
            'carpentry': {'enabled': False},
            'portaPotty': False,
            'duration': '5-7 days',
            'photos': {},
            'subtotal': '$1,000.00',
            'salesTaxAmt': '$83.75',
            'total': '$1,083.75',
            'deposit': '$361.25',
            'balance': '$722.50'
        }
        buf = generate_proposal(E)
        return jsonify({'success': True, 'size': len(buf.read())})
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print('TEST-GENERATE ERROR:\n' + tb, flush=True)
        return jsonify({'success': False, 'error': str(e), 'trace': tb})

@app.route('/generate', methods=['POST', 'OPTIONS'])
def generate():
    if flask_request.method == 'OPTIONS':
        return '', 200
    try:
        E = flask_request.get_json()
        if not E:
            return jsonify({'error': 'No data provided'}), 400
        buf = generate_proposal(E)
        doc_bytes = buf.read()
        client_name_raw = E.get('client', {}).get('name', 'Client')
        client_name_safe = client_name_raw.replace(' ', '_')
        date = E.get('dateIssued', '').replace('/', '-')
        filename = f"LUCAZProposal_{client_name_safe}_{date}.docx"
        status = E.get('jobStatus', 'Active')
        existing_file_id = E.get('driveFileId', None)

        # Save to Drive server-side
        drive_file_id = save_to_drive(doc_bytes, filename, client_name_raw, status, existing_file_id)

        # Send file back to client for download
        response = send_file(
            io.BytesIO(doc_bytes),
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )
        # Return Drive file ID in header so dashboard can store it
        if drive_file_id:
            response.headers['X-Drive-File-Id'] = drive_file_id
        return response
    except Exception as e:
        import traceback
        tb = traceback.format_exc()
        print('GENERATE ERROR:\n' + tb, flush=True)
        return jsonify({'error': str(e), 'trace': tb}), 500

@app.route('/move', methods=['POST', 'OPTIONS'])
def move():
    """Move a file between status folders in Drive"""
    if flask_request.method == 'OPTIONS':
        return '', 200
    try:
        data = flask_request.get_json()
        file_id = data.get('fileId')
        old_status = data.get('oldStatus')
        new_status = data.get('newStatus')
        if not all([file_id, old_status, new_status]):
            return jsonify({'error': 'Missing fields'}), 400
        success = move_drive_file(file_id, old_status, new_status)
        return jsonify({'success': success})
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/test-drive', methods=['GET'])
def test_drive():
    """Test Drive connection and return status"""
    try:
        token = get_drive_token()
        if not token:
            return jsonify({'success': False, 'error': 'Could not get access token'})
        # Use hardcoded Proposals folder ID
        proposals_id = PROPOSALS_FOLDER_ID
        active_id = get_or_create_folder(token, 'Active', proposals_id)
        return jsonify({
            'success': True,
            'message': 'Drive connected successfully',
            'proposals_folder_id': proposals_id,
            'active_folder_id': active_id,
            'service_account': 'lucaz-drive@lucaz-proposals.iam.gserviceaccount.com'
        })
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})

@app.route('/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok', 'template': os.path.exists(TEMPLATE_PATH)})

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
