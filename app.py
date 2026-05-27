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
STATUS_FOLDER_IDS = {
    'Active': '1qcWcpTDiY6gQDJlDhr76cNh9R5qD38dG',
    'Completed': '1gqxIiZN7i8ts-D0b0B98INm6sxa8vWTp',
    'Rejected': '14JZv7q4lRk2I2A-5tk3FEwJI5-beCXjx'
}

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
        print(f'DRIVE: Getting token...', flush=True)
        token = get_drive_token()
        print(f'DRIVE: Token ok, finding/creating client folder for {client_name} in {status}...', flush=True)
        status_id = STATUS_FOLDER_IDS.get(status, STATUS_FOLDER_IDS['Active'])
        client_id = get_or_create_folder(token, client_name, status_id)
        print(f'DRIVE: Client folder id={client_id}', flush=True)

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
        old_folder_id = STATUS_FOLDER_IDS.get(old_status)
        new_folder_id = STATUS_FOLDER_IDS.get(new_status)
        if not old_folder_id or not new_folder_id:
            return False
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

    # Fix footer — replace plain "Page" text with actual page number field
    for sect in doc.sections:
        ftr = sect.footer
        for para in ftr.paragraphs:
            for run in para.runs:
                if run.text == 'Page':
                    # Replace with: "Page " + PAGE field + " of " + NUMPAGES field
                    run.text = 'Page '
                    # Insert PAGE field after this run
                    from docx.oxml import OxmlElement
                    from docx.oxml.ns import qn
                    def make_page_field(instr):
                        r = OxmlElement('w:r')
                        # Copy run properties
                        rpr = run._element.find(qn('w:rPr'))
                        if rpr is not None:
                            import copy
                            r.append(copy.deepcopy(rpr))
                        fld_begin = OxmlElement('w:fldChar')
                        fld_begin.set(qn('w:fldCharType'), 'begin')
                        r.append(fld_begin)
                        r2 = OxmlElement('w:r')
                        if rpr is not None:
                            import copy
                            r2.append(copy.deepcopy(rpr))
                        instr_el = OxmlElement('w:instrText')
                        instr_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
                        instr_el.text = f' {instr} '
                        r2.append(instr_el)
                        r3 = OxmlElement('w:r')
                        if rpr is not None:
                            import copy
                            r3.append(copy.deepcopy(rpr))
                        fld_end = OxmlElement('w:fldChar')
                        fld_end.set(qn('w:fldCharType'), 'end')
                        r3.append(fld_end)
                        return r, r2, r3
                    # Insert PAGE field after the "Page " run
                    parent = run._element.getparent()
                    idx = list(parent).index(run._element)
                    for el in reversed(make_page_field('PAGE')):
                        parent.insert(idx + 1, el)
                    break

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

    photo_labels = E.get('photoLabels', {})

    # Build label to key map
    # Template has: Back, Left, Garage (=add1), Front, Right
    # Dashboard additional slots: add1, add2, add3
    label_to_key = {
        'Back': 'Back', 'Left': 'Left',
        'Garage': 'add1',  # template default for add1 slot
        'Front': 'Front', 'Right': 'Right',
        'Additional 1': 'add1', 'Additional 2': 'add2', 'Additional 3': 'add3'
    }
    # Add reverse mapping from custom labels typed by user
    for key, custom_label in photo_labels.items():
        if custom_label:
            label_to_key[custom_label] = key

    # Update template cell labels to show custom text AND fix Garage label
    for tbl in doc.tables:
        for row in tbl.rows:
            for cell in row.cells:
                if '[ Insert Photo Here ]' not in cell.text:
                    continue
                for para in cell.paragraphs:
                    txt = para.text.strip()
                    if not txt or '[ Insert Photo Here ]' in txt:
                        continue
                    # Find which slot key this cell belongs to
                    slot_key = label_to_key.get(txt)
                    if slot_key and slot_key in photo_labels and photo_labels[slot_key]:
                        # User typed a custom label for this slot — update cell text
                        if para.runs:
                            para.runs[0].text = photo_labels[slot_key]
                            for r in para.runs[1:]: r.text = ''

    # Build dynamic photo grid — only include photos that exist
    import base64
    from docx.shared import Inches, Pt
    from docx.oxml import OxmlElement
    from docx.oxml.ns import qn as _qn
    from docx.enum.text import WD_ALIGN_PARAGRAPH
    from docx.enum.table import WD_TABLE_ALIGNMENT
    import copy

    photo_labels = E.get('photoLabels', {})

    # Build ordered list of (key, label, photo_data) for photos that exist
    label_to_key = {
        'Back':'Back','Left':'Left','Garage':'add1',
        'Front':'Front','Right':'Right',
        'Additional 1':'add1','Additional 2':'add2','Additional 3':'add3'
    }
    for key, custom_label in photo_labels.items():
        if custom_label:
            label_to_key[custom_label] = key

    # Collect photos that actually have data
    active_photos = []
    # Check side photos first (in order of sides)
    side_labels = [s['label'] for s in sides_data]
    for side_label in side_labels:
        key = side_label
        photo_data = photos.get(key)
        display_label = photo_labels.get(key) or f'{side_label} side'
        if photo_data and photo_data.startswith('data:image'):
            active_photos.append((key, display_label, photo_data))

    # Then additional photos
    for slot_key in ['add1','add2','add3']:
        photo_data = photos.get(slot_key)
        display_label = photo_labels.get(slot_key) or f'Additional {slot_key[-1]}'
        if photo_data and photo_data.startswith('data:image'):
            active_photos.append((slot_key, display_label, photo_data))

    # Remove existing photo tables from document
    body = doc.element.body
    photo_tbls_to_remove = []
    for tbl in doc.tables:
        if any(
            any('[ Insert Photo Here ]' in para.text for para in cell.paragraphs)
            for row in tbl.rows for cell in row.cells
        ):
            photo_tbls_to_remove.append(tbl)

    # Remember position of first photo table to insert new ones there
    insert_after_el = None
    for tbl in photo_tbls_to_remove:
        if insert_after_el is None:
            insert_after_el = tbl._element
        body.remove(tbl._element)

    # If no photos at all, just remove the "PROJECT PHOTOS" section header too
    if not active_photos:
        # Remove the PROJECT PHOTOS heading and intro paragraph
        children = list(body)
        for i, child in enumerate(children):
            if child.tag.split('}')[-1] == 'p':
                from docx.text.paragraph import Paragraph
                txt = Paragraph(child, doc).text.strip()
                if 'PROJECT PHOTOS' in txt:
                    # Remove heading and next 2 paragraphs
                    to_rm = [child]
                    j = i + 1
                    while j < len(children) and len(to_rm) < 3:
                        nc = children[j]
                        if nc.tag.split('}')[-1] == 'p':
                            nc_txt = Paragraph(nc, doc).text.strip()
                            if nc_txt and any(x in nc_txt for x in ['PROJECT INFORMATION','WARRANTY','PAYMENT','COST']):
                                break
                            to_rm.append(nc)
                        j += 1
                    for el in to_rm:
                        try: body.remove(el)
                        except: pass
                    break
    else:
        # Build new photo tables with only active photos
        # Arrange in rows of up to 3
        W = 9360  # content width in DXA

        def make_photo_cell(width, photo_data, label, doc_ref):
            """Build a table cell with an embedded photo and label"""
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn as _qn
            from docx.shared import Inches
            from docx.text.paragraph import Paragraph

            tc = OxmlElement('w:tc')
            tcPr = OxmlElement('w:tcPr')
            tcW = OxmlElement('w:tcW')
            tcW.set(_qn('w:w'), str(width))
            tcW.set(_qn('w:type'), 'dxa')
            tcPr.append(tcW)
            tcBorders = OxmlElement('w:tcBorders')
            for side in ['top','left','bottom','right']:
                b = OxmlElement(f'w:{side}')
                b.set(_qn('w:val'), 'single')
                b.set(_qn('w:sz'), '4')
                b.set(_qn('w:color'), 'CCCCCC')
                tcBorders.append(b)
            tcPr.append(tcBorders)
            tc.append(tcPr)

            # Add photo using a temp paragraph on the actual document
            # This ensures the image relationship is stored in the right part
            tmp_para = doc_ref.add_paragraph()
            tmp_para.alignment = 1  # CENTER
            if photo_data and photo_data.startswith('data:image'):
                try:
                    header, b64 = photo_data.split(',', 1)
                    img_bytes = base64.b64decode(b64)
                    img_buf = io.BytesIO(img_bytes)
                    run = tmp_para.add_run()
                    run.add_picture(img_buf, width=Inches(min(2.3, width/1440.0)))
                except Exception as e:
                    print(f'Photo embed error: {e}')

            # Set paragraph spacing
            pPr = tmp_para._element.find(_qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                tmp_para._element.insert(0, pPr)
            sp = OxmlElement('w:spacing')
            sp.set(_qn('w:before'), '60')
            sp.set(_qn('w:after'), '40')
            pPr.append(sp)
            jc = OxmlElement('w:jc')
            jc.set(_qn('w:val'), 'center')
            pPr.append(jc)

            # Move paragraph element to cell
            p1_el = tmp_para._element
            doc_ref.paragraphs[-1]._element.getparent().remove(p1_el)
            tc.append(p1_el)

            # Label paragraph
            p2 = OxmlElement('w:p')
            pPr2 = OxmlElement('w:pPr')
            jc2 = OxmlElement('w:jc')
            jc2.set(_qn('w:val'), 'center')
            sp2 = OxmlElement('w:spacing')
            sp2.set(_qn('w:before'), '0')
            sp2.set(_qn('w:after'), '60')
            pPr2.append(jc2)
            pPr2.append(sp2)
            p2.append(pPr2)
            r2 = OxmlElement('w:r')
            rPr2 = OxmlElement('w:rPr')
            b_el2 = OxmlElement('w:b')
            fn2 = OxmlElement('w:rFonts')
            fn2.set(_qn('w:ascii'), 'Calibri')
            rPr2.append(b_el2)
            rPr2.append(fn2)
            r2.append(rPr2)
            t_el = OxmlElement('w:t')
            t_el.text = label
            t_el.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r2.append(t_el)
            p2.append(r2)
            tc.append(p2)
            return tc

        def make_photo_table(photo_row):
            n = len(photo_row)
            col_w = W // n
            tbl = OxmlElement('w:tbl')
            tblPr = OxmlElement('w:tblPr')
            tblW = OxmlElement('w:tblW')
            tblW.set(_qn('w:w'), str(W))
            tblW.set(_qn('w:type'), 'dxa')
            tblPr.append(tblW)
            tblLook = OxmlElement('w:tblLook')
            tblLook.set(_qn('w:val'), '0000')
            tblPr.append(tblLook)
            tbl.append(tblPr)
            tblGrid = OxmlElement('w:tblGrid')
            for _ in photo_row:
                gc = OxmlElement('w:gridCol')
                gc.set(_qn('w:w'), str(col_w))
                tblGrid.append(gc)
            tbl.append(tblGrid)
            tr = OxmlElement('w:tr')
            for key, label, photo_data in photo_row:
                tc = make_photo_cell(col_w, photo_data, label, doc)
                tr.append(tc)
            tbl.append(tr)
            return tbl

        # Split into rows of max 3
        rows = [active_photos[i:i+3] for i in range(0, len(active_photos), 3)]

        # Insert new tables where old photo tables were
        # insert_after_el was already removed from body, so insert at end of body or find by position
        parent = doc.element.body
        children_list = list(parent)
        # Find the PROJECT INFORMATION section to insert before it
        idx = len(children_list) - 1  # default to end
        for i, child in enumerate(children_list):
            if child.tag.split('}')[-1] == 'p':
                try:
                    from docx.text.paragraph import Paragraph
                    txt = Paragraph(child, doc).text.strip()
                    if 'PROJECT INFORMATION' in txt:
                        idx = i
                        break
                except: pass
        if True:

            # Add spacer paragraph
            sp_para = OxmlElement('w:p')
            sp_pPr = OxmlElement('w:pPr')
            sp_spacing = OxmlElement('w:spacing')
            sp_spacing.set(_qn('w:before'), '80')
            sp_spacing.set(_qn('w:after'), '0')
            sp_pPr.append(sp_spacing)
            sp_para.append(sp_pPr)

            for ri, row_photos in enumerate(rows):
                new_tbl = make_photo_table(row_photos)
                parent.insert(idx, new_tbl)
                idx += 1
                if ri < len(rows) - 1:
                    parent.insert(idx, copy.deepcopy(sp_para))
                    idx += 1
    # Fix spacing after "Photos of the areas..." paragraph
    for para in doc.paragraphs:
        if 'Photos of the areas' in para.text:
            from docx.oxml import OxmlElement
            from docx.oxml.ns import qn
            pPr = para._element.find(qn('w:pPr'))
            if pPr is None:
                pPr = OxmlElement('w:pPr')
                para._element.insert(0, pPr)
            sp = pPr.find(qn('w:spacing'))
            if sp is None:
                sp = OxmlElement('w:spacing')
                pPr.append(sp)
            sp.set(qn('w:after'), '40')
            break

    # 10. Porta Potty
    if not E.get('portaPotty', False):
        for tbl in doc.tables:
            for row in list(tbl.rows):
                if 'Porta Potty' in row.cells[0].text:
                    row._element.getparent().remove(row._element)
                    break

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

        def make_sig_line(space_before=160):
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            pBdr = OxmlElement('w:pBdr')
            bot = OxmlElement('w:bottom')
            bot.set(qn('w:val'), 'single')
            bot.set(qn('w:sz'), '6')
            bot.set(qn('w:space'), '1')
            bot.set(qn('w:color'), '000000')
            pBdr.append(bot)
            sp = OxmlElement('w:spacing')
            sp.set(qn('w:before'), str(space_before))
            sp.set(qn('w:after'), '40')
            pPr.append(pBdr)
            pPr.append(sp)
            p.append(pPr)
            r = OxmlElement('w:r')
            t = OxmlElement('w:t')
            t.text = ' '
            r.append(t)
            p.append(r)
            return p

        def make_sig_label(text, bold=False):
            p = OxmlElement('w:p')
            pPr = OxmlElement('w:pPr')
            sp = OxmlElement('w:spacing')
            sp.set(qn('w:before'), '40')
            sp.set(qn('w:after'), '20')
            pPr.append(sp)
            p.append(pPr)
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            fn = OxmlElement('w:rFonts')
            fn.set(qn('w:ascii'), 'Calibri')
            fn.set(qn('w:hAnsi'), 'Calibri')
            rPr.append(fn)
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), '18')
            rPr.append(sz)
            if bold:
                b = OxmlElement('w:b')
                rPr.append(b)
            r.append(rPr)
            t = OxmlElement('w:t')
            t.text = text
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            r.append(t)
            p.append(r)
            return p

        for col_idx, name_label, sig_label in [(0,'Client Name','Client Signature'),(2,'Contractor','Authorized Signature')]:
            if col_idx >= len(sig_tbl.rows[0].cells):
                continue
            cell = sig_tbl.rows[0].cells[col_idx]
            for p_el in list(cell._element.findall(qn('w:p'))):
                cell._element.remove(p_el)
            cell._element.append(make_sig_label('', bold=False))
            cell._element.append(make_sig_line(120))
            cell._element.append(make_sig_label(name_label, bold=True))
            cell._element.append(make_sig_line(180))
            cell._element.append(make_sig_label(sig_label))
            cell._element.append(make_sig_line(180))
            cell._element.append(make_sig_label('Date'))

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
        print(f'DRIVE SAVE: client={client_name_raw} status={status} file_id={drive_file_id}', flush=True)

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
        # Use hardcoded folder IDs
        proposals_id = PROPOSALS_FOLDER_ID
        active_id = STATUS_FOLDER_IDS.get('Active')
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
