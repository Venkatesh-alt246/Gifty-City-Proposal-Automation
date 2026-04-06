from flask import Flask, request, send_file, jsonify
from flask_cors import CORS
from docx import Document
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from datetime import datetime
import copy
import io
import os
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
import os

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── Register Roboto fonts in the Word document ──────────────────
def _embed_roboto_fonts(doc):
    """Embed Roboto font declarations into the document settings"""
    try:
        fonts_part = doc.part.fonts_part
    except:
        pass  # fonts embedding handled via rFonts in each run

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

app = Flask(__name__)
CORS(app)
application = app


# ── All scope item content ──────────────────────────────────────────────────

SECTION1_ITEMS = {
    's1a': {
        'label': 'Company/ LLP/ Entity formation with other activities',
        'bullets': [
            'Application for name reservation for a setup of proposed entity under Gift City-IFSC (with MCA)',
            'Identification of office space premise/co-working space in GIFT Special Economic Zone (SEZ) - IFSC',
            'Liaison to secure a Provisional Letter of Allotment (PLOA) from GIFT – SEZ Developers on finalization of the premise',
            'Application for incorporation of Company/Limited Liability Partnership entity with MCA, obtaining PAN (Permanent Account Number) and Tax Account Number (WHT)',
        ]
    },
    's1b': {
        'label': 'SEZ Unit application',
        'bullets': [
            'Preparation of documents for SEZ Unit in consultation with the client.',
            'Preparation of forms, declaration and support documents for IFSCA in consultation with the client.',
            'Preparation and submission of application online for setting up a unit in FORM FA along with requisite documents.',
            'Arranging for inclusion of Proposal in Unit Approval Committee agenda.',
            'Assisting in Interview to your representative for Approval Committee.',
            'Follow-up for Issuance of Letter of Approval from Development Commissioner, SEZ/ IFSCA (Administrator).',
            'Preparation and submission of Letter of acceptance of terms and condition within 45 days.',
            'Regularization of SEZ Online registration with unit and maker creation after payment of registration fees.',
            'Submission of Application to the concerned regulatory division office of IFSCA.',
            'Assisting in Interview to your representative for Approval Committee.',
            'Preparing draft of Query reply, if any, in consultation with the client.',
            'Assisting in understanding requirements for issuance of final letter of approval from IFSCA.',
            'Following up for issuance of Letter of Approval (in-principle) from IFSCA.',
            'Preparing submission along with the relevant documentation for application of letter of approval from IFSCA.',
            'Arranging for issuance of Letter of Approval (final) from IFSCA.',
        ]
    },
    's1c': {
        'label': 'Bond-Cum-Legal Undertaking',
        'bullets': [
            'Preparation of Application for Bond cum legal undertaking with annexure in consultation with the legal advisor.',
            'Submission of Bond cum Legal Undertaking to Specified Officer for procurement of Duty-free Goods.',
            'Getting the BLUT accepted from Specified officer of the SEZ.',
            'Submission of the duly accepted BLUT to the Development Commissioner office in physical as well on SEZ Online.',
            'Obtaining approval for BLUT from Development Commissioner office/ IFSCA (Admin).',
        ]
    },
    's1d': {
        'label': 'Obtaining Eligibility certificate for exemption',
        'bullets': [
            'Preparation of application for obtaining eligibility certificate.',
            'Submission of the same to the Development Commissioner office.',
            'Obtaining approval for eligibility certificate.',
        ]
    },
    's1e': {
        'label': 'Obtaining Import Export Code registration',
        'bullets': [
            'Providing detailed checklist of documents required for obtaining IEC Code in new entity name.',
            'Preparation and filing of application online.',
            'Submission of the same to respective authorities.',
            'Obtaining allotment of IEC code.',
        ]
    },
    's1f': {
        'label': 'Obtaining GST registration',
        'bullets': [
            'Providing detailed checklist of documents required for obtaining GST registration in new entity name.',
            'Preparation and filing of application online.',
            'Submission of the same to respective authorities.',
            'Obtaining GST registration.',
        ]
    },
    's1g': {
        'label': 'Obtaining RCMC',
        'bullets': [
            'Providing detailed checklist of documents required for obtaining RCMC.',
            'Preparation and filing of application online.',
            'Submission of the same to respective authorities.',
            'Follow up for obtaining RCMC.',
        ]
    },
    's1h': {
        'label': 'FIU (Financial Intelligence Unit) Registration',
        'bullets': [
            'Providing a detailed checklist of documents required for FIU registration under the Prevention of Money Laundering Act (PMLA).',
            'Preparation and filing of the online FIU registration application on the regulatory gateway in prescribed format.',
            'Submission of documents to FIU-IND, coordination with FIU officials, and follow-up for obtaining the FIU registration credentials (FINnet login ID, reporting entity code, etc.).',
        ]
    },
    's1i': {
        'label': 'Professional Tax Registration (PTEC & PTRC)',
        'bullets': [
            'Providing a checklist of documents required for obtaining Professional Tax Enrolment Certificate (PTEC) and Professional Tax Registration Certificate (PTRC).',
            'Preparation and online filing of applications for PTEC and PTRC with the Commercial Tax Department, along with supporting annexures. Liaison with the authorities for verification, addressing clarifications, and obtaining approvals for PTEC and PTRC registrations.',
        ]
    },
}

SECTION2_ITEMS = {
    's2a': {
        'label': 'Assistance in preparation of projections and business plan',
        'bullets': [
            'Collating the information on the promoters of the group.',
            'Collating the information on the background group.',
            'Collating the information on the directors/KMP of the entity.',
            'Preparing the presentation on the business operations based on inputs provided.',
            'Preparing the presentation of organizational chart based on inputs provided.',
            'Preparing the presentation of business plan based on inputs provided.',
        ]
    },
    's2b': {
        'label': 'Support on application of commencement of operations in SEZ/IFSCA',
        'bullets': [
            'Issuance of share certificates in compliance with MCA requirement.',
            'Assistance in company secretarial compliance for conducting 1st Board meeting as per companies Act.',
            'Preparation of documentation required and filing of Stat. Auditor appointment with MCA (if applicable)',
            'Preparation and processing of Letters of Acceptance with SEZ department.',
            'Preparation of application for commencement of operation before SEZ in coordination with legal advisors.',
            'Preparation of application for commencement of operation before IFSCA in coordination with legal advisors.',
            'Assistance in the registration process with the Financial Intelligence Unit Portal to ensure compliance with KYC, AML-CFT regulations.',
            'Regularization of SEZ portal with unit and maker creation after payment of registration fees.',
        ]
    },
    's2c': {
        'label': 'Liaison for Bank Account opening (Foreign currency account and INR account)',
        'bullets': [
            'Draft all necessary forms required for account opening and related transactions with Bank.',
            "Ensure accuracy and completeness of information in the forms, adhering to the bank's specifications.",
            'Collaborate and collate all Know Your Customer (KYC) documents required by Bank for account opening and other transactions.',
            "Verify that all submitted documents meet the bank's KYC requirements and standards.",
            'Guide and assist the client in completing the account opening process with Bank.',
            'Act as an intermediary between the client and Bank in resolving any queries, concerns, or clarifications related to account opening, transactions, or other banking matters.',
            'Provide guidance to the client on the documentation required by Bank and assist in obtaining any additional information needed.',
        ]
    },
    's2d': {
        'label': 'Assisting in additional capital infusion (to meet regulatory requirements)',
        'bullets': [
            'Preparing of relevant set of documents for holding Board meeting with agenda to increase authorized capital.',
            "Preparing of relevant set of documents for holding Shareholder's meeting (EGM) with agenda to increase authorized capital.",
            'Preparation and filing of relevant forms on MCA portal.',
            'Assistance in payment of MCA fees & other stamp duty etc.',
        ]
    },
    's2e': {
        'label': 'ICEGATE Registration',
        'bullets': [
            'Providing a detailed checklist of documents required for registration on the Indian Customs Electronic Gateway (ICEGATE) portal for facilitating import/export transactions.',
            'Preparation and Filing of the application on the ICEGATE portal, coordination with customs authorities, and follow-up for verification and approval of registration.',
        ]
    },
    's2f': {
        'label': 'Shop & Establishment Intimation/Registration.',
        'bullets': [
            'Providing a detailed checklist of documents required for registration / intimation for Shop & Establishment.',
            'Filing of the application, coordination with concerned authorities, and follow-up for verification and letter of approval/intimation.',
        ]
    },
    's2g': {
        'label': 'Obtaining SEZ ID cards',
        'bullets': [
            'Preparation and submission of applications for SEZ ID cards through the relevant SEZ online/offline portal.',
            "Coordination with the SEZ administration / Development Commissioner's office for processing of the applications.",
            'Follow-ups and assistance until issuance of the SEZ ID cards for the approved personnel.',
        ]
    },
}

ALPHA = 'abcdefghijklmnopqrstuvwxyz'


# ── XML helpers ─────────────────────────────────────────────────────────────

def _make_rPr(bold=False, italic=False, size_pt=11, color_hex=None, font_name='Roboto', underline=False):
    rPr = OxmlElement('w:rPr')
    if font_name:
        rFonts = OxmlElement('w:rFonts')
        rFonts.set(qn('w:ascii'), font_name)
        rFonts.set(qn('w:hAnsi'), font_name)
        rPr.append(rFonts)
    if bold:
        rPr.append(OxmlElement('w:b'))
    if italic:
        rPr.append(OxmlElement('w:i'))
    if underline:
        u = OxmlElement('w:u'); u.set(qn('w:val'), 'single'); rPr.append(u)
    sz = OxmlElement('w:sz'); sz.set(qn('w:val'), str(int(size_pt * 2))); rPr.append(sz)
    szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), str(int(size_pt * 2))); rPr.append(szCs)
    if color_hex:
        col = OxmlElement('w:color'); col.set(qn('w:val'), color_hex.lstrip('#')); rPr.append(col)
    return rPr


def _make_pPr(align='left', sb=0, sa=0, li=0, hanging=0, style_id=None):
    pPr = OxmlElement('w:pPr')
    if style_id:
        pStyle = OxmlElement('w:pStyle')
        pStyle.set(qn('w:val'), style_id)
        pPr.append(pStyle)
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), str(int(sb * 20)))
    spacing.set(qn('w:after'), str(int(sa * 20)))
    spacing.set(qn('w:line'), '276')
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    if li or hanging:
        ind = OxmlElement('w:ind')
        ind.set(qn('w:left'), str(int(li * 20)))
        if hanging:
            ind.set(qn('w:hanging'), str(int(hanging * 20)))
        pPr.append(ind)
    jc_map = {'left': 'left', 'center': 'center', 'right': 'right', 'justify': 'both', 'both': 'both'}
    jc_val = jc_map.get(align, 'left')
    if jc_val not in ('left',):
        jc = OxmlElement('w:jc'); jc.set(qn('w:val'), jc_val); pPr.append(jc)
    return pPr


def _p(text='', bold=False, italic=False, size_pt=11, color_hex=None,
       font='Roboto', align='left', sb=0, sa=0, li=0, hanging=0,
       underline=False, style_id=None):
    p = OxmlElement('w:p')
    pPr = _make_pPr(align=align, sb=sb, sa=sa, li=li, hanging=hanging, style_id=style_id)
    p.append(pPr)
    if text:
        r = OxmlElement('w:r')
        r.append(_make_rPr(bold=bold, italic=italic, size_pt=size_pt,
                           color_hex=color_hex, font_name=font, underline=underline))
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = text
        r.append(t); p.append(r)
    return p


def _p_multi(runs, align='left', sb=0, sa=0, li=0, hanging=0, style_id=None):
    """Paragraph with multiple runs (for mixed bold/normal in one line)"""
    p = OxmlElement('w:p')
    pPr = _make_pPr(align=align, sb=sb, sa=sa, li=li, hanging=hanging, style_id=style_id)
    p.append(pPr)
    for run_cfg in runs:
        r = OxmlElement('w:r')
        r.append(_make_rPr(
            bold=run_cfg.get('bold', False),
            italic=run_cfg.get('italic', False),
            size_pt=run_cfg.get('size_pt', 11),
            color_hex=run_cfg.get('color_hex'),
            font_name=run_cfg.get('font', 'Roboto'),
            underline=run_cfg.get('underline', False)
        ))
        t = OxmlElement('w:t')
        t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
        t.text = run_cfg.get('text', '')
        r.append(t); p.append(r)
    return p


def _sp(sa=6):
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), str(int(sa * 20)))
    pPr.append(spacing)
    p.append(pPr)
    return p
def _page_break():
    """Hard page break"""
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    return p


def _bul(text, size_pt=11, color_hex=None, font='Roboto'):
    """Bullet paragraph with proper hanging indent"""
    clean = text.lstrip('•').lstrip('\u2022').strip()
    p = OxmlElement('w:p')
    pPr = OxmlElement('w:pPr')
    spacing = OxmlElement('w:spacing')
    spacing.set(qn('w:before'), '0')
    spacing.set(qn('w:after'), '40')
    spacing.set(qn('w:line'), '276')
    spacing.set(qn('w:lineRule'), 'auto')
    pPr.append(spacing)
    ind = OxmlElement('w:ind')
    ind.set(qn('w:left'), '360')
    ind.set(qn('w:hanging'), '180')
    pPr.append(ind)
    p.append(pPr)
    r = OxmlElement('w:r')
    r.append(_make_rPr(size_pt=size_pt, color_hex=color_hex, font_name=font))
    t = OxmlElement('w:t')
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
    t.text = f'\u2022  {clean}'
    r.append(t); p.append(r)
    return p


def _heading(text, size_pt=12, color_hex='002060', bold=True, sb=10, sa=4):
    return _p(text, bold=bold, size_pt=size_pt, color_hex=color_hex,
              font='Roboto Medium', sb=sb, sa=sa)


def _sub_heading(text, letter, size_pt=11, bold=True):
    """Like: 'a)  Company/ LLP/ Entity formation...' """
    return _p(f'{letter})  {text}', bold=bold, size_pt=size_pt,
              color_hex='000000', font='Roboto', sb=4, sa=2, li=0)


def _make_commercials_table(rows):
    """Build the fee table — full width, compact rows, matching original"""
    tbl = OxmlElement('w:tbl')
    tblPr = OxmlElement('w:tblPr')
    tblStyle = OxmlElement('w:tblStyle')
    tblStyle.set(qn('w:val'), 'TableGrid')
    tblPr.append(tblStyle)

    # Full page width — 8640 twips (6 inches at 1440 per inch)
    tblW = OxmlElement('w:tblW')
    tblW.set(qn('w:w'), '8640')
    tblW.set(qn('w:type'), 'dxa')
    tblPr.append(tblW)

    # Zero table indent so it aligns with page margins
    tblInd = OxmlElement('w:tblInd')
    tblInd.set(qn('w:w'), '0')
    tblInd.set(qn('w:type'), 'dxa')
    tblPr.append(tblInd)

    # Borders
    tblBorders = OxmlElement('w:tblBorders')
    for s in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        b = OxmlElement(f'w:{s}')
        b.set(qn('w:val'), 'single')
        b.set(qn('w:sz'), '4')
        b.set(qn('w:color'), '000000')
        tblBorders.append(b)
    tblPr.append(tblBorders)

    # Table layout fixed
    tblLayout = OxmlElement('w:tblLayout')
    tblLayout.set(qn('w:type'), 'fixed')
    tblPr.append(tblLayout)

    tblJc = OxmlElement('w:jc')
    tblJc.set(qn('w:val'), 'center')
    tblPr.append(tblJc)

    tbl.append(tblPr)

    # Column widths: No.(540) | Scope(6300) | Fees(1800) = 8640 total
    col_widths = [540, 6300, 1800]
    tblGrid = OxmlElement('w:tblGrid')
    for w in col_widths:
        gc = OxmlElement('w:gridCol')
        gc.set(qn('w:w'), str(w))
        tblGrid.append(gc)
    tbl.append(tblGrid)

    def _tc(text, bold=False, align='left', width=None, shade=None, size_pt=10):
        tc = OxmlElement('w:tc')
        tcPr = OxmlElement('w:tcPr')
        if width:
            tcW = OxmlElement('w:tcW')
            tcW.set(qn('w:w'), str(width))
            tcW.set(qn('w:type'), 'dxa')
            tcPr.append(tcW)
        if shade:
            shd = OxmlElement('w:shd')
            shd.set(qn('w:val'), 'clear')
            shd.set(qn('w:color'), 'auto')
            shd.set(qn('w:fill'), shade)
            tcPr.append(shd)
        # Compact cell margins — tight like Excel
        mar = OxmlElement('w:tcMar')
        for side, val in zip(['top', 'bottom', 'left', 'right'], [60, 60, 108, 108]):
            m = OxmlElement(f'w:{side}')
            m.set(qn('w:w'), str(val))
            m.set(qn('w:type'), 'dxa')
            mar.append(m)
        tcPr.append(mar)
        va = OxmlElement('w:vAlign')
        va.set(qn('w:val'), 'center')
        tcPr.append(va)
        tc.append(tcPr)

        # Paragraph inside cell
        p = OxmlElement('w:p')
        pPr = OxmlElement('w:pPr')
        # Compact line spacing
        spacing = OxmlElement('w:spacing')
        spacing.set(qn('w:before'), '0')
        spacing.set(qn('w:after'), '0')
        spacing.set(qn('w:line'), '240')
        spacing.set(qn('w:lineRule'), 'auto')
        pPr.append(spacing)
        if align != 'left':
            jc = OxmlElement('w:jc')
            jc.set(qn('w:val'), 'center' if align == 'center' else 'right')
            pPr.append(jc)
        p.append(pPr)
        if text:
            r = OxmlElement('w:r')
            rPr = OxmlElement('w:rPr')
            rFonts = OxmlElement('w:rFonts')
            rFonts.set(qn('w:ascii'), 'Roboto')
            rFonts.set(qn('w:hAnsi'), 'Roboto')
            rPr.append(rFonts)
            if bold:
                rPr.append(OxmlElement('w:b'))
            sz = OxmlElement('w:sz')
            sz.set(qn('w:val'), str(size_pt * 2))
            rPr.append(sz)
            szCs = OxmlElement('w:szCs')
            szCs.set(qn('w:val'), str(size_pt * 2))
            rPr.append(szCs)
            # White text for header
            if shade == '1F3864':
                col = OxmlElement('w:color')
                col.set(qn('w:val'), 'FFFFFF')
                rPr.append(col)
            r.append(rPr)
            t = OxmlElement('w:t')
            t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')
            t.text = text
            r.append(t)
            p.append(r)
        tc.append(p)
        return tc

    for ri, row in enumerate(rows):
        tr = OxmlElement('w:tr')
        # Compact row height — 360 twips (0.25 inch), exact like Excel
        trPr = OxmlElement('w:trPr')
        trHeight = OxmlElement('w:trHeight')
        trHeight.set(qn('w:val'), '360')
        trHeight.set(qn('w:hRule'), 'atLeast')
        trPr.append(trHeight)
        tr.append(trPr)

        is_header = (ri == 0)
        shade = '1F3864' if is_header else None
        aligns = ['center', 'left', 'right']
        for ci, cell_text in enumerate(row):
            tc = _tc(cell_text, bold=is_header, align=aligns[ci],
                     width=col_widths[ci], shade=shade, size_pt=10)
            tr.append(tc)
        tbl.append(tr)
    return tbl

# ── Main route ───────────────────────────────────────────────────────────────

@app.route('/')
def index():
    from flask import make_response
    response = make_response(send_file(os.path.join(BASE_DIR, 'index.html')))
    response.headers['Cache-Control'] = 'no-store, no-cache, must-revalidate, max-age=0'
    return response


@app.route('/generate_giftcity_word', methods=['POST'])
def generate_giftcity_word():
    try:
        data = request.json or {}

        template_path = os.path.join(BASE_DIR, 'InCorp Proposal Gift City.docx')
        if not os.path.exists(template_path):
            return jsonify({'error': f'Template not found: {template_path}'}), 500

        # ── Parse inputs ────────────────────────────────────────
        proposal_date = data.get('proposalDate', '')
        try:
            fd = datetime.strptime(proposal_date, '%Y-%m-%d').strftime('%d.%m.%Y')
        except:
            fd = datetime.now().strftime('%d.%m.%Y')

        client_name       = data.get('clientName', '')
        client_desig      = data.get('clientDesignation', '')
        client_company    = data.get('clientCompany', '')
        client_addr1      = data.get('clientAddress1', '')
        client_addr2      = data.get('clientAddress2', '')
        client_addr3      = data.get('clientAddress3', '')
        subject_line      = f'Scope & Proposal for {client_company} - GIFT City Setup & Services'
        letter_body       = ''
        summary_text = data.get('summaryOfRequirements', '').strip()
        if not summary_text:
            summary_text = f'{client_company} are keen to set up a Fintech entity at GIFT City under existing IFSC regulations and has requested InCorp to provide a proposal for the same. Our scope of services will be as follows:'

        # First name for "Dear X,"
        first_name = client_name.split()[-1] if client_name else 'Sir/Madam'
        # Remove salutation for "Dear" if name includes Mr./Ms. etc.
        name_parts = client_name.split()
        if name_parts and name_parts[0] in ('Mr.', 'Ms.', 'Mrs.', 'Dr.', 'Prof.'):
            first_name = ' '.join(name_parts[1:])

        # Section 1 selected items
        s1_keys = ['s1a', 's1b', 's1c', 's1d', 's1e', 's1f', 's1g', 's1h', 's1i']
        s1_selected = [k for k in s1_keys if data.get(k)]

        # Section 2 selected items
        s2_keys = ['s2a', 's2b', 's2c', 's2d', 's2e', 's2f', 's2g']
        s2_selected = [k for k in s2_keys if data.get(k)]

        fee_s1      = data.get('fee_s1', '5,000')
        fee_s2_abcd = data.get('fee_s2_abcd', '750')
        fee_s2_efg  = data.get('fee_s2_efg', '750')

        sig1_name  = data.get('sig1Name', 'Meet Thakkar')
        sig1_title = data.get('sig1Title', 'Head – GIFT IFSC Practice')
        sig2_name  = data.get('sig2Name', 'Nikhil Joshi')
        sig2_title = data.get('sig2Title', 'Director Sales – Managed Services')

        # ── Build dynamic elements ───────────────────────────────
        elems = []

        # DATE
        elems.append(_p(fd, size_pt=11, font='Roboto', sb=0, sa=4))
        elems.append(_sp(2))

        # ADDRESS BLOCK
        elems.append(_p(client_name, bold=False, size_pt=11, font='Roboto', sa=0))
        if client_desig:
            elems.append(_p(client_desig, size_pt=11, font='Roboto', sa=0))
        elems.append(_p(client_company, size_pt=11, font='Roboto', sa=0))
        for addr in [client_addr1, client_addr2, client_addr3]:
            if addr.strip():
                elems.append(_p(addr, size_pt=11, font='Roboto', sa=0))
        elems.append(_sp(6))

        # DEAR
        elems.append(_p(f'Dear {first_name},', size_pt=11, font='Roboto', sb=0, sa=6))

        # SUBJECT
        elems.append(_p_multi([
            {'text': 'Re: ', 'bold': True, 'size_pt': 11, 'font': 'Roboto'},
            {'text': subject_line, 'bold': True, 'size_pt': 11, 'font': 'Roboto'},
        ], sb=2, sa=8))

        # LETTER BODY
        if letter_body:
            for line in letter_body.split('\n'):
                elems.append(_p(line.strip(), size_pt=11, font='Roboto', align='justify', sa=4))
        else:
            default_paras = [
                'We are pleased to present our proposal to you.',
                '',
                'At InCorp (now Ascentium India), we offer turnkey advisory and implementation services for businesses expanding into GIFT IFSC covering entity setup, regulatory licensing, obtaining various registration and supporting ongoing compliance like accounting, tax, regulatory compliances and corporate secretarial services. We support entity covered under various regulations including fund management, capital market intermediaries, fintech, NBFCs, banks, foreign universities and other IFSC-regulated activities.',
                '',
                f'This proposal outlines our approach to the GIFT IFSC related requirements that you have shared for {client_company}. We trust that it aligns with your expectations, and we look forward to establishing a long-term, mutually beneficial relationship with you and your company.',
            ]
            for para in default_paras:
                if para:
                    elems.append(_p(para, size_pt=11, font='Roboto', align='justify', sa=4))
                else:
                    elems.append(_sp(4))

        elems.append(_sp(6))
        elems.append(_p('Yours Sincerely and on behalf of InCorp (India),', size_pt=11, font='Roboto', sa=2))
        elems.append(_sp(12))
        elems.append(_p(sig2_name, bold=True, size_pt=11, font='Roboto', sa=0))
        elems.append(_p(sig2_title, size_pt=11, font='Roboto', sa=0))
        elems.append(_p('InCorp Advisory Services Pvt. Ltd.', size_pt=11, font='Roboto', sa=0))
        elems.append(_sp(12))

        elems.append(_page_break())
        # ── SUMMARY OF REQUIREMENTS ──────────────────────────────
        elems.append(_p_multi([
            {'text': '\u2756  SUMMARY OF REQUIREMENTS', 'bold': True, 'underline': True,
             'size_pt': 12, 'color_hex': 'C00000', 'font': 'Roboto'}
        ], sb=10, sa=6))
        elems.append(_p(summary_text, size_pt=11, font='Roboto', align='justify', sb=0, sa=8))

        # ── SCOPE OF SERVICES HEADING ────────────────────────────
        elems.append(_p_multi([
            {'text': '\u2756  SCOPE OF SERVICES', 'bold': True, 'underline': True,
             'size_pt': 12, 'color_hex': 'C00000', 'font': 'Roboto'}
        ], sb=8, sa=6))

        # ── SECTION 1 ────────────────────────────────────────────
        if s1_selected:
            elems.append(_heading('Entity Setup Services', size_pt=12, color_hex='002060', sb=8, sa=4))
            for idx, key in enumerate(s1_selected):
                item = SECTION1_ITEMS[key]
                letter = ALPHA[idx]
                # Sub-heading with auto letter
                elems.append(_sub_heading(item['label'], letter))
                for bullet in item['bullets']:
                    elems.append(_bul(bullet))
                elems.append(_sp(4))

        # ── SECTION 2 ────────────────────────────────────────────
        if s2_selected:
            elems.append(_page_break())
        if s2_selected:
            elems.append(_heading('Other Support Services (Optional)', size_pt=12, color_hex='002060', sb=8, sa=4))
            for idx, key in enumerate(s2_selected):
                item = SECTION2_ITEMS[key]
                letter = ALPHA[idx]
                elems.append(_sub_heading(item['label'], letter))
                for bullet in item['bullets']:
                    elems.append(_bul(bullet))
                elems.append(_sp(4))

        # ── COMMERCIALS TABLE ────────────────────────────────────
        elems.append(_page_break())
        elems.append(_sp(6))
        elems.append(_heading('Commercials with terms & conditions:', size_pt=13, color_hex='C00000', sb=8, sa=6))

        # Build fee rows dynamically
        comm_rows = [['No.', 'Scope', 'Fees (USD)']]
        row_num = 1

        if s1_selected:
            comm_rows.append([f'{row_num}.', 'Entity Setup Services', f'USD {fee_s1}'])
            row_num += 1

        if s2_selected:
            comm_rows.append([f'{row_num}.', 'Other Support Services (Optional)', ''])
            row_num += 1

            # Group a-d vs e-g based on what's selected
            s2_abcd = [k for k in s2_selected if k in ('s2a', 's2b', 's2c', 's2d')]
            s2_efg  = [k for k in s2_selected if k in ('s2e', 's2f', 's2g')]

            if s2_abcd:
                # List which sub-items are selected
                labels = ', '.join(
                    f'{ALPHA[s2_selected.index(k)]}'
                    for k in s2_abcd
                )
                comm_rows.append(['', f'  {labels} (as selected above)', f'USD {fee_s2_abcd} per service'])

            if s2_efg:
                labels = ', '.join(
                    f'{ALPHA[s2_selected.index(k)]}'
                    for k in s2_efg
                )
                comm_rows.append(['', f'  {labels} (as selected above)', f'USD {fee_s2_efg} (lumpsum)'])

        elems.append(_make_commercials_table(comm_rows))
        elems.append(_sp(6))

        # NOTE block
        elems.append(_p_multi([
            {'text': 'Note:', 'bold': True, 'underline': True, 'size_pt': 11, 'font': 'Roboto'},
        ], sb=4, sa=4))
        for note in [
            'All Fees quoted are exclusive of applicable GST and do not include any out-of-pocket expenses charged at actuals.',
            'Fees exclude expenses related to any government/statutory filing fees, levies and taxes.',
        ]:
            elems.append(_bul(note, size_pt=11))
        elems.append(_sp(6))

        # Force contact page to always start on new page
        from docx.oxml import OxmlElement
        from docx.oxml.ns import qn

        # ── INJECT INTO TEMPLATE ─────────────────────────────────
        with open(template_path, 'rb') as f:
            doc = Document(io.BytesIO(f.read()))

        body = doc.element.body
        children = list(body)

        # Dynamic content: indices 71-226 based on analysis
        # (date at 71, letter through commercials, notes at ~226)
        # Remove old dynamic content
        for el in children[71:228]:
            body.remove(el)

        # Insert new elements after index 70
        ref = list(body)[70]
        for elem in elems:
            ref.addnext(elem)
            ref = elem

        # Update company name on cover (child index 22 = "LAKESHORE INDIA")
        # Update company name on cover — same logic as InCorp app.py
        # Update company name on cover
        try:
            cover_p = list(body)[22]
            if cover_p.tag == qn('w:p'):
                # Remove left indent (template has 5760 twips = 4 inch indent)
                pPr = cover_p.find(qn('w:pPr'))
                if pPr is None:
                    pPr = OxmlElement('w:pPr')
                    cover_p.insert(0, pPr)
                for old_ind in pPr.findall(qn('w:ind')): pPr.remove(old_ind)
                ind = OxmlElement('w:ind')
                ind.set(qn('w:left'), '2800')
                pPr.append(ind)
                for old_jc in pPr.findall(qn('w:jc')): pPr.remove(old_jc)
                jc = OxmlElement('w:jc'); jc.set(qn('w:val'), 'left'); pPr.append(jc)
                # Auto font size based on company name length
                name_upper = client_company.upper()
                name_len = len(name_upper)
                if name_len <= 15:   font_val = '52'
                elif name_len <= 25: font_val = '48'
                elif name_len <= 35: font_val = '40'
                else:                font_val = '32'
                # Only update the text run (not the image run)
                for r_elem in cover_p.iter(qn('w:r')):
                    t_elem = r_elem.find(qn('w:t'))
                    if t_elem is not None and t_elem.text and t_elem.text.strip():
                        t_elem.text = name_upper
                        rPr = r_elem.find(qn('w:rPr'))
                        if rPr is None:
                            rPr = OxmlElement('w:rPr'); r_elem.insert(0, rPr)
                        for old in rPr.findall(qn('w:sz')): rPr.remove(old)
                        for old in rPr.findall(qn('w:szCs')): rPr.remove(old)
                        sz = OxmlElement('w:sz'); sz.set(qn('w:val'), font_val); rPr.append(sz)
                        szCs = OxmlElement('w:szCs'); szCs.set(qn('w:val'), font_val); rPr.append(szCs)
        except:
            pass

        # Update signatory contact page (children after commercials)
        # sig1 name is around index 247 in original — just update text nodes
        try:
            all_children = list(body)
            for child in all_children:
                if child.tag == qn('w:p'):
                    text = ''.join(t.text or '' for t in child.iter(qn('w:t')))
                    if 'Meet Thakkar' in text or 'Head \u2013 GIFT IFSC Practice' in text:
                        for t_elem in child.iter(qn('w:t')):
                            if t_elem.text and 'Meet Thakkar' in t_elem.text:
                                t_elem.text = t_elem.text.replace('Meet Thakkar', sig1_name)
                            if t_elem.text and 'Head \u2013 GIFT IFSC Practice' in t_elem.text:
                                t_elem.text = t_elem.text.replace('Head \u2013 GIFT IFSC Practice', sig1_title)
                    if 'Nikhil Joshi' in text:
                        for t_elem in child.iter(qn('w:t')):
                            if t_elem.text and 'Nikhil Joshi' in t_elem.text:
                                t_elem.text = t_elem.text.replace('Nikhil Joshi', sig2_name)
                    if 'Director Sales' in text and 'Managed Services' in text:
                        for t_elem in child.iter(qn('w:t')):
                            if t_elem.text and 'Director Sales' in t_elem.text:
                                t_elem.text = t_elem.text.replace(
                                    'Director Sales- Managed Services', sig2_title
                                ).replace('Director Sales – Managed Services', sig2_title)
        except:
            pass

        # Save
        out_buffer = io.BytesIO()
        doc.save(out_buffer)
        out_buffer.seek(0)

        company_safe = client_company.replace(' ', '_')[:40]
        filename = f'InCorp_GiftCity_Proposal_{company_safe}_{datetime.now().strftime("%Y%m%d")}.docx'
        print(f'✅ Gift City Word document generated: {filename}')

        return send_file(
            out_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
            as_attachment=True,
            download_name=filename
        )

    except Exception as e:
        import traceback; traceback.print_exc()
        return jsonify({'error': str(e)}), 500


if __name__ == '__main__':
    print("=" * 60)
    print("InCorp Gift City Proposal Generator")
    print("=" * 60)
    print(f"Template expected at: {os.path.join(BASE_DIR, 'InCorp_Proposal_for_LSI.docx')}")
    print("\n🚀 Starting server on http://localhost:5001")
    print("=" * 60)
    port = int(os.environ.get('PORT', 5001))
    app.run(debug=False, host='0.0.0.0', port=port)
