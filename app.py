from flask import Flask, request, jsonify, render_template, send_file, session, redirect, url_for
from werkzeug.security import generate_password_hash, check_password_hash
import sqlite3
import os
import io
import csv
import json
import functools
from datetime import datetime, date
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 32 * 1024 * 1024  # 32MB max upload size

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'chheepa_samaj.db')

# Ensure consistent persistent secret key for admin session cookies
SECRET_KEY_FILE = os.path.join(BASE_DIR, '.flask_secret_key')
if os.path.exists(SECRET_KEY_FILE):
    try:
        with open(SECRET_KEY_FILE, 'rb') as f:
            app.secret_key = f.read()
    except Exception:
        app.secret_key = b'chheepa_samaj_khanda_falsa_admin_secret_key_2026'
else:
    try:
        new_key = os.urandom(32)
        with open(SECRET_KEY_FILE, 'wb') as f:
            f.write(new_key)
        app.secret_key = new_key
    except Exception:
        app.secret_key = b'chheepa_samaj_khanda_falsa_admin_secret_key_2026'


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def calculate_age(dob_str):
    if not dob_str:
        return None
    try:
        dob = datetime.strptime(str(dob_str).strip()[:10], '%Y-%m-%d').date()
        today = date.today()
        return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    except Exception:
        return None


SAMPLE_MEMBERS = [
    (
        'CSJ-0001', 'FAM-KF-101', 1, 'Self (Head)',
        'Haji Mohammed Farooq Chheepa', 'हाजी मोहम्मद फारूक छीपा',
        'Male', '1966-04-12', 60, 'B+', 'Married',
        'Late Haji Abdul Rahim Chheepa', 'Late Hajjan Fatima Begum',
        'Rashida Begum', 4, 6,
        'Imran (Son), Irfan (Son), Rukhsar (Daughter-Married), Amina (Daughter-in-law)',
        'Solanki Chheepa', 'Khanda Falsa, Jodhpur', 'Adaa Bazaar / Khanda Falsa Chowk',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9829011223', '9829011223', 'farooq.textiles@example.com',
        'Graduate', 'B.Com', 'Jai Narain Vyas University (JNVU), Jodhpur',
        'Business / Self-Employed', 'Traditional Handblock & Textile Merchant',
        'Farooq & Sons Chheepa Handloom Exports', 'Main Road, Khanda Falsa, Jodhpur',
        '5 - 10 Lakh', 'Dabu & Bagru Handblock Printing, Textile Dyeing',
        '4821', 'JAN-29384711', 'Samaj Patron / Elder Committee',
        'Active contributor to Chheepa Jamaat education scholarship fund.', ''
    ),
    (
        'CSJ-0002', 'FAM-KF-101', 0, 'Wife',
        'Rashida Begum', 'रशीदा बेगम',
        'Female', '1970-08-20', 56, 'O+', 'Married',
        'Abdul Sattar Chheepa', 'Zainab Bibi',
        'Haji Mohammed Farooq Chheepa', 4, 6,
        'Wife of Head of Family (FAM-KF-101)',
        'Chauhan Chheepa', 'Pipar City / Jodhpur', 'Adaa Bazaar / Khanda Falsa Chowk',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9829011224', '9829011224', '',
        'Secondary (10th)', 'Arts', 'Govt. Girls School, Jodhpur',
        'Homemaker', 'Homemaker & Bandhej Artisan',
        'Home Boutique', 'Khanda Falsa, Jodhpur',
        'Below 2.5 Lakh', 'Tie-Dye (Bandhani) Craft',
        '9912', 'JAN-29384711', 'Women Wing Member',
        '', ''
    ),
    (
        'CSJ-0003', 'FAM-KF-101', 0, 'Son',
        'Er. Imran Ahmed Chheepa', 'इमरान अहमद छीपा',
        'Male', '1993-11-05', 32, 'B+', 'Married',
        'Haji Mohammed Farooq Chheepa', 'Rashida Begum',
        'Dr. Amina Parveen', 1, 6,
        'Son of Haji Mohammed Farooq Chheepa',
        'Solanki Chheepa', 'Khanda Falsa, Jodhpur', 'Adaa Bazaar / Khanda Falsa Chowk',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9414122334', '9414122334', 'imran.chheepa@example.com',
        'Post-Graduate / Professional', 'B.Tech (Computer Science)', 'MBM Engineering College, Jodhpur',
        'Private Sector Job', 'Senior Software Architect',
        'Marwar Tech Solutions Pvt Ltd', 'IT Park, Mia Basni, Jodhpur',
        'Above 10 Lakh', 'Software Engineering, E-Commerce for Handloom',
        '3319', 'JAN-29384711', 'Youth Wing Coordinator',
        'Organizes free career counseling camps for Samaj students.', ''
    ),
    (
        'CSJ-0004', 'FAM-KF-101', 0, 'Daughter-in-Law',
        'Dr. Amina Parveen Chheepa', 'डॉ. अमीना परवीन छीपा',
        'Female', '1995-02-18', 31, 'A+', 'Married',
        'Mohammed Yusuf Chheepa', 'Salma Bano',
        'Er. Imran Ahmed Chheepa', 1, 6,
        'Wife of Er. Imran Ahmed Chheepa',
        'Bhati Chheepa', 'Sojat Gate, Jodhpur', 'Adaa Bazaar / Khanda Falsa Chowk',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'House No. 42, Near Jama Masjid, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9414199887', '9414199887', 'dr.amina@example.com',
        'Doctorate / Medical', 'MBBS, MD (Pediatrics)', 'Dr. S.N. Medical College, Jodhpur',
        'Government Job', 'Medical Officer (Pediatrician)',
        'Umaid Hospital, Jodhpur', 'Siwanchi Gate Rd, Jodhpur',
        'Above 10 Lakh', 'Pediatrics, Community Health Camps',
        '7741', 'JAN-29384711', 'Medical Committee Lead',
        'Leads annual blood donation and health checkup camp in Khanda Falsa.', ''
    ),
    (
        'CSJ-0005', 'FAM-KF-102', 1, 'Self (Head)',
        'Abdul Qadir Chheepa', 'अब्दुल कादिर छीपा',
        'Male', '1978-06-15', 48, 'AB+', 'Married',
        'Haji Ghulam Nabi Chheepa', 'Razia Sultan',
        'Shabnam Bano', 3, 5,
        'Zaid (Son - Student), Sana (Daughter - Student), Rehan (Son)',
        'Panwar Chheepa', 'Khanda Falsa, Jodhpur', 'Bhistiyon Ka Bas / Khanda Falsa',
        'Plot 19, Gali No. 3, Khanda Falsa, Jodhpur',
        'Plot 19, Gali No. 3, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9928055667', '9928055667', 'qadir.dyers@example.com',
        'Higher Secondary (12th)', 'Commerce', 'Mahesh Senior Secondary School, Jodhpur',
        'Business / Self-Employed', 'Screen Printing & Fabric Processing Owner',
        'Marwar Chheepa Tex Prints', 'Basni Phase II, Jodhpur',
        '5 - 10 Lakh', 'Textile Screen Printing & Export Processing',
        '6104', 'JAN-88412093', 'General Secretary',
        'Secretary of Khanda Falsa Chheepa Jamaat Committee.', ''
    ),
    (
        'CSJ-0006', 'FAM-KF-102', 0, 'Daughter',
        'Sana Qadir Chheepa', 'सना कादिर छीपा',
        'Female', '2004-09-10', 22, 'O+', 'Single',
        'Abdul Qadir Chheepa', 'Shabnam Bano',
        '', 0, 5,
        'Daughter of Abdul Qadir Chheepa',
        'Panwar Chheepa', 'Khanda Falsa, Jodhpur', 'Bhistiyon Ka Bas / Khanda Falsa',
        'Plot 19, Gali No. 3, Khanda Falsa, Jodhpur',
        'Plot 19, Gali No. 3, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9928055668', '9928055668', 'sana.chheepa@example.com',
        'Graduate', 'CA Intermediate & B.Com (Hons)', 'JNVU / ICAI Jodhpur Branch',
        'Student', 'Chartered Accountancy Aspirant',
        'ICAI Jodhpur', 'Jodhpur',
        'Not Applicable', 'Accounting, Taxation, Gold Medalist in 12th Board',
        '5520', 'JAN-88412093', 'Youth Merit Scholar',
        'Awarded Samaj Pratibha Samman 2025.', ''
    )
]


def init_db():
    conn = get_db()
    c = conn.cursor()

    # Members Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS members (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            member_code TEXT UNIQUE,
            family_id TEXT NOT NULL,
            is_head INTEGER DEFAULT 0,
            relation_to_head TEXT DEFAULT 'Self',
            full_name TEXT NOT NULL,
            full_name_hi TEXT DEFAULT '',
            gender TEXT DEFAULT 'Male',
            dob TEXT DEFAULT '',
            age INTEGER DEFAULT 0,
            blood_group TEXT DEFAULT '',
            marital_status TEXT DEFAULT 'Single',
            father_name TEXT DEFAULT '',
            mother_name TEXT DEFAULT '',
            spouse_name TEXT DEFAULT '',
            children_count INTEGER DEFAULT 0,
            total_family_members INTEGER DEFAULT 1,
            family_members_notes TEXT DEFAULT '',
            gotra_subcaste TEXT DEFAULT '',
            native_origin TEXT DEFAULT 'Khanda Falsa, Jodhpur',
            mohalla_area TEXT DEFAULT 'Khanda Falsa',
            current_address TEXT DEFAULT '',
            permanent_address TEXT DEFAULT '',
            city TEXT DEFAULT 'Jodhpur',
            pincode TEXT DEFAULT '342001',
            phone_primary TEXT DEFAULT '',
            phone_whatsapp TEXT DEFAULT '',
            email TEXT DEFAULT '',
            education_level TEXT DEFAULT '',
            education_stream TEXT DEFAULT '',
            institution_name TEXT DEFAULT '',
            employment_status TEXT DEFAULT '',
            occupation_title TEXT DEFAULT '',
            business_or_employer TEXT DEFAULT '',
            work_address TEXT DEFAULT '',
            annual_income_range TEXT DEFAULT '',
            skills_specialization TEXT DEFAULT '',
            aadhaar_last4 TEXT DEFAULT '',
            bhamashah_jan_aadhaar TEXT DEFAULT '',
            samaj_role TEXT DEFAULT 'Member',
            special_notes TEXT DEFAULT '',
            photo_data TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP,
            updated_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Admin Users Table
    c.execute('''
        CREATE TABLE IF NOT EXISTS admin_users (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            username TEXT UNIQUE NOT NULL,
            password_hash TEXT NOT NULL,
            full_name TEXT DEFAULT 'Chief Administrator',
            email TEXT DEFAULT 'admin@chheepasamaj.org',
            last_login TEXT DEFAULT '',
            created_at TEXT DEFAULT CURRENT_TIMESTAMP
        )
    ''')

    # Seed default administrator if not present
    c.execute("SELECT id FROM admin_users WHERE username = 'admin'")
    if not c.fetchone():
        default_pwd_hash = generate_password_hash('admin123')
        c.execute('''
            INSERT INTO admin_users (username, password_hash, full_name, email, last_login)
            VALUES ('admin', ?, 'Chief Samaj Administrator', 'admin@chheepasamaj.org', '')
        ''', (default_pwd_hash,))
        conn.commit()

    # Seed members if empty
    c.execute('SELECT COUNT(*) as cnt FROM members')
    count = c.fetchone()['cnt']
    if count == 0:
        c.executemany('''
            INSERT INTO members (
                member_code, family_id, is_head, relation_to_head,
                full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
                father_name, mother_name, spouse_name, children_count, total_family_members,
                family_members_notes, gotra_subcaste, native_origin, mohalla_area,
                current_address, permanent_address, city, pincode,
                phone_primary, phone_whatsapp, email,
                education_level, education_stream, institution_name,
                employment_status, occupation_title, business_or_employer, work_address,
                annual_income_range, skills_specialization,
                aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, photo_data
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
        ''', SAMPLE_MEMBERS)
        conn.commit()

    conn.close()


def generate_next_code(conn):
    c = conn.cursor()
    c.execute('SELECT MAX(id) as max_id FROM members')
    row = c.fetchone()
    next_id = (row['max_id'] or 0) + 1
    return f"CSJ-{next_id:04d}"


def admin_required(f):
    @functools.wraps(f)
    def decorated_function(*args, **kwargs):
        if not session.get('admin_id'):
            if request.path.startswith('/admin/api/') or request.is_json:
                return jsonify({'error': 'Unauthorized. Admin login required.', 'auth_required': True}), 401
            return redirect(url_for('admin_login_page', next=request.path))
        return f(*args, **kwargs)
    return decorated_function


# ==========================================
# PUBLIC ROUTES
# ==========================================
@app.route('/')
def index():
    return render_template('index.html')


@app.route('/api/members', methods=['GET'])
def list_members():
    conn = get_db()
    c = conn.cursor()

    q = request.args.get('q', '').strip()
    family_id = request.args.get('family_id', '').strip()
    gender = request.args.get('gender', '').strip()
    marital_status = request.args.get('marital_status', '').strip()
    education_level = request.args.get('education_level', '').strip()
    employment_status = request.args.get('employment_status', '').strip()
    mohalla_area = (request.args.get('mohalla_area') or request.args.get('area') or '').strip()
    blood_group = request.args.get('blood_group', '').strip()
    only_heads = request.args.get('only_heads', '').strip()
    min_age = request.args.get('min_age', '').strip()
    max_age = request.args.get('max_age', '').strip()

    query = "SELECT * FROM members WHERE 1=1"
    params = []

    if q:
        like_q = f"%{q}%"
        query += """ AND (
            full_name LIKE ? OR
            full_name_hi LIKE ? OR
            member_code LIKE ? OR
            family_id LIKE ? OR
            father_name LIKE ? OR
            mother_name LIKE ? OR
            spouse_name LIKE ? OR
            phone_primary LIKE ? OR
            phone_whatsapp LIKE ? OR
            current_address LIKE ? OR
            mohalla_area LIKE ? OR
            occupation_title LIKE ? OR
            business_or_employer LIKE ? OR
            education_stream LIKE ? OR
            gotra_subcaste LIKE ? OR
            family_members_notes LIKE ?
        )"""
        params.extend([like_q] * 16)

    if family_id:
        query += " AND family_id = ?"
        params.append(family_id)
    if gender:
        query += " AND gender = ?"
        params.append(gender)
    if marital_status:
        query += " AND marital_status = ?"
        params.append(marital_status)
    if education_level:
        query += " AND education_level = ?"
        params.append(education_level)
    if employment_status:
        query += " AND employment_status = ?"
        params.append(employment_status)
    if mohalla_area:
        space_flexible_like = "%" + "%".join(mohalla_area.split()) + "%"
        query += " AND (mohalla_area = ? OR mohalla_area LIKE ? OR mohalla_area LIKE ?)"
        params.extend([mohalla_area, f"%{mohalla_area}%", space_flexible_like])
    if blood_group:
        query += " AND blood_group = ?"
        params.append(blood_group)
    if only_heads == '1':
        query += " AND is_head = 1"
    if min_age.isdigit():
        query += " AND age >= ?"
        params.append(int(min_age))
    if max_age.isdigit():
        query += " AND age <= ?"
        params.append(int(max_age))

    query += " ORDER BY family_id ASC, is_head DESC, age DESC"

    # Pagination support
    page = request.args.get('page', '1').strip()
    per_page = request.args.get('per_page', '').strip()

    if per_page.isdigit() and int(per_page) > 0:
        per_page_int = int(per_page)
        page_int = max(1, int(page)) if page.isdigit() else 1

        # First get total count
        count_query = query.replace("SELECT *", "SELECT COUNT(*)", 1)
        c.execute(count_query, params)
        total_count = c.fetchone()[0]
        total_pages = max(1, (total_count + per_page_int - 1) // per_page_int)
        page_int = min(page_int, total_pages)

        offset = (page_int - 1) * per_page_int
        query += " LIMIT ? OFFSET ?"
        params.extend([per_page_int, offset])

        c.execute(query, params)
        rows = [dict(r) for r in c.fetchall()]
        conn.close()
        return jsonify({
            'members': rows,
            'count': len(rows),
            'total_count': total_count,
            'total_pages': total_pages,
            'current_page': page_int,
            'per_page': per_page_int
        })
    else:
        c.execute(query, params)
        rows = [dict(r) for r in c.fetchall()]
        conn.close()
        return jsonify({'members': rows, 'count': len(rows)})


@app.route('/api/areas', methods=['GET'])
def list_areas():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT mohalla_area as name, COUNT(*) as count 
        FROM members 
        WHERE mohalla_area IS NOT NULL AND TRIM(mohalla_area) != ''
        GROUP BY mohalla_area 
        ORDER BY mohalla_area ASC
    """)
    areas = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'areas': areas, 'count': len(areas)})


@app.route('/api/members/<int:member_id>', methods=['GET'])
def get_member(member_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM members WHERE id = ?", (member_id,))
    row = c.fetchone()
    if not row:
        conn.close()
        return jsonify({'error': 'Member not found'}), 404

    member = dict(row)
    c.execute("SELECT id, member_code, full_name, full_name_hi, relation_to_head, age, gender, marital_status, occupation_title, is_head, phone_primary FROM members WHERE family_id = ? ORDER BY is_head DESC, age DESC", (member['family_id'],))
    family_members = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'member': member, 'family_members': family_members})


@app.route('/api/members', methods=['POST'])
def create_member():
    data = request.json or {}
    full_name = (data.get('full_name') or '').strip()
    if not full_name:
        return jsonify({'error': 'Full Name is required'}), 400

    conn = get_db()
    c = conn.cursor()

    member_code = (data.get('member_code') or '').strip()
    if not member_code:
        member_code = generate_next_code(conn)

    family_id = (data.get('family_id') or '').strip()
    if not family_id:
        family_id = f"FAM-KF-{datetime.now().strftime('%y%m%d%H%M')}"

    dob = (data.get('dob') or '').strip()
    age = data.get('age')
    if (age is None or str(age).strip() == '') and dob:
        calc = calculate_age(dob)
        age = calc if calc is not None else 0
    else:
        try:
            age = int(age)
        except Exception:
            age = 0

    is_head = 1 if data.get('is_head') in [True, 1, '1', 'true'] else 0

    try:
        children_count = int(data.get('children_count') or 0)
    except Exception:
        children_count = 0

    try:
        total_family_members = int(data.get('total_family_members') or 1)
    except Exception:
        total_family_members = 1

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    c.execute('''
        INSERT INTO members (
            member_code, family_id, is_head, relation_to_head,
            full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
            father_name, mother_name, spouse_name, children_count, total_family_members,
            family_members_notes, gotra_subcaste, native_origin, mohalla_area,
            current_address, permanent_address, city, pincode,
            phone_primary, phone_whatsapp, email,
            education_level, education_stream, institution_name,
            employment_status, occupation_title, business_or_employer, work_address,
            annual_income_range, skills_specialization,
            aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, photo_data,
            created_at, updated_at
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
    ''', (
        member_code, family_id, is_head, data.get('relation_to_head', 'Self'),
        full_name, data.get('full_name_hi', ''), data.get('gender', 'Male'), dob, age,
        data.get('blood_group', ''), data.get('marital_status', 'Single'),
        data.get('father_name', ''), data.get('mother_name', ''), data.get('spouse_name', ''),
        children_count, total_family_members, data.get('family_members_notes', ''),
        data.get('gotra_subcaste', ''), data.get('native_origin', 'Khanda Falsa, Jodhpur'),
        data.get('mohalla_area', 'Khanda Falsa'), data.get('current_address', ''),
        data.get('permanent_address', ''), data.get('city', 'Jodhpur'), data.get('pincode', '342001'),
        data.get('phone_primary', ''), data.get('phone_whatsapp', ''), data.get('email', ''),
        data.get('education_level', ''), data.get('education_stream', ''), data.get('institution_name', ''),
        data.get('employment_status', ''), data.get('occupation_title', ''),
        data.get('business_or_employer', ''), data.get('work_address', ''),
        data.get('annual_income_range', ''), data.get('skills_specialization', ''),
        data.get('aadhaar_last4', ''), data.get('bhamashah_jan_aadhaar', ''),
        data.get('samaj_role', 'Member'), data.get('special_notes', ''),
        data.get('photo_data', ''), now_str, now_str
    ))
    new_id = c.lastrowid
    conn.commit()
    conn.close()
    return jsonify({'message': 'Member record created successfully', 'id': new_id, 'member_code': member_code}), 201


@app.route('/api/members/<int:member_id>', methods=['PUT'])
def update_member(member_id):
    data = request.json or {}
    full_name = (data.get('full_name') or '').strip()
    if not full_name:
        return jsonify({'error': 'Full Name is required'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT id FROM members WHERE id = ?", (member_id,))
    if not c.fetchone():
        conn.close()
        return jsonify({'error': 'Member not found'}), 404

    dob = (data.get('dob') or '').strip()
    age = data.get('age')
    if (age is None or str(age).strip() == '') and dob:
        calc = calculate_age(dob)
        age = calc if calc is not None else 0
    else:
        try:
            age = int(age)
        except Exception:
            age = 0

    is_head = 1 if data.get('is_head') in [True, 1, '1', 'true'] else 0

    try:
        children_count = int(data.get('children_count') or 0)
    except Exception:
        children_count = 0

    try:
        total_family_members = int(data.get('total_family_members') or 1)
    except Exception:
        total_family_members = 1

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    c.execute('''
        UPDATE members SET
            family_id = ?, is_head = ?, relation_to_head = ?,
            full_name = ?, full_name_hi = ?, gender = ?, dob = ?, age = ?,
            blood_group = ?, marital_status = ?, father_name = ?, mother_name = ?,
            spouse_name = ?, children_count = ?, total_family_members = ?,
            family_members_notes = ?, gotra_subcaste = ?, native_origin = ?,
            mohalla_area = ?, current_address = ?, permanent_address = ?,
            city = ?, pincode = ?, phone_primary = ?, phone_whatsapp = ?, email = ?,
            education_level = ?, education_stream = ?, institution_name = ?,
            employment_status = ?, occupation_title = ?, business_or_employer = ?,
            work_address = ?, annual_income_range = ?, skills_specialization = ?,
            aadhaar_last4 = ?, bhamashah_jan_aadhaar = ?, samaj_role = ?,
            special_notes = ?, photo_data = ?, updated_at = ?
        WHERE id = ?
    ''', (
        data.get('family_id', ''), is_head, data.get('relation_to_head', 'Self'),
        full_name, data.get('full_name_hi', ''), data.get('gender', 'Male'), dob, age,
        data.get('blood_group', ''), data.get('marital_status', 'Single'),
        data.get('father_name', ''), data.get('mother_name', ''), data.get('spouse_name', ''),
        children_count, total_family_members, data.get('family_members_notes', ''),
        data.get('gotra_subcaste', ''), data.get('native_origin', 'Khanda Falsa, Jodhpur'),
        data.get('mohalla_area', 'Khanda Falsa'), data.get('current_address', ''),
        data.get('permanent_address', ''), data.get('city', 'Jodhpur'), data.get('pincode', '342001'),
        data.get('phone_primary', ''), data.get('phone_whatsapp', ''), data.get('email', ''),
        data.get('education_level', ''), data.get('education_stream', ''), data.get('institution_name', ''),
        data.get('employment_status', ''), data.get('occupation_title', ''),
        data.get('business_or_employer', ''), data.get('work_address', ''),
        data.get('annual_income_range', ''), data.get('skills_specialization', ''),
        data.get('aadhaar_last4', ''), data.get('bhamashah_jan_aadhaar', ''),
        data.get('samaj_role', 'Member'), data.get('special_notes', ''),
        data.get('photo_data', ''), now_str, member_id
    ))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Member updated successfully'})


@app.route('/api/members/<int:member_id>', methods=['DELETE'])
def delete_member(member_id):
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM members WHERE id = ?", (member_id,))
    conn.commit()
    conn.close()
    return jsonify({'message': 'Member record deleted'})


@app.route('/api/stats', methods=['GET'])
def get_stats():
    conn = get_db()
    c = conn.cursor()

    c.execute("SELECT COUNT(*) as total FROM members")
    total_members = c.fetchone()['total']

    c.execute("SELECT COUNT(DISTINCT family_id) as families FROM members")
    total_families = c.fetchone()['families']

    c.execute("SELECT gender, COUNT(*) as count FROM members GROUP BY gender")
    gender_counts = {r['gender'] or 'Other': r['count'] for r in c.fetchall()}

    c.execute("SELECT marital_status, COUNT(*) as count FROM members GROUP BY marital_status")
    marital_counts = {r['marital_status'] or 'Unspecified': r['count'] for r in c.fetchall()}

    c.execute("SELECT education_level, COUNT(*) as count FROM members GROUP BY education_level")
    education_counts = {r['education_level'] or 'Unspecified': r['count'] for r in c.fetchall()}

    c.execute("SELECT employment_status, COUNT(*) as count FROM members GROUP BY employment_status")
    employment_counts = {r['employment_status'] or 'Unspecified': r['count'] for r in c.fetchall()}

    c.execute("SELECT mohalla_area, COUNT(*) as count FROM members GROUP BY mohalla_area ORDER BY count DESC LIMIT 8")
    mohalla_counts = {r['mohalla_area'] or 'Khanda Falsa': r['count'] for r in c.fetchall()}

    c.execute("SELECT blood_group, COUNT(*) as count FROM members WHERE blood_group != '' GROUP BY blood_group")
    blood_counts = {r['blood_group']: r['count'] for r in c.fetchall()}

    c.execute("""
        SELECT
            SUM(CASE WHEN age < 18 THEN 1 ELSE 0 END) as under_18,
            SUM(CASE WHEN age >= 18 AND age <= 35 THEN 1 ELSE 0 END) as youth_18_35,
            SUM(CASE WHEN age >= 36 AND age <= 59 THEN 1 ELSE 0 END) as adult_36_59,
            SUM(CASE WHEN age >= 60 THEN 1 ELSE 0 END) as senior_60_plus
        FROM members
    """)
    age_row = dict(c.fetchone())

    c.execute("""
        SELECT
            COUNT(*) as total_pop,
            SUM(CASE WHEN education_level != '' AND education_level != 'Illiterate / Non-Formal' THEN 1 ELSE 0 END) as literate_total,
            SUM(CASE WHEN gender = 'Male' THEN 1 ELSE 0 END) as male_total,
            SUM(CASE WHEN gender = 'Male' AND education_level != '' AND education_level != 'Illiterate / Non-Formal' THEN 1 ELSE 0 END) as male_literate,
            SUM(CASE WHEN gender = 'Female' THEN 1 ELSE 0 END) as female_total,
            SUM(CASE WHEN gender = 'Female' AND education_level != '' AND education_level != 'Illiterate / Non-Formal' THEN 1 ELSE 0 END) as female_literate
        FROM members
    """)
    lit_row = dict(c.fetchone())
    total_pop = lit_row.get('total_pop') or 0
    literate_total = lit_row.get('literate_total') or 0
    illiterate_total = max(0, total_pop - literate_total)
    male_total = lit_row.get('male_total') or 0
    male_literate = lit_row.get('male_literate') or 0
    female_total = lit_row.get('female_total') or 0
    female_literate = lit_row.get('female_literate') or 0

    literacy_summary = {
        'literate_count': literate_total,
        'non_formal_or_unspecified': illiterate_total,
        'overall_rate': round((literate_total / total_pop) * 100, 1) if total_pop > 0 else 0.0,
        'male_rate': round((male_literate / male_total) * 100, 1) if male_total > 0 else 0.0,
        'female_rate': round((female_literate / female_total) * 100, 1) if female_total > 0 else 0.0,
        'male_literate': male_literate,
        'male_total': male_total,
        'female_literate': female_literate,
        'female_total': female_total
    }

    c.execute("""
        SELECT mohalla_area as name, COUNT(*) as count 
        FROM members 
        WHERE mohalla_area IS NOT NULL AND TRIM(mohalla_area) != ''
        GROUP BY mohalla_area 
        ORDER BY mohalla_area ASC
    """)
    all_areas = [dict(r) for r in c.fetchall()]

    c.execute("""
        SELECT family_id,
               MAX(CASE WHEN is_head = 1 THEN full_name ELSE NULL END) as head_name,
               MAX(mohalla_area) as mohalla_area,
               MAX(current_address) as current_address,
               MAX(gotra_subcaste) as gotra_subcaste,
               COUNT(*) as registered_count,
               MAX(total_family_members) as total_family_members
        FROM members
        GROUP BY family_id
        ORDER BY family_id ASC
    """)
    families = [dict(r) for r in c.fetchall()]

    conn.close()
    return jsonify({
        'total_members': total_members,
        'total_families': total_families,
        'gender_counts': gender_counts,
        'marital_counts': marital_counts,
        'education_counts': education_counts,
        'employment_counts': employment_counts,
        'mohalla_counts': mohalla_counts,
        'areas': all_areas,
        'blood_counts': blood_counts,
        'age_groups': {
            'Children (<18)': age_row.get('under_18') or 0,
            'Youth (18-35)': age_row.get('youth_18_35') or 0,
            'Adults (36-59)': age_row.get('adult_36_59') or 0,
            'Seniors (60+)': age_row.get('senior_60_plus') or 0
        },
        'literacy_summary': literacy_summary,
        'families': families
    })


@app.route('/api/export/csv', methods=['GET'])
@admin_required
def export_csv():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT
            member_code, family_id, is_head, relation_to_head,
            full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
            father_name, mother_name, spouse_name, children_count, total_family_members,
            family_members_notes, gotra_subcaste, native_origin, mohalla_area,
            current_address, permanent_address, city, pincode,
            phone_primary, phone_whatsapp, email,
            education_level, education_stream, institution_name,
            employment_status, occupation_title, business_or_employer, work_address,
            annual_income_range, skills_specialization,
            aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, created_at
        FROM members
        ORDER BY family_id ASC, is_head DESC, id ASC
    """)
    rows = c.fetchall()
    conn.close()

    output = io.StringIO()
    output.write('\ufeff')  # UTF-8 BOM for Excel Hindi support
    writer = csv.writer(output)
    headers = [
        'Member Code', 'Family ID', 'Is Head of Family (1/0)', 'Relation to Head',
        'Full Name (English)', 'Full Name (Hindi)', 'Gender', 'Date of Birth (YYYY-MM-DD)', 'Age',
        'Blood Group', 'Marital Status', 'Father Name', 'Mother Name', 'Husband / Wife Name',
        'Children Count', 'Total Family Members', 'Family Members List / Notes',
        'Gotra / Sub-caste', 'Native Origin', 'Mohalla / Area', 'Current Address',
        'Permanent Address', 'City', 'Pincode', 'Primary Phone', 'WhatsApp Number', 'Email',
        'Education Level', 'Degree / Stream', 'School / College / University',
        'Employment Status', 'Occupation / Designation', 'Business / Employer Name',
        'Work Address', 'Annual Income Range', 'Skills / Craft Specialization',
        'Aadhaar Last 4 Digits', 'Jan Aadhaar / Samaj ID', 'Samaj Role', 'Special Notes', 'Created At'
    ]
    writer.writerow(headers)
    for r in rows:
        writer.writerow(list(r))

    mem = io.BytesIO()
    mem.write(output.getvalue().encode('utf-8'))
    mem.seek(0)
    filename = f"Chheepa_Samaj_Khanda_Falsa_Jodhpur_{datetime.now().strftime('%Y%m%d')}.csv"
    return send_file(mem, mimetype='text/csv', as_attachment=True, download_name=filename)


@app.route('/api/export/excel', methods=['GET'])
@admin_required
def export_excel():
    conn = get_db()
    c = conn.cursor()
    c.execute("""
        SELECT
            member_code, family_id, is_head, relation_to_head,
            full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
            father_name, mother_name, spouse_name, children_count, total_family_members,
            family_members_notes, gotra_subcaste, native_origin, mohalla_area,
            current_address, permanent_address, city, pincode,
            phone_primary, phone_whatsapp, email,
            education_level, education_stream, institution_name,
            employment_status, occupation_title, business_or_employer, work_address,
            annual_income_range, skills_specialization,
            aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, created_at
        FROM members
        ORDER BY family_id ASC, is_head DESC, id ASC
    """)
    rows = c.fetchall()
    conn.close()

    headers = [
        'Member Code', 'Family ID', 'Is Head (1/0)', 'Relation to Head',
        'Full Name (English)', 'Full Name (Hindi)', 'Gender', 'Date of Birth', 'Age',
        'Blood Group', 'Marital Status', 'Father Name', 'Mother Name', 'Spouse Name',
        'Children Count', 'Total Family Members', 'Family Notes',
        'Gotra / Sub-caste', 'Native Origin', 'Mohalla / Area', 'Current Address',
        'Permanent Address', 'City', 'Pincode', 'Primary Phone', 'WhatsApp Number', 'Email',
        'Education Level', 'Degree / Stream', 'Institution Name',
        'Employment Status', 'Occupation / Title', 'Business / Employer',
        'Work Address', 'Annual Income Range', 'Skills / Specialization',
        'Aadhaar Last 4', 'Jan Aadhaar / Samaj ID', 'Samaj Role', 'Special Notes', 'Created At'
    ]

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Samaj Census'

    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws.append(headers)
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    row_font = Font(name='Arial', size=10)
    for r in rows:
        row_values = list(r)
        ws.append(row_values)
        row_num = ws.max_row
        for col_num in range(1, len(row_values) + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.font = row_font
            cell.border = thin_border
            if col_num in [1, 2, 7, 8, 9, 10, 24, 25, 26]:
                cell.alignment = Alignment(horizontal='center', vertical='center')

    # Auto adjust column widths
    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(12, min(max_len + 3, 40))

    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    filename = f"Chheepa_Samaj_Khanda_Falsa_Jodhpur_{datetime.now().strftime('%Y%m%d')}.xlsx"
    return send_file(
        mem,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name=filename
    )


@app.route('/api/export/json', methods=['GET'])
@admin_required
def export_json():
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM members ORDER BY family_id ASC, is_head DESC")
    rows = [dict(r) for r in c.fetchall()]
    conn.close()

    mem = io.BytesIO()
    mem.write(json.dumps({'samaj': 'Chheepa Samaj Khanda Falsa, Jodhpur', 'exported_at': datetime.now().isoformat(), 'members': rows}, ensure_ascii=False, indent=2).encode('utf-8'))
    mem.seek(0)
    filename = f"Chheepa_Samaj_Khanda_Falsa_Backup_{datetime.now().strftime('%Y%m%d')}.json"
    return send_file(mem, mimetype='application/json', as_attachment=True, download_name=filename)


@app.route('/api/import/json', methods=['POST'])
@admin_required
def import_json():
    payload = request.json or {}
    members_list = payload.get('members', [])
    if not isinstance(members_list, list) or not members_list:
        return jsonify({'error': 'Invalid or empty members array in JSON'}), 400

    conn = get_db()
    c = conn.cursor()
    imported = 0
    for item in members_list:
        full_name = (item.get('full_name') or '').strip()
        if not full_name:
            continue
        member_code = item.get('member_code') or generate_next_code(conn)
        c.execute("SELECT id FROM members WHERE member_code = ?", (member_code,))
        if c.fetchone():
            member_code = f"{member_code}-{datetime.now().strftime('%S%f')[:4]}"

        c.execute('''
            INSERT INTO members (
                member_code, family_id, is_head, relation_to_head,
                full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
                father_name, mother_name, spouse_name, children_count, total_family_members,
                family_members_notes, gotra_subcaste, native_origin, mohalla_area,
                current_address, permanent_address, city, pincode,
                phone_primary, phone_whatsapp, email,
                education_level, education_stream, institution_name,
                employment_status, occupation_title, business_or_employer, work_address,
                annual_income_range, skills_specialization,
                aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, photo_data
            ) VALUES (
                ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
            )
        ''', (
            member_code, item.get('family_id', 'FAM-KF-IMP'), int(item.get('is_head') or 0),
            item.get('relation_to_head', 'Self'), full_name, item.get('full_name_hi', ''),
            item.get('gender', 'Male'), item.get('dob', ''), int(item.get('age') or 0),
            item.get('blood_group', ''), item.get('marital_status', 'Single'),
            item.get('father_name', ''), item.get('mother_name', ''), item.get('spouse_name', ''),
            int(item.get('children_count') or 0), int(item.get('total_family_members') or 1),
            item.get('family_members_notes', ''), item.get('gotra_subcaste', ''),
            item.get('native_origin', 'Khanda Falsa, Jodhpur'), item.get('mohalla_area', 'Khanda Falsa'),
            item.get('current_address', ''), item.get('permanent_address', ''),
            item.get('city', 'Jodhpur'), item.get('pincode', '342001'),
            item.get('phone_primary', ''), item.get('phone_whatsapp', ''), item.get('email', ''),
            item.get('education_level', ''), item.get('education_stream', ''), item.get('institution_name', ''),
            item.get('employment_status', ''), item.get('occupation_title', ''),
            item.get('business_or_employer', ''), item.get('work_address', ''),
            item.get('annual_income_range', ''), item.get('skills_specialization', ''),
            item.get('aadhaar_last4', ''), item.get('bhamashah_jan_aadhaar', ''),
            item.get('samaj_role', 'Member'), item.get('special_notes', ''), item.get('photo_data', '')
        ))
        imported += 1

    conn.commit()
    conn.close()
    return jsonify({'message': f'Successfully imported {imported} member records.', 'imported': imported})


# ==========================================
# ADMIN AUTHENTICATION ROUTES
# ==========================================
@app.route('/admin/login')
def admin_login_page():
    if session.get('admin_id'):
        return redirect(url_for('admin_dashboard'))
    return render_template('admin_login.html')


@app.route('/admin')
@admin_required
def admin_dashboard():
    return render_template('admin.html')


@app.route('/admin/api/login', methods=['POST'])
def admin_api_login():
    data = request.json or request.form or {}
    username = (data.get('username') or '').strip()
    password = data.get('password') or ''

    if not username or not password:
        return jsonify({'error': 'Username and password are required'}), 400

    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM admin_users WHERE username = ?", (username,))
    admin = c.fetchone()

    if not admin or not check_password_hash(admin['password_hash'], password):
        conn.close()
        return jsonify({'error': 'Invalid username or password'}), 401

    session['admin_id'] = admin['id']
    session['admin_username'] = admin['username']
    session['admin_name'] = admin['full_name']

    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
    c.execute("UPDATE admin_users SET last_login = ? WHERE id = ?", (now_str, admin['id']))
    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'message': 'Authentication successful',
        'username': admin['username'],
        'full_name': admin['full_name']
    })


@app.route('/admin/api/logout', methods=['POST'])
def admin_api_logout():
    session.clear()
    return jsonify({'success': True, 'message': 'Logged out successfully'})


@app.route('/admin/api/session', methods=['GET'])
def admin_api_session():
    if session.get('admin_id'):
        conn = get_db()
        c = conn.cursor()
        c.execute("SELECT id, username, full_name, email, last_login FROM admin_users WHERE id = ?", (session['admin_id'],))
        admin = c.fetchone()
        conn.close()
        if admin:
            return jsonify({
                'logged_in': True,
                'username': admin['username'],
                'full_name': admin['full_name'],
                'email': admin['email'],
                'last_login': admin['last_login']
            })
    return jsonify({'logged_in': False})


@app.route('/admin/api/change-password', methods=['POST'])
@admin_required
def admin_api_change_password():
    data = request.json or {}
    old_password = data.get('old_password') or ''
    new_password = data.get('new_password') or ''

    if len(new_password) < 4:
        return jsonify({'error': 'New password must be at least 4 characters long'}), 400

    admin_id = session.get('admin_id')
    conn = get_db()
    c = conn.cursor()
    c.execute("SELECT * FROM admin_users WHERE id = ?", (admin_id,))
    admin = c.fetchone()

    if not admin or not check_password_hash(admin['password_hash'], old_password):
        conn.close()
        return jsonify({'error': 'Current password does not match'}), 400

    new_hash = generate_password_hash(new_password)
    c.execute("UPDATE admin_users SET password_hash = ? WHERE id = ?", (new_hash, admin_id))
    conn.commit()
    conn.close()

    return jsonify({'success': True, 'message': 'Admin password updated successfully'})


# ==========================================
# DATA INGESTION ENGINE (CSV, EXCEL, JSON)
# ==========================================
HEADER_MAPPINGS = {
    # full_name
    'full_name': 'full_name', 'name': 'full_name', 'full name': 'full_name',
    'fullname': 'full_name', 'member_name': 'full_name', 'person_name': 'full_name',
    'full name (english)': 'full_name', 'name (english)': 'full_name',
    'नाम': 'full_name', 'पूरा नाम': 'full_name', 'सदस्य का नाम': 'full_name',
    'नाम (अंग्रेजी)': 'full_name',

    # full_name_hi
    'full_name_hi': 'full_name_hi', 'name_hi': 'full_name_hi', 'name_hindi': 'full_name_hi',
    'hindi_name': 'full_name_hi', 'hindi name': 'full_name_hi', 'full name (hindi)': 'full_name_hi',
    'name (hindi)': 'full_name_hi', 'नाम (हिंदी)': 'full_name_hi', 'हिंदी नाम': 'full_name_hi',

    # member_code
    'member_code': 'member_code', 'code': 'member_code', 'member id': 'member_code',
    'member_id': 'member_code', 'id_code': 'member_code', 'सदस्य कोड': 'member_code',

    # family_id
    'family_id': 'family_id', 'family id': 'family_id', 'family no': 'family_id',
    'family_no': 'family_id', 'fam_id': 'family_id', 'family code': 'family_id',
    'परिवार आईडी': 'family_id', 'परिवार संख्या': 'family_id',

    # is_head
    'is_head': 'is_head', 'head': 'is_head', 'is head': 'is_head',
    'is head of family': 'is_head', 'is head of family (1/0)': 'is_head',
    'मुखिया': 'is_head', 'परिवार मुखिया': 'is_head', 'मुखिया (हाँ/ना)': 'is_head',

    # relation_to_head
    'relation_to_head': 'relation_to_head', 'relation to head': 'relation_to_head', 'relation': 'relation_to_head',
    'relationship': 'relation_to_head', 'मुखिया से संबंध': 'relation_to_head', 'संबंध': 'relation_to_head',

    # gender
    'gender': 'gender', 'sex': 'gender', 'लिंग': 'gender',

    # dob
    'dob': 'dob', 'date of birth': 'dob', 'birth date': 'dob', 'birthdate': 'dob',
    'date of birth (yyyy-mm-dd)': 'dob', 'जन्म तिथि': 'dob', 'जन्म तारीख': 'dob',

    # age
    'age': 'age', 'उम्र': 'age', 'आयु': 'age',

    # blood_group
    'blood_group': 'blood_group', 'blood group': 'blood_group', 'blood': 'blood_group',
    'रक्त समूह': 'blood_group', 'ब्लड ग्रुप': 'blood_group',

    # marital_status
    'marital_status': 'marital_status', 'marital status': 'marital_status',
    'marital': 'marital_status', 'वैवाहिक स्थिति': 'marital_status',

    # father_name
    'father_name': 'father_name', 'father': 'father_name', 'father name': 'father_name',
    'पिता का नाम': 'father_name', 'पिता': 'father_name',

    # mother_name
    'mother_name': 'mother_name', 'mother': 'mother_name', 'mother name': 'mother_name',
    'माता का नाम': 'mother_name', 'माता': 'mother_name',

    # spouse_name
    'spouse_name': 'spouse_name', 'spouse': 'spouse_name', 'spouse name': 'spouse_name',
    'husband / wife name': 'spouse_name', 'husband name': 'spouse_name', 'wife name': 'spouse_name',
    'पति / पत्नी का नाम': 'spouse_name', 'पति का नाम': 'spouse_name', 'पत्नी का नाम': 'spouse_name',

    # children_count
    'children_count': 'children_count', 'children': 'children_count', 'children count': 'children_count',
    'no of children': 'children_count', 'बच्चों की संख्या': 'children_count',

    # total_family_members
    'total_family_members': 'total_family_members', 'total family members': 'total_family_members',
    'family members count': 'total_family_members', 'परिवार के कुल सदस्य': 'total_family_members',

    # family_members_notes
    'family_members_notes': 'family_members_notes', 'family members list / notes': 'family_members_notes',
    'family notes': 'family_members_notes', 'परिवार सदस्य विवरण': 'family_members_notes',

    # gotra_subcaste
    'gotra_subcaste': 'gotra_subcaste', 'gotra': 'gotra_subcaste', 'subcaste': 'gotra_subcaste',
    'gotra / sub-caste': 'gotra_subcaste', 'गोत्र': 'gotra_subcaste', 'गोत्र / उपजाति': 'gotra_subcaste',

    # native_origin
    'native_origin': 'native_origin', 'native': 'native_origin', 'origin': 'native_origin',
    'मूल निवास': 'native_origin', 'मूल वतन': 'native_origin',

    # mohalla_area
    'mohalla_area': 'mohalla_area', 'mohalla': 'mohalla_area', 'area': 'mohalla_area',
    'mohalla / area': 'mohalla_area', 'मोहल्ला': 'mohalla_area', 'मोहल्ला / क्षेत्र': 'mohalla_area',

    # current_address
    'current_address': 'current_address', 'address': 'current_address', 'present address': 'current_address',
    'वर्तमान पता': 'current_address', 'पता': 'current_address',

    # permanent_address
    'permanent_address': 'permanent_address', 'स्थायी पता': 'permanent_address',

    # city
    'city': 'city', 'शहर': 'city',

    # pincode
    'pincode': 'pincode', 'pin': 'pincode', 'postal code': 'pincode', 'पिनकोड': 'pincode',

    # phone_primary
    'phone_primary': 'phone_primary', 'phone': 'phone_primary', 'mobile': 'phone_primary',
    'primary phone': 'phone_primary', 'mobile number': 'phone_primary',
    'मोबाइल': 'phone_primary', 'मोबाइल नंबर': 'phone_primary', 'फोन': 'phone_primary',

    # phone_whatsapp
    'phone_whatsapp': 'phone_whatsapp', 'whatsapp': 'phone_whatsapp', 'whatsapp number': 'phone_whatsapp',
    'व्हाट्सएप': 'phone_whatsapp', 'व्हाट्सएप नंबर': 'phone_whatsapp',

    # email
    'email': 'email', 'email id': 'email', 'ईमेल': 'email',

    # education_level
    'education_level': 'education_level', 'education': 'education_level', 'qualification': 'education_level',
    'शिक्षा': 'education_level', 'शिक्षा का स्तर': 'education_level',

    # education_stream
    'education_stream': 'education_stream', 'degree / stream': 'education_stream', 'stream': 'education_stream',
    'degree': 'education_stream', 'कोर्स': 'education_stream', 'डिग्री': 'education_stream',

    # institution_name
    'institution_name': 'institution_name', 'school / college / university': 'institution_name',
    'institution': 'institution_name', 'college': 'institution_name', 'school': 'institution_name', 'संस्थान': 'institution_name',

    # employment_status
    'employment_status': 'employment_status', 'employment': 'employment_status',
    'रोजगार': 'employment_status', 'रोजगार की स्थिति': 'employment_status',

    # occupation_title
    'occupation_title': 'occupation_title', 'occupation': 'occupation_title',
    'occupation / designation': 'occupation_title', 'designation': 'occupation_title',
    'व्यवसाय': 'occupation_title', 'पद': 'occupation_title', 'व्यवसाय / पद': 'occupation_title',

    # business_or_employer
    'business_or_employer': 'business_or_employer', 'business / employer name': 'business_or_employer',
    'employer': 'business_or_employer', 'company': 'business_or_employer', 'फर्म का नाम': 'business_or_employer',

    # work_address
    'work_address': 'work_address', 'कार्यस्थल पता': 'work_address',

    # annual_income_range
    'annual_income_range': 'annual_income_range', 'annual income': 'annual_income_range',
    'income': 'annual_income_range', 'वार्षिक आय': 'annual_income_range',

    # skills_specialization
    'skills_specialization': 'skills_specialization', 'skills / craft specialization': 'skills_specialization',
    'skills': 'skills_specialization', 'हुनर': 'skills_specialization', 'विशेषज्ञता': 'skills_specialization',

    # aadhaar_last4
    'aadhaar_last4': 'aadhaar_last4', 'aadhaar last 4 digits': 'aadhaar_last4',
    'aadhaar': 'aadhaar_last4', 'आधार': 'aadhaar_last4', 'आधार अंतिम 4 अंक': 'aadhaar_last4',

    # bhamashah_jan_aadhaar
    'bhamashah_jan_aadhaar': 'bhamashah_jan_aadhaar', 'jan aadhaar / samaj id': 'bhamashah_jan_aadhaar',
    'jan aadhaar': 'bhamashah_jan_aadhaar', 'जन आधार': 'bhamashah_jan_aadhaar',

    # samaj_role
    'samaj_role': 'samaj_role', 'role': 'samaj_role', 'samaj role': 'samaj_role',
    'समाज पद': 'samaj_role', 'समाज में भूमिका': 'samaj_role',

    # special_notes
    'special_notes': 'special_notes', 'notes': 'special_notes', 'special notes': 'special_notes',
    'टिप्पणी': 'special_notes', 'विशेष विवरण': 'special_notes'
}


def clean_cell_value(val):
    if val is None:
        return ""
    if isinstance(val, (datetime, date)):
        return val.strftime('%Y-%m-%d')
    if isinstance(val, float):
        if val.is_integer():
            return str(int(val))
        return str(val)
    return str(val).strip()


def parse_file_to_records(file_storage):
    filename = file_storage.filename or ''
    ext = filename.split('.')[-1].lower() if '.' in filename else ''
    content = file_storage.read()

    raw_rows = []

    if ext == 'csv':
        # Decode CSV handling UTF-8 with BOM and latin-1 fallback
        text = None
        for enc in ['utf-8-sig', 'utf-8', 'latin-1', 'cp1252']:
            try:
                text = content.decode(enc)
                break
            except Exception:
                continue

        if text is None:
            raise ValueError("Unable to decode CSV file. Please ensure it is saved in UTF-8 format.")

        # Detect delimiter (comma, tab, semicolon)
        delimiter = ','
        first_line = text.splitlines()[0] if text.splitlines() else ''
        if '\t' in first_line and first_line.count('\t') > first_line.count(','):
            delimiter = '\t'
        elif ';' in first_line and first_line.count(';') > first_line.count(','):
            delimiter = ';'

        reader = csv.reader(io.StringIO(text), delimiter=delimiter)
        rows_list = list(reader)
        if not rows_list:
            return [], [], [], 0
        raw_headers = [h.strip() for h in rows_list[0]]
        for row in rows_list[1:]:
            if not any(row):
                continue
            row_dict = {}
            for idx, h in enumerate(raw_headers):
                if idx < len(row):
                    row_dict[h] = row[idx].strip()
            raw_rows.append(row_dict)

    elif ext in ['xlsx', 'xls']:
        wb = openpyxl.load_workbook(io.BytesIO(content), data_only=True)
        ws = wb.active
        all_rows = list(ws.iter_rows(values_only=True))
        if not all_rows:
            return [], [], [], 0

        # Find first non-empty header row
        header_row_idx = 0
        while header_row_idx < len(all_rows) and not any(all_rows[header_row_idx]):
            header_row_idx += 1

        if header_row_idx >= len(all_rows):
            return [], [], [], 0

        raw_headers = [clean_cell_value(h) for h in all_rows[header_row_idx]]
        for row in all_rows[header_row_idx + 1:]:
            if not any(row):
                continue
            row_dict = {}
            for idx, h in enumerate(raw_headers):
                if idx < len(row) and h:
                    row_dict[h] = clean_cell_value(row[idx])
            raw_rows.append(row_dict)

    elif ext == 'json':
        text = content.decode('utf-8', errors='ignore')
        parsed_json = json.loads(text)
        if isinstance(parsed_json, list):
            raw_rows = parsed_json
        elif isinstance(parsed_json, dict):
            if 'members' in parsed_json and isinstance(parsed_json['members'], list):
                raw_rows = parsed_json['members']
            elif 'data' in parsed_json and isinstance(parsed_json['data'], list):
                raw_rows = parsed_json['data']
            elif 'rows' in parsed_json and isinstance(parsed_json['rows'], list):
                raw_rows = parsed_json['rows']
            else:
                raw_rows = [parsed_json]
        else:
            raise ValueError("Invalid JSON format. Expected JSON array or object with 'members' key.")
        raw_headers = list(raw_rows[0].keys()) if raw_rows else []
    else:
        raise ValueError(f"Unsupported file format: '.{ext}'. Supported formats are CSV, Excel (.xlsx/.xls), and JSON.")

    # Standardize headers into schema keys
    standardized_records = []
    mapped_fields_set = set()

    for row in raw_rows:
        std_row = {}
        for key, val in row.items():
            norm_key = str(key).strip().lower()
            target_field = HEADER_MAPPINGS.get(norm_key)
            if not target_field:
                target_field = HEADER_MAPPINGS.get(norm_key.replace(' ', '_'))
            if not target_field:
                target_field = HEADER_MAPPINGS.get(norm_key.replace('_', ' '))
            if not target_field:
                target_field = norm_key
            std_row[target_field] = clean_cell_value(val)
            if target_field in HEADER_MAPPINGS.values():
                mapped_fields_set.add(target_field)
        standardized_records.append(std_row)

    return standardized_records, sorted(list(mapped_fields_set)), raw_headers, len(standardized_records)


def normalize_record(item, conn):
    full_name = (item.get('full_name') or '').strip()
    extracted_gotra = ''
    extracted_phone = ''
    extracted_mohalla = ''

    if not full_name:
        # Pass 1: Look for hyphen-separated format e.g. 'Name - Gotra - Mobile'
        for k, v in item.items():
            s = str(v or '').strip()
            if '-' in s:
                parts = [p.strip() for p in s.split('-') if p.strip()]
                if parts and not parts[0].replace(' ', '').isdigit() and len(parts[0]) >= 2:
                    full_name = parts[0]
                    if len(parts) > 1 and not parts[1].replace(' ', '').isdigit():
                        extracted_gotra = parts[1]
                    if len(parts) > 2 and parts[2].replace(' ', '').isdigit():
                        extracted_phone = parts[2].replace(' ', '')
                    if 'गली' in k or 'bazaar' in k.lower() or 'chowk' in k.lower() or 'area' in k.lower():
                        extracted_mohalla = k.strip()
                    break

        # Pass 2: Look for non-numeric candidate string if Pass 1 didn't match
        if not full_name:
            for k, v in item.items():
                s = str(v or '').strip()
                if not s or s.replace('.', '').replace(' ', '').isdigit() or len(s) < 3:
                    continue
                full_name = s
                if 'गली' in k or 'bazaar' in k.lower() or 'chowk' in k.lower() or 'area' in k.lower():
                    extracted_mohalla = k.strip()
                break

    if not full_name:
        return None, "Full name is required"

    member_code = (item.get('member_code') or '').strip()
    family_id = (item.get('family_id') or '').strip()
    if not family_id:
        family_id = f"FAM-KF-{datetime.now().strftime('%y%m%d%H%M')}"

    dob = (item.get('dob') or '').strip()
    if len(dob) > 10:
        dob = dob[:10]

    age_val = item.get('age')
    if (age_val is None or str(age_val).strip() == '' or str(age_val) == '0') and dob:
        calc = calculate_age(dob)
        age = calc if calc is not None else 0
    else:
        try:
            age = int(float(age_val))
        except Exception:
            age = 0

    is_head_raw = str(item.get('is_head', '')).strip().lower()
    is_head = 1 if is_head_raw in ['1', 'true', 'yes', 'y', 'हाँ', 'head', 'self'] else 0

    try:
        children_count = int(float(item.get('children_count') or 0))
    except Exception:
        children_count = 0

    try:
        total_family_members = int(float(item.get('total_family_members') or 1))
    except Exception:
        total_family_members = 1

    gender = (item.get('gender') or 'Male').strip().capitalize()
    if gender.startswith('M') or 'पुरुष' in gender:
        gender = 'Male'
    elif gender.startswith('F') or 'महिला' in gender or 'स्त्री' in gender:
        gender = 'Female'
    elif gender not in ['Male', 'Female', 'Other']:
        gender = 'Male'

    phone_primary = (item.get('phone_primary') or extracted_phone or '').strip()
    if not phone_primary:
        for v in item.values():
            s = str(v or '').replace(' ', '').strip()
            if len(s) == 10 and s.isdigit():
                phone_primary = s
                break

    gotra_subcaste = (item.get('gotra_subcaste') or extracted_gotra or '').strip()
    mohalla_area = ' '.join((item.get('mohalla_area') or extracted_mohalla or 'Khanda Falsa').split())
    special_notes = (item.get('special_notes') or item.get('remark') or item.get('remarks') or '').strip()

    return {
        'member_code': member_code,
        'family_id': family_id,
        'is_head': is_head,
        'relation_to_head': item.get('relation_to_head') or ('Self (Head)' if is_head else 'Member'),
        'full_name': full_name,
        'full_name_hi': item.get('full_name_hi') or (full_name if any(ord(c) >= 2304 and ord(c) <= 2431 for c in full_name) else ''),
        'gender': gender,
        'dob': dob,
        'age': age,
        'blood_group': item.get('blood_group') or '',
        'marital_status': item.get('marital_status') or 'Single',
        'father_name': item.get('father_name') or '',
        'mother_name': item.get('mother_name') or '',
        'spouse_name': item.get('spouse_name') or '',
        'children_count': children_count,
        'total_family_members': total_family_members,
        'family_members_notes': item.get('family_members_notes') or '',
        'gotra_subcaste': gotra_subcaste,
        'native_origin': item.get('native_origin') or 'Khanda Falsa, Jodhpur',
        'mohalla_area': mohalla_area,
        'current_address': item.get('current_address') or '',
        'permanent_address': item.get('permanent_address') or '',
        'city': item.get('city') or 'Jodhpur',
        'pincode': item.get('pincode') or '342001',
        'phone_primary': phone_primary,
        'phone_whatsapp': item.get('phone_whatsapp') or phone_primary,
        'email': item.get('email') or '',
        'education_level': item.get('education_level') or '',
        'education_stream': item.get('education_stream') or '',
        'institution_name': item.get('institution_name') or '',
        'employment_status': item.get('employment_status') or '',
        'occupation_title': item.get('occupation_title') or '',
        'business_or_employer': item.get('business_or_employer') or '',
        'work_address': item.get('work_address') or '',
        'annual_income_range': item.get('annual_income_range') or '',
        'skills_specialization': item.get('skills_specialization') or '',
        'aadhaar_last4': item.get('aadhaar_last4') or '',
        'bhamashah_jan_aadhaar': item.get('bhamashah_jan_aadhaar') or '',
        'samaj_role': item.get('samaj_role') or 'Member',
        'special_notes': special_notes,
        'photo_data': item.get('photo_data') or ''
    }, None


@app.route('/admin/api/upload-preview', methods=['POST'])
@admin_required
def admin_upload_preview():
    if 'file' not in request.files:
        return jsonify({'error': 'No file was uploaded'}), 400

    file = request.files['file']
    if not file or not file.filename:
        return jsonify({'error': 'Empty file selected'}), 400

    try:
        records, mapped_fields, raw_headers, total = parse_file_to_records(file)
        return jsonify({
            'success': True,
            'total_rows': total,
            'mapped_fields': mapped_fields,
            'detected_headers': raw_headers,
            'sample_rows': records[:5]
        })
    except Exception as e:
        return jsonify({'error': f"Failed to preview file: {str(e)}"}), 400


@app.route('/admin/api/upload', methods=['POST'])
@admin_required
def admin_upload_commit():
    if 'file' not in request.files:
        return jsonify({'error': 'No file uploaded'}), 400

    file = request.files['file']
    mode = (request.form.get('mode') or 'append').lower()

    if not file or not file.filename:
        return jsonify({'error': 'Empty file selected'}), 400

    try:
        records, mapped_fields, raw_headers, total = parse_file_to_records(file)
    except Exception as e:
        return jsonify({'error': f"Parsing error: {str(e)}"}), 400

    if not records:
        return jsonify({'error': 'Uploaded file has no data rows'}), 400

    conn = get_db()
    c = conn.cursor()
    now_str = datetime.now().strftime('%Y-%m-%d %H:%M:%S')

    if mode == 'replace':
        c.execute("DELETE FROM members")
        conn.commit()

    inserted = 0
    updated = 0
    skipped = 0
    errors = []

    for idx, item in enumerate(records):
        row_num = idx + 1
        rec, err = normalize_record(item, conn)
        if err:
            errors.append({'row': row_num, 'reason': err})
            skipped += 1
            continue

        member_code = rec['member_code']
        existing_id = None

        if mode == 'update':
            if member_code:
                c.execute("SELECT id FROM members WHERE member_code = ?", (member_code,))
                row = c.fetchone()
                if row:
                    existing_id = row['id']

            if not existing_id and rec['phone_primary']:
                c.execute("SELECT id FROM members WHERE full_name = ? AND phone_primary = ?", (rec['full_name'], rec['phone_primary']))
                row = c.fetchone()
                if row:
                    existing_id = row['id']

        if existing_id:
            # Update existing record
            c.execute('''
                UPDATE members SET
                    family_id = ?, is_head = ?, relation_to_head = ?,
                    full_name = ?, full_name_hi = ?, gender = ?, dob = ?, age = ?,
                    blood_group = ?, marital_status = ?, father_name = ?, mother_name = ?,
                    spouse_name = ?, children_count = ?, total_family_members = ?,
                    family_members_notes = ?, gotra_subcaste = ?, native_origin = ?,
                    mohalla_area = ?, current_address = ?, permanent_address = ?,
                    city = ?, pincode = ?, phone_primary = ?, phone_whatsapp = ?, email = ?,
                    education_level = ?, education_stream = ?, institution_name = ?,
                    employment_status = ?, occupation_title = ?, business_or_employer = ?,
                    work_address = ?, annual_income_range = ?, skills_specialization = ?,
                    aadhaar_last4 = ?, bhamashah_jan_aadhaar = ?, samaj_role = ?,
                    special_notes = ?, updated_at = ?
                WHERE id = ?
            ''', (
                rec['family_id'], rec['is_head'], rec['relation_to_head'],
                rec['full_name'], rec['full_name_hi'], rec['gender'], rec['dob'], rec['age'],
                rec['blood_group'], rec['marital_status'], rec['father_name'], rec['mother_name'],
                rec['spouse_name'], rec['children_count'], rec['total_family_members'],
                rec['family_members_notes'], rec['gotra_subcaste'], rec['native_origin'],
                rec['mohalla_area'], rec['current_address'], rec['permanent_address'],
                rec['city'], rec['pincode'], rec['phone_primary'], rec['phone_whatsapp'], rec['email'],
                rec['education_level'], rec['education_stream'], rec['institution_name'],
                rec['employment_status'], rec['occupation_title'], rec['business_or_employer'],
                rec['work_address'], rec['annual_income_range'], rec['skills_specialization'],
                rec['aadhaar_last4'], rec['bhamashah_jan_aadhaar'], rec['samaj_role'],
                rec['special_notes'], now_str, existing_id
            ))
            updated += 1
        else:
            # Insert as new record
            if not member_code:
                member_code = generate_next_code(conn)
            else:
                c.execute("SELECT id FROM members WHERE member_code = ?", (member_code,))
                if c.fetchone():
                    member_code = generate_next_code(conn)

            c.execute('''
                INSERT INTO members (
                    member_code, family_id, is_head, relation_to_head,
                    full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
                    father_name, mother_name, spouse_name, children_count, total_family_members,
                    family_members_notes, gotra_subcaste, native_origin, mohalla_area,
                    current_address, permanent_address, city, pincode,
                    phone_primary, phone_whatsapp, email,
                    education_level, education_stream, institution_name,
                    employment_status, occupation_title, business_or_employer, work_address,
                    annual_income_range, skills_specialization,
                    aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, photo_data,
                    created_at, updated_at
                ) VALUES (
                    ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
                )
            ''', (
                member_code, rec['family_id'], rec['is_head'], rec['relation_to_head'],
                rec['full_name'], rec['full_name_hi'], rec['gender'], rec['dob'], rec['age'],
                rec['blood_group'], rec['marital_status'], rec['father_name'], rec['mother_name'],
                rec['spouse_name'], rec['children_count'], rec['total_family_members'],
                rec['family_members_notes'], rec['gotra_subcaste'], rec['native_origin'],
                rec['mohalla_area'], rec['current_address'], rec['permanent_address'],
                rec['city'], rec['pincode'], rec['phone_primary'], rec['phone_whatsapp'], rec['email'],
                rec['education_level'], rec['education_stream'], rec['institution_name'],
                rec['employment_status'], rec['occupation_title'], rec['business_or_employer'],
                rec['work_address'], rec['annual_income_range'], rec['skills_specialization'],
                rec['aadhaar_last4'], rec['bhamashah_jan_aadhaar'], rec['samaj_role'],
                rec['special_notes'], rec['photo_data'], now_str, now_str
            ))
            inserted += 1

    conn.commit()
    conn.close()

    return jsonify({
        'success': True,
        'message': f"Import finished: {inserted} inserted, {updated} updated, {skipped} skipped.",
        'total_rows': total,
        'inserted': inserted,
        'updated': updated,
        'skipped': skipped,
        'errors': errors
    })


# ==========================================
# DOWNLOADABLE SAMPLE TEMPLATES
# ==========================================
SAMPLE_TEMPLATE_DATA = [
    [
        'CSJ-0001', 'FAM-KF-101', 1, 'Self (Head)',
        'Haji Mohammed Farooq Chheepa', 'हाजी मोहम्मद फारूक छीपा', 'Male', '1966-04-12', 60,
        'B+', 'Married', 'Late Abdul Rahim Chheepa', 'Fatima Begum', 'Rashida Begum',
        4, 6, 'Elder family member in Khanda Falsa', 'Solanki Chheepa', 'Khanda Falsa, Jodhpur',
        'Adaa Bazaar / Khanda Falsa Chowk', 'House 42, Near Jama Masjid, Khanda Falsa',
        'House 42, Khanda Falsa, Jodhpur', 'Jodhpur', '342001', '9829011223', '9829011223',
        'farooq.textiles@example.com', 'Graduate', 'B.Com', 'JNVU Jodhpur',
        'Business / Self-Employed', 'Traditional Handblock Artisan & Merchant',
        'Farooq Chheepa Textiles', 'Main Road Khanda Falsa', '5 - 10 Lakh',
        'Dabu & Bagru Handblock Printing', '4821', 'JAN-29384711', 'Elder Patron', 'Active Jamaat Member'
    ],
    [
        'CSJ-0002', 'FAM-KF-101', 0, 'Daughter-in-Law',
        'Dr. Amina Parveen Chheepa', 'डॉ. अमीना परवीन छीपा', 'Female', '1995-02-18', 31,
        'A+', 'Married', 'Mohammed Yusuf Chheepa', 'Salma Bano', 'Er. Imran Ahmed Chheepa',
        1, 6, 'Pediatrician', 'Bhati Chheepa', 'Sojat Gate, Jodhpur', 'Adaa Bazaar / Khanda Falsa Chowk',
        'House 42, Near Jama Masjid, Khanda Falsa', 'House 42, Khanda Falsa, Jodhpur',
        'Jodhpur', '342001', '9414199887', '9414199887', 'dr.amina@example.com',
        'Doctorate / Medical', 'MBBS, MD (Pediatrics)', 'Dr. S.N. Medical College Jodhpur',
        'Government Job', 'Medical Officer', 'Umaid Hospital, Jodhpur', 'Siwanchi Gate Rd, Jodhpur',
        'Above 10 Lakh', 'Pediatrics & Neonatal Care', '7741', 'JAN-29384711',
        'Medical Committee Lead', 'Free Health Camp Doctor'
    ]
]

TEMPLATE_HEADERS = [
    'Member Code', 'Family ID', 'Is Head (1/0)', 'Relation to Head',
    'Full Name (English)', 'Full Name (Hindi)', 'Gender', 'Date of Birth (YYYY-MM-DD)', 'Age',
    'Blood Group', 'Marital Status', 'Father Name', 'Mother Name', 'Spouse Name',
    'Children Count', 'Total Family Members', 'Family Notes',
    'Gotra / Sub-caste', 'Native Origin', 'Mohalla / Area', 'Current Address',
    'Permanent Address', 'City', 'Pincode', 'Primary Phone', 'WhatsApp Number', 'Email',
    'Education Level', 'Degree / Stream', 'Institution Name',
    'Employment Status', 'Occupation / Title', 'Business / Employer Name',
    'Work Address', 'Annual Income Range', 'Skills / Specialization',
    'Aadhaar Last 4 Digits', 'Jan Aadhaar / Samaj ID', 'Samaj Role', 'Special Notes'
]


@app.route('/admin/api/templates/csv')
def download_csv_template():
    output = io.StringIO()
    output.write('\ufeff')  # UTF-8 BOM
    writer = csv.writer(output)
    writer.writerow(TEMPLATE_HEADERS)
    for row in SAMPLE_TEMPLATE_DATA:
        writer.writerow(row)

    mem = io.BytesIO()
    mem.write(output.getvalue().encode('utf-8'))
    mem.seek(0)
    return send_file(mem, mimetype='text/csv', as_attachment=True, download_name='Chheepa_Samaj_Template.csv')


@app.route('/admin/api/templates/excel')
def download_excel_template():
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = 'Samaj Upload Template'

    header_font = Font(name='Arial', size=11, bold=True, color='FFFFFF')
    header_fill = PatternFill(start_color='1E3A8A', end_color='1E3A8A', fill_type='solid')
    thin_border = Border(
        left=Side(style='thin', color='CBD5E1'),
        right=Side(style='thin', color='CBD5E1'),
        top=Side(style='thin', color='CBD5E1'),
        bottom=Side(style='thin', color='CBD5E1')
    )

    ws.append(TEMPLATE_HEADERS)
    for col_num in range(1, len(TEMPLATE_HEADERS) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)

    for row in SAMPLE_TEMPLATE_DATA:
        ws.append(row)
        row_num = ws.max_row
        for col_num in range(1, len(row) + 1):
            cell = ws.cell(row=row_num, column=col_num)
            cell.border = thin_border

    for col in ws.columns:
        max_len = 0
        col_letter = get_column_letter(col[0].column)
        for cell in col:
            val_str = str(cell.value or '')
            if len(val_str) > max_len:
                max_len = len(val_str)
        ws.column_dimensions[col_letter].width = max(14, min(max_len + 3, 38))

    mem = io.BytesIO()
    wb.save(mem)
    mem.seek(0)
    return send_file(
        mem,
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
        as_attachment=True,
        download_name='Chheepa_Samaj_Template.xlsx'
    )


@app.route('/admin/api/templates/json')
def download_json_template():
    keys = [
        'member_code', 'family_id', 'is_head', 'relation_to_head',
        'full_name', 'full_name_hi', 'gender', 'dob', 'age',
        'blood_group', 'marital_status', 'father_name', 'mother_name', 'spouse_name',
        'children_count', 'total_family_members', 'family_members_notes',
        'gotra_subcaste', 'native_origin', 'mohalla_area', 'current_address',
        'permanent_address', 'city', 'pincode', 'phone_primary', 'phone_whatsapp', 'email',
        'education_level', 'education_stream', 'institution_name',
        'employment_status', 'occupation_title', 'business_or_employer',
        'work_address', 'annual_income_range', 'skills_specialization',
        'aadhaar_last4', 'bhamashah_jan_aadhaar', 'samaj_role', 'special_notes'
    ]
    sample_list = []
    for r in SAMPLE_TEMPLATE_DATA:
        sample_list.append(dict(zip(keys, r)))

    data = {
        'samaj': 'Chheepa Samaj Khanda Falsa, Jodhpur',
        'description': 'Sample JSON census template for admin upload',
        'members': sample_list
    }
    mem = io.BytesIO()
    mem.write(json.dumps(data, ensure_ascii=False, indent=2).encode('utf-8'))
    mem.seek(0)
    return send_file(mem, mimetype='application/json', as_attachment=True, download_name='Chheepa_Samaj_Template.json')


# ==========================================
# DATABASE MAINTENANCE & RESET ENDPOINTS
# ==========================================
@app.route('/admin/api/reset-data', methods=['POST'])
@admin_required
def admin_reset_data():
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM members")
    try:
        c.execute("DELETE FROM sqlite_sequence WHERE name = 'members'")
    except Exception:
        pass
    c.executemany('''
        INSERT INTO members (
            member_code, family_id, is_head, relation_to_head,
            full_name, full_name_hi, gender, dob, age, blood_group, marital_status,
            father_name, mother_name, spouse_name, children_count, total_family_members,
            family_members_notes, gotra_subcaste, native_origin, mohalla_area,
            current_address, permanent_address, city, pincode,
            phone_primary, phone_whatsapp, email,
            education_level, education_stream, institution_name,
            employment_status, occupation_title, business_or_employer, work_address,
            annual_income_range, skills_specialization,
            aadhaar_last4, bhamashah_jan_aadhaar, samaj_role, special_notes, photo_data
        ) VALUES (
            ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?
        )
    ''', SAMPLE_MEMBERS)
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'Database reset to original 6 sample records.'})


@app.route('/admin/api/clear-data', methods=['POST'])
@admin_required
def admin_clear_data():
    conn = get_db()
    c = conn.cursor()
    c.execute("DELETE FROM members")
    try:
        c.execute("DELETE FROM sqlite_sequence WHERE name = 'members'")
    except Exception:
        pass
    conn.commit()
    conn.close()
    return jsonify({'success': True, 'message': 'All member records deleted.'})


if __name__ == '__main__':
    init_db()
    host = os.environ.get('HOST', '0.0.0.0')
    port = int(os.environ.get('PORT', 5000))
    print("=" * 65)
    print("  CHHEEPA SAMAJ KHANDA FALSA, JODHPUR - CENSUS & DIRECTORY PORTAL")
    print(f"  Public Portal: http://{host}:{port}")
    print(f"  Admin Panel:   http://{host}:{port}/admin")
    print(f"  Admin Login:   http://{host}:{port}/admin/login")
    print("=" * 65)
    app.run(host=host, port=port, debug=False)
