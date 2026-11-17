from flask import Flask, request, jsonify, render_template, send_file
import sqlite3
import os
import io
import csv
import json
from datetime import datetime, date

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max upload

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
DB_PATH = os.path.join(BASE_DIR, 'chheepa_samaj.db')


def get_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row
    return conn


def calculate_age(dob_str):
    if not dob_str:
        return None
    try:
        dob = datetime.strptime(dob_str, '%Y-%m-%d').date()
        today = date.today()
        return today.year - dob.year - ((today.month, today.day) < (dob.month, dob.day))
    except Exception:
        return None


def init_db():
    conn = get_db()
    c = conn.cursor()
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

    # Check if database is empty; if so, seed realistic sample records for Khanda Falsa, Jodhpur
    c.execute('SELECT COUNT(*) as cnt FROM members')
    count = c.fetchone()['cnt']
    if count == 0:
        sample_members = [
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
        ''', sample_members)
        conn.commit()
    conn.close()


def generate_next_code(conn):
    c = conn.cursor()
    c.execute('SELECT MAX(id) as max_id FROM members')
    row = c.fetchone()
    next_id = (row['max_id'] or 0) + 1
    return f"CSJ-{next_id:04d}"


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
    mohalla_area = request.args.get('mohalla_area', '').strip()
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
        query += " AND mohalla_area LIKE ?"
        params.append(f"%{mohalla_area}%")
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
    c.execute(query, params)
    rows = [dict(r) for r in c.fetchall()]
    conn.close()
    return jsonify({'members': rows, 'count': len(rows)})


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
    # Also fetch other members of the same family
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
        member_code,
        family_id,
        is_head,
        data.get('relation_to_head', 'Self'),
        full_name,
        data.get('full_name_hi', ''),
        data.get('gender', 'Male'),
        dob,
        age,
        data.get('blood_group', ''),
        data.get('marital_status', 'Single'),
        data.get('father_name', ''),
        data.get('mother_name', ''),
        data.get('spouse_name', ''),
        children_count,
        total_family_members,
        data.get('family_members_notes', ''),
        data.get('gotra_subcaste', ''),
        data.get('native_origin', 'Khanda Falsa, Jodhpur'),
        data.get('mohalla_area', 'Khanda Falsa'),
        data.get('current_address', ''),
        data.get('permanent_address', ''),
        data.get('city', 'Jodhpur'),
        data.get('pincode', '342001'),
        data.get('phone_primary', ''),
        data.get('phone_whatsapp', ''),
        data.get('email', ''),
        data.get('education_level', ''),
        data.get('education_stream', ''),
        data.get('institution_name', ''),
        data.get('employment_status', ''),
        data.get('occupation_title', ''),
        data.get('business_or_employer', ''),
        data.get('work_address', ''),
        data.get('annual_income_range', ''),
        data.get('skills_specialization', ''),
        data.get('aadhaar_last4', ''),
        data.get('bhamashah_jan_aadhaar', ''),
        data.get('samaj_role', 'Member'),
        data.get('special_notes', ''),
        data.get('photo_data', ''),
        now_str,
        now_str
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
            family_id = ?,
            is_head = ?,
            relation_to_head = ?,
            full_name = ?,
            full_name_hi = ?,
            gender = ?,
            dob = ?,
            age = ?,
            blood_group = ?,
            marital_status = ?,
            father_name = ?,
            mother_name = ?,
            spouse_name = ?,
            children_count = ?,
            total_family_members = ?,
            family_members_notes = ?,
            gotra_subcaste = ?,
            native_origin = ?,
            mohalla_area = ?,
            current_address = ?,
            permanent_address = ?,
            city = ?,
            pincode = ?,
            phone_primary = ?,
            phone_whatsapp = ?,
            email = ?,
            education_level = ?,
            education_stream = ?,
            institution_name = ?,
            employment_status = ?,
            occupation_title = ?,
            business_or_employer = ?,
            work_address = ?,
            annual_income_range = ?,
            skills_specialization = ?,
            aadhaar_last4 = ?,
            bhamashah_jan_aadhaar = ?,
            samaj_role = ?,
            special_notes = ?,
            photo_data = ?,
            updated_at = ?
        WHERE id = ?
    ''', (
        data.get('family_id', ''),
        is_head,
        data.get('relation_to_head', 'Self'),
        full_name,
        data.get('full_name_hi', ''),
        data.get('gender', 'Male'),
        dob,
        age,
        data.get('blood_group', ''),
        data.get('marital_status', 'Single'),
        data.get('father_name', ''),
        data.get('mother_name', ''),
        data.get('spouse_name', ''),
        children_count,
        total_family_members,
        data.get('family_members_notes', ''),
        data.get('gotra_subcaste', ''),
        data.get('native_origin', 'Khanda Falsa, Jodhpur'),
        data.get('mohalla_area', 'Khanda Falsa'),
        data.get('current_address', ''),
        data.get('permanent_address', ''),
        data.get('city', 'Jodhpur'),
        data.get('pincode', '342001'),
        data.get('phone_primary', ''),
        data.get('phone_whatsapp', ''),
        data.get('email', ''),
        data.get('education_level', ''),
        data.get('education_stream', ''),
        data.get('institution_name', ''),
        data.get('employment_status', ''),
        data.get('occupation_title', ''),
        data.get('business_or_employer', ''),
        data.get('work_address', ''),
        data.get('annual_income_range', ''),
        data.get('skills_specialization', ''),
        data.get('aadhaar_last4', ''),
        data.get('bhamashah_jan_aadhaar', ''),
        data.get('samaj_role', 'Member'),
        data.get('special_notes', ''),
        data.get('photo_data', ''),
        now_str,
        member_id
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

    # Age groups
    c.execute("""
        SELECT
            SUM(CASE WHEN age < 18 THEN 1 ELSE 0 END) as under_18,
            SUM(CASE WHEN age >= 18 AND age <= 35 THEN 1 ELSE 0 END) as youth_18_35,
            SUM(CASE WHEN age >= 36 AND age <= 59 THEN 1 ELSE 0 END) as adult_36_59,
            SUM(CASE WHEN age >= 60 THEN 1 ELSE 0 END) as senior_60_plus
        FROM members
    """)
    age_row = dict(c.fetchone())

    # Literacy summary (overall and by gender)
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

    # List of existing families for dropdown autocomplete
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


@app.route('/api/export/json', methods=['GET'])
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
        # Check if member_code exists
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
            member_code,
            item.get('family_id', 'FAM-KF-IMP'),
            int(item.get('is_head') or 0),
            item.get('relation_to_head', 'Self'),
            full_name,
            item.get('full_name_hi', ''),
            item.get('gender', 'Male'),
            item.get('dob', ''),
            int(item.get('age') or 0),
            item.get('blood_group', ''),
            item.get('marital_status', 'Single'),
            item.get('father_name', ''),
            item.get('mother_name', ''),
            item.get('spouse_name', ''),
            int(item.get('children_count') or 0),
            int(item.get('total_family_members') or 1),
            item.get('family_members_notes', ''),
            item.get('gotra_subcaste', ''),
            item.get('native_origin', 'Khanda Falsa, Jodhpur'),
            item.get('mohalla_area', 'Khanda Falsa'),
            item.get('current_address', ''),
            item.get('permanent_address', ''),
            item.get('city', 'Jodhpur'),
            item.get('pincode', '342001'),
            item.get('phone_primary', ''),
            item.get('phone_whatsapp', ''),
            item.get('email', ''),
            item.get('education_level', ''),
            item.get('education_stream', ''),
            item.get('institution_name', ''),
            item.get('employment_status', ''),
            item.get('occupation_title', ''),
            item.get('business_or_employer', ''),
            item.get('work_address', ''),
            item.get('annual_income_range', ''),
            item.get('skills_specialization', ''),
            item.get('aadhaar_last4', ''),
            item.get('bhamashah_jan_aadhaar', ''),
            item.get('samaj_role', 'Member'),
            item.get('special_notes', ''),
            item.get('photo_data', '')
        ))
        imported += 1

    conn.commit()
    conn.close()
    return jsonify({'message': f'Successfully imported {imported} member records.', 'imported': imported})


if __name__ == '__main__':
    init_db()
    print("=" * 65)
    print("  CHHEEPA SAMAJ KHANDA FALSA, JODHPUR - CENSUS & DIRECTORY PORTAL")
    print("  Open your browser at: http://127.0.0.1:5000")
    print("=" * 65)
    app.run(host='127.0.0.1', port=5000, debug=True)
