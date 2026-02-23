#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
استخراج بيانات الطلاب من ملفات Excel (HTML) وتحويلها لصيغة CSV للموقع
"""

import sys
import os
import csv
from collections import defaultdict
from bs4 import BeautifulSoup
import xlrd

sys.stdout.reconfigure(encoding='utf-8')

# ============================================================
# الإعدادات
# ============================================================
DATA_DIR = "data"
OUTPUT_CSV = os.path.join("KPI_TaifShare3h-main", "data", "data.csv")
EXISTING_CSV = os.path.join("KPI_TaifShare3h-main", "data", "data.csv")

# تسلسل السنوات (بدون 43 لأنها اندمجت مع 42)
# نضيف 38 كسنة أساس لحساب مستجدي 39
YEAR_SEQUENCE = [38, 39, 40, 41, 42, 44, 45, 46, 47]

def get_previous_year(year):
    """الحصول على السنة السابقة في التسلسل"""
    try:
        idx = YEAR_SEQUENCE.index(year)
        if idx > 0:
            return YEAR_SEQUENCE[idx - 1]
    except ValueError:
        pass
    return None

def get_year_n_before(year, n):
    """الحصول على السنة قبل n سنوات حقيقية (بالحساب الفعلي وليس بالفهرس)
    مثال: سنة 1446 قبل 4 سنوات = 1442 (وليس 1441)
    لأن 1443 غير موجودة (مدمجة) لكن الفارق الزمني الحقيقي هو 4 سنوات
    إذا وقعت السنة المستهدفة على سنة غير موجودة (مثل 1443)
    نرجع لأقرب سنة متاحة قبلها (1442)
    """
    target = year - n
    if target in YEAR_SEQUENCE:
        return target
    # سنة 1443 مدمجة → نرجع لأقرب سنة متاحة قبلها
    candidates = [y for y in YEAR_SEQUENCE if y <= target]
    if candidates:
        return max(candidates)
    return None

# تطبيع الدرجة العلمية
DEGREE_MAP = {
    'البكالوريوس': 'بكالوريوس',
    'بكالوريوس': 'بكالوريوس',
    'الماجستير': 'الماجستير',
    'ماجستير': 'الماجستير',
    'الدكتوراه': 'دكتوراه',
    'دكتوراه': 'دكتوراه',
}

# تحديد القسم من التخصص
DEPT_MAP = {
    'الأنظمة': 'الأنظمة',
    'القانون': 'الأنظمة',
    'الشريعة': 'الشريعة',
    'الفقه': 'الشريعة',
    'أصول الفقه': 'الشريعة',
    'العقيدة': 'الشريعة',
    'الدراسات الإسلامية': 'الدراسات الإسلامية',
    'القراءات': 'القراءات',
    'القرآن وعلومه': 'القراءات',
    'الدراسات القرآنية': 'القراءات',
    'الدراسات القرآنية المعاصرة': 'القراءات',
}

# الحالة الوحيدة لاحتساب الطالب ضمن الإجمالي
ENROLLED_STATUS = 'منتظم'
GRADUATED_STATUS = 'متخرج'

# البرامج المستبعدة (قديمة/ملغاة)
EXCLUDED_PROGRAMS = {
    'الشريعة والدراسات الاسلامية',
    'الشريعة والدراسات الإسلامية',
}

# مدة التخرج بالوقت حسب الدرجة العلمية
DEGREE_YEARS = {
    'بكالوريوس': 4,
    'الماجستير': 2,
    'دكتوراه': 3,
}

# ============================================================
# تحديد نوع الملف
# ============================================================
def is_real_xls(filepath):
    """التحقق إذا كان الملف Excel حقيقي (OLE2) وليس HTML"""
    with open(filepath, 'rb') as f:
        header = f.read(8)
    return header == b'\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1'


# ============================================================
# استخراج البيانات من ملف XLS حقيقي (OLE2)
# ============================================================
def parse_real_xls(filepath):
    """تحليل ملف XLS حقيقي واستخراج سجلات الطلاب"""
    wb = xlrd.open_workbook(filepath)
    sh = wb.sheet_by_index(0)
    students = []
    current_dept = ''

    # تحديد أعمدة الرأس
    col_map = {}  # column_name -> column_index

    for r in range(sh.nrows):
        # قراءة جميع القيم في الصف
        row_vals = {}
        for c in range(sh.ncols):
            v = sh.cell_value(r, c)
            if isinstance(v, str):
                v = v.replace('\xa0', ' ').strip()
            if isinstance(v, float) and v == int(v):
                v = int(v)
            if v:
                row_vals[c] = v

        if not row_vals:
            continue

        vals_list = list(row_vals.values())
        vals_str = [str(v) for v in vals_list]

        # التقاط القسم
        if len(vals_list) == 2 and any('القسم' in str(v) for v in vals_list):
            for v in vals_list:
                v_str = str(v).replace(':', '').strip()
                if 'القسم' not in v_str and v_str:
                    current_dept = v_str
            continue

        # تجاهل صف الكلية
        if any('الكلية' in str(v) for v in vals_list) and len(vals_list) <= 3:
            continue

        # التقاط صف الرأس
        if any('الرقم الجامعي' in str(v) for v in vals_list):
            col_map = {}
            for c, v in row_vals.items():
                v_str = str(v).strip()
                if v_str == 'م':
                    col_map['idx'] = c
                elif 'الرقم الجامعي' in v_str:
                    col_map['student_id'] = c
                elif v_str == 'الاسم':
                    col_map['name'] = c
                elif v_str == 'التخصص':
                    col_map['program'] = c
                elif v_str == 'الحالة':
                    col_map['status'] = c
                elif v_str == 'الجنس':
                    col_map['gender'] = c
                elif v_str == 'العمر':
                    col_map['age'] = c
                elif v_str == 'الجنسية':
                    col_map['nationality'] = c
                elif 'الدرجة العلمية' in v_str:
                    col_map['degree'] = c
                elif 'نوع الدراسة' in v_str:
                    col_map['study_type'] = c
                elif 'تاريخ القبول' in v_str:
                    col_map['admission_date'] = c
                elif 'المتوقع' in v_str:
                    col_map['expected_grad'] = c
                elif 'تاريخ التخرج' in v_str:
                    col_map['grad_date'] = c
                elif v_str == 'المعدل':
                    col_map['gpa'] = c
            continue

        if not col_map:
            continue

        # استخراج بيانات الطالب
        def get_val(key):
            if key in col_map and col_map[key] in row_vals:
                return str(row_vals[col_map[key]])
            return ''

        idx_val = get_val('idx')
        if not idx_val or not str(idx_val).replace('.0', '').replace('.', '').isdigit():
            continue

        student = {
            'student_id': get_val('student_id').replace('.0', ''),
            'name': get_val('name'),
            'program': get_val('program'),
            'status': get_val('status'),
            'gender': get_val('gender'),
            'age': get_val('age'),
            'nationality': get_val('nationality'),
            'degree': get_val('degree'),
            'study_type': get_val('study_type'),
            'admission_date': get_val('admission_date'),
            'expected_grad': get_val('expected_grad'),
            'grad_date': get_val('grad_date'),
            'gpa': get_val('gpa'),
            'dept': current_dept,
        }
        students.append(student)

    return students


# ============================================================
# استخراج البيانات من ملف HTML/XLS
# ============================================================
def parse_html_xls(filepath):
    """تحليل ملف XLS (HTML) واستخراج سجلات الطلاب"""
    with open(filepath, 'rb') as f:
        content = f.read()

    text = content.decode('cp1256')
    soup = BeautifulSoup(text, 'html.parser')
    tables = soup.find_all('table')

    if not tables:
        return []

    rows = tables[0].find_all('tr')
    students = []
    current_dept = ''

    for row in rows:
        cells = row.find_all(['td', 'th'])
        cell_texts = [c.get_text(strip=True).replace('\xa0', ' ').strip() for c in cells]
        non_empty = [t for t in cell_texts if t]

        if not non_empty:
            continue

        # التقاط اسم القسم من صفوف الرأس
        if len(non_empty) == 2 and 'القسم' in non_empty[0]:
            current_dept = non_empty[1].replace(':', '').strip()
            continue

        # تجاهل صفوف الرأس الأخرى
        if len(non_empty) <= 3:
            continue

        # تجاهل صف العناوين
        if 'الرقم الجامعي' in non_empty or 'م' == non_empty[0] and 'الاسم' in non_empty:
            continue

        # استخراج بيانات الطالب
        # الأعمدة: م, الرقم الجامعي, الاسم, التخصص, الحالة, الجنس, العمر, الجنسية, الدرجة العلمية, نوع الدراسة, تاريخ القبول, [تاريخ متوقع], [تاريخ تخرج], المعدل
        if len(non_empty) >= 10:
            try:
                # التحقق من أن العمود الأول رقم (م)
                idx_num = non_empty[0]
                if not idx_num.isdigit():
                    continue

                student_id = non_empty[1]
                name = non_empty[2]
                program = non_empty[3]
                status = non_empty[4]
                gender = non_empty[5]
                age = non_empty[6]
                nationality = non_empty[7]
                degree = non_empty[8] if len(non_empty) > 8 else ''
                study_type = non_empty[9] if len(non_empty) > 9 else ''
                admission_date = non_empty[10] if len(non_empty) > 10 else ''

                # الحقول الاختيارية (قد تكون مفقودة)
                expected_grad = ''
                grad_date = ''
                gpa = ''

                if len(non_empty) >= 14:
                    expected_grad = non_empty[11]
                    grad_date = non_empty[12]
                    gpa = non_empty[13]
                elif len(non_empty) == 13:
                    # قد يكون تاريخ التخرج المتوقع مفقوداً
                    # نحتاج التمييز: هل هو (متوقع, تخرج, معدل) أم (تخرج, معدل, ?)
                    # إذا كان العنصر الأخير يبدو كمعدل (رقم عشري)
                    last = non_empty[12]
                    second_last = non_empty[11]
                    try:
                        float(last)
                        gpa = last
                        # هل second_last تاريخ؟
                        if '-' in second_last and len(second_last) >= 8:
                            grad_date = second_last
                        else:
                            expected_grad = second_last
                    except ValueError:
                        expected_grad = non_empty[11]
                        grad_date = non_empty[12]
                elif len(non_empty) == 12:
                    last = non_empty[11]
                    try:
                        float(last)
                        gpa = last
                    except ValueError:
                        expected_grad = last
                elif len(non_empty) == 11:
                    last = non_empty[10]
                    if '-' in last:
                        admission_date = last

                students.append({
                    'student_id': student_id,
                    'name': name,
                    'program': program,
                    'status': status,
                    'gender': gender,
                    'age': age,
                    'nationality': nationality,
                    'degree': degree,
                    'study_type': study_type,
                    'admission_date': admission_date,
                    'expected_grad': expected_grad,
                    'grad_date': grad_date,
                    'gpa': gpa,
                    'dept': current_dept,
                })
            except (IndexError, ValueError) as e:
                continue

    return students


# ============================================================
# المعالجة الرئيسية
# ============================================================
def main():
    print("=" * 70)
    print("بدء استخراج البيانات من ملفات Excel")
    print("المنهجية: إجمالي الطلاب = منتظم فقط من الفصل الأول")
    print("         الخريجين = متخرج من أي فصل في السنة")
    print("=" * 70)

    # 1. قراءة جميع الملفات وتنظيمها حسب السنة والفصل
    all_files = sorted([f for f in os.listdir(DATA_DIR) if f.endswith('.xls')])

    year_files = defaultdict(list)
    for fname in all_files:
        num = fname.replace('.xls', '')
        year = int(num[:2])
        semester = int(num[2])
        year_files[year].append((fname, semester))

    print(f"\nالسنوات المتاحة: {sorted(year_files.keys())}")
    for year in sorted(year_files.keys()):
        files = year_files[year]
        print(f"  سنة {year} (14{year:02d}): {[f[0] for f in files]}")

    # 2. استخراج الطلاب لكل ملف حسب الفصل
    # sem1_students[year] = {student_id: student_record}  ← الفصل الأول فقط
    # all_semesters_students[year] = {student_id: student_record}  ← كل الفصول (للخريجين)
    sem1_students = defaultdict(dict)
    all_semesters_students = defaultdict(dict)

    total_records = 0
    for year in sorted(year_files.keys()):
        for fname, semester in sorted(year_files[year], key=lambda x: x[1]):
            filepath = os.path.join(DATA_DIR, fname)
            print(f"\n  معالجة {fname}...", end=' ')
            if is_real_xls(filepath):
                students = parse_real_xls(filepath)
            else:
                students = parse_html_xls(filepath)
            print(f"{len(students)} سجل")
            total_records += len(students)

            for s in students:
                sid = s['student_id']

                # الفصل الأول فقط: لحساب إجمالي الطلاب
                if semester == 1:
                    sem1_students[year][sid] = s

                # كل الفصول: لتتبع حالة التخرج
                # إذا تخرج في أي فصل، نسجّل ذلك
                existing = all_semesters_students[year].get(sid)
                if existing:
                    if s['status'] == GRADUATED_STATUS:
                        all_semesters_students[year][sid] = s
                else:
                    all_semesters_students[year][sid] = s

    print(f"\n{'='*70}")
    print(f"إجمالي السجلات المقروءة: {total_records:,}")
    print(f"{'='*70}")

    # 3. طباعة ملخص
    for year in sorted(sem1_students.keys()):
        s1 = sem1_students[year]
        enrolled = sum(1 for s in s1.values() if s['status'] == ENROLLED_STATUS)
        all_s = all_semesters_students[year]
        graduated = sum(1 for s in all_s.values() if s['status'] == GRADUATED_STATUS)
        print(f"\n  سنة {year} (14{year:02d}): فصل1={len(s1):,} | منتظم={enrolled:,} | متخرج(كل الفصول)={graduated:,}")

    # 4. حساب المستجدين: منتظم في فصل1 للسنة الحالية ولم يكن في فصل1 للسنة السابقة
    # new_students[year] = set of student_ids (منتظم في الفصل الأول وجديد)
    new_students = {}
    for year in YEAR_SEQUENCE:
        if year not in sem1_students:
            continue

        # أرقام المنتظمين فقط في الفصل الأول
        current_enrolled_ids = set(
            sid for sid, s in sem1_students[year].items()
            if s['status'] == ENROLLED_STATUS
        )

        prev_year = get_previous_year(year)
        if prev_year and prev_year in sem1_students:
            # أرقام كل الطلاب في الفصل الأول للسنة السابقة (بأي حالة)
            prev_all_ids = set(sem1_students[prev_year].keys())
            new_students[year] = current_enrolled_ids - prev_all_ids
        else:
            new_students[year] = set()

        if new_students[year]:
            print(f"\n  المستجدون سنة {year}: {len(new_students[year]):,}")

    # 5. تجميع البيانات حسب (سنة، تخصص، درجة)
    print(f"\n{'='*70}")
    print("تجميع البيانات...")
    print(f"{'='*70}")

    aggregated = {}

    for year in sorted(sem1_students.keys()):
        if year not in YEAR_SEQUENCE:
            print(f"  تحذير: سنة {year} ليست في التسلسل المعروف")
            continue

        # ── جمع "المنتظمين" من الفصل الأول حسب (تخصص، درجة) ──
        sem1_groups = defaultdict(list)
        for s in sem1_students[year].values():
            prog = s['program']
            degree = DEGREE_MAP.get(s['degree'], s['degree'])
            if not degree or (degree == s['degree'] and degree not in DEGREE_MAP.values()):
                continue
            sem1_groups[(prog, degree)].append(s)

        # ── جمع "المتخرجين" من كل الفصول حسب (تخصص، درجة) ──
        grad_groups = defaultdict(list)
        for s in all_semesters_students[year].values():
            if s['status'] == GRADUATED_STATUS:
                prog = s['program']
                degree = DEGREE_MAP.get(s['degree'], s['degree'])
                if not degree or (degree == s['degree'] and degree not in DEGREE_MAP.values()):
                    continue
                grad_groups[(prog, degree)].append(s)

        # ── جمع كل المفاتيح (قد يوجد برنامج فيه خريجون بدون منتظمين أو العكس) ──
        all_keys = set(sem1_groups.keys()) | set(grad_groups.keys())

        for (prog, degree) in all_keys:
            # استبعاد البرامج القديمة/الملغاة
            if prog in EXCLUDED_PROGRAMS:
                continue
            dept = DEPT_MAP.get(prog, prog)
            key = f"{dept}|{prog}|{degree}|{year}"

            # المنتظمون من الفصل الأول فقط
            sem1_list = sem1_groups.get((prog, degree), [])
            enrolled = [s for s in sem1_list if s['status'] == ENROLLED_STATUS]

            students_total = len(enrolled)
            students_male = sum(1 for s in enrolled if s['gender'] == 'ذكر')
            students_female = sum(1 for s in enrolled if s['gender'] == 'أنثى')
            students_saudi = sum(1 for s in enrolled if s['nationality'] == 'سعودي')
            students_intl = sum(1 for s in enrolled if s['nationality'] != 'سعودي' and s['nationality'])

            # الخريجين من كل الفصول
            graduated = grad_groups.get((prog, degree), [])
            graduates_total = len(graduated)

            # المستجدون: منتظم في فصل1 وجديد
            prog_new = 0
            if year in new_students:
                enrolled_ids = set(s['student_id'] for s in enrolled)
                prog_new = len(enrolled_ids & new_students[year])

            # الاستبقاء: مستجدو السنة السابقة الذين لا يزالون منتظمين في فصل1 الحالي
            students_retained = 0
            prev_year = get_previous_year(year)
            if prev_year and prev_year in new_students:
                current_enrolled_ids = set(s['student_id'] for s in enrolled)
                for sid in new_students[prev_year]:
                    if sid in sem1_students[prev_year]:
                        prev_s = sem1_students[prev_year][sid]
                        prev_prog = prev_s['program']
                        prev_degree = DEGREE_MAP.get(prev_s['degree'], prev_s['degree'])
                        if prev_prog == prog and prev_degree == degree:
                            if sid in current_enrolled_ids:
                                students_retained += 1

            # التخرج بالوقت المحدد حسب الدرجة العلمية:
            # بكالوريوس = 4 سنوات، ماجستير = سنتين، دكتوراه = 3 سنوات
            graduates_ontime = 0
            n_years = DEGREE_YEARS.get(degree, 4)
            year_n_ago = get_year_n_before(year, n_years)
            if year_n_ago and year_n_ago in new_students:
                for sid in new_students[year_n_ago]:
                    if sid in sem1_students[year_n_ago]:
                        old_s = sem1_students[year_n_ago][sid]
                        old_prog = old_s['program']
                        old_degree = DEGREE_MAP.get(old_s['degree'], old_s['degree'])
                        if old_prog == prog and old_degree == degree:
                            # البحث عن التخرج في أي فصل خلال المدة المحددة
                            found_graduated = False
                            idx_start = YEAR_SEQUENCE.index(year_n_ago) + 1
                            idx_end = YEAR_SEQUENCE.index(year) + 1
                            for check_idx in range(idx_start, idx_end):
                                check_year = YEAR_SEQUENCE[check_idx]
                                if check_year in all_semesters_students:
                                    check_s = all_semesters_students[check_year].get(sid)
                                    if check_s and check_s['status'] == GRADUATED_STATUS:
                                        found_graduated = True
                                        break
                            if found_graduated:
                                graduates_ontime += 1

            # عدد مستجدي n سنوات سابقة في هذا البرنامج (مقام نسبة التخرج بالوقت)
            new_n_ago_count = 0
            if year_n_ago and year_n_ago in new_students:
                for sid in new_students[year_n_ago]:
                    if sid in sem1_students[year_n_ago]:
                        old_s = sem1_students[year_n_ago][sid]
                        old_prog = old_s['program']
                        old_degree = DEGREE_MAP.get(old_s['degree'], old_s['degree'])
                        if old_prog == prog and old_degree == degree:
                            new_n_ago_count += 1

            # عدد مستجدي السنة السابقة في هذا البرنامج (مقام نسبة الاستبقاء)
            prev_new_count = 0
            if prev_year and prev_year in new_students:
                for sid in new_students[prev_year]:
                    if sid in sem1_students[prev_year]:
                        prev_s = sem1_students[prev_year][sid]
                        prev_prog = prev_s['program']
                        prev_degree = DEGREE_MAP.get(prev_s['degree'], prev_s['degree'])
                        if prev_prog == prog and prev_degree == degree:
                            prev_new_count += 1

            aggregated[key] = {
                'dept': dept,
                'prog': prog,
                'degree': degree,
                'semester': year,
                'students_total': students_total,
                'students_male': students_male,
                'students_female': students_female,
                'students_saudi': students_saudi,
                'students_international': students_intl,
                'students_new': prog_new,
                'students_retained': students_retained,
                'graduates_total': graduates_total,
                'graduates_ontime': graduates_ontime,
                'prev_new_count': prev_new_count,
                'new_4_ago_count': new_n_ago_count,
            }

    # 6. طباعة الملخص
    print(f"\n{'='*70}")
    print(f"إجمالي السجلات المجمعة: {len(aggregated)}")
    print(f"{'='*70}")

    for key in sorted(aggregated.keys()):
        d = aggregated[key]
        print(f"\n  {d['dept']} | {d['prog']} | {d['degree']} | 14{d['semester']:02d}")
        print(f"    منتظمون(فصل1): {d['students_total']} (ذكور: {d['students_male']}, إناث: {d['students_female']})")
        print(f"    سعوديون: {d['students_saudi']}, دوليون: {d['students_international']}")
        print(f"    مستجدون: {d['students_new']}, مستبقون: {d['students_retained']} من {d['prev_new_count']}")
        n_yrs = DEGREE_YEARS.get(d['degree'], 4)
        print(f"    خريجون(كل الفصول): {d['graduates_total']}, بالوقت({n_yrs} سنوات): {d['graduates_ontime']} من {d['new_4_ago_count']}")
        if d['prev_new_count'] > 0:
            retention = round(d['students_retained'] / d['prev_new_count'] * 100, 1)
            print(f"    معدل الاستبقاء: {retention}% ({d['students_retained']} من {d['prev_new_count']})")
        if d['new_4_ago_count'] > 0:
            grad_rate = round(d['graduates_ontime'] / d['new_4_ago_count'] * 100, 1)
            print(f"    معدل التخرج بالوقت: {grad_rate}% ({d['graduates_ontime']} من {d['new_4_ago_count']})")

    # 7. استخراج سجلات الخريجين الفردية
    print(f"\n{'='*70}")
    print("استخراج سجلات الخريجين الفردية...")
    print(f"{'='*70}")

    graduates_list = []
    for year in sorted(all_semesters_students.keys()):
        if year not in YEAR_SEQUENCE or year == 38:
            continue
        for sid, s in all_semesters_students[year].items():
            if s['status'] == GRADUATED_STATUS:
                prog = s['program']
                if prog in EXCLUDED_PROGRAMS:
                    continue
                degree = DEGREE_MAP.get(s['degree'], s['degree'])
                dept = DEPT_MAP.get(prog, prog)
                graduates_list.append({
                    'year': year,
                    'student_id': sid,
                    'name': s['name'],
                    'program': prog,
                    'degree': degree,
                    'dept': dept,
                    'gender': s['gender'],
                    'nationality': s['nationality'],
                    'admission_date': s.get('admission_date', ''),
                    'grad_date': s.get('grad_date', ''),
                    'expected_grad': s.get('expected_grad', ''),
                    'gpa': s.get('gpa', ''),
                })

    GRADUATES_CSV = os.path.join("KPI_TaifShare3h-main", "data", "graduates_detail.csv")
    with open(GRADUATES_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            'السنة', 'الرقم_الجامعي', 'الاسم', 'التخصص', 'الدرجة', 'القسم',
            'الجنس', 'الجنسية', 'تاريخ_القبول', 'تاريخ_التخرج',
            'تاريخ_التخرج_المتوقع', 'المعدل'
        ])
        for g in sorted(graduates_list, key=lambda x: (x['year'], x['dept'], x['program'])):
            writer.writerow([
                g['year'], g['student_id'], g['name'], g['program'],
                g['degree'], g['dept'], g['gender'], g['nationality'],
                g['admission_date'], g['grad_date'],
                g['expected_grad'], g['gpa']
            ])
    print(f"  تم كتابة {len(graduates_list)} سجل خريج في {GRADUATES_CSV}")

    # 8. استخراج سجلات غير المكملين (جميع الحالات عدا منتظم ومتخرج)
    print(f"\n{'='*70}")
    print("استخراج سجلات غير المكملين...")
    print(f"{'='*70}")

    # نجمع كل الحالات الفريدة أولاً للطباعة
    all_statuses = set()
    non_completers_dict = {}  # (student_id, program, degree) -> latest record

    for year in sorted(all_semesters_students.keys()):
        if year not in YEAR_SEQUENCE or year == 38:
            continue
        for sid, s in all_semesters_students[year].items():
            status = s['status']
            all_statuses.add(status)

            if status == ENROLLED_STATUS or status == GRADUATED_STATUS:
                continue

            prog = s['program']
            if prog in EXCLUDED_PROGRAMS:
                continue
            degree = DEGREE_MAP.get(s['degree'], s['degree'])
            dept = DEPT_MAP.get(prog, prog)

            rec_key = (sid, prog, degree)
            existing = non_completers_dict.get(rec_key)
            # نحتفظ بأحدث سجل (أكبر سنة)
            if not existing or year > existing['year']:
                non_completers_dict[rec_key] = {
                    'year': year,
                    'student_id': sid,
                    'name': s['name'],
                    'program': prog,
                    'degree': degree,
                    'dept': dept,
                    'status': status,
                    'gender': s['gender'],
                    'nationality': s['nationality'],
                    'admission_date': s.get('admission_date', ''),
                    'gpa': s.get('gpa', ''),
                    'study_type': s.get('study_type', ''),
                }

    # استبعاد من تخرج لاحقاً (قد يكون غير مكمل في سنة ثم تخرج لاحقاً)
    graduated_ids = set()
    for year in all_semesters_students:
        for sid, s in all_semesters_students[year].items():
            if s['status'] == GRADUATED_STATUS:
                prog = s['program']
                degree = DEGREE_MAP.get(s['degree'], s['degree'])
                graduated_ids.add((sid, prog, degree))

    non_completers_list = [
        rec for key, rec in non_completers_dict.items()
        if key not in graduated_ids
    ]

    print(f"  جميع الحالات الموجودة: {all_statuses}")
    non_comp_statuses = set(r['status'] for r in non_completers_list)
    print(f"  حالات غير المكملين: {non_comp_statuses}")

    NON_COMP_CSV = os.path.join("KPI_TaifShare3h-main", "data", "non_completers.csv")
    with open(NON_COMP_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow([
            'آخر_سنة', 'الرقم_الجامعي', 'الاسم', 'التخصص', 'الدرجة', 'القسم',
            'الحالة', 'الجنس', 'الجنسية', 'تاريخ_القبول', 'المعدل', 'نوع_الدراسة'
        ])
        for nc in sorted(non_completers_list, key=lambda x: (x['year'], x['dept'], x['program'], x['status'])):
            writer.writerow([
                nc['year'], nc['student_id'], nc['name'], nc['program'],
                nc['degree'], nc['dept'], nc['status'], nc['gender'],
                nc['nationality'], nc['admission_date'], nc['gpa'],
                nc['study_type']
            ])
    print(f"  تم كتابة {len(non_completers_list)} سجل غير مكمل في {NON_COMP_CSV}")

    # 9. كتابة CSV النهائي
    print(f"\n{'='*70}")
    print(f"كتابة الملف النهائي: {OUTPUT_CSV}")
    print(f"{'='*70}")

    csv_headers = [
        'Dept_aName', 'Major_aName', 'Degree_aName', 'Semester',
        'students_total', 'students_male', 'students_female',
        'students_saudi', 'students_international',
        'students_new', 'students_retained',
        'graduates_total', 'graduates_ontime',
        'prev_new_count', 'new_4_ago_count',
        'sections_total', 'sections_male', 'sections_female',
        'faculty_total', 'faculty_phd', 'faculty_male', 'faculty_female',
        'faculty_published', 'research_count', 'citations',
        'eval_courses', 'eval_experience', 'eval_employers',
        'performance_rate', 'employment_rate'
    ]

    rows_written = 0
    with open(OUTPUT_CSV, 'w', encoding='utf-8', newline='') as f:
        writer = csv.writer(f, delimiter=';')
        writer.writerow(csv_headers)

        for key in sorted(aggregated.keys()):
            d = aggregated[key]

            row = [
                d['dept'],
                d['prog'],
                d['degree'],
                d['semester'],
                d['students_total'],
                d['students_male'],
                d['students_female'],
                d['students_saudi'],
                d['students_international'],
                d['students_new'],
                d['students_retained'],
                d['graduates_total'],
                d['graduates_ontime'],
                d['prev_new_count'],
                d['new_4_ago_count'],
                '', '', '',  # sections
                '', '', '', '',  # faculty
                '', '', '',  # research
                '', '', '',  # eval
                '', '',  # rates
            ]
            writer.writerow(row)
            rows_written += 1

    print(f"\n  تم كتابة {rows_written} صف في {OUTPUT_CSV}")
    print(f"\n{'='*70}")
    print("تم الانتهاء بنجاح!")
    print(f"{'='*70}")


if __name__ == '__main__':
    main()
