import csv
import html
import re
from pathlib import Path


ROOT = Path(__file__).resolve().parent
SOURCE_DIR = ROOT / 'بيانات الطلاب'
DATA_DIR = ROOT / 'data'
OUTPUT_CSV = DATA_DIR / 'students_detail.csv'
PROGRAMS_CSV = DATA_DIR / 'data.csv'

HEADER = [
    'السنة',
    'الرقم_الجامعي',
    'الاسم',
    'التخصص',
    'الدرجة',
    'القسم',
    'الحالة',
    'الجنس',
    'العمر',
    'الجنسية',
    'نوع_الدراسة',
    'تاريخ_القبول',
    'تاريخ_التخرج_المتوقع',
    'تاريخ_التخرج',
    'المعدل',
    'مصدر_الملف',
]


def clean_html_text(value: str) -> str:
    text = html.unescape(value or '')
    text = re.sub(r'<br\s*/?>', ' ', text, flags=re.IGNORECASE)
    text = re.sub(r'<[^>]+>', ' ', text)
    text = text.replace('\xa0', ' ')
    text = re.sub(r'\s+', ' ', text).strip()
    return text


def normalize_degree(value: str) -> str:
    text = (value or '').strip()
    if text in {'البكالوريوس', 'بكالوريوس', 'بكالوريوس انتساب'}:
        return 'بكالوريوس'
    if text in {'الماجستير', 'ماجستير'}:
        return 'الماجستير'
    if text in {'الدكتوراه', 'دكتوراه'}:
        return 'دكتوراه'
    return text


def normalize_program(value: str) -> str:
    text = (value or '').strip()
    aliases = {
        'القران وعلومه': 'القرآن وعلومه',
        'الدراسات القرانيه': 'الدراسات القرآنية',
        'الانظمة': 'الأنظمة',
    }
    return aliases.get(text, text)


def load_department_map():
    mapping = {}
    with PROGRAMS_CSV.open(encoding='utf-8-sig', newline='') as handle:
        reader = csv.DictReader(handle, delimiter=';')
        for row in reader:
            program = normalize_program(row.get('Major_aName', ''))
            degree = normalize_degree(row.get('Degree_aName', ''))
            dept = (row.get('Dept_aName') or '').strip()
            if not program or not degree or not dept:
                continue
            mapping[f'{program}|{degree}'] = dept
    return mapping


def parse_student_rows(path: Path):
    text = path.read_text(encoding='cp1256', errors='ignore')
    rows = []
    for tr_html in re.findall(r'<tr[^>]*>(.*?)</tr>', text, flags=re.IGNORECASE | re.DOTALL):
        cells = [
            clean_html_text(cell)
            for cell in re.findall(r'<t[dh][^>]*>(.*?)</t[dh]>', tr_html, flags=re.IGNORECASE | re.DOTALL)
        ]
        if len(cells) != 15 or not cells[0].isdigit():
            continue
        rows.append({
            'الرقم_الجامعي': cells[1],
            'الاسم': cells[2],
            'التخصص': normalize_program(cells[3]),
            'الحالة': cells[4],
            'الجنس': cells[5],
            'العمر': cells[6],
            'الجنسية': cells[7],
            'الدرجة': normalize_degree(cells[8]),
            'نوع_الدراسة': cells[9],
            'تاريخ_القبول': cells[10],
            'تاريخ_التخرج_المتوقع': cells[11],
            'تاريخ_التخرج': cells[12],
            'المعدل': cells[13],
            'مصدر_الملف': path.name,
        })
    return rows


def merge_year_46_rows(base_rows, override_rows):
    merged = {
        (row['الرقم_الجامعي'], row['التخصص'], row['الدرجة']): row
        for row in base_rows
    }
    for row in override_rows:
        merged[(row['الرقم_الجامعي'], row['التخصص'], row['الدرجة'])] = row
    return list(merged.values())


def attach_year_and_department(rows, year, dept_map):
    enriched = []
    for row in rows:
        record = dict(row)
        record['السنة'] = str(year)
        record['القسم'] = dept_map.get(f"{record['التخصص']}|{record['الدرجة']}", '')
        enriched.append(record)
    return enriched


def main():
    dept_map = load_department_map()

    rows_45 = parse_student_rows(SOURCE_DIR / '452.xls')
    rows_46_base = parse_student_rows(SOURCE_DIR / '462.xls')
    rows_46_updates = parse_student_rows(SOURCE_DIR / '465.xls')
    rows_46 = merge_year_46_rows(rows_46_base, rows_46_updates)

    final_rows = (
        attach_year_and_department(rows_45, 45, dept_map) +
        attach_year_and_department(rows_46, 46, dept_map)
    )

    final_rows.sort(key=lambda row: (
        int(row['السنة']),
        row['القسم'],
        row['التخصص'],
        row['الدرجة'],
        row['الرقم_الجامعي'],
    ))

    with OUTPUT_CSV.open('w', encoding='utf-8-sig', newline='') as handle:
        writer = csv.DictWriter(handle, fieldnames=HEADER, delimiter=';')
        writer.writeheader()
        writer.writerows(final_rows)

    print(f'Wrote {len(final_rows)} student detail rows to {OUTPUT_CSV}')


if __name__ == '__main__':
    main()
