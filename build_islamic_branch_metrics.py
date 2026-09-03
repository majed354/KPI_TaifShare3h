#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""Build branch-level Islamic Studies student metrics from result rosters.

The result files do not contain an official program, branch, admission, or
graduation field.  This builder therefore keeps the source of every decision
explicit and accepts only three internal confidence tiers:

``official``
    The student/event exists in an official detail file and the branch has at
    least ``high`` evidence in the result rosters.
``very_high``
    The deterministic evidence meets the strict program, branch, and event
    rules below.
``high``
    The evidence meets the conservative fallback rules below.

The UI may aggregate all three tiers without rendering the tier.  The tier is
retained in this JSON so every published number remains auditable.

Important limitations
---------------------
* ``students_total`` inferred from results is an active-results proxy.  Only
  rows matched to ``students_detail.csv`` with status ``منتظم`` are official.
* Plan 47 is used to identify students and first-level courses, but never to
  infer graduation because its elective-group metadata is incomplete.
* An inferred Plan-39 graduate must pass every fixed course plus the required
  elective hours, then be absent from a following term that is itself fully
  observed and covered for the assigned branch.
* Ambiguous grades are never counted as passing.
* Campus is read from ``sections.campus``; the source file's parent folder is
  deliberately ignored.
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import math
import os
import re
import sqlite3
import unicodedata
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import date, datetime, timedelta, timezone
from pathlib import Path
from typing import Any, Iterable, Mapping, Sequence


ROOT = Path(__file__).resolve().parent
ANALYSIS_ROOT = ROOT.parent

DEFAULT_DATABASE = ANALYSIS_ROOT / "automation_import_campus.sqlite"
DEFAULT_PLAN_DATA = ANALYSIS_ROOT / "courseOfShar3ah" / "data.json"
DEFAULT_STUDENTS = ROOT / "data" / "students_detail.csv"
DEFAULT_GRADUATES = ROOT / "data" / "graduates_detail.csv"
DEFAULT_NON_COMPLETERS = ROOT / "data" / "non_completers.csv"
DEFAULT_KPI = ROOT / "data" / "data.csv"
DEFAULT_OUTPUT = ROOT / "data" / "islamic_branch_students.json"

PROGRAM_AR = "الدراسات الإسلامية"
DEGREE_AR = "بكالوريوس"
DEPARTMENT_AR = "الدراسات الإسلامية"
TARGET_YEARS = (1446, 1447)
BRANCH_ORDER = ("الحوية", "تربة", "رنية", "الخرمة")
SEMESTER_LABELS = {1: "الفصل الأول", 2: "الفصل الثاني", 3: "الفصل الصيفي"}
CONFIDENCE_RANK = {"high": 1, "very_high": 2, "official": 3}

PASS_GRADES = frozenset({"أ+", "أ", "ب+", "ب", "ج+", "ج", "د+", "د"})
FAIL_GRADES = frozenset({"هـ", "ه"})
NON_COMPLETION_GRADES = frozenset({"ع", "ح", "غ"})
AMBIGUOUS_GRADES = frozenset({"ق", "ل", "هد", "ند", "هت"})
AWARDED_GRADES = PASS_GRADES | FAIL_GRADES

# This is an explicit quality policy for the supplied 42-file snapshot.  It is
# not guessed from volume alone: a large general-course file is not evidence
# that Islamic Studies is covered.  The observed counts are added separately
# to the emitted coverage object.
COVERAGE_POLICY: Mapping[int, Mapping[str, Mapping[str, Any]]] = {
    1446: {
        "الحوية": {
            "status": "not_covered",
            "allow_regular": False,
            "allow_entrant": False,
            "allow_inferred_graduate": False,
            "history": "not_covered",
            "reasons": ["ملفات الفروع لا تمثل الحوية تمثيلًا شاملاً."],
        },
        "تربة": {
            "status": "usable",
            "allow_regular": True,
            "allow_entrant": True,
            "allow_inferred_graduate": True,
            "history": "partial_missing_1443",
            "reasons": ["الفصلان الحاليان متاحان؛ سنة 1443 مفقودة من الأرشيف."],
        },
        "رنية": {
            "status": "usable_with_history_gaps",
            "allow_regular": True,
            "allow_entrant": True,
            "allow_inferred_graduate": True,
            "history": "partial_missing_1442_1443",
            "reasons": ["الفصلان الحاليان متاحان؛ سنتا 1442 و1443 مفقودتان."],
        },
        "الخرمة": {
            "status": "partial",
            "allow_regular": False,
            "allow_entrant": False,
            "allow_inferred_graduate": False,
            "history": "partial_truncated_1446_1",
            "reasons": ["تغطية الفصل الأول 1446 مبتورة وغير متوازنة."],
        },
    },
    1447: {
        "الحوية": {
            "status": "not_covered",
            "allow_regular": False,
            "allow_entrant": False,
            "allow_inferred_graduate": False,
            "history": "not_covered",
            "reasons": ["ملفات الفروع لا تمثل الحوية تمثيلًا شاملاً."],
        },
        "تربة": {
            "status": "usable",
            "allow_regular": True,
            "allow_entrant": True,
            "allow_inferred_graduate": True,
            "history": "partial_missing_1443",
            "reasons": ["الفصلان الحاليان متاحان؛ سنة 1443 مفقودة من الأرشيف."],
        },
        "رنية": {
            "status": "usable_with_history_gaps",
            "allow_regular": True,
            "allow_entrant": True,
            "allow_inferred_graduate": True,
            "history": "partial_missing_1442_1443",
            "reasons": ["الفصلان الحاليان متاحان؛ سنتا 1442 و1443 مفقودتان."],
        },
        "الخرمة": {
            "status": "insufficient_program_coverage",
            "allow_regular": False,
            "allow_entrant": False,
            "allow_inferred_graduate": False,
            "history": "partial",
            "reasons": ["لا توجد تغطية تخصصية كافية للدراسات الإسلامية في 1447."],
        },
    },
}


ARABIC_DIGITS = str.maketrans("٠١٢٣٤٥٦٧٨٩۰۱۲۳۴۵۶۷۸۹", "01234567890123456789")
ARABIC_LETTERS = str.maketrans({"إ": "ا", "أ": "ا", "آ": "ا", "ى": "ي", "ة": "ه", "ـ": ""})


def clean(value: Any) -> str:
    return re.sub(r"\s+", " ", str(value or "").replace("\xa0", " ")).strip()


def normalized(value: Any) -> str:
    text = unicodedata.normalize("NFKC", clean(value)).translate(ARABIC_DIGITS)
    text = text.translate(ARABIC_LETTERS)
    text = re.sub(r"[\u064b-\u065f\u0670]", "", text)
    return re.sub(r"\s+", " ", text).strip().casefold()


def normalize_code(value: Any) -> str:
    text = unicodedata.normalize("NFKC", clean(value)).translate(ARABIC_DIGITS)
    text = text.replace("–", "-").replace("—", "-").replace("−", "-")
    match = re.search(r"(?<!\d)(\d{5,9})\s*-\s*(\d{1,3})(?!\d)", text)
    return f"{match.group(1)}-{match.group(2)}" if match else re.sub(r"\s+", "", text)


def normalize_student_id(value: Any) -> str:
    text = clean(value).translate(ARABIC_DIGITS)
    if text.endswith(".0"):
        text = text[:-2]
    digits = re.sub(r"\D", "", text)
    return digits if 6 <= len(digits) <= 14 else ""


def normalize_grade(value: Any) -> str:
    return clean(value).replace("ه‍", "هـ")


def semester_number(value: Any) -> int | None:
    text = normalized(value)
    if "الاول" in text:
        return 1
    if "الثاني" in text:
        return 2
    if "الصيف" in text or "الثالث" in text:
        return 3
    return None


def short_year(year: int) -> str:
    return str(year - 1400 if year >= 1400 else year)


def full_year(value: Any) -> int | None:
    token = clean(value).translate(ARABIC_DIGITS)
    match = re.search(r"\d+", token)
    if not match:
        return None
    year = int(match.group())
    return year + 1400 if 0 <= year < 100 else year


def comparable_date(value: Any) -> tuple[int, int, int] | None:
    """Parse an ISO/slashed date or an Excel serial without changing calendars."""

    token = clean(value).translate(ARABIC_DIGITS)
    if not token:
        return None
    if re.fullmatch(r"\d{5}(?:\.0+)?", token):
        parsed = date(1899, 12, 30) + timedelta(days=int(float(token)))
        return parsed.year, parsed.month, parsed.day
    parts = [int(part) for part in re.findall(r"\d+", token)]
    if len(parts) != 3:
        return None
    if len(str(parts[0])) == 4:
        year, month, day = parts
    elif len(str(parts[2])) == 4:
        day, month, year = parts
    else:
        return None
    if not (1 <= month <= 12 and 1 <= day <= 31 and 1 <= year <= 9999):
        return None
    return year, month, day


def official_on_time(actual: Any, expected: Any) -> bool | None:
    actual_date = comparable_date(actual)
    expected_date = comparable_date(expected)
    if actual_date is None or expected_date is None:
        return None
    # Never compare a Hijri date to a Gregorian date without an explicit
    # calendar conversion supplied by the source.
    if (actual_date[0] >= 1700) != (expected_date[0] >= 1700):
        return None
    return actual_date <= expected_date


def term_key(year: int, semester: int) -> int:
    return year * 10 + semester


def split_term(key: int) -> tuple[int, int]:
    return key // 10, key % 10


def next_regular_term(key: int) -> int:
    year, semester = split_term(key)
    if semester == 1:
        return term_key(year, 2)
    return term_key(year + 1, 1)


def canonical_branch(campus: Any) -> str | None:
    value = normalized(campus)
    if value.startswith("تربه"):
        return "تربة"
    if value.startswith("رنيه"):
        return "رنية"
    if value.startswith("الخرمه"):
        return "الخرمة"
    if value.startswith("حويه"):
        return "الحوية"
    return None


def gender_from_campus(campus: Any) -> str:
    value = normalized(campus)
    if "طالبات" in value:
        return "أنثى"
    if "طلاب" in value:
        return "ذكر"
    return ""


def sha256_file(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as stream:
        for block in iter(lambda: stream.read(1024 * 1024), b""):
            digest.update(block)
    return digest.hexdigest()


@dataclass(frozen=True)
class Plan:
    version: int
    total_hours: int
    courses: Mapping[str, Mapping[str, Any]]
    all_codes: frozenset[str]
    level_one_core: frozenset[str]
    specialty_core: frozenset[str]
    optional_pool: frozenset[str]


@dataclass
class StudentProfile:
    student_id: str
    names: Counter[str] = field(default_factory=Counter)
    all_terms: set[int] = field(default_factory=set)
    program_evidence: dict[str, set[str]] = field(default_factory=lambda: defaultdict(set))
    islamic_codes: set[str] = field(default_factory=set)
    term_branch_codes: dict[int, dict[str, set[str]]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(set))
    )
    term_branch_exclusive: dict[int, dict[str, set[str]]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(set))
    )
    term_branch_awarded: dict[int, dict[str, set[str]]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(set))
    )
    term_branch_campuses: dict[int, dict[str, Counter[str]]] = field(
        default_factory=lambda: defaultdict(lambda: defaultdict(Counter))
    )
    passed_terms: dict[str, set[int]] = field(default_factory=lambda: defaultdict(set))

    @property
    def name(self) -> str:
        if not self.names:
            return ""
        return sorted(self.names.items(), key=lambda item: (-item[1], item[0]))[0][0]


@dataclass(frozen=True)
class PlanBundle:
    plans: Mapping[int, Plan]
    program_courses: Mapping[str, frozenset[str]]
    exclusive_by_program: Mapping[str, frozenset[str]]
    islamic_exclusive: frozenset[str]
    islamic_union: frozenset[str]
    old_only: frozenset[str]
    new_only: frozenset[str]
    ambiguous_identity_codes: frozenset[str]
    equivalencies: Mapping[int, Mapping[str, frozenset[str]]]


def load_plan_bundle(path: Path) -> PlanBundle:
    payload = json.loads(path.read_text(encoding="utf-8"))
    program_courses: dict[str, set[str]] = defaultdict(set)
    code_names: dict[str, set[str]] = defaultdict(set)
    selected: dict[int, Plan] = {}

    for program in payload.get("programs") or []:
        if "بكالوريوس" not in normalized(program.get("degree")):
            continue
        program_name = clean(program.get("name"))
        if not program_name:
            continue
        course_map: dict[str, dict[str, Any]] = {}
        for course in program.get("courses") or []:
            code = normalize_code(course.get("code"))
            if not code:
                continue
            name_key = normalized(course.get("name"))
            if name_key:
                code_names[code].add(name_key)
            program_courses[program_name].add(code)
            course_map.setdefault(code, dict(course))

        if normalized(program_name) == normalized(PROGRAM_AR):
            version = int(program.get("version") or 0)
            all_codes = frozenset(course_map)
            level_one_core = frozenset(
                code
                for code, row in course_map.items()
                if int(row.get("level") or 0) == 1
                and normalized(row.get("category")) == normalized("إجبارية القسم")
            )
            specialty_core = frozenset(
                code
                for code, row in course_map.items()
                if normalized(row.get("category")) == normalized("إجبارية القسم")
            )
            optional_pool = frozenset(
                code for code, row in course_map.items() if row.get("optional_pool") is True
            )
            selected[version] = Plan(
                version=version,
                total_hours=int(program.get("total_hours") or 0),
                courses=course_map,
                all_codes=all_codes,
                level_one_core=level_one_core,
                specialty_core=specialty_core,
                optional_pool=optional_pool,
            )

    if 39 not in selected or 47 not in selected:
        raise RuntimeError("لم تُوجد خطتا الدراسات الإسلامية 39 و47 في data.json.")

    ambiguous = frozenset(code for code, names in code_names.items() if len(names) > 1)
    clean_program_courses = {
        program: frozenset(codes - set(ambiguous)) for program, codes in program_courses.items()
    }
    course_programs: dict[str, set[str]] = defaultdict(set)
    for program, codes in clean_program_courses.items():
        for code in codes:
            course_programs[code].add(program)
    exclusive_by_program = {
        program: frozenset(code for code in codes if course_programs[code] == {program})
        for program, codes in clean_program_courses.items()
    }
    islamic_exclusive = exclusive_by_program.get(PROGRAM_AR, frozenset())
    islamic_union = selected[39].all_codes | selected[47].all_codes

    equivalencies: dict[int, dict[str, set[str]]] = defaultdict(lambda: defaultdict(set))
    for source_code, rules in (payload.get("equivalencies") or {}).items():
        source = normalize_code(source_code)
        for rule in rules or []:
            if normalized(rule.get("program")) != normalized(PROGRAM_AR):
                continue
            if "بكالوريوس" not in normalized(rule.get("degree")):
                continue
            try:
                version = int(rule.get("version") or 0)
            except (TypeError, ValueError):
                continue
            for option in rule.get("options") or []:
                code = normalize_code(option.get("code") if isinstance(option, Mapping) else option)
                if code:
                    equivalencies[version][source].add(code)

    return PlanBundle(
        plans=selected,
        program_courses=clean_program_courses,
        exclusive_by_program=exclusive_by_program,
        islamic_exclusive=islamic_exclusive,
        islamic_union=islamic_union,
        old_only=frozenset(selected[39].all_codes - selected[47].all_codes),
        new_only=frozenset(selected[47].all_codes - selected[39].all_codes),
        ambiguous_identity_codes=ambiguous,
        equivalencies={
            version: {code: frozenset(options) for code, options in rules.items()}
            for version, rules in equivalencies.items()
        },
    )


def open_readonly_database(path: Path) -> sqlite3.Connection:
    # SQLite's URI parser cannot reliably reopen decomposed Arabic macOS paths.
    # A normal path plus query_only keeps the source immutable while remaining
    # portable across those filesystems.
    connection = sqlite3.connect(str(path.resolve()))
    connection.execute("PRAGMA query_only = ON")
    connection.row_factory = sqlite3.Row
    return connection


def read_profiles(
    database: Path,
    plans: PlanBundle,
) -> tuple[dict[str, StudentProfile], dict[tuple[int, str, int], dict[str, Any]], dict[str, Any]]:
    profiles: dict[str, StudentProfile] = {}
    coverage: dict[tuple[int, str, int], dict[str, Any]] = defaultdict(
        lambda: {
            "section_ids": set(),
            "enrollments": 0,
            "student_ids": set(),
            "islamic_course_codes": set(),
            "islamic_student_ids": set(),
        }
    )
    logical_rows: set[tuple[Any, ...]] = set()
    grade_counts: Counter[str] = Counter()
    database_summary: dict[str, Any] = {}

    connection = open_readonly_database(database)
    try:
        database_summary = {
            "sources": int(connection.execute("SELECT COUNT(*) FROM sources").fetchone()[0]),
            "sections": int(connection.execute("SELECT COUNT(*) FROM sections").fetchone()[0]),
            "enrollments": int(connection.execute("SELECT COUNT(*) FROM enrollments").fetchone()[0]),
            "students": int(connection.execute("SELECT COUNT(DISTINCT student_id) FROM enrollments").fetchone()[0]),
            "missing_total_score": int(
                connection.execute("SELECT COUNT(*) FROM enrollments WHERE total_score IS NULL").fetchone()[0]
            ),
            "missing_semester_sections": int(
                connection.execute(
                    "SELECT COUNT(*) FROM sections WHERE semester IS NULL OR TRIM(semester) = ''"
                ).fetchone()[0]
            ),
        }
        query = """
            SELECT
                s.id AS section_id,
                s.source_id,
                s.academic_year,
                s.semester,
                s.campus,
                s.activity,
                s.section_no,
                s.course_code,
                s.course_name,
                e.student_id,
                e.student_name,
                e.total_score,
                e.grade
            FROM sections s
            JOIN enrollments e ON e.section_id = s.id
            ORDER BY s.academic_year, s.semester, s.id, e.id
        """
        for row in connection.execute(query):
            sid = normalize_student_id(row["student_id"])
            code = normalize_code(row["course_code"])
            year = full_year(row["academic_year"])
            semester = semester_number(row["semester"])
            branch = canonical_branch(row["campus"])
            grade = normalize_grade(row["grade"])
            if not sid or not code:
                continue

            # Defensive cross-source de-duplication.  The current snapshot has
            # no duplicates at this grain, but the invariant protects reruns
            # after additional overlapping source files are supplied.
            logical_key = (
                year,
                semester,
                normalized(row["campus"]),
                normalized(row["activity"]),
                clean(row["section_no"]),
                code,
                sid,
            )
            if logical_key in logical_rows:
                continue
            logical_rows.add(logical_key)
            grade_counts[grade] += 1

            profile = profiles.setdefault(sid, StudentProfile(student_id=sid))
            student_name = clean(row["student_name"])
            if student_name:
                profile.names[student_name] += 1

            key: int | None = None
            if year is not None and semester is not None:
                key = term_key(year, semester)
                profile.all_terms.add(key)

            for program, exclusive_codes in plans.exclusive_by_program.items():
                if code in exclusive_codes:
                    profile.program_evidence[program].add(code)

            if code in plans.islamic_union:
                profile.islamic_codes.add(code)
                if key is not None and branch is not None:
                    profile.term_branch_codes[key][branch].add(code)
                    profile.term_branch_campuses[key][branch][clean(row["campus"])] += 1
                    if code in plans.islamic_exclusive:
                        profile.term_branch_exclusive[key][branch].add(code)
                    if grade in AWARDED_GRADES:
                        profile.term_branch_awarded[key][branch].add(code)
                if key is not None and grade in PASS_GRADES:
                    profile.passed_terms[code].add(key)

            if year in TARGET_YEARS and semester is not None and branch is not None:
                cell = coverage[(year, branch, semester)]
                cell["section_ids"].add(int(row["section_id"]))
                cell["enrollments"] += 1
                cell["student_ids"].add(sid)
                if code in plans.islamic_union:
                    cell["islamic_course_codes"].add(code)
                    cell["islamic_student_ids"].add(sid)
    finally:
        connection.close()

    database_summary["logical_enrollments"] = len(logical_rows)
    database_summary["grade_counts"] = dict(sorted(grade_counts.items()))
    database_summary["ambiguous_grade_rows"] = sum(grade_counts[grade] for grade in AMBIGUOUS_GRADES)
    return profiles, coverage, database_summary


def program_membership(profile: StudentProfile) -> dict[str, Any]:
    counts = {program: len(codes) for program, codes in profile.program_evidence.items() if codes}
    islamic = counts.get(PROGRAM_AR, 0)
    competitors = sorted(
        ((count, program) for program, count in counts.items() if program != PROGRAM_AR),
        reverse=True,
    )
    second = competitors[0][0] if competitors else 0
    max_score = max(counts.values(), default=0)
    winners = sorted(program for program, count in counts.items() if count == max_score and count > 0)
    is_winner = winners == [PROGRAM_AR]
    margin = islamic - second
    confidence = None
    if is_winner and islamic >= 3 and margin >= 2:
        confidence = "very_high"
    elif is_winner and islamic >= 2 and margin >= 1:
        confidence = "high"
    return {
        "confidence": confidence,
        "islamic_evidence": islamic,
        "runner_up_evidence": second,
        "margin": margin,
        "winner": winners[0] if len(winners) == 1 else None,
        "ambiguous": len(winners) > 1,
    }


def assign_plan(profile: StudentProfile, plans: PlanBundle) -> dict[str, Any]:
    scores = {
        39: len(profile.islamic_codes & plans.old_only),
        47: len(profile.islamic_codes & plans.new_only),
    }
    winner = 39 if scores[39] > scores[47] else 47 if scores[47] > scores[39] else None
    if winner is None:
        return {"version": None, "confidence": None, "scores": scores, "margin": 0}
    loser = 47 if winner == 39 else 39
    margin = scores[winner] - scores[loser]
    confidence = None
    if scores[winner] >= 3 and margin >= 2:
        confidence = "very_high"
    elif scores[winner] >= 2 and margin >= 1:
        confidence = "high"
    return {"version": winner, "confidence": confidence, "scores": scores, "margin": margin}


def branch_evidence(profile: StudentProfile, terms: Sequence[int]) -> dict[str, Any]:
    all_items: dict[str, set[tuple[int, str]]] = defaultdict(set)
    exclusive_items: dict[str, set[tuple[int, str]]] = defaultdict(set)
    awarded_items: dict[str, set[tuple[int, str]]] = defaultdict(set)
    campuses: dict[str, Counter[str]] = defaultdict(Counter)
    for key in terms:
        for branch, codes in profile.term_branch_codes.get(key, {}).items():
            all_items[branch].update((key, code) for code in codes)
        for branch, codes in profile.term_branch_exclusive.get(key, {}).items():
            exclusive_items[branch].update((key, code) for code in codes)
        for branch, codes in profile.term_branch_awarded.get(key, {}).items():
            awarded_items[branch].update((key, code) for code in codes)
        for branch, values in profile.term_branch_campuses.get(key, {}).items():
            campuses[branch].update(values)

    counts = {branch: len(items) for branch, items in all_items.items()}
    if not counts:
        return {
            "branch": None,
            "confidence": None,
            "course_count": 0,
            "exclusive_count": 0,
            "awarded_count": 0,
            "share": 0.0,
            "margin": 0,
            "gender": "",
        }
    ordered = sorted(counts.items(), key=lambda item: (-item[1], BRANCH_ORDER.index(item[0])))
    branch, top = ordered[0]
    second = ordered[1][1] if len(ordered) > 1 else 0
    total = sum(counts.values())
    share = top / total if total else 0.0
    exclusive_count = len(exclusive_items.get(branch, set()))
    awarded_count = len(awarded_items.get(branch, set()))
    margin = top - second
    confidence = None
    if top >= 3 and exclusive_count >= 2 and share >= 0.80 and margin >= 2:
        confidence = "very_high"
    elif top >= 2 and exclusive_count >= 1 and share >= (2 / 3) and margin >= 1:
        confidence = "high"

    common_campus = ""
    if campuses.get(branch):
        common_campus = campuses[branch].most_common(1)[0][0]
    return {
        "branch": branch,
        "confidence": confidence,
        "course_count": top,
        "exclusive_count": exclusive_count,
        "awarded_count": awarded_count,
        "share": round(share, 4),
        "margin": margin,
        "gender": gender_from_campus(common_campus),
    }


def official_rows(path: Path) -> list[dict[str, str]]:
    if not path.is_file():
        return []
    with path.open("r", encoding="utf-8-sig", newline="") as stream:
        return [dict(row) for row in csv.DictReader(stream, delimiter=";")]


def official_regular_by_year(path: Path) -> dict[int, dict[str, dict[str, str]]]:
    result: dict[int, dict[str, dict[str, str]]] = defaultdict(dict)
    for row in official_rows(path):
        year = full_year(row.get("السنة"))
        sid = normalize_student_id(row.get("الرقم_الجامعي"))
        if year not in TARGET_YEARS or not sid:
            continue
        if normalized(row.get("التخصص")) != normalized(PROGRAM_AR):
            continue
        if "بكالوريوس" not in normalized(row.get("الدرجة")):
            continue
        if normalized(row.get("الحالة")) != normalized("منتظم"):
            continue
        result[year][sid] = row
    return result


def official_graduates_by_year(path: Path) -> dict[int, dict[str, dict[str, str]]]:
    result: dict[int, dict[str, dict[str, str]]] = defaultdict(dict)
    for row in official_rows(path):
        year = full_year(row.get("السنة"))
        sid = normalize_student_id(row.get("الرقم_الجامعي"))
        if year not in TARGET_YEARS or not sid:
            continue
        if normalized(row.get("التخصص")) != normalized(PROGRAM_AR):
            continue
        if "بكالوريوس" not in normalized(row.get("الدرجة")):
            continue
        result[year][sid] = row
    return result


def deduplicate_official_program_rows(
    path: Path,
    *,
    year_column: str,
) -> tuple[dict[int, dict[str, dict[str, str]]], dict[str, Any]]:
    """Filter and deduplicate official program rows by academic year and ID.

    Conflicting duplicates are excluded instead of choosing an official status,
    name, or gender arbitrarily.  Complementary blank fields are allowed and
    the most complete otherwise-equivalent row is retained deterministically.
    """

    source = official_rows(path)
    groups: dict[tuple[int, str], list[dict[str, str]]] = defaultdict(list)
    eligible_by_year: Counter[int] = Counter()
    missing_id_by_year: Counter[int] = Counter()
    for row in source:
        year = full_year(row.get(year_column))
        if year not in TARGET_YEARS:
            continue
        if normalized(row.get("التخصص")) != normalized(PROGRAM_AR):
            continue
        if "بكالوريوس" not in normalized(row.get("الدرجة")):
            continue
        eligible_by_year[year] += 1
        sid = normalize_student_id(row.get("الرقم_الجامعي"))
        if not sid:
            missing_id_by_year[year] += 1
            continue
        groups[(year, sid)].append(row)

    selected: dict[int, dict[str, dict[str, str]]] = defaultdict(dict)
    conflicting_keys_by_year: Counter[int] = Counter()
    conflicting_rows_by_year: Counter[int] = Counter()
    identical_dropped_by_year: Counter[int] = Counter()
    duplicate_rows_by_year: Counter[int] = Counter()
    unique_keys_by_year: Counter[int] = Counter()
    material_fields = ("الاسم", "الحالة", "الجنس")

    for (year, sid), rows in sorted(groups.items()):
        unique_keys_by_year[year] += 1
        duplicate_rows_by_year[year] += len(rows) - 1
        conflicts = any(
            len({normalized(row.get(field)) for row in rows if clean(row.get(field))}) > 1
            for field in material_fields
        )
        if conflicts:
            conflicting_keys_by_year[year] += 1
            conflicting_rows_by_year[year] += len(rows)
            continue
        chosen = sorted(
            rows,
            key=lambda row: (
                -sum(bool(clean(value)) for value in row.values()),
                json.dumps(row, ensure_ascii=False, sort_keys=True),
            ),
        )[0]
        selected[year][sid] = chosen
        identical_dropped_by_year[year] += len(rows) - 1

    by_year: dict[str, dict[str, int]] = {}
    for year in TARGET_YEARS:
        keyed_rows = sum(len(rows) for (row_year, _sid), rows in groups.items() if row_year == year)
        reconciled = (
            missing_id_by_year[year]
            + conflicting_rows_by_year[year]
            + len(selected.get(year, {}))
            + identical_dropped_by_year[year]
        )
        if reconciled != eligible_by_year[year]:
            raise RuntimeError(f"فشلت مصالحة التكرار الرسمية سنة {short_year(year)}.")
        by_year[short_year(year)] = {
            "eligible_program_rows": eligible_by_year[year],
            "missing_student_id_rows": missing_id_by_year[year],
            "keyed_rows": keyed_rows,
            "unique_keys_before_conflicts": unique_keys_by_year[year],
            "duplicate_rows_total": duplicate_rows_by_year[year],
            "conflicting_keys_excluded": conflicting_keys_by_year[year],
            "conflicting_rows_excluded": conflicting_rows_by_year[year],
            "identical_duplicate_rows_dropped": identical_dropped_by_year[year],
            "deduplicated_records": len(selected.get(year, {})),
        }

    totals = {
        field: sum(year_stats[field] for year_stats in by_year.values())
        for field in next(iter(by_year.values()), {})
    }
    return selected, {
        "source_available": path.is_file(),
        "source_rows_total": len(source),
        "year_column": year_column,
        "by_year": by_year,
        "totals": totals,
    }


def link_official_rows_to_branches(
    rows_by_year: Mapping[int, Mapping[str, Mapping[str, str]]],
    profiles: Mapping[str, StudentProfile],
) -> tuple[list[dict[str, Any]], dict[str, Any]]:
    """Attach a partial, evidence-backed branch to official event rows."""

    records: list[dict[str, Any]] = []
    reason_by_year: dict[str, Counter[str]] = {
        short_year(year): Counter() for year in TARGET_YEARS
    }
    for year in TARGET_YEARS:
        year_key = short_year(year)
        through = term_key(year, 3)
        for sid, official in rows_by_year.get(year, {}).items():
            profile = profiles.get(sid)
            if profile is None:
                reason_by_year[year_key]["missing_from_result_rosters"] += 1
                continue
            terms = observed_plan_terms(profile, through)
            if not terms:
                reason_by_year[year_key]["no_observed_plan_term_by_year_end"] += 1
                continue
            window = terms[-2:]
            branch = branch_evidence(profile, window)
            branch_name = branch["branch"]
            if branch_name not in BRANCH_ORDER or branch["confidence"] is None:
                reason_by_year[year_key]["no_strong_branch_evidence"] += 1
                continue
            records.append(
                {
                    "year": year - 1400,
                    "student_id": sid,
                    "name": clean(official.get("الاسم")) or profile.name,
                    "program": PROGRAM_AR,
                    "degree": DEGREE_AR,
                    "department": clean(official.get("القسم")) or DEPARTMENT_AR,
                    "branch": branch_name,
                    "gender": clean(official.get("الجنس")) or branch["gender"],
                    "status": clean(official.get("الحالة")),
                    "gpa": clean(official.get("المعدل")),
                    "event_confidence": "official",
                    "branch_confidence": branch["confidence"],
                    "confidence": "official",
                    "_evidence": {
                        "branch_terms": window,
                        "branch": branch,
                    },
                }
            )

    records.sort(
        key=lambda row: (
            int(row["year"]),
            BRANCH_ORDER.index(str(row["branch"])),
            normalized(row.get("status")),
            str(row["student_id"]),
        )
    )
    by_year: dict[str, dict[str, Any]] = {}
    for year in TARGET_YEARS:
        year_key = short_year(year)
        source_records = len(rows_by_year.get(year, {}))
        matched = sum(row["year"] == year - 1400 for row in records)
        reasons = dict(sorted(reason_by_year[year_key].items()))
        if matched + sum(reasons.values()) != source_records:
            raise RuntimeError(f"فشلت مصالحة ربط الفرع سنة {year_key}.")
        by_year[year_key] = {
            "deduplicated_official_records": source_records,
            "branch_matched": matched,
            "without_strong_branch": source_records - matched,
            "unmatched_reasons": reasons,
        }
    return records, {
        "semantics": "partial_official_records_with_inferred_branch_not_an_official_total",
        "by_year": by_year,
        "totals": {
            "deduplicated_official_records": sum(
                row["deduplicated_official_records"] for row in by_year.values()
            ),
            "branch_matched": len(records),
            "without_strong_branch": sum(row["without_strong_branch"] for row in by_year.values()),
        },
    }


def base_student_record(
    *,
    year: int,
    semester: int,
    sid: str,
    name: str,
    branch: str,
    gender: str,
    confidence: str,
    event_confidence: str,
    branch_confidence: str,
    plan_version: int | None,
) -> dict[str, Any]:
    return {
        "year": year - 1400,
        "semester": semester,
        "student_id": sid,
        "name": name,
        "program": PROGRAM_AR,
        "degree": DEGREE_AR,
        "department": DEPARTMENT_AR,
        "branch": branch,
        "gender": gender,
        "confidence": confidence,
        "event_confidence": event_confidence,
        "branch_confidence": branch_confidence,
        "plan_version": plan_version,
    }


def choose_confidence(*values: str | None) -> str | None:
    if any(value is None for value in values):
        return None
    return "very_high" if all(value == "very_high" for value in values) else "high"


def infer_regular_students(
    profiles: Mapping[str, StudentProfile],
    plans: PlanBundle,
    official_regular: Mapping[int, Mapping[str, Mapping[str, str]]],
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for year in TARGET_YEARS:
        key = term_key(year, 1)
        for sid, profile in profiles.items():
            membership = program_membership(profile)
            official = official_regular.get(year, {}).get(sid)
            branch = branch_evidence(profile, [key])
            branch_name = branch["branch"]
            if branch_name not in BRANCH_ORDER or branch["confidence"] is None:
                continue
            policy = COVERAGE_POLICY[year][branch_name]
            if not policy["allow_regular"]:
                continue
            if official:
                confidence = "official"
            else:
                if membership["confidence"] is None:
                    continue
                if branch["course_count"] < 2 or branch["awarded_count"] < 1:
                    continue
                confidence = choose_confidence(membership["confidence"], branch["confidence"])
                if confidence == "very_high" and (
                    branch["course_count"] < 3 or branch["awarded_count"] < 2
                ):
                    confidence = "high"
            plan = assign_plan(profile, plans)
            name = clean(official.get("الاسم")) if official else profile.name
            gender = clean(official.get("الجنس")) if official else branch["gender"]
            record = base_student_record(
                year=year,
                semester=1,
                sid=sid,
                name=name,
                branch=branch_name,
                gender=gender,
                confidence=confidence,
                event_confidence="official" if official else confidence,
                branch_confidence=branch["confidence"],
                plan_version=plan["version"],
            )
            record["official_status"] = clean(official.get("الحالة")) if official else ""
            record["_evidence"] = {
                "program": membership,
                "branch": branch,
                "plan": plan,
                "metric_semantics": "official_regular" if official else "active_results_proxy",
            }
            records.append(record)
    return deduplicate(records)


def infer_entrants(
    profiles: Mapping[str, StudentProfile],
    plans: PlanBundle,
) -> list[dict[str, Any]]:
    records: list[dict[str, Any]] = []
    for sid, profile in profiles.items():
        if not profile.all_terms:
            continue
        first = min(profile.all_terms)
        year, semester = split_term(first)
        if year not in TARGET_YEARS or semester != 1:
            continue
        membership = program_membership(profile)
        plan_evidence = assign_plan(profile, plans)
        version = plan_evidence["version"]
        if membership["confidence"] is None or plan_evidence["confidence"] is None or version is None:
            continue
        branch = branch_evidence(profile, [first])
        branch_name = branch["branch"]
        if branch_name not in BRANCH_ORDER or branch["confidence"] is None:
            continue
        if not COVERAGE_POLICY[year][branch_name]["allow_entrant"]:
            continue
        plan = plans.plans[version]
        level_one_at_branch = len(
            profile.term_branch_codes.get(first, {}).get(branch_name, set()) & plan.level_one_core
        )
        if level_one_at_branch < 1 or branch["course_count"] < 2:
            continue
        next_key = next_regular_term(first)
        next_branch = branch_evidence(profile, [next_key])
        continued_same_branch = (
            next_branch["branch"] == branch_name and next_branch["course_count"] >= 1
        )
        confidence = choose_confidence(
            membership["confidence"], plan_evidence["confidence"], branch["confidence"]
        )
        if confidence == "very_high" and not (
            level_one_at_branch >= 2 and continued_same_branch
        ):
            confidence = "high"
        if confidence is None:
            continue
        record = base_student_record(
            year=year,
            semester=1,
            sid=sid,
            name=profile.name,
            branch=branch_name,
            gender=branch["gender"],
            confidence=confidence,
            event_confidence=confidence,
            branch_confidence=branch["confidence"],
            plan_version=version,
        )
        record["official"] = False
        record["_evidence"] = {
            "program": membership,
            "branch": branch,
            "plan": plan_evidence,
            "first_seen_term": first,
            "level_one_core_courses": level_one_at_branch,
            "continued_in_next_term_same_branch": continued_same_branch,
        }
        records.append(record)
    return deduplicate(records)


def observed_plan_terms(profile: StudentProfile, through: int) -> list[int]:
    return sorted(
        key
        for key, branches in profile.term_branch_codes.items()
        if key <= through and any(branches.values())
    )


def official_graduate_records(
    profiles: Mapping[str, StudentProfile],
    plans: PlanBundle,
    official_graduates: Mapping[int, Mapping[str, Mapping[str, str]]],
) -> tuple[list[dict[str, Any]], dict[int, int]]:
    records: list[dict[str, Any]] = []
    unassigned: dict[int, int] = defaultdict(int)
    for year in TARGET_YEARS:
        for sid, official in official_graduates.get(year, {}).items():
            profile = profiles.get(sid)
            if profile is None:
                unassigned[year] += 1
                continue
            terms = observed_plan_terms(profile, term_key(year, 3))
            window = terms[-2:]
            branch = branch_evidence(profile, window)
            branch_name = branch["branch"]
            if branch_name not in BRANCH_ORDER or branch["confidence"] is None:
                unassigned[year] += 1
                continue
            semester = split_term(window[-1])[1] if window else 2
            plan = assign_plan(profile, plans)
            record = base_student_record(
                year=year,
                semester=semester,
                sid=sid,
                name=clean(official.get("الاسم")) or profile.name,
                branch=branch_name,
                gender=clean(official.get("الجنس")) or branch["gender"],
                confidence="official",
                event_confidence="official",
                branch_confidence=branch["confidence"],
                plan_version=plan["version"],
            )
            record.update(
                {
                    "official": True,
                    "official_year": year - 1400,
                    "official_admission_date": clean(official.get("تاريخ_القبول")),
                    "official_graduation_date": clean(official.get("تاريخ_التخرج")),
                    "official_graduation_expected_date": clean(
                        official.get("تاريخ_التخرج_المتوقع")
                    ),
                    "official_on_time": official_on_time(
                        official.get("تاريخ_التخرج"),
                        official.get("تاريخ_التخرج_المتوقع"),
                    ),
                    "official_gpa": clean(official.get("المعدل")),
                    "_evidence": {
                        "branch": branch,
                        "branch_terms": window,
                        "plan": plan,
                    },
                }
            )
            records.append(record)
    return deduplicate(records), dict(unassigned)


def satisfaction_term(
    profile: StudentProfile,
    requirement: str,
    equivalencies: Mapping[str, frozenset[str]],
) -> int | None:
    accepted = {requirement, *equivalencies.get(requirement, frozenset())}
    terms = [key for code in accepted for key in profile.passed_terms.get(code, set())]
    return min(terms) if terms else None


def inferred_plan39_graduates(
    profiles: Mapping[str, StudentProfile],
    plans: PlanBundle,
    official_ids: set[str],
    available_terms: set[int],
    observed_coverage: Mapping[tuple[int, str, int], Mapping[str, Any]],
) -> list[dict[str, Any]]:
    plan = plans.plans[39]
    rules = plans.equivalencies.get(39, {})
    # The old plan has 142 fixed hours and a 12-hour optional bank.  Its
    # published total is 146, hence two two-hour electives are required.
    non_optional_hours = sum(
        int(row.get("hours") or 0)
        for code, row in plan.courses.items()
        if code not in plan.optional_pool
    )
    elective_hours_required = max(0, plan.total_hours - non_optional_hours)
    elective_course_hours = sorted(
        {int(plan.courses[code].get("hours") or 0) for code in plan.optional_pool if code in plan.courses}
    )
    if elective_hours_required and elective_course_hours != [2]:
        raise RuntimeError("تعذّر اشتقاق عدد اختيارات الخطة 39 بصورة وحيدة.")
    elective_required = math.ceil(elective_hours_required / 2) if elective_hours_required else 0

    records: list[dict[str, Any]] = []
    for sid, profile in profiles.items():
        if sid in official_ids:
            continue
        membership = program_membership(profile)
        plan_evidence = assign_plan(profile, plans)
        if membership["confidence"] is None:
            continue
        if plan_evidence["version"] != 39 or plan_evidence["confidence"] is None:
            continue

        required_terms: dict[str, int] = {}
        missing: list[str] = []
        fixed_requirements = plan.all_codes - plan.optional_pool
        for code in sorted(fixed_requirements):
            key = satisfaction_term(profile, code, rules)
            if key is None:
                missing.append(code)
            else:
                required_terms[code] = key
        if missing:
            continue

        elective_terms = sorted(
            key
            for code in plan.optional_pool
            if (key := satisfaction_term(profile, code, rules)) is not None
        )
        if len(elective_terms) < elective_required:
            continue
        completion_candidates = list(required_terms.values())
        if elective_required:
            completion_candidates.append(elective_terms[elective_required - 1])
        if not completion_candidates:
            continue
        completion = max(completion_candidates)
        year, semester = split_term(completion)
        if year not in TARGET_YEARS:
            continue
        terms = observed_plan_terms(profile, completion)
        branch = branch_evidence(profile, terms[-2:])
        branch_name = branch["branch"]
        if branch_name not in BRANCH_ORDER or branch["confidence"] is None:
            continue
        if not COVERAGE_POLICY[year][branch_name]["allow_inferred_graduate"]:
            continue
        next_key = next_regular_term(completion)
        if next_key in profile.all_terms:
            continue
        next_term_available = next_key in available_terms
        next_year, next_semester = split_term(next_key)
        next_cell = observed_coverage.get((next_year, branch_name, next_semester), {})
        next_term_branch_observed = bool(next_cell.get("section_ids")) and bool(
            next_cell.get("islamic_course_codes")
        )
        next_coverage_usable = (
            next_year in COVERAGE_POLICY
            and COVERAGE_POLICY[next_year][branch_name]["allow_regular"]
        )
        if not (next_term_available and next_term_branch_observed and next_coverage_usable):
            continue
        confidence = choose_confidence(
            membership["confidence"], plan_evidence["confidence"], branch["confidence"]
        )
        if confidence is None:
            continue
        record = base_student_record(
            year=year,
            semester=semester,
            sid=sid,
            name=profile.name,
            branch=branch_name,
            gender=branch["gender"],
            confidence=confidence,
            event_confidence=confidence,
            branch_confidence=branch["confidence"],
            plan_version=39,
        )
        record.update(
            {
                "official": False,
                "official_year": None,
                "official_admission_date": "",
                "official_graduation_date": "",
                "official_graduation_expected_date": "",
                "official_on_time": None,
                "official_gpa": "",
                "_evidence": {
                    "program": membership,
                    "branch": branch,
                    "plan": plan_evidence,
                    "fixed_plan_courses_required": len(fixed_requirements),
                    "fixed_plan_courses_passed": len(required_terms),
                    "electives_required": elective_required,
                    "electives_passed": len(elective_terms),
                    "completion_term": completion,
                    "next_term": next_key,
                    "present_next_term_any_course": False,
                    "next_term_available": next_term_available,
                    "next_term_branch_observed": next_term_branch_observed,
                    "next_term_branch_coverage_usable": next_coverage_usable,
                    "plan47_graduation_inference_disabled": True,
                },
            }
        )
        records.append(record)
    return deduplicate(records)


def deduplicate(records: Iterable[dict[str, Any]]) -> list[dict[str, Any]]:
    selected: dict[tuple[int, str], dict[str, Any]] = {}
    for record in records:
        key = (int(record["year"]), str(record["student_id"]))
        previous = selected.get(key)
        if previous is None or CONFIDENCE_RANK[record["confidence"]] > CONFIDENCE_RANK[previous["confidence"]]:
            selected[key] = record
    return sorted(
        selected.values(),
        key=lambda row: (
            int(row["year"]),
            BRANCH_ORDER.index(row["branch"]),
            int(row.get("semester") or 0),
            str(row["student_id"]),
        ),
    )


def build_coverage(
    observed: Mapping[tuple[int, str, int], Mapping[str, Any]],
    profiles: Mapping[str, StudentProfile],
) -> dict[str, dict[str, Any]]:
    output: dict[str, dict[str, Any]] = {}
    for year in TARGET_YEARS:
        year_output: dict[str, Any] = {}
        for branch in BRANCH_ORDER:
            policy = dict(COVERAGE_POLICY[year][branch])
            semesters: dict[str, Any] = {}
            for semester in (1, 2):
                cell = observed.get((year, branch, semester), {})
                student_ids = set(cell.get("student_ids") or set())
                islamic_ids = set(cell.get("islamic_student_ids") or set())
                scoped = 0
                for sid in islamic_ids:
                    profile = profiles.get(sid)
                    if profile and program_membership(profile)["confidence"] is not None:
                        scoped += 1
                semesters[str(semester)] = {
                    "sections": len(cell.get("section_ids") or set()),
                    "enrollments": int(cell.get("enrollments") or 0),
                    "students": len(student_ids),
                    "islamic_plan_course_codes": len(cell.get("islamic_course_codes") or set()),
                    "islamic_scope_students": scoped,
                }
            policy["semesters"] = semesters
            year_output[branch] = policy
        output[short_year(year)] = year_output
    return output


def confidence_counts(records: Sequence[Mapping[str, Any]]) -> dict[str, int]:
    values = Counter(str(row.get("confidence") or "") for row in records)
    return {name: values.get(name, 0) for name in ("official", "very_high", "high")}


def metrics_from_records(
    regular_students: Sequence[Mapping[str, Any]],
    entrants: Sequence[Mapping[str, Any]],
    graduates: Sequence[Mapping[str, Any]],
) -> dict[str, dict[str, dict[str, Any]]]:
    result: dict[str, dict[str, dict[str, Any]]] = {}
    for year in (46, 47):
        year_key = str(year)
        result[year_key] = {}
        for branch in BRANCH_ORDER:
            policy = COVERAGE_POLICY[year + 1400][branch]
            regular_slice = [row for row in regular_students if row["year"] == year and row["branch"] == branch]
            entrant_slice = [row for row in entrants if row["year"] == year and row["branch"] == branch]
            graduate_slice = [row for row in graduates if row["year"] == year and row["branch"] == branch]
            male = sum(normalized(row.get("gender")) == normalized("ذكر") for row in regular_slice)
            female = sum(normalized(row.get("gender")) == normalized("أنثى") for row in regular_slice)
            regular_allowed = bool(policy["allow_regular"])
            entrants_allowed = bool(policy["allow_entrant"])
            evaluable_ontime = [
                row for row in graduate_slice if row.get("official_on_time") is not None
            ]
            result[year_key][branch] = {
                "students_total": len(regular_slice) if regular_allowed else None,
                "students_male": male if regular_allowed else None,
                "students_female": female if regular_allowed else None,
                "students_gender_unknown": len(regular_slice) - male - female if regular_allowed else None,
                "students_new": len(entrant_slice) if entrants_allowed else None,
                # The official source has no branch field, so the matched subset
                # must never be presented as the branch's complete graduate total.
                "graduates_total": None,
                "graduates_branch_matched": len(graduate_slice),
                "graduates_ontime": (
                    sum(row["official_on_time"] is True for row in evaluable_ontime)
                    if evaluable_ontime
                    else None
                ),
                "graduates_ontime_denominator": (
                    len(evaluable_ontime) if evaluable_ontime else None
                ),
                "_confidence": {
                    "students_total": confidence_counts(regular_slice),
                    "students_new": confidence_counts(entrant_slice),
                    "graduates_branch_matched": confidence_counts(graduate_slice),
                },
            }
    return result


def load_reference_metrics(path: Path) -> dict[str, dict[str, int]]:
    output: dict[str, dict[str, int]] = {}
    for row in official_rows(path):
        if normalized(row.get("Major_aName")) != normalized(PROGRAM_AR):
            continue
        if "بكالوريوس" not in normalized(row.get("Degree_aName")):
            continue
        year = full_year(row.get("Semester"))
        if year not in TARGET_YEARS:
            continue
        output[short_year(year)] = {
            "students_total": int(float(row.get("students_total") or 0)),
            "students_new": int(float(row.get("students_new") or 0)),
            "graduates_total": int(float(row.get("graduates_total") or 0)),
        }
    return output


def validate_payload(payload: Mapping[str, Any]) -> None:
    if list(payload.get("branch_order") or []) != list(BRANCH_ORDER):
        raise RuntimeError("ترتيب الفروع في الناتج غير صحيح.")
    allowed_confidence = set(CONFIDENCE_RANK)
    event_lists = {
        "students_total": payload.get("regular_students") or [],
        "students_new": payload.get("entrants") or [],
        "graduates_branch_matched": payload.get("graduates") or [],
    }
    for label, records in event_lists.items():
        seen: set[tuple[int, str]] = set()
        for row in records:
            if row.get("confidence") not in allowed_confidence:
                raise RuntimeError(f"درجة ثقة غير مقبولة في {label}.")
            if row.get("event_confidence") not in allowed_confidence:
                raise RuntimeError(f"درجة ثقة حدث غير مقبولة في {label}.")
            if row.get("branch_confidence") not in {"very_high", "high"}:
                raise RuntimeError(f"درجة ثقة فرع غير مقبولة في {label}.")
            if (row.get("confidence") == "official") != (
                row.get("event_confidence") == "official"
            ):
                raise RuntimeError(f"لم تفصل ثقة الحدث عن ثقة الفرع في {label}.")
            if row.get("branch") not in BRANCH_ORDER:
                raise RuntimeError(f"فرع غير مقبول في {label}.")
            key = (int(row.get("year")), str(row.get("student_id")))
            if key in seen:
                raise RuntimeError(f"تكرار طالب في {label}: {key[0]}.")
            seen.add(key)
        for year in (46, 47):
            for branch in BRANCH_ORDER:
                expected = sum(1 for row in records if row["year"] == year and row["branch"] == branch)
                policy = COVERAGE_POLICY[year + 1400][branch]
                if label == "students_total":
                    reportable = bool(policy["allow_regular"])
                elif label == "students_new":
                    reportable = bool(policy["allow_entrant"])
                else:
                    reportable = True
                actual = payload["metrics"][str(year)][branch][label]
                if not reportable:
                    if expected:
                        raise RuntimeError(
                            f"وجدت سجلات {label} في فرع غير مغطى: {year} {branch}."
                        )
                    if actual is not None:
                        raise RuntimeError(
                            f"يجب أن تكون {label} غير متاحة سنة {year} في {branch}."
                        )
                    continue
                if actual is None or int(actual) != expected:
                    raise RuntimeError(
                        f"فشل مصالحة {label} سنة {year} في {branch}: {actual} != {expected}."
                    )

    for year in (46, 47):
        for branch in BRANCH_ORDER:
            cell = payload["metrics"][str(year)][branch]
            if cell["graduates_total"] is not None:
                raise RuntimeError(f"لا يجوز عرض إجمالي خريجي الفرع بلا مصدر رسمي: {year} {branch}.")
            if not COVERAGE_POLICY[year + 1400][branch]["allow_regular"]:
                if any(
                    cell[field] is not None
                    for field in (
                        "students_total",
                        "students_male",
                        "students_female",
                        "students_gender_unknown",
                    )
                ):
                    raise RuntimeError(f"فروع غير المغطاة يجب ألا تظهر كصفر: {year} {branch}.")
            else:
                gender_total = sum(
                    int(cell[field])
                    for field in ("students_male", "students_female", "students_gender_unknown")
                )
                if gender_total != int(cell["students_total"]):
                    raise RuntimeError(f"فشلت مصالحة الجنس سنة {year} في {branch}.")

            graduate_slice = [
                row
                for row in payload.get("graduates") or []
                if row["year"] == year and row["branch"] == branch
            ]
            evaluable = [row for row in graduate_slice if row.get("official_on_time") is not None]
            expected_ontime = sum(row.get("official_on_time") is True for row in evaluable)
            if evaluable:
                if cell["graduates_ontime"] != expected_ontime:
                    raise RuntimeError(f"فشلت مصالحة التخرج في الوقت: {year} {branch}.")
                if cell["graduates_ontime_denominator"] != len(evaluable):
                    raise RuntimeError(f"فشل مقام التخرج في الوقت: {year} {branch}.")
            elif cell["graduates_ontime"] is not None or cell["graduates_ontime_denominator"] is not None:
                raise RuntimeError(f"لا يجوز إظهار التخرج في الوقت بلا تاريخ متوقع: {year} {branch}.")

    for row in payload.get("graduates") or []:
        if not row.get("official") and row.get("plan_version") != 39:
            raise RuntimeError("وُجد خريج مستدل من خطة غير الخطة 39.")
        if not row.get("official") and not COVERAGE_POLICY[int(row["year"]) + 1400][row["branch"]][
            "allow_inferred_graduate"
        ]:
            raise RuntimeError("وُجد خريج مستدل في فرع محظور تغطيته.")
        if not row.get("official"):
            evidence = row.get("_evidence") or {}
            fixed_required = evidence.get("fixed_plan_courses_required")
            fixed_passed = evidence.get("fixed_plan_courses_passed")
            plan_fixed = payload["_internal"]["plan_summary"]["39"]["fixed_courses"]
            if fixed_required != plan_fixed or fixed_passed != fixed_required:
                raise RuntimeError("خريج مستدل لم يكمل كل مقررات الخطة 39 الثابتة.")
            if evidence.get("electives_passed", 0) < evidence.get("electives_required", 0):
                raise RuntimeError("خريج مستدل لم يكمل اختيارات الخطة 39.")
            if not all(
                evidence.get(field) is True
                for field in (
                    "next_term_available",
                    "next_term_branch_observed",
                    "next_term_branch_coverage_usable",
                )
            ):
                raise RuntimeError("خريج مستدل بلا فصل تالٍ مغطى لإثبات الغياب.")

    reconciliation = payload.get("reconciliation") or {}
    partial_linkages = reconciliation.get("partial_official_branch_linkages") or {}
    for list_name in ("noncompleters", "student_detail_branches"):
        rows = payload.get(list_name) or []
        seen: set[tuple[int, str]] = set()
        for row in rows:
            required = (
                "year",
                "student_id",
                "name",
                "program",
                "degree",
                "department",
                "branch",
                "gender",
                "status",
                "gpa",
                "event_confidence",
                "branch_confidence",
                "confidence",
            )
            if any(field not in row for field in required):
                raise RuntimeError(f"حقول ناقصة في {list_name}.")
            if row["year"] not in (46, 47) or row["branch"] not in BRANCH_ORDER:
                raise RuntimeError(f"سنة أو فرع غير مقبول في {list_name}.")
            if row["event_confidence"] != "official" or row["confidence"] != "official":
                raise RuntimeError(f"سجل غير رسمي في {list_name}.")
            if row["branch_confidence"] not in {"very_high", "high"}:
                raise RuntimeError(f"دليل فرع غير كافٍ في {list_name}.")
            key = (int(row["year"]), str(row["student_id"]))
            if key in seen:
                raise RuntimeError(f"تكرار المفتاح year+student_id في {list_name}.")
            seen.add(key)

        linkage = partial_linkages.get(list_name) or {}
        deduplication = linkage.get("deduplication") or {}
        branch_linkage = linkage.get("branch_linkage") or {}
        if branch_linkage.get("totals", {}).get("branch_matched") != len(rows):
            raise RuntimeError(f"فشل إجمالي ربط {list_name}.")
        for year in (46, 47):
            year_key = str(year)
            matched = sum(row["year"] == year for row in rows)
            link_year = branch_linkage.get("by_year", {}).get(year_key, {})
            dedup_year = deduplication.get("by_year", {}).get(year_key, {})
            if link_year.get("branch_matched") != matched:
                raise RuntimeError(f"فشل ربط {list_name} سنة {year}.")
            if (
                int(link_year.get("branch_matched", 0))
                + int(link_year.get("without_strong_branch", 0))
                != int(dedup_year.get("deduplicated_records", 0))
            ):
                raise RuntimeError(f"فشلت مصالحة سجلات {list_name} سنة {year}.")

    for year in (46, 47):
        key = str(year)
        distributed = reconciliation["distributed_totals"][key]
        for metric in event_lists:
            available_values = [
                payload["metrics"][key][branch][metric]
                for branch in BRANCH_ORDER
                if payload["metrics"][key][branch][metric] is not None
            ]
            expected = sum(available_values)
            if distributed[metric] != expected:
                raise RuntimeError(f"فشل إجمالي المصالحة لـ{metric} سنة {year}.")
        if distributed["graduates_total"] is not None:
            raise RuntimeError("إجمالي الخريجين بحسب الفرع غير متاح وليس صفرًا.")


def build_payload(args: argparse.Namespace) -> dict[str, Any]:
    plans = load_plan_bundle(args.plan_data)
    profiles, observed_coverage, database_summary = read_profiles(args.database, plans)
    official_regular = official_regular_by_year(args.students_detail)
    official_graduates = official_graduates_by_year(args.graduates_detail)
    official_noncompleters, noncompleter_deduplication = deduplicate_official_program_rows(
        args.non_completers,
        year_column="آخر_سنة",
    )
    official_student_details, student_detail_deduplication = deduplicate_official_program_rows(
        args.students_detail,
        year_column="السنة",
    )

    regular_students = infer_regular_students(profiles, plans, official_regular)
    entrants = infer_entrants(profiles, plans)
    official_graduate_list, official_unassigned = official_graduate_records(
        profiles, plans, official_graduates
    )
    all_official_graduate_ids = {
        sid for rows in official_graduates.values() for sid in rows
    }
    available_terms = {key for profile in profiles.values() for key in profile.all_terms}
    inferred_graduates = inferred_plan39_graduates(
        profiles,
        plans,
        all_official_graduate_ids,
        available_terms,
        observed_coverage,
    )
    graduates = deduplicate([*official_graduate_list, *inferred_graduates])
    noncompleters, noncompleter_linkage = link_official_rows_to_branches(
        official_noncompleters,
        profiles,
    )
    student_detail_branches, student_detail_linkage = link_official_rows_to_branches(
        official_student_details,
        profiles,
    )
    metrics = metrics_from_records(regular_students, entrants, graduates)
    reference = load_reference_metrics(args.kpi_data)

    official_linkage: dict[str, dict[str, int]] = {}
    for year in TARGET_YEARS:
        year_key = short_year(year)
        first_semester = term_key(year, 1)
        regular_ids = set(official_regular.get(year, {}))
        graduate_ids = set(official_graduates.get(year, {}))
        official_linkage[year_key] = {
            "regular_in_result_rosters": len(regular_ids & set(profiles)),
            "regular_in_first_semester_rosters": sum(
                first_semester in profiles[sid].all_terms for sid in regular_ids & set(profiles)
            ),
            "regular_with_strong_first_semester_branch": sum(
                branch_evidence(profiles[sid], [first_semester])["confidence"] is not None
                for sid in regular_ids & set(profiles)
            ),
            "regular_distributed_as_official": sum(
                row["year"] == year - 1400 and row["confidence"] == "official"
                for row in regular_students
            ),
            "graduates_in_result_rosters": len(graduate_ids & set(profiles)),
            "graduates_with_strong_branch": sum(
                row["year"] == year - 1400 and row.get("official") is True
                for row in official_graduate_list
            ),
            "graduates_branch_matched": sum(
                row["year"] == year - 1400 and row.get("official") is True
                for row in official_graduate_list
            ),
            "graduates_ontime_evaluable": sum(
                row["year"] == year - 1400 and row.get("official_on_time") is not None
                for row in official_graduate_list
            ),
        }

    detail_counts_by_branch = {
        str(year): {
            branch: {
                "regular_students": sum(
                    row["year"] == year and row["branch"] == branch for row in regular_students
                ),
                "entrants": sum(
                    row["year"] == year and row["branch"] == branch for row in entrants
                ),
                "graduates": sum(
                    row["year"] == year and row["branch"] == branch for row in graduates
                ),
                "noncompleters": sum(
                    row["year"] == year and row["branch"] == branch for row in noncompleters
                ),
                "student_detail_branches": sum(
                    row["year"] == year and row["branch"] == branch
                    for row in student_detail_branches
                ),
            }
            for branch in BRANCH_ORDER
        }
        for year in (46, 47)
    }

    distributed_totals: dict[str, dict[str, Any]] = {}
    distributed_availability: dict[str, dict[str, dict[str, Any]]] = {}
    for year in (46, 47):
        year_key = str(year)
        distributed_totals[year_key] = {}
        distributed_availability[year_key] = {}
        for metric in (
            "students_total",
            "students_new",
            "graduates_total",
            "graduates_branch_matched",
        ):
            available = {
                branch: metrics[year_key][branch][metric]
                for branch in BRANCH_ORDER
                if metrics[year_key][branch][metric] is not None
            }
            unavailable = [branch for branch in BRANCH_ORDER if branch not in available]
            distributed_totals[year_key][metric] = (
                sum(available.values()) if available else None
            )
            distributed_availability[year_key][metric] = {
                "available_branches": list(available),
                "unavailable_branches": unavailable,
                "is_complete_all_branches": not unavailable,
            }

    source_import_time = ""
    connection = open_readonly_database(args.database)
    try:
        source_import_time = clean(
            connection.execute("SELECT MAX(imported_at) FROM sources").fetchone()[0]
        )
    finally:
        connection.close()

    payload: dict[str, Any] = {
        "schema_version": 2,
        "generated_at_utc": datetime.now(timezone.utc).replace(microsecond=0).isoformat(),
        "source_snapshot_imported_at": source_import_time,
        "program": PROGRAM_AR,
        "degree": DEGREE_AR,
        "department": DEPARTMENT_AR,
        "branch_order": list(BRANCH_ORDER),
        "coverage": build_coverage(observed_coverage, profiles),
        "metrics": metrics,
        "regular_students": regular_students,
        "entrants": entrants,
        "graduates": graduates,
        "noncompleters": noncompleters,
        "student_detail_branches": student_detail_branches,
        "reconciliation": {
            "status": "passed",
            "distributed_totals": distributed_totals,
            "distributed_availability": distributed_availability,
            "whole_program_reference": reference,
            "partial_official_branch_linkages": {
                "noncompleters": {
                    "deduplication": noncompleter_deduplication,
                    "branch_linkage": noncompleter_linkage,
                },
                "student_detail_branches": {
                    "deduplication": student_detail_deduplication,
                    "branch_linkage": student_detail_linkage,
                },
            },
            "official_detail_rows": {
                short_year(year): {
                    "regular": len(official_regular.get(year, {})),
                    "graduates": len(official_graduates.get(year, {})),
                    "graduates_without_strong_branch_evidence": official_unassigned.get(year, 0),
                    **official_linkage[short_year(year)],
                }
                for year in TARGET_YEARS
            },
            "note": (
                "الإجماليات الموزعة لا يلزم أن تساوي مرجع البرنامج الكامل؛ "
                "السجل الرسمي لا يحوي فرعًا، وملفات النتائج لا تغطي الحوية والخرمة بصورة كاملة. "
                "graduates_total غير متاح على مستوى الفرع؛ graduates_branch_matched هو عدد الحالات "
                "الرسمية التي أمكن إسنادها بدليل فرع قوي، وليس إجمالي خريجي الفرع."
            ),
        },
        "_internal": {
            "rules_version": "1.1.0",
            "accepted_confidence": ["official", "very_high", "high"],
            "confidence_semantics": {
                "confidence": "compatibility field mirroring event confidence",
                "event_confidence": ["official", "very_high", "high"],
                "branch_confidence": ["very_high", "high"],
                "official_event_does_not_make_branch_official": True,
            },
            "detail_counts_by_branch": detail_counts_by_branch,
            "regular_definition": {
                "term": "first_semester_of_academic_year",
                "official_when": "official regular roster row plus strong branch evidence",
                "otherwise": "active-results proxy with strong program and branch evidence",
            },
            "entrant_policy": {
                "first_global_result_appearance": "first_semester_of_target_year",
                "requires_level_one_specialty_course": True,
                "continuation_required_for_high": False,
                "continuation_used_to_confirm_very_high": True,
                "year_1447_high_does_not_require_next_term": True,
            },
            "partial_official_linkage_policy": {
                "lists_are_not_official_branch_totals": True,
                "branch_window": "last_two_observed_plan_terms_through_academic_year_end",
                "event_confidence": "official",
                "branch_confidence_allowed": ["very_high", "high"],
                "deduplication_key": ["year", "student_id"],
                "cross_source_rows_kept_in_separate_lists": True,
            },
            "program_membership_thresholds": {
                "very_high": {"exclusive_courses_min": 3, "margin_min": 2},
                "high": {"exclusive_courses_min": 2, "margin_min": 1},
            },
            "branch_thresholds": {
                "very_high": {
                    "plan_courses_min": 3,
                    "exclusive_courses_min": 2,
                    "share_min": 0.80,
                    "margin_min": 2,
                },
                "high": {
                    "plan_courses_min": 2,
                    "exclusive_courses_min": 1,
                    "share_min": round(2 / 3, 6),
                    "margin_min": 1,
                },
            },
            "grade_policy": {
                "pass": sorted(PASS_GRADES),
                "fail": sorted(FAIL_GRADES),
                "non_completion": sorted(NON_COMPLETION_GRADES),
                "ambiguous_not_passed": sorted(AMBIGUOUS_GRADES),
                "score_alone_never_proves_pass": True,
            },
            "graduation_policy": {
                "official_records_allowed_any_plan": True,
                "inferred_plan_versions": [39],
                "plan47_inference_disabled": True,
                "requires_all_fixed_plan_courses_passed": True,
                "requires_required_elective_hours_passed": True,
                "requires_next_term_absence": True,
                "requires_next_term_fully_observed_for_branch": True,
            },
            "plan_summary": {
                "39": {
                    "courses": len(plans.plans[39].all_codes),
                    "fixed_courses": len(
                        plans.plans[39].all_codes - plans.plans[39].optional_pool
                    ),
                    "total_hours": plans.plans[39].total_hours,
                    "specialty_core": len(plans.plans[39].specialty_core),
                    "optional_pool": len(plans.plans[39].optional_pool),
                },
                "47": {
                    "courses": len(plans.plans[47].all_codes),
                    "total_hours": plans.plans[47].total_hours,
                    "specialty_core": len(plans.plans[47].specialty_core),
                    "graduation_inference": "disabled_incomplete_elective_groups",
                },
                "common_codes": len(plans.plans[39].all_codes & plans.plans[47].all_codes),
                "old_only_codes": len(plans.old_only),
                "new_only_codes": len(plans.new_only),
                "ambiguous_identity_codes_excluded": len(plans.ambiguous_identity_codes),
            },
            "source_files": {
                "database": {"sha256": sha256_file(args.database), **database_summary},
                "plan_data": {"sha256": sha256_file(args.plan_data)},
                "students_detail": {
                    "sha256": sha256_file(args.students_detail) if args.students_detail.is_file() else ""
                },
                "graduates_detail": {
                    "sha256": sha256_file(args.graduates_detail) if args.graduates_detail.is_file() else ""
                },
                "non_completers": {
                    "sha256": sha256_file(args.non_completers) if args.non_completers.is_file() else ""
                },
                "kpi_data": {"sha256": sha256_file(args.kpi_data) if args.kpi_data.is_file() else ""},
            },
        },
    }
    validate_payload(payload)
    return payload


def parse_args(argv: Sequence[str] | None = None) -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="استخراج مؤشرات طلاب الدراسات الإسلامية حسب الفرع."
    )
    parser.add_argument("--database", type=Path, default=DEFAULT_DATABASE)
    parser.add_argument("--plan-data", type=Path, default=DEFAULT_PLAN_DATA)
    parser.add_argument("--students-detail", type=Path, default=DEFAULT_STUDENTS)
    parser.add_argument("--graduates-detail", type=Path, default=DEFAULT_GRADUATES)
    parser.add_argument("--non-completers", type=Path, default=DEFAULT_NON_COMPLETERS)
    parser.add_argument("--kpi-data", type=Path, default=DEFAULT_KPI)
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT)
    parser.add_argument(
        "--check-only",
        action="store_true",
        help="نفّذ الاستخراج والمصالحة دون كتابة ملف الناتج.",
    )
    return parser.parse_args(argv)


def main(argv: Sequence[str] | None = None) -> int:
    args = parse_args(argv)
    required = (args.database, args.plan_data)
    missing = [path for path in required if not path.is_file()]
    if missing:
        raise SystemExit("مدخلات مفقودة: " + "، ".join(str(path) for path in missing))

    payload = build_payload(args)
    if not args.check_only:
        args.output.parent.mkdir(parents=True, exist_ok=True)
        temporary = args.output.with_suffix(args.output.suffix + ".tmp")
        temporary.write_text(
            json.dumps(payload, ensure_ascii=False, indent=2) + "\n",
            encoding="utf-8",
        )
        os.replace(temporary, args.output)

    totals = payload["reconciliation"]["distributed_totals"]
    partial_counts = {
        name: {
            str(year): sum(row["year"] == year for row in payload[name])
            for year in (46, 47)
        }
        for name in ("noncompleters", "student_detail_branches")
    }
    print(
        json.dumps(
            {
                "status": "passed",
                "output": str(args.output),
                "written": not args.check_only,
                "distributed_totals": totals,
                "partial_branch_linkages": partial_counts,
                "official_detail_rows": payload["reconciliation"]["official_detail_rows"],
            },
            ensure_ascii=False,
            indent=2,
        )
    )
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
