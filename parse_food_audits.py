import json
import re
import shutil
from pathlib import Path
from typing import Any, Dict, List, Optional, Tuple

import pdfplumber


DEPARTAMENTOS: set[str] = {
    "alta verapaz", "baja verapaz", "chimaltenango", "chiquimula",
    "el progreso", "escuintla", "guatemala", "huehuetenango", "izabal",
    "jalapa", "jutiapa", "peten", "quetzaltenango", "quiche",
    "retalhuleu", "sacatepequez", "san marcos", "santa rosa",
    "solola", "suchitepequez", "totonicapan", "zacapa",
}

# Municipios cuyo departamento no aparece explícito en el PDF — se usa como último recurso
MUNICIPIO_DEPTO_FALLBACK: dict[str, str] = {
    "coban": "Alta Verapaz", "cobán": "Alta Verapaz",
    "mazatenango": "Suchitepéquez",
    "san martin jilotepeque": "Chimaltenango", "san martín jilotepeque": "Chimaltenango",
    "patzun": "Chimaltenango", "patzún": "Chimaltenango",
    "mixco": "Guatemala",
    "santa catarina pinula": "Guatemala",
    "san jose pinula": "Guatemala", "san josé pinula": "Guatemala",
    "villa nueva": "Guatemala",
    "villa canales": "Guatemala",
    "san pedro ayampuc": "Guatemala",
    "amatitlan": "Guatemala", "amatitlán": "Guatemala",
    "san pedro soloma": "Huehuetenango",
    "aguacatan": "Huehuetenango", "aguacatán": "Huehuetenango",
    "santa eulalia": "Huehuetenango",
    "joyabaj": "Quiché",
    "momostenango": "Totonicapán",
    "barberena": "Santa Rosa",
    "la gomera": "Escuintla",
    "morales": "Izabal",
    "puerto barrios": "Izabal",
    "el chal": "Petén",
    "san benito": "Petén",
    "gualan": "Zacapa", "gualán": "Zacapa",
    "jocotenango": "Sacatepéquez",
    "ciudad vieja": "Sacatepéquez",
    "malacatan": "San Marcos", "malacatán": "San Marcos",
    "pajapita": "San Marcos",
    "ayutla": "San Marcos",
    "santa catarina mita": "Jutiapa",
}


DEPTO_NORM: dict[str, str] = {
    "peten": "Petén", "quiche": "Quiché", "sacatepequez": "Sacatepéquez",
    "solola": "Sololá", "suchitepequez": "Suchitepéquez", "totonicapan": "Totonicapán",
    "el progreso": "El Progreso", "alta verapaz": "Alta Verapaz",
    "baja verapaz": "Baja Verapaz", "san marcos": "San Marcos",
    "santa rosa": "Santa Rosa", "chimaltenango": "Chimaltenango",
    "chiquimula": "Chiquimula", "escuintla": "Escuintla",
    "guatemala": "Guatemala", "huehuetenango": "Huehuetenango",
    "izabal": "Izabal", "jalapa": "Jalapa", "jutiapa": "Jutiapa",
    "quetzaltenango": "Quetzaltenango", "retalhuleu": "Retalhuleu",
    "zacapa": "Zacapa",
}


SECTION_NAMES = [
    "Food Safety Risk Factors",
    "Cleanliness",
    "Maintenance and Facility",
    "Storage",
    "Knowledge and Compliance",
    "Critical Violations",
    "Dough (Back of Store)",
    "Vegetables (Back of Store)",
]


def clean_text(text: str) -> str:
    if not text:
        return ""
    text = text.replace("\u00a0", " ")
    text = text.replace("￾", "")
    text = re.sub(r"[ \t]+", " ", text)
    text = re.sub(r"\n{2,}", "\n", text)
    return text.strip()


def sql_escape(value: Optional[str]) -> str:
    if value is None:
        return "NULL"
    return "'" + str(value).replace("'", "''") + "'"


def read_pdf_text(pdf_path: Path) -> List[str]:
    pages: List[str] = []
    with pdfplumber.open(str(pdf_path)) as pdf:
        for page in pdf.pages:
            txt = page.extract_text() or ""
            pages.append(clean_text(txt))
    return pages


def normalize_datetime(value: str) -> Optional[str]:
    value = value.strip()
    m = re.match(
        r"(\d{1,2})/(\d{1,2})/(\d{4})\s+(\d{1,2}):(\d{2}):(\d{2})\s*(AM|PM)",
        value,
        re.IGNORECASE,
    )
    if not m:
        return value

    month, day, year, hh, mm, ss, ampm = m.groups()
    hour = int(hh)
    ampm = ampm.upper()

    if ampm == "PM" and hour != 12:
        hour += 12
    if ampm == "AM" and hour == 12:
        hour = 0

    return f"{year}-{int(month):02d}-{int(day):02d} {hour:02d}:{mm}:{ss}"


def _find_depto_in_text(text: str) -> Optional[str]:
    text_lower = text.lower()
    for depto in sorted(DEPARTAMENTOS, key=len, reverse=True):
        if re.search(r"\b" + re.escape(depto) + r"\b", text_lower):
            return DEPTO_NORM.get(depto, depto.title())
    return None


def extract_municipio_departamento(location: Optional[str]) -> Tuple[Optional[str], Optional[str]]:
    if not location:
        return None, None

    if re.search(r"guatemala\s+city", location, re.IGNORECASE):
        return "Guatemala", "Guatemala"

    m_start = re.search(r"Audit\s+Start\s+Time", location, re.IGNORECASE)
    m_end   = re.search(r"Audit\s+End\s+Time",   location, re.IGNORECASE)

    before_start = location[: m_start.start()] if m_start else location
    between = location[m_start.end() : m_end.start()] if (m_start and m_end) else ""

    text = re.sub(r"Auditor\s+[A-Z0-9]+", "", before_start, flags=re.IGNORECASE)
    text = re.sub(r",\s*,", ",", text).strip().strip(",").strip()
    tokens = [t.strip() for t in text.split(",") if t.strip()]

    departamento: Optional[str] = None
    municipio:    Optional[str] = None

    for i in range(len(tokens) - 1, -1, -1):
        tok       = tokens[i].strip()
        tok_lower = tok.lower()

        if tok_lower in DEPARTAMENTOS:
            departamento = DEPTO_NORM.get(tok_lower, tok.title())
            municipio    = tokens[i - 1].title() if i > 0 else None
            break

        for depto in sorted(DEPARTAMENTOS, key=len, reverse=True):
            if tok_lower.endswith(" " + depto) or tok_lower == depto:
                departamento = DEPTO_NORM.get(depto, depto.title())
                mun_part     = tok[: len(tok) - len(depto)].strip().strip(",").strip()
                municipio    = mun_part.title() if mun_part else (tokens[i - 1].title() if i > 0 else None)
                break

        if departamento:
            break

    # Si no encontramos departamento, el último token legible es el candidato a municipio
    if not departamento and tokens:
        candidate = tokens[-1].strip()
        # Descartar tokens que parecen direcciones (empiezan con número o son muy cortos)
        if not re.match(r"^\d", candidate) and len(candidate) > 2:
            municipio = candidate.title()

    # Fallback 1: municipio conocido → departamento por lookup
    if not departamento and municipio:
        departamento = MUNICIPIO_DEPTO_FALLBACK.get(municipio.lower())

    # Fallback 2: buscar entre Audit Start Time y Audit End Time (excluye 'guatemala' ambiguo)
    if not departamento and between:
        found = _find_depto_in_text(between)
        if found and found != "Guatemala":
            departamento = found
            if tokens:
                municipio = tokens[-1].title()

    return municipio, departamento


def extract_audit_id(page1: str) -> Optional[int]:
    patterns = [
        r"Audit\s*ID\s*:?\s*(\d+)",
        r"AuditID\s*:?\s*(\d+)",
        r"Audit\s*#\s*:?\s*(\d+)",
        r"Audit\s*Number\s*:?\s*(\d+)",
    ]

    for pattern in patterns:
        m = re.search(pattern, page1, re.IGNORECASE)
        if m:
            return int(m.group(1))

    return None


def extract_location(page1: str) -> Optional[str]:
    lines = [x.strip() for x in page1.split("\n") if x.strip()]

    ignore_prefixes = (
        "Store ",
        "AuditID",
        "Audit ID",
        "Audit #",
        "Audit Number",
        "Auditor ",
        "Audit Start Time",
        "Audit End Time",
        "Franchisee ",
        "Store Manager",
        "Manager In Charge",
        "Critical Violations",
        "Total Points",
        "Percentage Score",
        "Summary Scoring By Section",
        "Section Name",
        "Earned Points",
        "Possible Points",
        "Section Score",
        "Total Violations",
    )

    try:
        store_idx = next(i for i, line in enumerate(lines) if line.startswith("Store "))
    except StopIteration:
        return None

    location_lines: List[str] = []
    for line in lines[store_idx + 1:]:
        if line.startswith(ignore_prefixes):
            continue
        if line in SECTION_NAMES:
            break
        if re.match(r"^\d+\s+\d+\s+\d+(?:\.\d+)?%\s+\d+$", line):
            break
        if re.match(r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),", line):
            break
        if "Page " in line and " of " in line:
            break
        location_lines.append(line)

        # Cortar cuando ya tenemos una línea que parece país/ciudad final
        if len(location_lines) >= 3:
            break

    location = ", ".join(location_lines).strip(", ")
    return location or None


def extract_header(page1: str, source_file: str) -> Dict[str, Any]:
    data: Dict[str, Any] = {
        "store_id": None,
        "audit_id": None,
        "location": None,
        "departamento": None,
        "municipio": None,
        "audit_start": None,
        "audit_end": None,
        "auditor": None,
        "franchisee": None,
        "manager_in_charge": None,
        "points_earned": None,
        "points_possible": None,
        "percentage_score": None,
        "critical_violations": None,
        "total_violations": None,
        "source_file": source_file,
    }

    m = re.search(r"Store\s+(\d+)", page1, re.IGNORECASE)
    if m:
        data["store_id"] = int(m.group(1))

    data["audit_id"] = extract_audit_id(page1)

    m = re.search(r"Auditor\s+([A-Z0-9]+)", page1, re.IGNORECASE)
    if m:
        data["auditor"] = m.group(1).strip()

    m = re.search(r"Audit Start Time\s+([0-9/:\sAPMapm]+)", page1, re.IGNORECASE)
    if m:
        data["audit_start"] = normalize_datetime(m.group(1))

    m = re.search(r"Audit End Time\s+([0-9/:\sAPMapm]+)", page1, re.IGNORECASE)
    if m:
        data["audit_end"] = normalize_datetime(m.group(1))

    m = re.search(r"Franchisee\s+([A-Za-zÁÉÍÓÚáéíóúÑñ, ]+)", page1, re.IGNORECASE)
    if m:
        data["franchisee"] = m.group(1).replace(",", ", ").strip()

    m = re.search(
        r"Manager In Charge\s+([A-Za-zÁÉÍÓÚáéíóúÑñ.\- ]+)",
        page1,
        re.IGNORECASE,
    )
    if m:
        data["manager_in_charge"] = m.group(1).strip()

    m = re.search(r"Critical Violations\s+(\d+)", page1, re.IGNORECASE)
    if m:
        data["critical_violations"] = int(m.group(1))

    m = re.search(r"Total Points\s+(\d+)\s*/\s*(\d+)", page1, re.IGNORECASE)
    if m:
        data["points_earned"] = int(m.group(1))
        data["points_possible"] = int(m.group(2))

    m = re.search(r"Percentage Score\s+(\d+(?:\.\d+)?)%", page1, re.IGNORECASE)
    if m:
        data["percentage_score"] = float(m.group(1))

    section_rows = re.findall(
        r"(Food Safety Risk Factors|Cleanliness|Maintenance and Facility|Storage|Knowledge and Compliance|Critical Violations)\s+(\d+)\s+(\d+)\s+(\d+(?:\.\d+)?)%\s+(\d+)",
        page1,
        re.IGNORECASE,
    )
    if section_rows:
        data["total_violations"] = sum(int(r[4]) for r in section_rows)

    data["location"] = extract_location(page1)
    data["municipio"], data["departamento"] = extract_municipio_departamento(data["location"])

    return data


def extract_sections(page1: str, audit_id: int) -> List[Dict[str, Any]]:
    matches = re.findall(
        r"(Food Safety Risk Factors|Cleanliness|Maintenance and Facility|Storage|Knowledge and Compliance|Critical Violations)\s+(\d+)\s+(\d+)\s+(\d+(?:\.\d+)?)%\s+(\d+)",
        page1,
        re.IGNORECASE,
    )

    sections: List[Dict[str, Any]] = []
    for name, earned, possible, score, violations in matches:
        sections.append(
            {
                "audit_id": audit_id,
                "section_name": normalize_section_name(name),
                "points_earned": int(earned),
                "points_possible": int(possible),
                "section_score": float(score),
                "total_violations": int(violations),
            }
        )
    return sections


def normalize_section_name(name: str) -> str:
    normalized = " ".join(name.split())
    if normalized.lower() == "vegetables (back of store)":
        return "Vegetable processing"
    if normalized.lower() == "dough (back of store)":
        return "Dough processing"
    for s in SECTION_NAMES:
        if normalized.lower() == s.lower():
            return s
    return normalized


def infer_finding_type(question: str, comment: str) -> str:
    q = f"{question} {comment or ''}".lower()

    rules: List[Tuple[str, List[str]]] = [
        ("temperature", ["41°", "42.", "43.", "44.", "45.", "46.", "47.", "48.", "49.", "50°", "51.", "tcs", "temperature", "temperatura"]),
        ("servsafe", ["servsafe", "certification", "certificado", "certificación"]),
        ("pest", ["cucaracha", "mosca", "plaga", "pest", "infestation"]),
        ("sanitizer", ["sanitizer", "desinfectante", "ppm", "concentración"]),
        ("chemical_control", ["químico", "quimico", "etiqueta", "sds", "chemical"]),
        ("equipment_damage", ["roto", "deteriorado", "óxido", "oxido", "gasket", "empaque", "fuga", "quebrado", "poor repair"]),
        ("cleanliness", ["suciedad", "dirty", "moho", "grasa", "acumulación", "acumulacion", "mold"]),
        ("cross_contamination", ["cross", "contamination", "contaminación", "contaminacion", "food exposed", "hilos", "uncovered"]),
        ("policy_gap", ["policy", "política", "politica", "manual", "allergen", "global store food safety"]),
        ("process_gap", ["procedure", "procedimiento", "knowledge not demostrated", "knowledge not demonstrated"]),
        ("data_integrity", ["falsification", "registros", "fecha futura", "future"]),
        ("personal_behavior", ["eating in back of house", "team member eating", "consumir alimentos", "drinking"]),
    ]

    for label, words in rules:
        if any(w in q for w in words):
            return label

    return "other"


def extract_non_compliance(pages: List[str], audit_id: int) -> List[Dict[str, Any]]:
    findings: List[Dict[str, Any]] = []
    in_nc = False
    buffer: List[Tuple[int, str]] = []

    for page_no, page_text in enumerate(pages, start=1):
        if "Non-Compliance Summary" in page_text:
            in_nc = True
        if "Question Summary" in page_text:
            in_nc = False

        if in_nc:
            buffer.append((page_no, page_text))

    current: Optional[Dict[str, Any]] = None
    entries: List[Dict[str, Any]] = []

    entry_pattern = re.compile(
        r"^(Food Safety Risk Factors|Cleanliness|Maintenance and Facility|Storage|Knowledge and Compliance|Critical Violations|Dough \(Back of Store\)|Vegetables \(Back of Store\))-(.+?)\s+0/(\d+)$",
        re.IGNORECASE,
    )

    for page_no, text in buffer:
        lines = [l.strip() for l in text.split("\n") if l.strip()]
        for line in lines:
            m = entry_pattern.match(line)
            if m:
                if current:
                    entries.append(current)

                current = {
                    "section_name": normalize_section_name(m.group(1)),
                    "question_text": m.group(2).strip(),
                    "answer_value": "No",
                    "points_earned": 0,
                    "points_possible": int(m.group(3)),
                    "comment_lines": [],
                    "evidence_page": page_no,
                }
                continue

            if current:
                if line in {"Non-Compliance Summary", "Earned Points"}:
                    continue
                if re.match(r"^(Monday|Tuesday|Wednesday|Thursday|Friday|Saturday|Sunday),", line):
                    continue
                if re.match(r"^Page \d+ of \d+$", line):
                    continue
                current["comment_lines"].append(line)

    if current:
        entries.append(current)

    for entry in entries:
        comment_text = " ".join(entry["comment_lines"]).strip()
        finding_type = infer_finding_type(entry["question_text"], comment_text)

        findings.append(
            {
                "audit_id": audit_id,
                "section_name": entry["section_name"],
                "question_text": entry["question_text"],
                "answer_value": entry["answer_value"],
                "points_earned": entry["points_earned"],
                "points_possible": entry["points_possible"],
                "finding_type": finding_type,
                "comment_text": comment_text[:4000] if comment_text else None,
                "evidence_page": entry["evidence_page"],
            }
        )

    return findings


def build_sql(audit: Dict[str, Any], sections: List[Dict[str, Any]], findings: List[Dict[str, Any]]) -> str:
    sql_parts: List[str] = []

    sql_parts.append(
        f"""INSERT INTO audits (
    store_id,
    audit_id,
    location,
    departamento,
    municipio,
    audit_start,
    audit_end,
    auditor,
    franchisee,
    manager_in_charge,
    points_earned,
    points_possible,
    percentage_score,
    critical_violations,
    total_violations,
    source_file
) VALUES (
    {audit['store_id']},
    {audit['audit_id']},
    {sql_escape(audit['location'])},
    {sql_escape(audit['departamento'])},
    {sql_escape(audit['municipio'])},
    {sql_escape(audit['audit_start'])},
    {sql_escape(audit['audit_end'])},
    {sql_escape(audit['auditor'])},
    {sql_escape(audit['franchisee'])},
    {sql_escape(audit['manager_in_charge'])},
    {audit['points_earned']},
    {audit['points_possible']},
    {audit['percentage_score']},
    {audit['critical_violations']},
    {audit['total_violations']},
    {sql_escape(audit['source_file'])}
);"""
    )

    if sections:
        values = []
        for s in sections:
            values.append(
                f"({s['audit_id']}, {sql_escape(s['section_name'])}, {s['points_earned']}, {s['points_possible']}, {s['section_score']}, {s['total_violations']})"
            )
        sql_parts.append(
            "INSERT INTO audit_sections (\n"
            "    audit_id,\n"
            "    section_name,\n"
            "    points_earned,\n"
            "    points_possible,\n"
            "    section_score,\n"
            "    total_violations\n"
            ") VALUES\n"
            + ",\n".join(values)
            + ";"
        )

    if findings:
        values = []
        for f in findings:
            values.append(
                "("
                f"{f['audit_id']}, "
                f"{sql_escape(f['section_name'])}, "
                f"{sql_escape(f['question_text'])}, "
                f"{sql_escape(f['answer_value'])}, "
                f"{f['points_earned']}, "
                f"{f['points_possible']}, "
                f"{sql_escape(f['finding_type'])}, "
                f"{sql_escape(f['comment_text'])}, "
                f"{f['evidence_page']}"
                ")"
            )
        sql_parts.append(
            "INSERT INTO audit_findings (\n"
            "    audit_id,\n"
            "    section_name,\n"
            "    question_text,\n"
            "    answer_value,\n"
            "    points_earned,\n"
            "    points_possible,\n"
            "    finding_type,\n"
            "    comment_text,\n"
            "    evidence_page\n"
            ") VALUES\n"
            + ",\n".join(values)
            + ";"
        )

    return "\n\n".join(sql_parts)


def build_sql_update(audit: Dict[str, Any], sections: List[Dict[str, Any]], findings: List[Dict[str, Any]]) -> str:
    sql_parts: List[str] = []

    sql_parts.append(
        f"""UPDATE audits SET
    store_id = {audit['store_id']},
    location = {sql_escape(audit['location'])},
    departamento = {sql_escape(audit['departamento'])},
    municipio = {sql_escape(audit['municipio'])},
    audit_start = {sql_escape(audit['audit_start'])},
    audit_end = {sql_escape(audit['audit_end'])},
    auditor = {sql_escape(audit['auditor'])},
    franchisee = {sql_escape(audit['franchisee'])},
    manager_in_charge = {sql_escape(audit['manager_in_charge'])},
    points_earned = {audit['points_earned']},
    points_possible = {audit['points_possible']},
    percentage_score = {audit['percentage_score']},
    critical_violations = {audit['critical_violations']},
    total_violations = {audit['total_violations']},
    source_file = {sql_escape(audit['source_file'])}
WHERE audit_id = {audit['audit_id']};"""
    )

    for s in sections:
        sql_parts.append(
            f"""UPDATE audit_sections SET
    points_earned = {s['points_earned']},
    points_possible = {s['points_possible']},
    section_score = {s['section_score']},
    total_violations = {s['total_violations']}
WHERE audit_id = {s['audit_id']}
  AND section_name = {sql_escape(s['section_name'])};"""
        )

    for f in findings:
        sql_parts.append(
            f"""UPDATE audit_findings SET
    answer_value = {sql_escape(f['answer_value'])},
    points_earned = {f['points_earned']},
    points_possible = {f['points_possible']},
    finding_type = {sql_escape(f['finding_type'])},
    comment_text = {sql_escape(f['comment_text'])},
    evidence_page = {f['evidence_page']}
WHERE audit_id = {f['audit_id']}
  AND section_name = {sql_escape(f['section_name'])}
  AND question_text = {sql_escape(f['question_text'])};"""
        )

    return "\n\n".join(sql_parts)


def process_pdf(pdf_path: Path) -> Dict[str, Any]:
    pages = read_pdf_text(pdf_path)
    if not pages:
        raise ValueError(f"No se pudo extraer texto de {pdf_path.name}")

    audit = extract_header(pages[0], pdf_path.name)

    if not audit["audit_id"]:
        print("\n--- DEBUG PAGE 1 ---")
        print(pages[0][:4000])
        print("--- END DEBUG ---\n")
        raise ValueError(f"No se encontró audit_id en {pdf_path.name}")

    sections = extract_sections(pages[0], audit["audit_id"])
    findings = extract_non_compliance(pages, audit["audit_id"])

    return {
        "audit": audit,
        "sections": sections,
        "findings": findings,
    }


def dedupe_results(results: List[Dict[str, Any]]) -> List[Dict[str, Any]]:
    seen = set()
    clean = []

    for item in results:
        audit_id = item["audit"].get("audit_id")
        if audit_id in seen:
            continue
        seen.add(audit_id)
        clean.append(item)

    return clean


def load_existing_results(output_json: Path) -> List[Dict[str, Any]]:
    if output_json.exists():
        with open(output_json, "r", encoding="utf-8") as f:
            return json.load(f)
    return []


def process_folder(base_folder: str) -> None:
    base = Path(base_folder)
    in_folder = base / "in"
    out_folder = base / "out"
    error_folder = base / "error"

    in_folder.mkdir(parents=True, exist_ok=True)
    out_folder.mkdir(parents=True, exist_ok=True)
    error_folder.mkdir(parents=True, exist_ok=True)

    output_json = base / "audits.json"
    output_sql = base / "all_inserts.sql"
    output_updates_sql = base / "all_updates.sql"
    output_errors = base / "errors.log"

    existing_results = load_existing_results(output_json)
    existing_results = dedupe_results(existing_results)

    existing_index = {
        item["audit"]["audit_id"]: idx
        for idx, item in enumerate(existing_results)
        if item.get("audit", {}).get("audit_id") is not None
    }

    pdf_files = sorted(in_folder.glob("*.pdf"))
    errors: List[str] = []
    updates: List[Dict[str, Any]] = []

    for pdf_file in pdf_files:
        try:
            parsed = process_pdf(pdf_file)
            audit_id = parsed["audit"]["audit_id"]

            if audit_id in existing_index:
                existing_results[existing_index[audit_id]] = parsed
                updates.append(parsed)
                shutil.move(str(pdf_file), str(out_folder / pdf_file.name))
                print(f"UPDATE: {pdf_file.name} -> audit_id {audit_id}")
            else:
                existing_results.append(parsed)
                existing_index[audit_id] = len(existing_results) - 1
                shutil.move(str(pdf_file), str(out_folder / pdf_file.name))
                print(f"OK: {pdf_file.name}")

        except Exception as e:
            errors.append(f"{pdf_file.name} -> {e}")
            shutil.move(str(pdf_file), str(error_folder / pdf_file.name))
            print(f"ERROR: {pdf_file.name} -> {e}")

    existing_results = dedupe_results(existing_results)

    sql_blocks = [
        build_sql(item["audit"], item["sections"], item["findings"])
        for item in existing_results
    ]

    update_blocks = [
        build_sql_update(item["audit"], item["sections"], item["findings"])
        for item in updates
    ]

    with open(output_json, "w", encoding="utf-8") as f:
        json.dump(existing_results, f, ensure_ascii=False, indent=2)

    with open(output_sql, "w", encoding="utf-8") as f:
        f.write("\n\n-- ==============================\n\n".join(sql_blocks))

    if update_blocks:
        with open(output_updates_sql, "w", encoding="utf-8") as f:
            f.write("\n\n-- ==============================\n\n".join(update_blocks))

    with open(output_errors, "w", encoding="utf-8") as f:
        for err in errors:
            f.write(err + "\n")

    unique_stores = {
        item["audit"]["store_id"]
        for item in existing_results
        if item.get("audit", {}).get("store_id") is not None
    }

    print("\nResumen final")
    print(f"Auditorías únicas: {len(existing_results)}")
    print(f"Stores únicas: {len(unique_stores)}")
    print(f"Actualizadas  : {len(updates)}")
    print(f"Generado JSON: {output_json}")
    print(f"Generado SQL (INSERTs): {output_sql}")
    if update_blocks:
        print(f"Generado SQL (UPDATEs): {output_updates_sql}")
    print(f"Errores       : {output_errors}")


if __name__ == "__main__":
    process_folder("./pdfs")