"""
GMP Automation System - Offline OCR Output Parsers
Converts raw text/markdown/HTML produced by the local Offline OCR backend into
the structured JSON schema consumed by the Excel generator, so
excel_generator.py does not need to change.

The Offline OCR backend emits documents as HTML <table> blocks (with rowspan/colspan for
merged cells) wrapped in <|ref|> markers, e.g.:

    <|ref|>table<|/ref|><|det|>[[23, 66, 958, 831]]<|/det|>
    <table><tr><td rowspan="2">NO.</td> ... </tr></table>

The parsers below therefore:
  1. Extract every <table> block and expand rowspan/colspan into a flat grid.
  2. Merge multi-row (group + sub) headers into one header row.
  3. Locate identity columns (grade / room no / room name / volume / total / ACH)
     and treat every remaining column as measurement values.
  4. Fall back to classic markdown |...| tables if the model ever emits those.
"""

import re
from html.parser import HTMLParser


# =============================================================================
# GENERIC FIELD / TABLE EXTRACTION HELPERS
# =============================================================================

def _join_pages(pages_text):
    return "\n".join(pages_text)


# -----------------------------------------------------------------------------
# HTML table extraction (rowspan/colspan aware)
# -----------------------------------------------------------------------------

class _TableParser(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.tables = []
        self._table = None
        self._row = None
        self._cell = None
        self._rowspan = 1
        self._colspan = 1

    def handle_starttag(self, tag, attrs):
        if tag == 'table':
            self._table = []
        elif tag == 'tr' and self._table is not None:
            self._row = []
        elif tag in ('td', 'th') and self._row is not None:
            self._cell = []
            d = dict(attrs)
            self._rowspan = int(d.get('rowspan', 1) or 1)
            self._colspan = int(d.get('colspan', 1) or 1)
        elif tag == 'br' and self._cell is not None:
            self._cell.append(' ')

    def handle_data(self, data):
        if self._cell is not None:
            self._cell.append(data)

    def handle_endtag(self, tag):
        if tag == 'table':
            if self._table:
                self.tables.append(self._table)
            self._table = None
        elif tag == 'tr':
            if self._row is not None and self._table is not None:
                self._table.append(self._row)
            self._row = None
        elif tag in ('td', 'th') and self._cell is not None and self._row is not None:
            text = ' '.join(''.join(self._cell).split())
            self._row.append({'text': text, 'rowspan': self._rowspan, 'colspan': self._colspan})
            self._cell = None


def _expand_table(rows):
    """Expand rowspan/colspan cells into a rectangular grid of plain strings."""
    result = []
    pending = {}  # col -> (text, remaining rows after current)
    for row in rows:
        out_row = []
        col = 0
        cell_idx = 0
        while cell_idx < len(row) or any(c >= col for c in pending):
            if col in pending:
                text, remaining = pending[col]
                out_row.append(text)
                if remaining <= 1:
                    del pending[col]
                else:
                    pending[col] = (text, remaining - 1)
                col += 1
                continue
            if cell_idx >= len(row):
                out_row.append('')
                col += 1
                continue
            cell = row[cell_idx]
            cell_idx += 1
            for k in range(cell['colspan']):
                out_row.append(cell['text'])
                if cell['rowspan'] > 1:
                    pending[col + k] = (cell['text'], cell['rowspan'] - 1)
            col += cell['colspan']
        result.append(out_row)
    return result


def _to_number(s, as_int=False):
    if s is None:
        return None
    s = str(s).strip().replace(',', '').replace('%', '').replace(' ', '')
    if s in ('', '-', '—'):
        return None
    m = re.match(r'-?\d+\.?\d*', s)
    if not m:
        return None
    try:
        v = float(m.group(0))
        return int(v) if as_int else v
    except ValueError:
        return None


def _is_data_row(row):
    """A data row starts with an integer NO. and has at least one non-numeric
    cell in the following columns (grade letter, room number, room name...)."""
    if not row:
        return False
    first_is_number = bool(re.fullmatch(r'\d+', str(row[0]).strip()))
    numbers = sum(_to_number(cell) is not None for cell in row)
    has_grade = any(re.fullmatch(r'[ABCD]', str(cell).strip(), re.IGNORECASE) for cell in row)
    has_text = any(re.search(r'[A-Za-z가-힣]', str(c)) for c in row[1:6])
    return (first_is_number and has_text) or (numbers >= 2 and has_grade)


def _merge_header_rows(table):
    """Split a (possibly multi-row) header from the data rows and merge header
    cells so downstream code sees a single flat header row."""
    data_start = 0
    for idx, row in enumerate(table):
        if _is_data_row(row):
            data_start = idx
            break
    if data_start == 0:
        return table
    header_rows = table[:data_start]
    width = max(len(r) for r in header_rows)
    header = []
    for i in range(width):
        parts = [r[i] for r in header_rows if i < len(r) and r[i]]
        header.append(' '.join(parts))
    return [header] + table[data_start:]


def extract_html_tables(text):
    """All <table> blocks in text as list of tables; each table is a list of
    rows, each row a list of cell strings (merged cells repeated)."""
    parser = _TableParser()
    parser.feed(text)
    tables = []
    for raw in parser.tables:
        tables.append(_merge_header_rows(_expand_table(raw)))
    return tables


def extract_markdown_tables(text):
    """Fallback: classic markdown |...| tables. Each table is a list of rows."""
    tables = []
    current = []
    for line in text.splitlines():
        stripped = line.strip()
        if stripped.startswith('|') and stripped.endswith('|'):
            cells = [c.strip() for c in stripped.strip('|').split('|')]
            if all(re.fullmatch(r':?-+:?', c) for c in cells):
                continue
            current.append(cells)
        else:
            if current:
                tables.append(current)
                current = []
    if current:
        tables.append(current)
    return tables


def extract_tables(text):
    """All tables (HTML from Offline OCR, then markdown fallback)."""
    return extract_html_tables(text) + extract_markdown_tables(text)


# -----------------------------------------------------------------------------
# Text normalization for key:value field matching
# -----------------------------------------------------------------------------

class _TextExtractor(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []

    def handle_starttag(self, tag, attrs):
        if tag in ('br', 'tr', 'td', 'th', 'p', 'div', 'table'):
            self.parts.append(' ')

    def handle_data(self, data):
        self.parts.append(data)


def html_to_text(text):
    """HTML -> readable text (tags dropped, whitespace collapsed)."""
    parser = _TextExtractor()
    parser.feed(text)
    return ' '.join(''.join(parser.parts).split())


def extract_field(text, patterns, default=None):
    """Try a list of regexes in order, return the first captured group found."""
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return default


def extract_ahu(text):
    def candidate_number(value, allow_plain_number=False):
        compact_value = re.sub(r'\s+', '', str(value or ''))
        compact_value = compact_value.replace('－', '-').replace('−', '-').replace('—', '-')
        patterns = [
            r'(?:공조기|AHU)[_:\-|–]*'
            r'(?:(?:NO\.?)|번호|#)?[_:\-|–]*([1-9]\d{0,2})(?!\d)'
        ]
        if allow_plain_number:
            patterns.append(r'^([1-9]\d{0,2})(?:호기?)?$')
        return extract_field(compact_value, patterns)

    if '<table' in text.lower():
        for table in extract_tables(text):
            for row_index, row in enumerate(table):
                for index, cell in enumerate(row):
                    label = re.sub(r'[\s:|._\-]+', '', str(cell)).upper()
                    if not (
                        label.startswith('해당공조기')
                        or label in ('공조기', '공조기번호', 'AHU', 'AHUNO', '해당AHU')
                    ):
                        continue
                    candidates = list(row[index + 1:index + 3])
                    candidates.extend(
                        lower_row[index]
                        for lower_row in table[row_index + 1:]
                        if index < len(lower_row)
                    )
                    for candidate in candidates:
                        number = candidate_number(candidate, allow_plain_number=True)
                        if number:
                            return number
        text = html_to_text(text)

    return candidate_number(text) or 'unknown'


def extract_date(text):
    compact = re.sub(r'\s+', '', text)
    date = extract_field(
        compact,
        [
            r'측정일자[^\d]*(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})',
            r'(\d{4}[.\-/]\d{1,2}[.\-/]\d{1,2})',
        ],
        default='2025.08.01',
    )
    return date.replace('-', '.').replace('/', '.')


def extract_result(text):
    compact = re.sub(r'\s+', '', text)
    return extract_field(compact, [r'측정결과[:|]?(적합|부적합)'], default='적합')


def extract_standard(text):
    return extract_field(text, [r'측정\s*기준[^\S\n]*[:|]?\s*([\d.]+\s*%)'], default=None)


# -----------------------------------------------------------------------------
# Shared column helpers
# -----------------------------------------------------------------------------

def _find_col(header_row, keywords):
    """Return index of first header cell containing any keyword (whitespace-agnostic)."""
    for i, cell in enumerate(header_row):
        norm = re.sub(r'\s+', '', str(cell))
        for kw in keywords:
            if kw in norm:
                return i
    return None


_ID_COLS_KEYWORDS = {
    'grade': ['청정등급', '등급'],
    'room_no': ['실번호'],
    'room_name': ['실명', '설비명', '기기명'],
    'volume': ['체적'],
    'total': ['합계'],
    'ach': ['환기횟수', 'ACH', '회/hr'],
    'no': ['NO', 'No', 'no'],
    'point': ['측정번호', '측정횟수', '측정점'],
}


def _align_row(row, header_width):
    """Fix rows that are longer than the header because the OCR emitted a
    colspan artifact (a merged cell duplicated). Drop one duplicate cell."""
    if len(row) <= header_width:
        return row
    for i in range(1, len(row)):
        if row[i] == row[i - 1]:
            return row[:i] + row[i + 1:]
    return row


def _used_cols(header, *groups):
    """Set of header indices belonging to any of the given identity groups."""
    used = set()
    for group in groups:
        i = _find_col(header, _ID_COLS_KEYWORDS[group])
        if i is not None:
            used.add(i)
    return used


# =============================================================================
# A. AIRBORNE PARTICLE
# =============================================================================

def parse_airborne_particle(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    rooms = []
    no_counter = 1
    entries = []
    last_header = None
    previous_identity = None

    for table in extract_tables(text):
        if not table:
            continue
        candidate_header = table[0]

        room_no_i = _find_col(candidate_header, _ID_COLS_KEYWORDS['room_no'])
        room_name_i = _find_col(candidate_header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(candidate_header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and room_name_i is None:
            if last_header is None:
                continue
            header = last_header
            no_i = _find_col(header, _ID_COLS_KEYWORDS['no'])
            if no_i is None or not any(
                    no_i < len(row) and _to_number(row[no_i], as_int=True) is not None
                    for row in table):
                continue
            rows = table
            room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
            room_name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
            grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        else:
            header, rows = candidate_header, table[1:]
            last_header = header

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'no', 'point')
        measure_cols = [i for i in range(len(header)) if i not in used]
        # measurement columns come in pairs: 0.5um, 5.0um per point
        pairs = [measure_cols[i:i + 2] for i in range(0, len(measure_cols) - 1, 2)]

        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                row += [''] * (len(header) - len(row))
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not room_name and previous_identity:
                grade, room_number, room_name = previous_identity
            if not room_number and not room_name:
                continue
            previous_identity = (grade, room_number, room_name)

            measurements = []
            for (c05, c50) in pairs:
                v05 = _to_number(row[c05], as_int=True) if c05 < len(row) else None
                v50 = _to_number(row[c50], as_int=True) if c50 < len(row) else None
                if v05 is None and v50 is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value_05': v05 or 0, 'value_50': v50 or 0})

            if not measurements:
                continue
            entries.append({
                'grade': grade,
                'room_number': room_number,
                'room_name': room_name,
                'measurements': measurements,
            })

    # Group after all pages are parsed so a room can continue across a page break.
    i = 0
    while i < len(entries):
        e = entries[i]
        j = i
        while j + 1 < len(entries) and \
                entries[j + 1]['room_number'] == e['room_number'] and \
                entries[j + 1]['room_name'] == e['room_name'] and \
                entries[j + 1]['grade'] == e['grade']:
            j += 1
        group = []
        for en in entries[i:j + 1]:
            for m in en['measurements']:
                group.append(dict(m, point=len(group) + 1))
        rooms.append({
            'no_start': no_counter,
            'no_end': no_counter + len(group) - 1,
            'grade': e['grade'],
            'room_number': e['room_number'],
            'room_name': e['room_name'],
            'measurements': group,
        })
        no_counter += len(group)
        i = j + 1

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'rooms': rooms,
    }


# =============================================================================
# B. AIR VELOCITY
# =============================================================================

def parse_air_velocity(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    machines = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and name_i is None:
            continue

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'no', 'point')
        value_cols = [i for i in range(len(header)) if i not in used]

        entries = []
        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            machine_name = row[name_i] if name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            if not room_number and not machine_name:
                continue

            measurements = []
            for c in value_cols:
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value': v})

            if not measurements:
                continue
            entries.append({
                'grade': grade,
                'room_number': room_number,
                'machine_name': machine_name,
                'measurements': measurements,
            })

        # group consecutive entries for the same machine into one machine object
        i = 0
        while i < len(entries):
            e = entries[i]
            j = i
            while j + 1 < len(entries) and \
                    entries[j + 1]['room_number'] == e['room_number'] and \
                    entries[j + 1]['machine_name'] == e['machine_name'] and \
                    entries[j + 1]['grade'] == e['grade']:
                j += 1
            group = []
            for en in entries[i:j + 1]:
                for m in en['measurements']:
                    group.append(dict(m, point=len(group) + 1))
            machines.append({
                'no_start': no_counter,
                'no_end': no_counter + len(group) - 1,
                'grade': e['grade'],
                'room_number': e['room_number'],
                'machine_name': e['machine_name'],
                'measurements': group,
            })
            no_counter += len(group)
            i = j + 1

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'machines': machines,
    }


# =============================================================================
# C. AIR CHANGE RATE (ACH)
# =============================================================================

def parse_air_change_rate(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    rooms = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        room_name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        grade_i = _find_col(header, _ID_COLS_KEYWORDS['grade'])
        if room_no_i is None and room_name_i is None:
            continue

        used = _used_cols(header, 'grade', 'room_no', 'room_name', 'volume', 'total', 'ach', 'no', 'point')
        flow_cols = [i for i in range(len(header)) if i not in used]

        point_i = _find_col(header, _ID_COLS_KEYWORDS['point'])
        volume_i = _find_col(header, _ID_COLS_KEYWORDS['volume'])
        total_i = _find_col(header, _ID_COLS_KEYWORDS['total'])
        ach_i = _find_col(header, _ID_COLS_KEYWORDS['ach'])
        no_i = _find_col(header, _ID_COLS_KEYWORDS['no'])
        table_rooms = {}
        previous_identity = None

        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            room_name = row[room_name_i] if room_name_i is not None else ''
            grade = row[grade_i] if grade_i is not None else ''
            volume = _to_number(row[volume_i]) if volume_i is not None else None

            # Markdown OCR may leave merged identity cells blank on continuation rows.
            if not room_number and not room_name and previous_identity:
                grade, room_number, room_name, volume = previous_identity
            if not room_number and not room_name:
                continue
            previous_identity = (grade, room_number, room_name, volume)

            key = (grade, room_number, room_name, volume)
            if key not in table_rooms:
                source_no = _to_number(row[no_i], as_int=True) if no_i is not None else None
                table_rooms[key] = {
                    'no': source_no or no_counter,
                    'grade': grade,
                    'room_number': room_number,
                    'room_name': room_name,
                    'volume': volume,
                    'air_flow_measurements': [],
                    'total_air_flow': None,
                    'ach': None,
                }
                no_counter += 1
            room = table_rooms[key]

            point_label = str(row[point_i]).strip() if point_i is not None else ''
            flow_values = [_to_number(row[c]) for c in flow_cols if c < len(row)]
            flow_values = [value for value in flow_values if value is not None]
            explicit_total = _to_number(row[total_i]) if total_i is not None else None

            if '합계' in point_label:
                room['total_air_flow'] = explicit_total if explicit_total is not None else (flow_values[0] if flow_values else None)
            else:
                for value in flow_values:
                    room['air_flow_measurements'].append({
                        'point': len(room['air_flow_measurements']) + 1,
                        'air_flow': value,
                    })
                if explicit_total is not None:
                    room['total_air_flow'] = explicit_total

            ach = _to_number(row[ach_i], as_int=True) if ach_i is not None else None
            if ach is not None:
                room['ach'] = ach

        for room in table_rooms.values():
            if room['total_air_flow'] is None and room['air_flow_measurements']:
                room['total_air_flow'] = round(sum(m['air_flow'] for m in room['air_flow_measurements']), 1)
            if room['air_flow_measurements'] or room['ach'] is not None:
                rooms.append(room)

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'rooms': rooms,
    }


# =============================================================================
# D. HEPA FILTER
# =============================================================================

def parse_hepa_filter(pages_text):
    text = _join_pages(pages_text)
    readable = html_to_text(text)
    items = []
    no_counter = 1

    for table in extract_tables(text):
        if len(table) < 2:
            continue
        header, rows = table[0], table[1:]

        room_no_i = _find_col(header, _ID_COLS_KEYWORDS['room_no'])
        name_i = _find_col(header, _ID_COLS_KEYWORDS['room_name'])
        if room_no_i is None and name_i is None:
            continue

        used = _used_cols(header, 'room_no', 'room_name', 'no', 'point')
        value_cols = [i for i in range(len(header)) if i not in used]

        entries = []
        for row in rows:
            row = _align_row(row, len(header))
            if len(row) < len(header):
                continue
            room_number = row[room_no_i] if room_no_i is not None else ''
            item_name = row[name_i] if name_i is not None else ''
            if not room_number and not item_name:
                continue

            measurements = []
            for c in value_cols:
                v = _to_number(row[c]) if c < len(row) else None
                if v is None:
                    continue
                measurements.append({'point': len(measurements) + 1, 'value': v})

            if not measurements:
                continue
            entries.append({
                'room_number': room_number,
                'item_name': item_name,
                'measurements': measurements,
            })

        # group consecutive entries for the same item into one item object
        i = 0
        while i < len(entries):
            e = entries[i]
            j = i
            while j + 1 < len(entries) and \
                    entries[j + 1]['room_number'] == e['room_number'] and \
                    entries[j + 1]['item_name'] == e['item_name']:
                j += 1
            group = []
            for en in entries[i:j + 1]:
                for m in en['measurements']:
                    group.append(dict(m, point=len(group) + 1))
            items.append({
                'no_start': no_counter,
                'no_end': no_counter + len(group) - 1,
                'room_number': e['room_number'],
                'item_name': e['item_name'],
                'measurements': group,
            })
            no_counter += len(group)
            i = j + 1

    return {
        'ahu': extract_ahu(text),
        'date': extract_date(readable),
        'result': extract_result(readable),
        'standard': extract_standard(readable),
        'items': items,
    }


# =============================================================================
# E. AIRFLOW PATTERN (field-based, one item per page)
# =============================================================================

def parse_airflow_pattern(pages_text):
    items = []
    for page_text in pages_text:
        readable = html_to_text(page_text)
        tables = extract_tables(page_text)
        field_labels = {
            '측정대상', '측정일자', '결재', '측정자', '확인자',
            '측정사진', '측정기준', '측정결과', '동영상첨부', '판정결과',
        }

        def normalized(value):
            return re.sub(r'[\s:|]+', '', str(value or ''))

        def table_value(label, allowed_values=None):
            target = normalized(label)
            for table in tables:
                for row_index, row in enumerate(table):
                    for col_index, cell in enumerate(row):
                        if normalized(cell) != target:
                            continue
                        candidates = list(row[col_index + 1:])
                        candidates.extend(
                            lower_row[col_index]
                            for lower_row in table[row_index + 1:]
                            if col_index < len(lower_row)
                        )
                        for candidate in candidates:
                            candidate_norm = normalized(candidate)
                            if allowed_values:
                                for allowed in allowed_values:
                                    if candidate_norm == normalized(allowed):
                                        return allowed
                            elif candidate_norm and candidate_norm not in field_labels:
                                return str(candidate).strip()
            return ''

        name = table_value('측정대상')
        if not name:
            name = extract_field(
                readable,
                [r'측정\s*대상\s*[:|]?\s*(.+?)(?=측정\s*일자|결재|측정자|확인자|측정사진|측정기준|$)'],
                default='',
            )
        if not name:
            continue

        date = table_value('측정일자') or extract_field(
            readable,
            [r'측정\s*일자\s*[:|]?\s*(\d{4}\s*[.\-/]\s*\d{1,2}\s*[.\-/]\s*\d{1,2})'],
            default='',
        )
        criteria = table_value('측정기준') or extract_field(
            readable,
            [r'측정\s*기준\s*[:|]?\s*(.+?)(?=측정\s*결과|동영상\s*첨부|판정\s*결과|$)'],
            default='',
        )
        video = table_value('동영상 첨부', ('미첨부', '첨부'))
        judgment = table_value('판정결과', ('부적합', '적합'))

        result_start = re.search(r'측정\s*결과', readable)
        result_text = readable[result_start.start():] if result_start else readable
        if not video:
            video = '미첨부' if '미첨부' in result_text else ('첨부' if '첨부' in result_text else '')
        if not judgment:
            judgment = '부적합' if '부적합' in result_text else ('적합' if '적합' in result_text else '')

        name = re.split(r'측정\s*일자|결재|측정자|확인자', name, maxsplit=1)[0]
        name = ' '.join(name.split())
        date = re.sub(r'\s+', '', date).replace('-', '.').replace('/', '.')
        criteria = re.sub(r'\s*측정\s*결과.*$', '', criteria).strip()
        criteria = ' '.join(criteria.split())
        criteria = re.sub(r'\s*2\.\s*', '\n2. ', criteria)

        items.append({
            'name': name,
            'date': date,
            'criteria': criteria,
            'video_attached': video,
            'judgment': judgment,
        })

    ahu = extract_ahu(_join_pages(pages_text))
    return {'ahu': ahu, 'date': items[0]['date'] if items else '', 'items': items}
