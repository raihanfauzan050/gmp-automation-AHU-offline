import os
import re


_DEFAULT_AHU_BY_TEST = {
    'airborne_particle': '33',
    'airflow_pattern': '33',
}

_AHU_NUMBER_PATTERN = (
    r'(?:AHU|공조기)[_:\-|–—−]*'
    r'(?:(?:NO\.?)|번호|#)?[_:\-|–—−]*([1-9]\d*)'
)


def default_ahu_for_test(test_type):
    """Return the agreed fallback when a document has no readable AHU."""
    return _DEFAULT_AHU_BY_TEST.get(test_type, 'unknown')


def ahu_sort_key(value):
    """Sort numeric AHUs first without comparing integers and strings."""
    text = str(value).strip()
    if text.isdigit():
        return 0, int(text)
    return 1, text.casefold()


def extract_ahu_number(value, filename='', default='unknown'):
    """Prefer an explicit filename AHU, then normalize the OCR value."""
    filename_text = re.sub(r'\s+', '', os.path.basename(filename))
    filename_match = re.search(
        _AHU_NUMBER_PATTERN,
        filename_text,
        re.IGNORECASE,
    )
    if filename_match:
        return filename_match.group(1)

    value_text = str(value or '').strip()
    if re.fullmatch(r'\d+(?:\.0)?', value_text):
        number = int(float(value_text))
        if number > 0:
            return str(number)

    compact = re.sub(r'\s+', '', value_text)
    match = re.search(_AHU_NUMBER_PATTERN, compact, re.IGNORECASE)
    if match:
        return match.group(1)
    return default
