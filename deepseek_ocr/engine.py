"""Offline OCR engine for extracting structured data from scanned PDFs."""

from .client import ocr_pdf
from .parsers import (
    parse_airborne_particle,
    parse_air_velocity,
    parse_air_change_rate,
    parse_hepa_filter,
    parse_airflow_pattern,
)


def extract_airborne_particle(pdf_path, endpoint_url=None):
    return parse_airborne_particle(ocr_pdf(pdf_path, endpoint_url))


def extract_air_velocity(pdf_path, endpoint_url=None):
    return parse_air_velocity(ocr_pdf(pdf_path, endpoint_url))


def extract_air_change_rate(pdf_path, endpoint_url=None):
    return parse_air_change_rate(ocr_pdf(pdf_path, endpoint_url))


def extract_hepa_filter(pdf_path, endpoint_url=None):
    return parse_hepa_filter(ocr_pdf(pdf_path, endpoint_url))


def extract_airflow_pattern(pdf_path, endpoint_url=None):
    return parse_airflow_pattern(ocr_pdf(pdf_path, endpoint_url))


# Map test types to their extraction functions.
EXTRACTORS = {
    'airborne_particle': extract_airborne_particle,
    'air_velocity': extract_air_velocity,
    'air_change_rate': extract_air_change_rate,
    'hepa_filter': extract_hepa_filter,
    'airflow_pattern': extract_airflow_pattern,
}
