import re
import unicodedata
import csv
import io
import chardet


def detect_encoding(file_path: str) -> str:
    with open(file_path, 'rb') as f:
        raw = f.read()
    result = chardet.detect(raw)
    encoding = result.get('encoding', 'utf-8') or 'utf-8'
    confidence = result.get('confidence', 0)
    if confidence < 0.5:
        encoding = 'utf-8-sig'
    encoding_lower = encoding.lower().replace('-', '').replace('_', '')
    if encoding_lower in ('utf8', 'utf8sig', 'utf8bom'):
        if raw[:3] == b'\xef\xbb\xbf':
            return 'utf-8-sig'
        return 'utf-8'
    return encoding


def read_file_text(file_path: str, encoding: str = None) -> str:
    if encoding is None:
        encoding = detect_encoding(file_path)
    try:
        with open(file_path, 'r', encoding=encoding, errors='replace') as f:
            return f.read()
    except Exception:
        with open(file_path, 'r', encoding='utf-8-sig', errors='replace') as f:
            return f.read()


def read_file_bytes(file_path: str) -> bytes:
    with open(file_path, 'rb') as f:
        return f.read()


def normalize_doi(doi_str: str) -> str:
    if not doi_str:
        return ''
    doi = doi_str.strip()
    prefixes = [
        'https://doi.org/',
        'http://doi.org/',
        'https://dx.doi.org/',
        'http://dx.doi.org/',
        'doi.org/',
        'dx.doi.org/',
        'DOI:',
        'doi:',
        'DOI ',
        'doi ',
    ]
    for prefix in prefixes:
        if doi.lower().startswith(prefix.lower()):
            doi = doi[len(prefix):]
            break
    doi = doi.strip()
    return doi.lower()


def normalize_title(title: str) -> str:
    if not title:
        return ''
    title = unicodedata.normalize('NFKD', title)
    title = title.lower()
    title = re.sub(r'[{}\[\]()]', '', title)
    title = re.sub(r'[^\w\s]', '', title)
    title = re.sub(r'\s+', ' ', title).strip()
    return title


def extract_first_author(author_str: str, format_type: str = 'bibtex') -> str:
    if not author_str:
        return ''
    author_str = author_str.strip()

    if format_type == 'bibtex':
        first = author_str.split(' and ')[0].strip()
        first = re.sub(r'[{}]', '', first)
        if ',' in first:
            return first.split(',')[0].strip().lower()
        parts = first.split()
        return parts[-1].strip().lower() if parts else ''

    elif format_type in ('wos_txt', 'wos_xls'):
        first = author_str.split('\n')[0].strip()
        if ',' in first:
            return first.split(',')[0].strip().lower()
        parts = first.split()
        return parts[-1].strip().lower() if parts else ''

    elif format_type in ('scopus_csv', 'scopus_txt'):
        first = author_str.split(';')[0].strip()
        if ',' in first:
            return first.split(',')[0].strip().lower()
        parts = first.split()
        return parts[-1].strip().lower() if parts else ''

    elif format_type in ('ei_csv', 'ei_txt'):
        first = author_str.split(';')[0].strip()
        first = re.sub(r'\s*\([\d,\s]+\)\s*$', '', first)
        if ',' in first:
            return first.split(',')[0].strip().lower()
        parts = first.split()
        return parts[-1].strip().lower() if parts else ''

    else:
        if ',' in author_str:
            return author_str.split(',')[0].strip().lower()
        return author_str.split()[0].strip().lower() if author_str.split() else ''


def extract_first_two_authors(author_str: str, format_type: str = 'bibtex') -> list:
    if not author_str:
        return []
    results = []

    if format_type == 'bibtex':
        authors = author_str.split(' and ')
    elif format_type in ('wos_txt',):
        authors = author_str.split('\n')
    elif format_type in ('scopus_csv', 'scopus_txt', 'ei_csv', 'ei_txt', 'wos_xls'):
        authors = author_str.split(';')
    else:
        authors = [author_str]

    for a in authors[:2]:
        a = a.strip()
        if not a:
            continue
        a = re.sub(r'[{}]', '', a)
        a = re.sub(r'\s*\([\d,\s]+\)\s*$', '', a)
        if ',' in a:
            surname = a.split(',')[0].strip().lower()
        else:
            parts = a.split()
            surname = parts[-1].strip().lower() if parts else ''
        if surname:
            results.append(surname)

    return results


def detect_csv_dialect(file_path: str, encoding: str = None):
    if encoding is None:
        encoding = detect_encoding(file_path)
    with open(file_path, 'r', encoding=encoding, errors='replace') as f:
        sample = f.read(8192)
    try:
        dialect = csv.Sniffer().sniff(sample, delimiters=',;\t|')
        return dialect
    except csv.Error:
        return csv.excel


def detect_scopus_txt_language(text: str) -> str:
    if '导出日期:' in text[:200] or '摘要:' in text[:5000]:
        return 'chinese'
    return 'english'


def detect_scopus_csv_language(headers: list) -> str:
    for h in headers:
        if h.strip() in ('作者', '文献标题', '年份', '来源出版物名称'):
            return 'chinese'
    return 'english'
