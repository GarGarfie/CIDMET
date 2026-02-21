import re
import csv
import io
from dataclasses import dataclass, field
from typing import List, Dict, Optional, Tuple

import bibtexparser
import xlrd

from utils import (
    detect_encoding, read_file_text, normalize_doi, normalize_title,
    extract_first_author, extract_first_two_authors, detect_csv_dialect,
    detect_scopus_txt_language, detect_scopus_csv_language,
)


@dataclass
class ParsedRecord:
    doi: str = ''
    doi_normalized: str = ''
    title: str = ''
    title_normalized: str = ''
    year: str = ''
    authors: str = ''
    first_author: str = ''
    first_two_authors: list = field(default_factory=list)
    journal: str = ''
    volume: str = ''
    pages: str = ''
    issn: str = ''
    isbn: str = ''
    abstract: str = ''
    keywords: str = ''
    raw_fields: dict = field(default_factory=dict)
    source_db: str = ''
    source_format: str = ''
    raw_block_index: int = -1


# ---------------------------------------------------------------------------
# BibTeX Parser
# ---------------------------------------------------------------------------

def parse_bibtex(file_path: str) -> List[ParsedRecord]:
    encoding = detect_encoding(file_path)
    text = read_file_text(file_path, encoding)
    parser = bibtexparser.bparser.BibTexParser(common_strings=True)
    parser.ignore_nonstandard_types = False
    bib_db = bibtexparser.loads(text, parser=parser)

    records = []
    for entry in bib_db.entries:
        raw_title = entry.get('title', '')
        raw_title = re.sub(r'[{}]', '', raw_title).strip()
        raw_doi = entry.get('doi', '')
        raw_author = entry.get('author', '')
        raw_year = entry.get('year', '')
        if not raw_year:
            raw_date = entry.get('date', '')
            year_match = re.match(r'(\d{4})', str(raw_date))
            raw_year = year_match.group(1) if year_match else ''
        raw_year = re.sub(r'[{}]', '', str(raw_year)).strip()

        journal_field = entry.get('journal', '') or entry.get('journaltitle', '') or entry.get('booktitle', '')

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=raw_author,
            first_author=extract_first_author(raw_author, 'bibtex'),
            first_two_authors=extract_first_two_authors(raw_author, 'bibtex'),
            journal=re.sub(r'[{}]', '', journal_field).strip(),
            volume=entry.get('volume', ''),
            pages=entry.get('pages', ''),
            issn=entry.get('issn', ''),
            isbn=entry.get('isbn', ''),
            abstract=entry.get('abstract', ''),
            keywords=entry.get('keywords', ''),
            raw_fields=dict(entry),
            source_db='bibtex',
            source_format='bib',
            raw_block_index=-1,
        )
        records.append(rec)
    return records


# ---------------------------------------------------------------------------
# WoS TXT Parser
# ---------------------------------------------------------------------------

def _wos_txt_extract_field(block: str, tag: str) -> str:
    lines = block.split('\n')
    collecting = False
    parts = []
    for line in lines:
        if line.startswith(tag + ' ') or line.startswith(tag + '\t'):
            collecting = True
            val = line[len(tag):].strip()
            parts.append(val)
        elif collecting:
            if line.startswith('   '):
                parts.append(line.strip())
            else:
                collecting = False
    return ' '.join(parts)


def _wos_txt_extract_all_fields(block: str) -> dict:
    lines = block.split('\n')
    fields = {}
    current_tag = None
    current_parts = []
    for line in lines:
        if len(line) >= 3 and line[:2].strip() and line[2] == ' ':
            if current_tag is not None:
                val = '\n'.join(current_parts)
                if current_tag in fields:
                    fields[current_tag] += '\n' + val
                else:
                    fields[current_tag] = val
            current_tag = line[:2].strip()
            current_parts = [line[3:].strip()]
        elif line.startswith('   ') and current_tag is not None:
            current_parts.append(line.strip())
        elif line.strip() == '' and current_tag is not None:
            pass
        else:
            if current_tag is not None:
                val = '\n'.join(current_parts)
                if current_tag in fields:
                    fields[current_tag] += '\n' + val
                else:
                    fields[current_tag] = val
                current_tag = None
                current_parts = []
    if current_tag is not None:
        val = '\n'.join(current_parts)
        if current_tag in fields:
            fields[current_tag] += '\n' + val
        else:
            fields[current_tag] = val
    return fields


def parse_wos_txt(file_path: str) -> Tuple[List[ParsedRecord], List[str], str, str]:
    encoding = detect_encoding(file_path)
    text = read_file_text(file_path, encoding)

    parts = re.split(r'\nER\s*\n', text)
    header = ''
    raw_blocks = []
    footer = ''

    if parts:
        first = parts[0]
        fn_match = re.match(r'(.*?)(PT\s)', first, re.DOTALL)
        if fn_match:
            header = fn_match.group(1)
            first_record = fn_match.group(2) + first[fn_match.end():]
            raw_blocks.append(first_record.rstrip() + '\nER')
        else:
            raw_blocks.append(first.rstrip() + '\nER')

        for part in parts[1:]:
            stripped = part.strip()
            if not stripped:
                continue
            if stripped.startswith('EF'):
                footer = part
                continue
            raw_blocks.append(part.rstrip() + '\nER')

    records = []
    for i, block in enumerate(raw_blocks):
        fields = _wos_txt_extract_all_fields(block)
        raw_doi = fields.get('DI', '')
        raw_title = fields.get('TI', '')
        raw_year = fields.get('PY', '')
        raw_author_short = fields.get('AU', '')
        raw_author_full = fields.get('AF', '')
        authors_str = raw_author_full if raw_author_full else raw_author_short

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year.strip(),
            authors=authors_str,
            first_author=extract_first_author(raw_author_short.split('\n')[0] if raw_author_short else '', 'wos_txt'),
            first_two_authors=extract_first_two_authors(raw_author_short, 'wos_txt'),
            journal=fields.get('SO', ''),
            volume=fields.get('VL', ''),
            pages=fields.get('BP', '') + ('-' + fields.get('EP', '') if fields.get('EP', '') else ''),
            issn=fields.get('SN', ''),
            isbn=fields.get('BN', ''),
            abstract=fields.get('AB', ''),
            keywords=fields.get('DE', ''),
            raw_fields=fields,
            source_db='wos',
            source_format='txt',
            raw_block_index=i,
        )
        records.append(rec)

    return records, raw_blocks, header, footer


# ---------------------------------------------------------------------------
# WoS XLS Parser
# ---------------------------------------------------------------------------

def parse_wos_xls(file_path: str) -> Tuple[List[ParsedRecord], object, List[str]]:
    wb = xlrd.open_workbook(file_path)
    ws = wb.sheet_by_index(0)

    headers = [str(ws.cell_value(0, c)).strip() for c in range(ws.ncols)]

    col_map = {}
    for i, h in enumerate(headers):
        col_map[h] = i

    doi_col = col_map.get('DOI', -1)
    title_col = col_map.get('Article Title', -1)
    year_col = col_map.get('Publication Year', -1)
    author_col = col_map.get('Authors', -1)
    author_full_col = col_map.get('Author Full Names', -1)

    records = []
    for row_idx in range(1, ws.nrows):
        def cell_val(c):
            if c < 0 or c >= ws.ncols:
                return ''
            v = ws.cell_value(row_idx, c)
            if isinstance(v, float) and v == int(v):
                return str(int(v))
            return str(v).strip()

        raw_doi = cell_val(doi_col)
        raw_title = cell_val(title_col)
        raw_year = cell_val(year_col)
        raw_author = cell_val(author_col)
        raw_author_full = cell_val(author_full_col) if author_full_col >= 0 else ''

        raw_fields = {}
        for h_name, c_idx in col_map.items():
            raw_fields[h_name] = cell_val(c_idx)

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=raw_author_full if raw_author_full else raw_author,
            first_author=extract_first_author(raw_author, 'wos_xls'),
            first_two_authors=extract_first_two_authors(raw_author, 'wos_xls'),
            journal=raw_fields.get('Source Title', ''),
            volume=raw_fields.get('Volume', ''),
            pages=(raw_fields.get('Start Page', '') + ('-' + raw_fields.get('End Page', '') if raw_fields.get('End Page', '') else '')),
            issn=raw_fields.get('ISSN', ''),
            isbn=raw_fields.get('ISBN', ''),
            abstract=raw_fields.get('Abstract', ''),
            keywords=raw_fields.get('Author Keywords', ''),
            raw_fields=raw_fields,
            source_db='wos',
            source_format='xls',
            raw_block_index=row_idx,
        )
        records.append(rec)

    return records, wb, headers


# ---------------------------------------------------------------------------
# Scopus CSV Parser
# ---------------------------------------------------------------------------

def parse_scopus_csv(file_path: str) -> Tuple[List[ParsedRecord], List[List[str]], str, object, str]:
    encoding = detect_encoding(file_path)
    dialect = detect_csv_dialect(file_path, encoding)

    with open(file_path, 'r', encoding=encoding, errors='replace', newline='') as f:
        reader = csv.reader(f, dialect)
        rows = list(reader)

    if not rows:
        return [], [], '', dialect, encoding

    headers = [h.strip() for h in rows[0]]
    lang = detect_scopus_csv_language(headers)

    if lang == 'chinese':
        title_key, year_key, author_key = '文献标题', '年份', '作者'
    else:
        title_key, year_key, author_key = 'Title', 'Year', 'Authors'

    col_map = {}
    for i, h in enumerate(headers):
        col_map[h] = i

    doi_col = col_map.get('DOI', -1)
    title_col = col_map.get(title_key, -1)
    year_col = col_map.get(year_key, -1)
    author_col = col_map.get(author_key, -1)

    records = []
    raw_data_rows = []
    for row_idx, row in enumerate(rows[1:], start=1):
        raw_data_rows.append(row)

        def cell_val(c):
            if c < 0 or c >= len(row):
                return ''
            return row[c].strip()

        raw_doi = cell_val(doi_col)
        raw_title = cell_val(title_col)
        raw_year = cell_val(year_col)
        raw_author = cell_val(author_col)

        raw_fields = {}
        for h_name, c_idx in col_map.items():
            if c_idx < len(row):
                raw_fields[h_name] = row[c_idx]
            else:
                raw_fields[h_name] = ''

        src_key = 'Source title' if 'Source title' in col_map else ('来源出版物名称' if '来源出版物名称' in col_map else '')
        vol_key = 'Volume' if 'Volume' in col_map else ('卷' if '卷' in col_map else '')
        issn_key = 'ISSN' if 'ISSN' in col_map else 'ISSN'
        isbn_key = 'ISBN' if 'ISBN' in col_map else 'ISBN'
        abs_key = 'Abstract' if 'Abstract' in col_map else ('摘要' if '摘要' in col_map else '')
        kw_key = 'Author Keywords' if 'Author Keywords' in col_map else ('作者关键字' if '作者关键字' in col_map else '')

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=raw_author,
            first_author=extract_first_author(raw_author, 'scopus_csv'),
            first_two_authors=extract_first_two_authors(raw_author, 'scopus_csv'),
            journal=raw_fields.get(src_key, ''),
            volume=raw_fields.get(vol_key, ''),
            pages='',
            issn=raw_fields.get(issn_key, ''),
            isbn=raw_fields.get(isbn_key, ''),
            abstract=raw_fields.get(abs_key, ''),
            keywords=raw_fields.get(kw_key, ''),
            raw_fields=raw_fields,
            source_db='scopus',
            source_format='csv',
            raw_block_index=row_idx,
        )

        ps = raw_fields.get('Page start' if 'Page start' in col_map else ('起始页码' if '起始页码' in col_map else ''), '')
        pe = raw_fields.get('Page end' if 'Page end' in col_map else ('结束页码' if '结束页码' in col_map else ''), '')
        if ps:
            rec.pages = ps + ('-' + pe if pe else '')

        records.append(rec)

    return records, raw_data_rows, headers, dialect, encoding


# ---------------------------------------------------------------------------
# Scopus TXT Parser
# ---------------------------------------------------------------------------

def _split_scopus_txt_records(text: str) -> Tuple[str, List[str]]:
    lines = text.split('\n')
    header_lines = []
    body_start = 0

    for i, line in enumerate(lines):
        if line.strip().startswith('Scopus') or line.strip().startswith('EXPORT DATE:') or line.strip().startswith('导出日期:'):
            header_lines.append(line)
            body_start = i + 1
        elif line.strip() == '' and i <= body_start + 1:
            header_lines.append(line)
            body_start = i + 1
        else:
            break

    header = '\n'.join(header_lines)
    body = '\n'.join(lines[body_start:])

    raw_blocks = []
    current_block_lines = []
    blank_count = 0

    for line in lines[body_start:]:
        if line.strip() == '':
            blank_count += 1
            current_block_lines.append(line)
        else:
            if blank_count >= 1 and current_block_lines:
                content_lines = [l for l in current_block_lines if l.strip()]
                if content_lines:
                    block_text = '\n'.join(current_block_lines)
                    has_author = False
                    for cl in content_lines[:3]:
                        if re.match(r'^[A-Z][a-z]+,\s', cl) or cl.startswith('AUTHOR FULL NAMES:'):
                            has_author = True
                            break
                    if has_author and raw_blocks:
                        raw_blocks.append(block_text)
                        current_block_lines = [line]
                        blank_count = 0
                        continue
                    elif not raw_blocks and has_author:
                        raw_blocks.append(block_text)
                        current_block_lines = [line]
                        blank_count = 0
                        continue
            blank_count = 0
            current_block_lines.append(line)

    if current_block_lines:
        content_lines = [l for l in current_block_lines if l.strip()]
        if content_lines:
            raw_blocks.append('\n'.join(current_block_lines))

    if not raw_blocks:
        chunks = re.split(r'\n\n(?=[A-Z][a-z]+,\s)', body)
        raw_blocks = [c for c in chunks if c.strip()]

    return header, raw_blocks


def _parse_scopus_txt_block(block: str, lang: str) -> dict:
    lines = block.strip().split('\n')
    non_empty = [l for l in lines if l.strip()]
    if not non_empty:
        return {}

    result = {'_raw_lines': lines}
    author_short = ''
    author_full = ''
    title = ''
    doi = ''
    year = ''

    line_idx = 0
    for i, line in enumerate(non_empty):
        stripped = line.strip()
        if i == 0 and re.match(r'^[A-Z][a-z]', stripped):
            author_short = stripped
            result['author_short'] = author_short
            line_idx = i
            continue
        if stripped.startswith('AUTHOR FULL NAMES:'):
            author_full = stripped[len('AUTHOR FULL NAMES:'):].strip()
            result['author_full'] = author_full
            continue
        if re.match(r'^\d{8,};\s', stripped):
            continue
        if stripped.startswith('DOI:'):
            doi = stripped[4:].strip()
            result['DOI'] = doi
            continue
        if stripped.startswith('https://www.scopus.com/'):
            result['link'] = stripped
            continue

        label_map_en = {
            'AFFILIATIONS:': 'affiliations',
            'ABSTRACT:': 'abstract',
            'AUTHOR KEYWORDS:': 'author_keywords',
            'INDEX KEYWORDS:': 'index_keywords',
            'FUNDING DETAILS:': 'funding_details',
            'FUNDING TEXT': 'funding_text',
            'REFERENCES:': 'references',
            'CORRESPONDENCE ADDRESS:': 'correspondence',
            'PUBLISHER:': 'publisher',
            'LANGUAGE OF ORIGINAL DOCUMENT:': 'language',
            'ABBREVIATED SOURCE TITLE:': 'abbreviated_source',
            'DOCUMENT TYPE:': 'document_type',
            'PUBLICATION STAGE:': 'publication_stage',
            'OPEN ACCESS:': 'open_access',
            'ISSN:': 'issn',
            'ISBN:': 'isbn',
        }
        label_map_cn = {
            '归属机构:': 'affiliations',
            '摘要:': 'abstract',
            '作者关键字:': 'author_keywords',
            '索引关键字:': 'index_keywords',
            '出资详情:': 'funding_details',
            '资金资助文本': 'funding_text',
            '参考文献:': 'references',
            '通讯地址:': 'correspondence',
            '出版商:': 'publisher',
            '原始文献语言:': 'language',
            '来源出版物名称缩写:': 'abbreviated_source',
            '文献类型:': 'document_type',
            '出版阶段:': 'publication_stage',
            '开放获取:': 'open_access',
            'ISSN:': 'issn',
            'ISBN:': 'isbn',
        }
        label_map = label_map_cn if lang == 'chinese' else label_map_en

        matched_label = False
        for label, key in label_map.items():
            if stripped.startswith(label):
                result[key] = stripped[len(label):].strip()
                matched_label = True
                break

        if matched_label:
            continue

        year_source_match = re.match(r'^\((\d{4})\)\s+(.+?)(?:,|$)', stripped)
        if year_source_match:
            year = year_source_match.group(1)
            result['year'] = year
            result['source_line'] = stripped
            continue

        if not title:
            if i >= 1 and not stripped.startswith('DOI:') and not stripped.startswith('http'):
                if not any(stripped.startswith(l) for l in list(label_map_en.keys()) + list(label_map_cn.keys())):
                    if not re.match(r'^\d{8,}', stripped):
                        title = stripped
                        result['title'] = title

    return result


def parse_scopus_txt(file_path: str) -> Tuple[List[ParsedRecord], List[str], str]:
    encoding = detect_encoding(file_path)
    text = read_file_text(file_path, encoding)
    lang = detect_scopus_txt_language(text)

    lines = text.split('\n')
    header_end = 0
    for i, line in enumerate(lines):
        if line.strip() == '' and i > 0:
            header_end = i + 1
            break
    header_text = '\n'.join(lines[:header_end])

    body_text = '\n'.join(lines[header_end:])

    author_line_pattern = re.compile(r'^[A-Z][a-z]+(?:[-\'][A-Z])?.*?,\s+[A-Z]')

    record_starts = []
    body_lines = body_text.split('\n')
    for i, line in enumerate(body_lines):
        stripped = line.strip()
        if stripped and author_line_pattern.match(stripped):
            if i == 0 or body_lines[i - 1].strip() == '':
                record_starts.append(i)

    raw_blocks = []
    for idx, start in enumerate(record_starts):
        if idx + 1 < len(record_starts):
            end = record_starts[idx + 1]
            while end > start and body_lines[end - 1].strip() == '':
                end -= 1
            block = '\n'.join(body_lines[start:end])
        else:
            block = '\n'.join(body_lines[start:])
            block = block.rstrip('\n')
        raw_blocks.append(block)

    records = []
    for i, block in enumerate(raw_blocks):
        parsed = _parse_scopus_txt_block(block, lang)
        if not parsed:
            continue

        raw_doi = parsed.get('DOI', '')
        raw_title = parsed.get('title', '')
        raw_year = parsed.get('year', '')
        raw_author = parsed.get('author_short', '')

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=parsed.get('author_full', raw_author),
            first_author=extract_first_author(raw_author, 'scopus_txt'),
            first_two_authors=extract_first_two_authors(raw_author, 'scopus_txt'),
            journal=parsed.get('abbreviated_source', ''),
            volume='',
            pages='',
            issn=parsed.get('issn', ''),
            isbn=parsed.get('isbn', ''),
            abstract=parsed.get('abstract', ''),
            keywords=parsed.get('author_keywords', ''),
            raw_fields=parsed,
            source_db='scopus',
            source_format='txt',
            raw_block_index=i,
        )
        records.append(rec)

    return records, raw_blocks, header_text


# ---------------------------------------------------------------------------
# EI CSV Parser
# ---------------------------------------------------------------------------

def parse_ei_csv(file_path: str) -> Tuple[List[ParsedRecord], List[List[str]], List[str], object, str]:
    encoding = detect_encoding(file_path)
    dialect = detect_csv_dialect(file_path, encoding)

    with open(file_path, 'r', encoding=encoding, errors='replace', newline='') as f:
        reader = csv.reader(f, dialect)
        rows = list(reader)

    if not rows:
        return [], [], [], dialect, encoding

    headers = [h.strip() for h in rows[0]]

    col_map = {}
    for i, h in enumerate(headers):
        col_map[h] = i

    doi_col = col_map.get('DOI', -1)
    title_col = col_map.get('Title', -1)
    year_col = col_map.get('Publication year', -1)
    author_col = col_map.get('Author', -1)

    records = []
    raw_data_rows = []
    for row_idx, row in enumerate(rows[1:], start=1):
        raw_data_rows.append(row)

        def cell_val(c):
            if c < 0 or c >= len(row):
                return ''
            return row[c].strip()

        raw_doi = cell_val(doi_col)
        raw_title = cell_val(title_col)
        raw_year = cell_val(year_col)
        raw_author = cell_val(author_col)

        raw_fields = {}
        for h_name, c_idx in col_map.items():
            if c_idx < len(row):
                raw_fields[h_name] = row[c_idx]
            else:
                raw_fields[h_name] = ''

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=raw_author,
            first_author=extract_first_author(raw_author, 'ei_csv'),
            first_two_authors=extract_first_two_authors(raw_author, 'ei_csv'),
            journal=raw_fields.get('Source', raw_fields.get('Abbreviated source title', '')),
            volume=raw_fields.get('Volume', ''),
            pages=raw_fields.get('Pages', ''),
            issn=raw_fields.get('ISSN', ''),
            isbn=raw_fields.get('ISBN13', ''),
            abstract=raw_fields.get('Abstract', ''),
            keywords=raw_fields.get('Controlled/Subject terms', ''),
            raw_fields=raw_fields,
            source_db='ei',
            source_format='csv',
            raw_block_index=row_idx,
        )
        records.append(rec)

    return records, raw_data_rows, headers, dialect, encoding


# ---------------------------------------------------------------------------
# EI TXT Parser
# ---------------------------------------------------------------------------

def parse_ei_txt(file_path: str) -> Tuple[List[ParsedRecord], List[str]]:
    encoding = detect_encoding(file_path)
    text = read_file_text(file_path, encoding)

    record_pattern = re.compile(r'<RECORD\s+\d+>')
    parts = record_pattern.split(text)
    markers = record_pattern.findall(text)

    raw_blocks = []
    for i, marker in enumerate(markers):
        block_content = parts[i + 1] if i + 1 < len(parts) else ''
        raw_blocks.append(marker + block_content.rstrip('\n'))

    records = []
    for i, block in enumerate(raw_blocks):
        fields = {}
        lines = block.split('\n')
        current_key = None
        current_val_lines = []

        for line in lines:
            match = re.match(r'^([A-Za-z][A-Za-z /\-()]+?):(.*)', line)
            if match:
                if current_key is not None:
                    fields[current_key] = '\n'.join(current_val_lines).strip()
                current_key = match.group(1).strip()
                current_val_lines = [match.group(2)]
            elif current_key is not None:
                current_val_lines.append(line)

        if current_key is not None:
            fields[current_key] = '\n'.join(current_val_lines).strip()

        raw_doi = fields.get('DOI', '')
        raw_title = fields.get('Title', '')
        raw_year = fields.get('Publication year', '')
        raw_author = fields.get('Authors', '')

        rec = ParsedRecord(
            doi=raw_doi,
            doi_normalized=normalize_doi(raw_doi),
            title=raw_title,
            title_normalized=normalize_title(raw_title),
            year=raw_year,
            authors=raw_author,
            first_author=extract_first_author(raw_author, 'ei_txt'),
            first_two_authors=extract_first_two_authors(raw_author, 'ei_txt'),
            journal=fields.get('Source title', fields.get('Abbreviated source title', '')),
            volume=fields.get('Volume', ''),
            pages=fields.get('Pages', ''),
            issn=fields.get('ISSN', ''),
            isbn=fields.get('ISBN-13', ''),
            abstract=fields.get('Abstract', ''),
            keywords=fields.get('Controlled terms', ''),
            raw_fields=fields,
            source_db='ei',
            source_format='txt',
            raw_block_index=i,
        )
        records.append(rec)

    return records, raw_blocks
