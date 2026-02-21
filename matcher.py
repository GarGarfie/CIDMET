from dataclasses import dataclass, field
from typing import List, Optional, Dict, Tuple

from rapidfuzz import fuzz

from parsers import ParsedRecord
from utils import normalize_doi, normalize_title


@dataclass
class MatchResult:
    matched: bool = False
    bib_key: str = ''
    bib_index: int = -1
    match_method: str = ''
    confidence: float = 0.0
    bib_record: Optional[ParsedRecord] = None
    db_record: Optional[ParsedRecord] = None
    notes: str = ''


class MatchEngine:
    def __init__(self, bib_records: List[ParsedRecord], fuzzy_threshold: float = 90.0):
        self.bib_records = bib_records
        self.fuzzy_threshold = fuzzy_threshold

        self.doi_index: Dict[str, List[int]] = {}
        self.title_index: Dict[str, List[int]] = {}

        for i, rec in enumerate(bib_records):
            if rec.doi_normalized:
                self.doi_index.setdefault(rec.doi_normalized, []).append(i)
            if rec.title_normalized:
                self.title_index.setdefault(rec.title_normalized, []).append(i)

        self.match_count: Dict[int, List[MatchResult]] = {}
        self.unmatched_bib: set = set(range(len(bib_records)))

    def match(self, db_record: ParsedRecord) -> MatchResult:
        result = self._try_doi_match(db_record)
        if result.matched:
            self._record_match(result)
            return result

        result = self._try_title_exact_match(db_record)
        if result.matched:
            self._record_match(result)
            return result

        result = self._try_fuzzy_match(db_record)
        if result.matched:
            self._record_match(result)
            return result

        return MatchResult(matched=False, db_record=db_record, notes='No match found')

    def _try_doi_match(self, db_record: ParsedRecord) -> MatchResult:
        if not db_record.doi_normalized:
            return MatchResult(matched=False, notes='DB record has no DOI')

        indices = self.doi_index.get(db_record.doi_normalized, [])
        if not indices:
            return MatchResult(matched=False)

        bib_idx = indices[0]
        bib_rec = self.bib_records[bib_idx]
        bib_key = bib_rec.raw_fields.get('ID', bib_rec.raw_fields.get('id', str(bib_idx)))

        return MatchResult(
            matched=True,
            bib_key=bib_key,
            bib_index=bib_idx,
            match_method='DOI exact',
            confidence=100.0,
            bib_record=bib_rec,
            db_record=db_record,
        )

    def _try_title_exact_match(self, db_record: ParsedRecord) -> MatchResult:
        if not db_record.title_normalized:
            return MatchResult(matched=False, notes='DB record has no title')

        indices = self.title_index.get(db_record.title_normalized, [])
        if not indices:
            return MatchResult(matched=False)

        bib_idx = indices[0]
        bib_rec = self.bib_records[bib_idx]
        bib_key = bib_rec.raw_fields.get('ID', bib_rec.raw_fields.get('id', str(bib_idx)))

        return MatchResult(
            matched=True,
            bib_key=bib_key,
            bib_index=bib_idx,
            match_method='Title exact',
            confidence=99.0,
            bib_record=bib_rec,
            db_record=db_record,
        )

    def _try_fuzzy_match(self, db_record: ParsedRecord) -> MatchResult:
        if not db_record.title_normalized:
            return MatchResult(matched=False, notes='DB record has no title for fuzzy match')

        best_score = 0.0
        best_idx = -1

        for i, bib_rec in enumerate(self.bib_records):
            if not bib_rec.title_normalized:
                continue

            title_score = fuzz.ratio(db_record.title_normalized, bib_rec.title_normalized)
            if title_score < self.fuzzy_threshold:
                continue

            year_ok = True
            if db_record.year and bib_rec.year:
                if db_record.year.strip() != bib_rec.year.strip():
                    year_ok = False
            elif not db_record.year and not bib_rec.year:
                year_ok = True

            author_ok = True
            if db_record.first_author and bib_rec.first_author:
                author_score = fuzz.ratio(db_record.first_author, bib_rec.first_author)
                if author_score < 80:
                    db_authors = set(db_record.first_two_authors)
                    bib_authors = set(bib_rec.first_two_authors)
                    if not db_authors.intersection(bib_authors):
                        author_ok = False
            elif not db_record.first_author or not bib_rec.first_author:
                author_ok = True

            if not year_ok and not author_ok:
                continue

            combined = title_score
            if year_ok and db_record.year and bib_rec.year:
                combined += 2
            if author_ok and db_record.first_author and bib_rec.first_author:
                combined += 2

            if combined > best_score:
                best_score = combined
                best_idx = i

        if best_idx >= 0:
            bib_rec = self.bib_records[best_idx]
            bib_key = bib_rec.raw_fields.get('ID', bib_rec.raw_fields.get('id', str(best_idx)))
            title_score = fuzz.ratio(db_record.title_normalized, bib_rec.title_normalized)

            notes_parts = []
            if not db_record.year or not bib_rec.year:
                notes_parts.append('year missing in one record')
            if not db_record.first_author or not bib_rec.first_author:
                notes_parts.append('author missing in one record')

            return MatchResult(
                matched=True,
                bib_key=bib_key,
                bib_index=best_idx,
                match_method=f'Fuzzy (title={title_score:.1f}%)',
                confidence=title_score,
                bib_record=bib_rec,
                db_record=db_record,
                notes='; '.join(notes_parts) if notes_parts else '',
            )

        return MatchResult(matched=False, db_record=db_record)

    def _record_match(self, result: MatchResult):
        bib_idx = result.bib_index
        self.match_count.setdefault(bib_idx, []).append(result)
        self.unmatched_bib.discard(bib_idx)

    def get_duplicates(self) -> Dict[int, List[MatchResult]]:
        return {k: v for k, v in self.match_count.items() if len(v) > 1}

    def get_unmatched_bib_records(self) -> List[Tuple[int, ParsedRecord]]:
        return [(i, self.bib_records[i]) for i in sorted(self.unmatched_bib)]

    def get_match_stats(self) -> dict:
        total_bib = len(self.bib_records)
        matched_bib = total_bib - len(self.unmatched_bib)
        total_matches = sum(len(v) for v in self.match_count.values())
        duplicates = len(self.get_duplicates())
        return {
            'total_bib': total_bib,
            'matched_bib': matched_bib,
            'unmatched_bib': len(self.unmatched_bib),
            'total_db_matches': total_matches,
            'duplicate_bib_entries': duplicates,
            'fuzzy_threshold': self.fuzzy_threshold,
        }
