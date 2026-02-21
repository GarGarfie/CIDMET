import os
import sys
import csv
import traceback
from typing import List, Dict, Optional, Tuple

from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QGridLayout,
    QLabel, QPushButton, QLineEdit, QFileDialog, QTextEdit,
    QProgressBar, QSlider, QGroupBox, QTabWidget, QTableWidget,
    QTableWidgetItem, QHeaderView, QComboBox, QSplitter, QMessageBox,
    QDialog, QDialogButtonBox, QCheckBox, QApplication, QScrollArea,
)
from PySide6.QtCore import Qt, QThread, Signal, QMimeData, QUrl
from PySide6.QtGui import QDragEnterEvent, QDropEvent

from parsers import (
    ParsedRecord, parse_bibtex,
    parse_wos_txt, parse_wos_xls,
    parse_scopus_csv, parse_scopus_txt,
    parse_ei_csv, parse_ei_txt,
)
from matcher import MatchEngine, MatchResult
from writers import (
    write_wos_txt_subset, write_wos_xls_subset, write_csv_subset,
    write_scopus_txt_subset, write_ei_txt_subset,
    record_to_unified, deduplicate_records, export_merged,
)
from utils import detect_encoding


# ---------------------------------------------------------------------------
# Drag-drop line edit
# ---------------------------------------------------------------------------

class DragDropLineEdit(QLineEdit):
    def __init__(self, file_filter: str = '', parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.file_filter = file_filter
        self.setReadOnly(True)
        self.setPlaceholderText('Drag file here or click Browse...')

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            for url in event.mimeData().urls():
                path = url.toLocalFile()
                if self._check_extension(path):
                    event.acceptProposedAction()
                    return
        event.ignore()

    def dropEvent(self, event: QDropEvent):
        for url in event.mimeData().urls():
            path = url.toLocalFile()
            if self._check_extension(path):
                self.setText(path)
                event.acceptProposedAction()
                return

    def _check_extension(self, path: str) -> bool:
        if not self.file_filter:
            return True
        ext = os.path.splitext(path)[1].lower()
        allowed = [e.strip().lower() for e in self.file_filter.split(',')]
        return ext in allowed


# ---------------------------------------------------------------------------
# Worker thread
# ---------------------------------------------------------------------------

class ProcessingWorker(QThread):
    progress = Signal(int, str)
    log = Signal(str)
    finished_signal = Signal(dict)
    error = Signal(str)

    def __init__(self, params: dict):
        super().__init__()
        self.params = params
        self._cancelled = False

    def run(self):
        try:
            result = self._process()
            if not self._cancelled:
                self.finished_signal.emit(result)
        except Exception as e:
            self.error.emit(f'Error: {e}\n{traceback.format_exc()}')

    def cancel(self):
        self._cancelled = True

    def _process(self) -> dict:
        bibtex_path = self.params['bibtex_path']
        output_dir = self.params['output_dir']
        fuzzy_threshold = self.params['fuzzy_threshold']
        inputs = self.params['inputs']

        total_steps = 2 + len([v for v in inputs.values() if v])
        current_step = 0

        # Step 1: Parse BibTeX
        self.log.emit('Parsing BibTeX file...')
        self.progress.emit(int(current_step / total_steps * 100), 'Parsing BibTeX...')
        bib_records = parse_bibtex(bibtex_path)
        self.log.emit(f'  Found {len(bib_records)} BibTeX entries')
        current_step += 1

        # Step 2: Initialize matcher
        engine = MatchEngine(bib_records, fuzzy_threshold)

        all_match_results = {}
        all_subset_info = {}
        all_matched_records = []
        per_db_stats = {}

        db_configs = {
            'wos_xls': {'label': 'WoS XLS', 'db': 'wos'},
            'wos_txt': {'label': 'WoS TXT', 'db': 'wos'},
            'scopus_csv': {'label': 'Scopus CSV', 'db': 'scopus'},
            'scopus_txt': {'label': 'Scopus TXT', 'db': 'scopus'},
            'ei_csv': {'label': 'EI CSV', 'db': 'ei'},
            'ei_txt': {'label': 'EI TXT', 'db': 'ei'},
        }

        for db_key, config in db_configs.items():
            file_path = inputs.get(db_key, '')
            if not file_path or not os.path.isfile(file_path):
                continue

            if self._cancelled:
                return {}

            label = config['label']
            self.log.emit(f'\nProcessing {label}: {os.path.basename(file_path)}')
            self.progress.emit(int(current_step / total_steps * 100), f'Processing {label}...')

            try:
                match_results, subset_data = self._process_single_db(
                    db_key, file_path, engine, output_dir
                )
                all_match_results[db_key] = match_results
                all_subset_info[db_key] = subset_data

                matched_count = sum(1 for r in match_results if r.matched)
                total_count = len(match_results)
                self.log.emit(f'  Matched: {matched_count}/{total_count}')

                per_db_stats[db_key] = {
                    'label': label,
                    'total': total_count,
                    'matched': matched_count,
                    'unmatched': total_count - matched_count,
                }

                for r in match_results:
                    if r.matched and r.db_record:
                        all_matched_records.append(r.db_record)

            except Exception as e:
                self.log.emit(f'  ERROR processing {label}: {e}')
                self.log.emit(traceback.format_exc())

            current_step += 1

        # Final stats
        self.progress.emit(int(current_step / total_steps * 100), 'Generating summary...')
        stats = engine.get_match_stats()
        unmatched = engine.get_unmatched_bib_records()
        duplicates = engine.get_duplicates()

        self.log.emit(f'\n=== Summary ===')
        self.log.emit(f'BibTeX entries: {stats["total_bib"]}')
        self.log.emit(f'Matched BibTeX entries: {stats["matched_bib"]}')
        self.log.emit(f'Unmatched BibTeX entries: {stats["unmatched_bib"]}')
        self.log.emit(f'Total DB record matches: {stats["total_db_matches"]}')
        self.log.emit(f'Duplicate BibTeX entries (matched by multiple DB records): {stats["duplicate_bib_entries"]}')

        if unmatched:
            self.log.emit(f'\nUnmatched BibTeX entries:')
            for idx, rec in unmatched:
                self.log.emit(f'  [{idx}] {rec.title[:80]}...' if len(rec.title) > 80 else f'  [{idx}] {rec.title}')

        current_step += 1
        self.progress.emit(100, 'Done')

        return {
            'match_results': all_match_results,
            'per_db_stats': per_db_stats,
            'global_stats': stats,
            'unmatched_bib': unmatched,
            'duplicates': duplicates,
            'matched_records': all_matched_records,
            'engine': engine,
        }

    def _process_single_db(
        self, db_key: str, file_path: str, engine: MatchEngine, output_dir: str
    ) -> Tuple[List[MatchResult], dict]:

        encoding = detect_encoding(file_path)
        ext = os.path.splitext(file_path)[1].lower()
        base_name = os.path.splitext(os.path.basename(file_path))[0]
        subset_data = {'file_path': file_path, 'db_key': db_key}

        if db_key == 'wos_txt':
            records, raw_blocks, header, footer = parse_wos_txt(file_path)
            match_results = [engine.match(r) for r in records]
            matched_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.txt')
            write_wos_txt_subset(matched_indices, raw_blocks, header, footer, out_path, encoding)
            self.log.emit(f'  Subset written: {out_path}')
            subset_data['output_path'] = out_path

        elif db_key == 'wos_xls':
            records, wb, headers = parse_wos_xls(file_path)
            match_results = [engine.match(r) for r in records]
            matched_row_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.xls')
            actual_path = write_wos_xls_subset(matched_row_indices, wb, headers, out_path)
            if actual_path != out_path:
                self.log.emit(f'  Note: Output as .xlsx instead of .xls')
            self.log.emit(f'  Subset written: {actual_path}')
            subset_data['output_path'] = actual_path

        elif db_key == 'scopus_csv':
            records, raw_rows, headers, dialect, enc = parse_scopus_csv(file_path)
            match_results = [engine.match(r) for r in records]
            matched_row_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.csv')
            write_csv_subset(matched_row_indices, raw_rows, headers, dialect, enc, out_path)
            self.log.emit(f'  Subset written: {out_path}')
            subset_data['output_path'] = out_path

        elif db_key == 'scopus_txt':
            records, raw_blocks, header_text = parse_scopus_txt(file_path)
            match_results = [engine.match(r) for r in records]
            matched_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.txt')
            write_scopus_txt_subset(matched_indices, raw_blocks, header_text, out_path, encoding)
            self.log.emit(f'  Subset written: {out_path}')
            subset_data['output_path'] = out_path

        elif db_key == 'ei_csv':
            records, raw_rows, headers, dialect, enc = parse_ei_csv(file_path)
            match_results = [engine.match(r) for r in records]
            matched_row_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.csv')
            write_csv_subset(matched_row_indices, raw_rows, headers, dialect, enc, out_path)
            self.log.emit(f'  Subset written: {out_path}')
            subset_data['output_path'] = out_path

        elif db_key == 'ei_txt':
            records, raw_blocks = parse_ei_txt(file_path)
            match_results = [engine.match(r) for r in records]
            matched_indices = [records[i].raw_block_index for i, r in enumerate(match_results) if r.matched]
            out_path = os.path.join(output_dir, f'{base_name}_subset.txt')
            write_ei_txt_subset(matched_indices, raw_blocks, out_path, encoding)
            self.log.emit(f'  Subset written: {out_path}')
            subset_data['output_path'] = out_path

        else:
            match_results = []

        return match_results, subset_data


# ---------------------------------------------------------------------------
# Duplicate Dialog
# ---------------------------------------------------------------------------

class DuplicateDialog(QDialog):
    def __init__(self, duplicates: dict, bib_records: list, parent=None):
        super().__init__(parent)
        self.setWindowTitle('Duplicate Entries')
        self.setMinimumSize(900, 600)
        self.duplicates = duplicates
        self.bib_records = bib_records
        self.selections = {}
        self._build_ui()

    def _build_ui(self):
        layout = QVBoxLayout(self)
        label = QLabel(f'Found {len(self.duplicates)} BibTeX entries matched by multiple DB records.\n'
                       f'Review and select which records to keep:')
        layout.addWidget(label)

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        container = QWidget()
        container_layout = QVBoxLayout(container)

        for bib_idx, match_list in self.duplicates.items():
            bib_rec = self.bib_records[bib_idx]
            group = QGroupBox(f'BibTeX: {bib_rec.title[:80]}')
            group_layout = QVBoxLayout(group)

            table = QTableWidget(len(match_list), 5)
            table.setHorizontalHeaderLabels(['Keep', 'Source', 'Match Method', 'Title', 'DOI'])
            table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

            for row, mr in enumerate(match_list):
                cb = QCheckBox()
                cb.setChecked(True)
                table.setCellWidget(row, 0, cb)
                db_rec = mr.db_record
                src = f'{db_rec.source_db}/{db_rec.source_format}' if db_rec else '?'
                table.setItem(row, 1, QTableWidgetItem(src))
                table.setItem(row, 2, QTableWidgetItem(mr.match_method))
                table.setItem(row, 3, QTableWidgetItem(db_rec.title[:60] if db_rec else ''))
                table.setItem(row, 4, QTableWidgetItem(db_rec.doi if db_rec else ''))

            table.setMaximumHeight(40 + len(match_list) * 30)
            group_layout.addWidget(table)
            container_layout.addWidget(group)
            self.selections[bib_idx] = (table, match_list)

        container_layout.addStretch()
        scroll.setWidget(container)
        layout.addWidget(scroll)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def get_kept_records(self) -> dict:
        result = {}
        for bib_idx, (table, match_list) in self.selections.items():
            kept = []
            for row in range(table.rowCount()):
                cb = table.cellWidget(row, 0)
                if cb and cb.isChecked():
                    kept.append(match_list[row])
            result[bib_idx] = kept
        return result


# ---------------------------------------------------------------------------
# Main Window
# ---------------------------------------------------------------------------

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('CIDMET - Literature Database Cross-Matching Tool')
        self.setMinimumSize(1000, 800)
        self.worker = None
        self.last_result = None
        self._build_ui()

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        # --- BibTeX input ---
        bib_group = QGroupBox('BibTeX Target File')
        bib_layout = QHBoxLayout(bib_group)
        self.bib_input = DragDropLineEdit('.bib')
        bib_btn = QPushButton('Browse...')
        bib_btn.clicked.connect(lambda: self._browse_file(self.bib_input, 'BibTeX Files (*.bib)'))
        bib_layout.addWidget(self.bib_input, 1)
        bib_layout.addWidget(bib_btn)
        main_layout.addWidget(bib_group)

        # --- Database inputs ---
        db_group = QGroupBox('Database Export Files')
        db_grid = QGridLayout(db_group)

        self.db_inputs = {}
        db_fields = [
            ('wos_xls', 'WoS XLS:', '.xls,.xlsx', 'Excel Files (*.xls *.xlsx)', 0, 0),
            ('wos_txt', 'WoS TXT:', '.txt', 'Text Files (*.txt)', 0, 2),
            ('scopus_csv', 'Scopus CSV:', '.csv', 'CSV Files (*.csv)', 1, 0),
            ('scopus_txt', 'Scopus TXT:', '.txt', 'Text Files (*.txt)', 1, 2),
            ('ei_csv', 'EI CSV:', '.csv', 'CSV Files (*.csv)', 2, 0),
            ('ei_txt', 'EI TXT:', '.txt', 'Text Files (*.txt)', 2, 2),
        ]

        for key, label_text, ext_filter, file_filter, row, col in db_fields:
            lbl = QLabel(label_text)
            inp = DragDropLineEdit(ext_filter)
            btn = QPushButton('Browse')
            clear_btn = QPushButton('X')
            clear_btn.setFixedWidth(30)
            clear_btn.clicked.connect(lambda checked=False, i=inp: i.clear())

            btn.clicked.connect(lambda checked=False, i=inp, f=file_filter: self._browse_file(i, f))

            db_grid.addWidget(lbl, row, col)
            db_grid.addWidget(inp, row, col + 1)
            h_box = QHBoxLayout()
            h_box.addWidget(btn)
            h_box.addWidget(clear_btn)
            h_widget = QWidget()
            h_widget.setLayout(h_box)
            db_grid.addWidget(h_widget, row, col + 1)

            sub_layout = QHBoxLayout()
            sub_layout.setContentsMargins(0, 0, 0, 0)
            sub_layout.addWidget(inp, 1)
            sub_layout.addWidget(btn)
            sub_layout.addWidget(clear_btn)
            container = QWidget()
            container.setLayout(sub_layout)
            db_grid.addWidget(lbl, row, col)
            db_grid.addWidget(container, row, col + 1)

            self.db_inputs[key] = inp

        main_layout.addWidget(db_group)

        # --- Output dir ---
        out_group = QGroupBox('Output Directory')
        out_layout = QHBoxLayout(out_group)
        self.output_dir_input = QLineEdit()
        self.output_dir_input.setReadOnly(True)
        self.output_dir_input.setPlaceholderText('Select output directory...')
        out_btn = QPushButton('Browse...')
        out_btn.clicked.connect(self._browse_output_dir)
        out_layout.addWidget(self.output_dir_input, 1)
        out_layout.addWidget(out_btn)
        main_layout.addWidget(out_group)

        # --- Settings ---
        settings_group = QGroupBox('Settings')
        settings_layout = QHBoxLayout(settings_group)
        settings_layout.addWidget(QLabel('Fuzzy Threshold:'))
        self.threshold_slider = QSlider(Qt.Orientation.Horizontal)
        self.threshold_slider.setRange(50, 100)
        self.threshold_slider.setValue(90)
        self.threshold_label = QLabel('90')
        self.threshold_slider.valueChanged.connect(lambda v: self.threshold_label.setText(str(v)))
        settings_layout.addWidget(self.threshold_slider, 1)
        settings_layout.addWidget(self.threshold_label)
        main_layout.addWidget(settings_group)

        # --- Run + Progress ---
        run_layout = QHBoxLayout()
        self.run_btn = QPushButton('Run Matching')
        self.run_btn.setMinimumHeight(40)
        self.run_btn.setStyleSheet('font-weight: bold; font-size: 14px;')
        self.run_btn.clicked.connect(self._run_processing)
        run_layout.addWidget(self.run_btn)

        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        self.progress_bar.setFormat('%p% - %v')
        self.progress_label = QLabel('')
        run_layout.addWidget(self.progress_bar, 1)
        run_layout.addWidget(self.progress_label)
        main_layout.addLayout(run_layout)

        # --- Tabs: Log / Results / Merge ---
        self.tabs = QTabWidget()

        # Log tab
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFontFamily('Consolas')
        self.tabs.addTab(self.log_text, 'Log')

        # Results tab
        results_widget = QWidget()
        results_layout = QVBoxLayout(results_widget)

        self.stats_table = QTableWidget(0, 4)
        self.stats_table.setHorizontalHeaderLabels(['Database', 'Total Records', 'Matched', 'Unmatched'])
        self.stats_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.stats_table.setMaximumHeight(200)
        results_layout.addWidget(QLabel('Per-Database Matching Results:'))
        results_layout.addWidget(self.stats_table)

        btn_layout = QHBoxLayout()
        self.show_unmatched_btn = QPushButton('Show Unmatched BibTeX Entries')
        self.show_unmatched_btn.clicked.connect(self._show_unmatched)
        self.show_unmatched_btn.setEnabled(False)
        btn_layout.addWidget(self.show_unmatched_btn)

        self.show_duplicates_btn = QPushButton('Show Duplicates')
        self.show_duplicates_btn.clicked.connect(self._show_duplicates)
        self.show_duplicates_btn.setEnabled(False)
        btn_layout.addWidget(self.show_duplicates_btn)

        self.export_report_btn = QPushButton('Export Report (CSV)')
        self.export_report_btn.clicked.connect(self._export_report)
        self.export_report_btn.setEnabled(False)
        btn_layout.addWidget(self.export_report_btn)

        results_layout.addLayout(btn_layout)

        self.details_table = QTableWidget(0, 6)
        self.details_table.setHorizontalHeaderLabels(['Source', 'Title', 'DOI', 'Year', 'Match Method', 'Confidence'])
        self.details_table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        results_layout.addWidget(QLabel('Match Details:'))
        results_layout.addWidget(self.details_table)

        self.tabs.addTab(results_widget, 'Results')

        # Merge tab
        merge_widget = QWidget()
        merge_layout = QVBoxLayout(merge_widget)
        merge_layout.addWidget(QLabel('Select a template file to define output format:'))

        template_layout = QHBoxLayout()
        self.template_combo = QComboBox()
        self.template_combo.setMinimumWidth(400)
        template_layout.addWidget(self.template_combo, 1)
        merge_layout.addLayout(template_layout)

        merge_btn_layout = QHBoxLayout()
        self.merge_btn = QPushButton('Export Merged File')
        self.merge_btn.setEnabled(False)
        self.merge_btn.clicked.connect(self._export_merged)
        merge_btn_layout.addWidget(self.merge_btn)
        merge_layout.addLayout(merge_btn_layout)

        self.merge_log = QTextEdit()
        self.merge_log.setReadOnly(True)
        self.merge_log.setMaximumHeight(150)
        merge_layout.addWidget(self.merge_log)
        merge_layout.addStretch()

        self.tabs.addTab(merge_widget, 'Merge Export')

        main_layout.addWidget(self.tabs, 1)

    def _browse_file(self, line_edit: QLineEdit, file_filter: str):
        path, _ = QFileDialog.getOpenFileName(self, 'Select File', '', file_filter)
        if path:
            line_edit.setText(path)

    def _browse_output_dir(self):
        d = QFileDialog.getExistingDirectory(self, 'Select Output Directory')
        if d:
            self.output_dir_input.setText(d)

    def _run_processing(self):
        bibtex_path = self.bib_input.text().strip()
        if not bibtex_path or not os.path.isfile(bibtex_path):
            QMessageBox.warning(self, 'Error', 'Please select a valid BibTeX file.')
            return

        output_dir = self.output_dir_input.text().strip()
        if not output_dir:
            QMessageBox.warning(self, 'Error', 'Please select an output directory.')
            return

        os.makedirs(output_dir, exist_ok=True)

        inputs = {}
        for key, inp in self.db_inputs.items():
            path = inp.text().strip()
            if path and os.path.isfile(path):
                inputs[key] = path

        if not inputs:
            QMessageBox.warning(self, 'Error', 'Please provide at least one database export file.')
            return

        self.log_text.clear()
        self.progress_bar.setValue(0)
        self.progress_label.setText('')
        self.run_btn.setEnabled(False)
        self.stats_table.setRowCount(0)
        self.details_table.setRowCount(0)

        params = {
            'bibtex_path': bibtex_path,
            'output_dir': output_dir,
            'fuzzy_threshold': self.threshold_slider.value(),
            'inputs': inputs,
        }

        self.worker = ProcessingWorker(params)
        self.worker.progress.connect(self._on_progress)
        self.worker.log.connect(self._on_log)
        self.worker.finished_signal.connect(self._on_finished)
        self.worker.error.connect(self._on_error)
        self.worker.start()

    def _on_progress(self, value: int, text: str):
        self.progress_bar.setValue(value)
        self.progress_label.setText(text)

    def _on_log(self, text: str):
        self.log_text.append(text)

    def _on_error(self, text: str):
        self.log_text.append(f'\n!!! {text}')
        self.run_btn.setEnabled(True)
        QMessageBox.critical(self, 'Error', text[:500])

    def _on_finished(self, result: dict):
        self.run_btn.setEnabled(True)
        self.last_result = result

        if not result:
            return

        # Populate stats table
        per_db = result.get('per_db_stats', {})
        self.stats_table.setRowCount(len(per_db))
        for row, (key, stats) in enumerate(per_db.items()):
            self.stats_table.setItem(row, 0, QTableWidgetItem(stats['label']))
            self.stats_table.setItem(row, 1, QTableWidgetItem(str(stats['total'])))
            self.stats_table.setItem(row, 2, QTableWidgetItem(str(stats['matched'])))
            self.stats_table.setItem(row, 3, QTableWidgetItem(str(stats['unmatched'])))

        # Populate details table
        all_matches = result.get('match_results', {})
        detail_rows = []
        for db_key, matches in all_matches.items():
            for mr in matches:
                if mr.matched:
                    dr = mr.db_record
                    detail_rows.append((
                        f'{dr.source_db}/{dr.source_format}' if dr else db_key,
                        dr.title[:80] if dr else '',
                        dr.doi if dr else '',
                        dr.year if dr else '',
                        mr.match_method,
                        f'{mr.confidence:.1f}%',
                    ))

        self.details_table.setRowCount(len(detail_rows))
        for row, data in enumerate(detail_rows):
            for col, val in enumerate(data):
                self.details_table.setItem(row, col, QTableWidgetItem(str(val)))

        # Enable buttons
        self.show_unmatched_btn.setEnabled(bool(result.get('unmatched_bib')))
        self.show_duplicates_btn.setEnabled(bool(result.get('duplicates')))
        self.export_report_btn.setEnabled(True)

        # Populate merge template combo
        self.template_combo.clear()
        inputs = self.worker.params['inputs'] if self.worker else {}
        db_name_map = {
            'wos_xls': 'wos', 'wos_txt': 'wos',
            'scopus_csv': 'scopus', 'scopus_txt': 'scopus',
            'ei_csv': 'ei', 'ei_txt': 'ei',
        }
        for key, path in inputs.items():
            if path:
                db = db_name_map.get(key, '')
                self.template_combo.addItem(f'{key}: {os.path.basename(path)}', (path, db))
        self.merge_btn.setEnabled(self.template_combo.count() > 0)

        self.tabs.setCurrentIndex(1)

    def _show_unmatched(self):
        if not self.last_result:
            return
        unmatched = self.last_result.get('unmatched_bib', [])
        if not unmatched:
            QMessageBox.information(self, 'Info', 'No unmatched BibTeX entries.')
            return

        dlg = QDialog(self)
        dlg.setWindowTitle(f'Unmatched BibTeX Entries ({len(unmatched)})')
        dlg.setMinimumSize(700, 400)
        layout = QVBoxLayout(dlg)

        table = QTableWidget(len(unmatched), 4)
        table.setHorizontalHeaderLabels(['Index', 'Title', 'DOI', 'Year'])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)

        for row, (idx, rec) in enumerate(unmatched):
            table.setItem(row, 0, QTableWidgetItem(str(idx)))
            table.setItem(row, 1, QTableWidgetItem(rec.title[:80]))
            table.setItem(row, 2, QTableWidgetItem(rec.doi))
            table.setItem(row, 3, QTableWidgetItem(rec.year))

        layout.addWidget(table)

        btn = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok)
        btn.accepted.connect(dlg.accept)
        layout.addWidget(btn)
        dlg.exec()

    def _show_duplicates(self):
        if not self.last_result:
            return
        duplicates = self.last_result.get('duplicates', {})
        engine = self.last_result.get('engine')
        if not duplicates or not engine:
            QMessageBox.information(self, 'Info', 'No duplicates found.')
            return

        dlg = DuplicateDialog(duplicates, engine.bib_records, self)
        dlg.exec()

    def _export_report(self):
        if not self.last_result:
            return

        path, _ = QFileDialog.getSaveFileName(
            self, 'Save Report', 'match_report.csv', 'CSV Files (*.csv);;Excel Files (*.xlsx)'
        )
        if not path:
            return

        try:
            all_matches = self.last_result.get('match_results', {})
            rows = []
            for db_key, matches in all_matches.items():
                for mr in matches:
                    dr = mr.db_record
                    rows.append({
                        'Source DB': f'{dr.source_db}/{dr.source_format}' if dr else db_key,
                        'Matched': 'Yes' if mr.matched else 'No',
                        'Match Method': mr.match_method,
                        'Confidence': f'{mr.confidence:.1f}',
                        'BibTeX Key': mr.bib_key,
                        'DB Title': dr.title if dr else '',
                        'DB DOI': dr.doi if dr else '',
                        'DB Year': dr.year if dr else '',
                        'DB Authors': dr.authors[:100] if dr else '',
                        'Notes': mr.notes,
                    })

            if path.endswith('.xlsx'):
                import openpyxl
                wb = openpyxl.Workbook()
                ws = wb.active
                if rows:
                    ws.append(list(rows[0].keys()))
                    for r in rows:
                        ws.append(list(r.values()))
                wb.save(path)
            else:
                with open(path, 'w', newline='', encoding='utf-8-sig') as f:
                    if rows:
                        writer = csv.DictWriter(f, fieldnames=rows[0].keys())
                        writer.writeheader()
                        writer.writerows(rows)

            self.log_text.append(f'\nReport exported: {path}')
        except Exception as e:
            QMessageBox.critical(self, 'Error', f'Failed to export report: {e}')

    def _export_merged(self):
        if not self.last_result:
            return

        idx = self.template_combo.currentIndex()
        if idx < 0:
            QMessageBox.warning(self, 'Error', 'No template selected.')
            return

        template_path, template_db = self.template_combo.currentData()
        ext = os.path.splitext(template_path)[1].lower()

        output_dir = self.output_dir_input.text().strip()
        if not output_dir:
            QMessageBox.warning(self, 'Error', 'No output directory set.')
            return

        out_path = os.path.join(output_dir, f'merged_output{ext}')

        try:
            matched_records = self.last_result.get('matched_records', [])
            if not matched_records:
                QMessageBox.information(self, 'Info', 'No matched records to merge.')
                return

            unified = [record_to_unified(r) for r in matched_records]
            deduped = deduplicate_records(unified)

            self.merge_log.clear()
            self.merge_log.append(f'Total matched records: {len(unified)}')
            self.merge_log.append(f'After deduplication: {len(deduped)}')
            self.merge_log.append(f'Template: {os.path.basename(template_path)} ({template_db})')

            actual_path = export_merged(deduped, template_path, out_path, template_db)

            self.merge_log.append(f'Merged file exported: {actual_path}')
            self.log_text.append(f'\nMerged file exported: {actual_path}')
            QMessageBox.information(self, 'Success', f'Merged file exported:\n{actual_path}')

        except Exception as e:
            self.merge_log.append(f'Error: {e}\n{traceback.format_exc()}')
            QMessageBox.critical(self, 'Error', f'Failed to export merged file: {e}')
