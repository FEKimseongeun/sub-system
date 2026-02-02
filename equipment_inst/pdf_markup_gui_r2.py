# -*- coding: utf-8 -*-
"""
PDF 마크업 GUI - Excel 좌표 기반 PDF 박스 주석 도구 (폴더 기반)

Excel 파일(equipment_subsystem_instrument_tags.xlsx)의 bbox 좌표를 읽어
PDF 폴더 내 개별 PDF 파일들에 타입별 색상 마크업(annotation)을 추가합니다.

- Excel의 pdf_name 열을 기준으로 해당 PDF 파일을 찾아 마크업
- 마크업된 PDF는 출력 폴더에 {원본명}_marked.pdf로 저장

타입별 색상:
- special / special_item : 녹색 (Green)
- subsystem_name         : 빨간색 (Red)
- instrument             : 파란색 (Blue)
- valve                  : 주황색 (Orange)
- equipment              : 보라색 (Purple)
"""
from __future__ import annotations
from pathlib import Path
from typing import Dict, List, Optional, Tuple
import sys

import pandas as pd
import pymupdf as fitz

from PyQt6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QFileDialog, QTextEdit, QProgressBar,
    QGroupBox, QCheckBox, QDoubleSpinBox, QMessageBox, QFrame,
    QColorDialog
)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QFont, QColor


# ============== 색상 유틸리티 ==============
def rgb255_to_float(rgb: Tuple[int, int, int]) -> Tuple[float, float, float]:
    """RGB(0~255) → float(0~1) 변환"""
    return tuple(c / 255.0 for c in rgb)


def qcolor_to_rgb255(qcolor: QColor) -> Tuple[int, int, int]:
    """QColor → RGB(0~255) 변환"""
    return (qcolor.red(), qcolor.green(), qcolor.blue())


def rgb255_to_hex(rgb: Tuple[int, int, int]) -> str:
    """RGB(0~255) → HEX 문자열 변환"""
    return f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"


# ============== 기본 타입별 스타일 설정 ==============
# type → (stroke_rgb255, fill_rgb255, default_opacity)
DEFAULT_TYPE_STYLES: Dict[str, Tuple[Tuple[int, int, int], Tuple[int, int, int], float]] = {
    "special":       ((0, 150, 0),   (0, 220, 0),    0.30),
    "special_item":  ((0, 150, 0),   (0, 220, 0),    0.30),
    "subsystem_name":((220, 0, 0),   (255, 100, 100),0.30),
    "instrument":    ((0, 80, 180),  (0, 120, 255),  0.30),
    "valve":         ((200, 100, 0), (255, 165, 0),  0.30),
    "equipment":     ((128, 0, 128), (180, 100, 255),0.30),
}


# ============== bbox 컬럼 탐색 ==============
def pick_bbox_columns(df: pd.DataFrame) -> Tuple[str, str, str, str]:
    """
    DataFrame에서 bbox 컬럼을 찾아 반환
    지원 형식:
      - x0, y0, x1, y1
      - x1, y1, x2, y2
    """
    cols = {c.lower().strip() for c in df.columns}

    if {"x0", "y0", "x1", "y1"}.issubset(cols):
        return "x0", "y0", "x1", "y1"

    if {"x1", "y1", "x2", "y2"}.issubset(cols):
        return "x1", "y1", "x2", "y2"

    raise ValueError(
        "bbox 컬럼을 찾지 못했습니다. "
        "x0,y0,x1,y1 또는 x1,y1,x2,y2 중 하나는 있어야 합니다."
    )


def normalize_type(t) -> str:
    """타입 문자열 정규화"""
    if not isinstance(t, str):
        return ""
    return t.strip().lower()


# ============== 마크업 함수 ==============
def add_rect_markup(
    page: fitz.Page,
    rect: fitz.Rect,
    stroke_rgb: Tuple[int, int, int],
    fill_rgb: Tuple[int, int, int],
    opacity: float
):
    """
    페이지에 사각 annotation 추가 (채움+투명도)
    """
    stroke = rgb255_to_float(stroke_rgb)
    fill = rgb255_to_float(fill_rgb)

    annot = page.add_rect_annot(rect)
    annot.set_colors(stroke=stroke, fill=fill)
    annot.set_opacity(opacity)

    try:
        annot.set_border(width=0.5)
    except Exception:
        try:
            border = annot.border
            border["width"] = 0.5
            annot.set_border(border)
        except Exception:
            pass

    annot.update()


# ============== 처리 워커 스레드 ==============
class MarkupWorker(QThread):
    """백그라운드 마크업 처리 스레드 (폴더 기반)"""
    progress = pyqtSignal(int, int, str)  # current, total, message
    log = pyqtSignal(str)
    finished_signal = pyqtSignal(int, int)  # processed_count, total_count
    error = pyqtSignal(str)

    def __init__(
        self,
        pdf_folder: Path,
        excel_path: Path,
        output_folder: Path,
        type_styles: Dict[str, Tuple[Tuple[int, int, int], Tuple[int, int, int], float]],
        enabled_types: Dict[str, bool],
        min_area: float = 4.0
    ):
        super().__init__()
        self.pdf_folder = pdf_folder
        self.excel_path = excel_path
        self.output_folder = output_folder
        self.type_styles = type_styles
        self.enabled_types = enabled_types
        self.min_area = min_area
        self._is_cancelled = False

    def cancel(self):
        self._is_cancelled = True

    def run(self):
        try:
            self.log.emit(f"Excel 파일 로드 중: {self.excel_path.name}")

            # Excel 데이터 로드
            df = pd.read_excel(self.excel_path)
            if df.empty:
                self.error.emit("Excel 데이터가 비어있습니다.")
                return

            df.columns = [c.strip() for c in df.columns]

            # bbox 컬럼 찾기
            try:
                x_a, y_a, x_b, y_b = pick_bbox_columns(df)
            except ValueError as e:
                self.error.emit(str(e))
                return

            # 필수 컬럼 확인
            lower_cols = [c.lower() for c in df.columns]
            if "page" not in lower_cols:
                self.error.emit("필수 컬럼 없음: page")
                return
            if "type" not in lower_cols:
                self.error.emit("필수 컬럼 없음: type")
                return
            if "pdf_name" not in lower_cols:
                self.error.emit("필수 컬럼 없음: pdf_name")
                return

            df["page"] = df["page"].astype(int)
            df["type"] = df["type"].apply(normalize_type)
            
            # pdf_name이 비어있거나 NaN인 행 처리
            df["pdf_name"] = df["pdf_name"].fillna("").astype(str)

            # 활성화된 타입 필터링
            active_types = [t for t, enabled in self.enabled_types.items() if enabled]
            self.log.emit(f"활성화된 타입: {', '.join(active_types)}")

            # 출력 폴더 생성
            self.output_folder.mkdir(parents=True, exist_ok=True)

            # pdf_name으로 그룹핑 (빈 문자열 제외)
            valid_df = df[df["pdf_name"] != ""].copy()
            pdf_groups = valid_df.groupby("pdf_name")
            total_pdfs = len(pdf_groups)
            processed_count = 0
            skipped_pdfs = []

            # pdf_name이 비어있는 행 확인
            empty_pdf_name_count = len(df[df["pdf_name"] == ""])
            if empty_pdf_name_count > 0:
                self.log.emit(f"⚠️ pdf_name이 비어있는 행: {empty_pdf_name_count}개 (마크업 건너뜀)")

            self.log.emit(f"총 {total_pdfs}개 PDF 파일 처리 예정")
            self.log.emit("")

            for pdf_idx, (pdf_name, pdf_df) in enumerate(pdf_groups):
                if self._is_cancelled:
                    self.log.emit("작업이 취소되었습니다.")
                    return

                # PDF 파일 경로 찾기
                pdf_path = self.pdf_folder / pdf_name
                if not pdf_path.exists():
                    # 확장자 없이 시도
                    pdf_path = self.pdf_folder / f"{pdf_name}.pdf"
                    if not pdf_path.exists():
                        self.log.emit(f"  [스킵] PDF 파일 없음: {pdf_name}")
                        skipped_pdfs.append(pdf_name)
                        continue

                self.progress.emit(pdf_idx + 1, total_pdfs, f"처리 중: {pdf_name}")
                self.log.emit(f"[{pdf_idx + 1}/{total_pdfs}] {pdf_name}")

                try:
                    doc = fitz.open(pdf_path.as_posix())
                    markup_count = 0
                    skipped_count = 0

                    # 페이지별 처리
                    for page_num, page_df in pdf_df.groupby("page"):
                        if self._is_cancelled:
                            doc.close()
                            return

                        if page_num < 1 or page_num > doc.page_count:
                            continue

                        page = doc[page_num - 1]

                        for _, row in page_df.iterrows():
                            ttype = row["type"]

                            # special_item을 special로 통합 처리
                            lookup_type = "special" if ttype == "special_item" else ttype

                            # 활성화되지 않은 타입 스킵
                            if lookup_type not in active_types:
                                continue

                            # 스타일 가져오기
                            if lookup_type not in self.type_styles:
                                continue

                            stroke_rgb, fill_rgb, opacity = self.type_styles[lookup_type]

                            try:
                                x0 = float(row[x_a])
                                y0 = float(row[y_a])
                                x1 = float(row[x_b])
                                y1 = float(row[y_b])
                            except (ValueError, TypeError):
                                skipped_count += 1
                                continue

                            rect = fitz.Rect(min(x0, x1), min(y0, y1), max(x0, x1), max(y0, y1))

                            # 너무 작은 박스 스킵
                            area = rect.width * rect.height
                            if area < self.min_area:
                                skipped_count += 1
                                continue

                            add_rect_markup(page, rect, stroke_rgb, fill_rgb, opacity)
                            markup_count += 1

                    # 출력 파일명 생성
                    output_name = f"{pdf_path.stem}_marked.pdf"
                    output_path = self.output_folder / output_name

                    doc.save(output_path.as_posix())
                    doc.close()

                    self.log.emit(f"    -> {markup_count}개 마크업, 저장: {output_name}")
                    if skipped_count > 0:
                        self.log.emit(f"       ({skipped_count}개 스킵됨)")
                    processed_count += 1

                except Exception as e:
                    self.log.emit(f"    [오류] {str(e)}")
                    skipped_pdfs.append(pdf_name)
                    continue

            # 완료 로그
            self.log.emit("")
            self.log.emit("=" * 50)
            self.log.emit(f"마크업 완료!")
            self.log.emit(f"  - 처리된 PDF: {processed_count}개")
            self.log.emit(f"  - 스킵된 PDF: {len(skipped_pdfs)}개")
            self.log.emit(f"  - 출력 폴더: {self.output_folder}")

            if skipped_pdfs:
                self.log.emit("")
                self.log.emit("스킵된 파일 목록:")
                for name in skipped_pdfs[:10]:
                    self.log.emit(f"  - {name}")
                if len(skipped_pdfs) > 10:
                    self.log.emit(f"  ... 외 {len(skipped_pdfs) - 10}개")

            self.finished_signal.emit(processed_count, total_pdfs)

        except Exception as e:
            self.error.emit(f"처리 중 오류: {str(e)}")


# ============== 색상 버튼 위젯 ==============
class ColorButton(QPushButton):
    """색상 선택 버튼"""
    color_changed = pyqtSignal(tuple)

    def __init__(self, initial_color: Tuple[int, int, int], parent=None):
        super().__init__(parent)
        self._color = initial_color
        self.setFixedSize(30, 25)
        self.update_style()
        self.clicked.connect(self.pick_color)

    @property
    def color(self) -> Tuple[int, int, int]:
        return self._color

    @color.setter
    def color(self, rgb: Tuple[int, int, int]):
        self._color = rgb
        self.update_style()

    def update_style(self):
        hex_color = rgb255_to_hex(self._color)
        self.setStyleSheet(f"""
            QPushButton {{
                background-color: {hex_color};
                border: 1px solid #888;
                border-radius: 3px;
            }}
            QPushButton:hover {{
                border: 2px solid #333;
            }}
        """)

    def pick_color(self):
        qcolor = QColor(*self._color)
        new_color = QColorDialog.getColor(qcolor, self, "색상 선택")
        if new_color.isValid():
            self._color = qcolor_to_rgb255(new_color)
            self.update_style()
            self.color_changed.emit(self._color)


# ============== 타입별 설정 위젯 ==============
class TypeSettingWidget(QFrame):
    """개별 타입의 마크업 설정 위젯"""

    def __init__(
        self,
        type_name: str,
        display_name: str,
        stroke_color: Tuple[int, int, int],
        fill_color: Tuple[int, int, int],
        opacity: float,
        parent=None
    ):
        super().__init__(parent)
        self.type_name = type_name
        self.setFrameStyle(QFrame.Shape.StyledPanel)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(5, 3, 5, 3)

        # 체크박스
        self.checkbox = QCheckBox(display_name)
        self.checkbox.setChecked(True)
        self.checkbox.setMinimumWidth(150)
        layout.addWidget(self.checkbox)

        # Fill 색상
        layout.addWidget(QLabel("Fill:"))
        self.fill_btn = ColorButton(fill_color)
        layout.addWidget(self.fill_btn)

        # Stroke 색상
        layout.addWidget(QLabel("Stroke:"))
        self.stroke_btn = ColorButton(stroke_color)
        layout.addWidget(self.stroke_btn)

        # 투명도
        layout.addWidget(QLabel("투명도:"))
        self.opacity_spin = QDoubleSpinBox()
        self.opacity_spin.setRange(0.05, 1.0)
        self.opacity_spin.setSingleStep(0.05)
        self.opacity_spin.setValue(opacity)
        self.opacity_spin.setFixedWidth(65)
        layout.addWidget(self.opacity_spin)

        layout.addStretch()

    def is_enabled(self) -> bool:
        return self.checkbox.isChecked()

    def get_style(self) -> Tuple[Tuple[int, int, int], Tuple[int, int, int], float]:
        return (self.stroke_btn.color, self.fill_btn.color, self.opacity_spin.value())


# ============== 메인 GUI ==============
class PDFMarkupGUI(QMainWindow):
    def __init__(self):
        super().__init__()
        self.pdf_folder: Optional[Path] = None
        self.excel_path: Optional[Path] = None
        self.output_folder: Optional[Path] = None
        self.worker: Optional[MarkupWorker] = None
        self.type_widgets: Dict[str, TypeSettingWidget] = {}

        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("PDF 마크업 도구 (Equipment, Subsystem & Instrument)")
        self.setMinimumSize(750, 700)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)

        # ===== 파일 설정 =====
        file_group = QGroupBox("파일 설정")
        file_layout = QVBoxLayout(file_group)

        # PDF 폴더
        pdf_row = QHBoxLayout()
        pdf_row.addWidget(QLabel("PDF 폴더:"))
        self.pdf_label = QLabel("선택된 폴더 없음")
        self.pdf_label.setStyleSheet("color: gray;")
        pdf_row.addWidget(self.pdf_label, 1)
        btn_pdf = QPushButton("폴더 선택")
        btn_pdf.clicked.connect(self.select_pdf_folder)
        pdf_row.addWidget(btn_pdf)
        file_layout.addLayout(pdf_row)

        # Excel 파일
        excel_row = QHBoxLayout()
        excel_row.addWidget(QLabel("Excel 파일:"))
        self.excel_label = QLabel("선택된 파일 없음")
        self.excel_label.setStyleSheet("color: gray;")
        excel_row.addWidget(self.excel_label, 1)
        btn_excel = QPushButton("파일 선택")
        btn_excel.clicked.connect(self.select_excel)
        excel_row.addWidget(btn_excel)
        file_layout.addLayout(excel_row)

        # 출력 폴더
        out_row = QHBoxLayout()
        out_row.addWidget(QLabel("출력 폴더:"))
        self.out_label = QLabel("선택된 폴더 없음")
        self.out_label.setStyleSheet("color: gray;")
        out_row.addWidget(self.out_label, 1)
        btn_out = QPushButton("폴더 선택")
        btn_out.clicked.connect(self.select_output_folder)
        out_row.addWidget(btn_out)
        file_layout.addLayout(out_row)

        main_layout.addWidget(file_group)

        # ===== 마크업 설정 =====
        markup_group = QGroupBox("타입별 마크업 설정")
        markup_layout = QVBoxLayout(markup_group)

        # 타입별 설정 위젯 생성
        type_configs = [
            ("special", "Special / Special_item", DEFAULT_TYPE_STYLES["special"]),
            ("subsystem_name", "Subsystem Name", DEFAULT_TYPE_STYLES["subsystem_name"]),
            ("instrument", "Instrument", DEFAULT_TYPE_STYLES["instrument"]),
            ("valve", "Valve", DEFAULT_TYPE_STYLES["valve"]),
            ("equipment", "Equipment", DEFAULT_TYPE_STYLES["equipment"]),
        ]

        for type_name, display_name, (stroke, fill, opacity) in type_configs:
            widget = TypeSettingWidget(type_name, display_name, stroke, fill, opacity)
            self.type_widgets[type_name] = widget
            markup_layout.addWidget(widget)

        # 전체 선택/해제 버튼
        btn_row = QHBoxLayout()
        btn_select_all = QPushButton("전체 선택")
        btn_select_all.clicked.connect(lambda: self.set_all_types(True))
        btn_row.addWidget(btn_select_all)
        btn_deselect_all = QPushButton("전체 해제")
        btn_deselect_all.clicked.connect(lambda: self.set_all_types(False))
        btn_row.addWidget(btn_deselect_all)
        btn_row.addStretch()
        markup_layout.addLayout(btn_row)

        main_layout.addWidget(markup_group)

        # ===== 실행 버튼 =====
        exec_row = QHBoxLayout()
        self.btn_run = QPushButton("마크업 실행")
        self.btn_run.setMinimumHeight(40)
        self.btn_run.setStyleSheet("""
            QPushButton {
                background-color: #2196F3;
                color: white;
                font-size: 14px;
                font-weight: bold;
                border-radius: 5px;
            }
            QPushButton:hover {
                background-color: #1976D2;
            }
            QPushButton:disabled {
                background-color: #cccccc;
            }
        """)
        self.btn_run.clicked.connect(self.run_markup)
        exec_row.addWidget(self.btn_run)

        self.btn_cancel = QPushButton("취소")
        self.btn_cancel.setMinimumHeight(40)
        self.btn_cancel.setEnabled(False)
        self.btn_cancel.clicked.connect(self.cancel_markup)
        exec_row.addWidget(self.btn_cancel)

        main_layout.addLayout(exec_row)

        # ===== 진행 상황 =====
        self.progress_bar = QProgressBar()
        self.progress_bar.setTextVisible(True)
        main_layout.addWidget(self.progress_bar)

        self.status_label = QLabel("대기 중...")
        main_layout.addWidget(self.status_label)

        # ===== 로그 =====
        log_group = QGroupBox("로그")
        log_layout = QVBoxLayout(log_group)
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        self.log_text.setFont(QFont("Consolas", 9))
        log_layout.addWidget(self.log_text)
        main_layout.addWidget(log_group, 1)

    def set_all_types(self, enabled: bool):
        """모든 타입 체크박스 설정"""
        for widget in self.type_widgets.values():
            widget.checkbox.setChecked(enabled)

    def select_pdf_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "PDF 폴더 선택")
        if folder:
            self.pdf_folder = Path(folder)
            # PDF 파일 개수 확인
            pdf_count = len(list(self.pdf_folder.glob("*.pdf")))
            self.pdf_label.setText(f"{self.pdf_folder.name} ({pdf_count}개 PDF)")
            self.pdf_label.setStyleSheet("color: black;")
            self.log_text.append(f"PDF 폴더 선택: {self.pdf_folder}")
            self.log_text.append(f"  -> {pdf_count}개 PDF 파일 발견")

    def select_excel(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Excel 파일 선택", "",
            "Excel Files (*.xlsx *.xls);;All Files (*)"
        )
        if file_path:
            self.excel_path = Path(file_path)
            self.excel_label.setText(self.excel_path.name)
            self.excel_label.setStyleSheet("color: black;")
            self.log_text.append(f"Excel 선택: {self.excel_path}")

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "출력 폴더 선택")
        if folder:
            self.output_folder = Path(folder)
            self.out_label.setText(str(self.output_folder))
            self.out_label.setStyleSheet("color: black;")
            self.log_text.append(f"출력 폴더 선택: {self.output_folder}")

    def run_markup(self):
        # 유효성 검사
        if not self.pdf_folder or not self.pdf_folder.exists():
            QMessageBox.warning(self, "경고", "PDF 폴더를 선택해주세요.")
            return

        if not self.excel_path or not self.excel_path.exists():
            QMessageBox.warning(self, "경고", "Excel 파일을 선택해주세요.")
            return

        if not self.output_folder:
            QMessageBox.warning(self, "경고", "출력 폴더를 선택해주세요.")
            return

        # 활성화된 타입 확인
        enabled_types = {name: w.is_enabled() for name, w in self.type_widgets.items()}
        if not any(enabled_types.values()):
            QMessageBox.warning(self, "경고", "최소 하나의 마크업 타입을 선택해주세요.")
            return

        # 스타일 수집
        type_styles = {name: w.get_style() for name, w in self.type_widgets.items()}

        # UI 상태 변경
        self.btn_run.setEnabled(False)
        self.btn_cancel.setEnabled(True)
        self.progress_bar.setValue(0)
        self.log_text.clear()
        self.log_text.append("마크업 작업 시작...")
        self.log_text.append("")

        # 워커 시작
        self.worker = MarkupWorker(
            pdf_folder=self.pdf_folder,
            excel_path=self.excel_path,
            output_folder=self.output_folder,
            type_styles=type_styles,
            enabled_types=enabled_types
        )

        self.worker.progress.connect(self.on_progress)
        self.worker.log.connect(self.on_log)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.error.connect(self.on_error)

        self.worker.start()

    def cancel_markup(self):
        if self.worker:
            self.worker.cancel()
            self.btn_cancel.setEnabled(False)
            self.status_label.setText("취소 중...")

    def on_progress(self, current: int, total: int, message: str):
        if total > 0:
            percent = int(current / total * 100)
            self.progress_bar.setValue(percent)
        self.status_label.setText(message)

    def on_log(self, message: str):
        self.log_text.append(message)
        scrollbar = self.log_text.verticalScrollBar()
        scrollbar.setValue(scrollbar.maximum())

    def on_finished(self, processed: int, total: int):
        self.btn_run.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.progress_bar.setValue(100)
        self.status_label.setText("완료!")

        QMessageBox.information(
            self, "완료",
            f"마크업이 완료되었습니다!\n\n"
            f"처리된 PDF: {processed}/{total}개\n"
            f"출력 폴더: {self.output_folder}"
        )

    def on_error(self, error_msg: str):
        self.btn_run.setEnabled(True)
        self.btn_cancel.setEnabled(False)
        self.status_label.setText("오류 발생")
        self.log_text.append(f"오류: {error_msg}")
        QMessageBox.critical(self, "오류", f"처리 중 오류 발생:\n{error_msg}")


# ============== 메인 실행 ==============
def main():
    app = QApplication(sys.argv)
    app.setStyle("Fusion")

    window = PDFMarkupGUI()
    window.show()

    sys.exit(app.exec())


if __name__ == "__main__":
    main()