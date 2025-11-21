"""
Simple GUI tool to reorder 4-up invoices on A4 pages into cut-friendly order.
Assumptions:
- Each original page is evenly split into 4 parts from top to bottom with no margins.
- Total invoice count is divisible by 4.

Output order per page i (1-based): i, i+N, i+2N, i+3N where N = original page count.
After cutting the stack into four piles, each pile is already sequential.
"""

import os
import sys
from typing import List, Tuple

from pypdf import PdfReader, PdfWriter
from PySide6 import QtCore, QtWidgets


def slice_page_into_quarters(page) -> List:
    """Return 4 cropped copies from top to bottom."""
    w = float(page.mediabox.width)
    h = float(page.mediabox.height)
    step = h / 4.0
    parts = []
    for i in range(4):  # 0 = top
        part = page.copy()
        top = h - step * i
        bottom = top - step
        part.mediabox.lower_left = (0, bottom)
        part.mediabox.upper_right = (w, top)
        parts.append(part)
    return parts


def reorder_stride(parts: List) -> Tuple[List, int]:
    total_parts = len(parts)
    if total_parts % 4 != 0:
        raise ValueError("Total invoices must be divisible by 4.")
    num_pages = total_parts // 4
    reordered = []
    # Output page i: i, i+N, i+2N, i+3N (1-based concept, 0-based indices)
    for i in range(num_pages):
        for j in range(4):
            reordered.append(parts[i + j * num_pages])
    return reordered, num_pages


def rebuild_pages(parts: List, num_pages: int, template_page) -> PdfWriter:
    w = float(template_page.mediabox.width)
    h = float(template_page.mediabox.height)
    step = h / 4.0
    writer = PdfWriter()
    for i in range(num_pages):
        page = writer.add_blank_page(width=w, height=h)
        for idx, part in enumerate(parts[i * 4 : (i + 1) * 4]):
            y = h - step * (idx + 1)
            page.merge_translated_page(part, 0, y, expand=False)
    return writer


def process_pdf(src_path: str, dst_path: str) -> int:
    reader = PdfReader(src_path)
    if not reader.pages:
        raise ValueError("Source PDF is empty.")
    parts = []
    for p in reader.pages:
        parts.extend(slice_page_into_quarters(p))
    reordered, num_pages = reorder_stride(parts)
    writer = rebuild_pages(reordered, num_pages, reader.pages[0])
    with open(dst_path, "wb") as f:
        writer.write(f)
    return num_pages


class PdfWorker(QtCore.QThread):
    progress = QtCore.Signal(str)
    finished = QtCore.Signal(int)
    failed = QtCore.Signal(str)

    def __init__(self, src: str, dst: str, parent=None):
        super().__init__(parent)
        self.src = src
        self.dst = dst

    def run(self):
        try:
            self.progress.emit("Processing...")
            pages = process_pdf(self.src, self.dst)
            self.finished.emit(pages)
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(str(exc))


class MainWindow(QtWidgets.QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Invoice PDF Reorder")
        self.resize(500, 200)
        self.worker = None
        self._init_ui()

    def _init_ui(self):
        central = QtWidgets.QWidget()
        layout = QtWidgets.QVBoxLayout()

        self.src_edit = QtWidgets.QLineEdit()
        self.dst_edit = QtWidgets.QLineEdit()
        self.status_label = QtWidgets.QLabel("Ready.")
        self.status_label.setWordWrap(True)
        browse_src = QtWidgets.QPushButton("Browse...")
        browse_dst = QtWidgets.QPushButton("Browse...")
        run_button = QtWidgets.QPushButton("Start Reorder")

        browse_src.clicked.connect(self._choose_src)
        browse_dst.clicked.connect(self._choose_dst)
        run_button.clicked.connect(self._run)

        form_layout = QtWidgets.QGridLayout()
        form_layout.addWidget(QtWidgets.QLabel("Source PDF:"), 0, 0)
        form_layout.addWidget(self.src_edit, 0, 1)
        form_layout.addWidget(browse_src, 0, 2)
        form_layout.addWidget(QtWidgets.QLabel("Output PDF:"), 1, 0)
        form_layout.addWidget(self.dst_edit, 1, 1)
        form_layout.addWidget(browse_dst, 1, 2)

        layout.addLayout(form_layout)
        layout.addWidget(run_button)
        layout.addWidget(self.status_label)
        central.setLayout(layout)
        self.setCentralWidget(central)

        self.run_button = run_button

    def _choose_src(self):
        path, _ = QtWidgets.QFileDialog.getOpenFileName(
            self, "Select source PDF", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if path:
            self.src_edit.setText(path)
            # Auto-fill output path alongside source
            base, ext = os.path.splitext(path)
            self.dst_edit.setText(f"{base}_reordered{ext}")

    def _choose_dst(self):
        path, _ = QtWidgets.QFileDialog.getSaveFileName(
            self, "Save output PDF", "", "PDF Files (*.pdf);;All Files (*)"
        )
        if path:
            self.dst_edit.setText(path)

    def _run(self):
        src = self.src_edit.text().strip()
        dst = self.dst_edit.text().strip()
        try:
            self._validate_paths(src, dst)
        except ValueError as exc:
            QtWidgets.QMessageBox.warning(self, "Invalid input", str(exc))
            return

        self._set_running(True)
        self.status_label.setText("Starting...")

        self.worker = PdfWorker(src, dst)
        self.worker.progress.connect(self._on_progress)
        self.worker.finished.connect(self._on_finished)
        self.worker.failed.connect(self._on_failed)
        self.worker.finished.connect(lambda _: self._set_running(False))
        self.worker.failed.connect(lambda _: self._set_running(False))
        self.worker.start()

    def closeEvent(self, event):  # noqa: N802
        if self.worker and self.worker.isRunning():
            self.worker.quit()
            self.worker.wait(2000)
        super().closeEvent(event)

    def _validate_paths(self, src: str, dst: str):
        if not src or not dst:
            raise ValueError("Please select both source and output paths.")
        if not os.path.isfile(src):
            raise ValueError("Source file does not exist.")
        if os.path.abspath(src) == os.path.abspath(dst):
            raise ValueError("Source and output paths must be different.")
        dst_dir = os.path.dirname(dst) or "."
        if not os.path.exists(dst_dir):
            os.makedirs(dst_dir, exist_ok=True)
        if not os.access(dst_dir, os.W_OK):
            raise ValueError("Output directory is not writable.")

    @QtCore.Slot(str)
    def _on_progress(self, msg: str):
        self.status_label.setText(msg)

    @QtCore.Slot(int)
    def _on_finished(self, pages: int):
        self.status_label.setText(f"Done. Pages: {pages}")
        QtWidgets.QMessageBox.information(self, "Completed", f"Finished. Pages: {pages}")

    @QtCore.Slot(str)
    def _on_failed(self, msg: str):
        self.status_label.setText(f"Failed: {msg}")
        QtWidgets.QMessageBox.critical(self, "Error", msg)

    def _set_running(self, running: bool):
        self.run_button.setEnabled(not running)
        self.src_edit.setEnabled(not running)
        self.dst_edit.setEnabled(not running)


def main():
    app = QtWidgets.QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
