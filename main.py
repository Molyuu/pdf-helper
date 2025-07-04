import sys
from PyQt5.QtCore import Qt
from docx2pdf import convert as docx2pdf_convert
from PyPDF2 import PdfMerger
from PyQt5.QtWidgets import QProgressBar
from os import system
import os

from PyQt5.QtWidgets import (
    QApplication, QWidget, QVBoxLayout, QPushButton, QFileDialog,
    QListWidget, QLabel, QMessageBox, QHBoxLayout, QListWidgetItem
)

class PDFHelper(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("PDF Helper")
        self.setGeometry(100, 100, 600, 500)
        self.setStyleSheet("""
            QWidget {
                background-color: #f5f7fa;
            }
            QLabel {
                color: #22223b;
                font-weight: medium;
            }
            QPushButton {
                background-color: #4f8cff;
                color: white;
                border-radius: 8px;
                padding: 8px 16px;
                font-weight: medium;
            }
            QPushButton:hover {
                background-color: #2563eb;
            }
            QListWidget {
                background-color: #e9ecef;
                border-radius: 6px;
                padding: 4px;
            }
            QProgressBar {
                border: 1px solid #bbb;
                border-radius: 6px;
                text-align: center;
                background: #e9ecef;
            }
            QProgressBar::chunk {
                background-color: #4f8cff;
                border-radius: 6px;
            }
        """)

        # Set global font
        font = self.font()
        font.setFamily("Consolas")
        font.setPointSize(16)
        self.setFont(font)

        layout = QVBoxLayout()
        layout.setSpacing(18)
        layout.setContentsMargins(24, 24, 24, 24)

        # DOCX to PDF
        self.docx_label = QLabel("Convert DOCX files to PDF:")
        self.docx_label.setFont(font)
        self.docx_btn = QPushButton("Select DOCX files")
        self.docx_btn.setFont(font)
        self.docx_btn.clicked.connect(self.convert_docx_to_pdf)
        layout.addWidget(self.docx_label)
        layout.addWidget(self.docx_btn)

        layout.addSpacing(18)

        # PDF Combiner
        self.pdf_label = QLabel("Combine PDF files (drag to reorder, press Del to delete)")
        self.pdf_label.setFont(font)
        self.pdf_list = QListWidget()
        self.pdf_list.setFont(font)
        self.pdf_list.setSelectionMode(QListWidget.ExtendedSelection)
        self.pdf_list.setDragDropMode(QListWidget.InternalMove)
        self.pdf_list.setMinimumHeight(140)
        self.add_pdf_btn = QPushButton("Add PDFs")
        self.add_pdf_btn.setFont(font)
        self.add_pdf_btn.clicked.connect(self.add_pdfs)
        self.combine_pdf_btn = QPushButton("Combine PDFs")
        self.combine_pdf_btn.setFont(font)
        self.combine_pdf_btn.clicked.connect(self.combine_pdfs)

        pdf_btn_layout = QHBoxLayout()
        pdf_btn_layout.addWidget(self.add_pdf_btn)
        pdf_btn_layout.addWidget(self.combine_pdf_btn)
        pdf_btn_layout.setSpacing(12)

        layout.addWidget(self.pdf_label)
        layout.addWidget(self.pdf_list)
        layout.addLayout(pdf_btn_layout)

        self.setLayout(layout)

    def convert_docx_to_pdf(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select DOCX files", "", "Word Documents (*.docx)"
        )
        if not files:
            return
        out_dir = QFileDialog.getExistingDirectory(self, "Select Output Directory")
        if not out_dir:
            return

        progress = QProgressBar(self)
        font = self.font()
        progress.setFont(font)
        progress.setMinimum(0)
        progress.setMaximum(len(files))
        progress.setValue(0)
        progress.setFormat("Converting %v of %m")
        progress.setTextVisible(True)
        self.layout().addWidget(progress)

        try:
            for idx, f in enumerate(files, 1):
                docx2pdf_convert(f, out_dir)
                progress.setValue(idx)
                QApplication.processEvents()
            QMessageBox.information(self, "Success", "DOCX files converted to PDF!")
            system(f'start "" "{out_dir}"')

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to convert: {e}")
        finally:
            self.layout().removeWidget(progress)
            progress.deleteLater()

    def add_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF files", "", "PDF Files (*.pdf)"
        )
        for f in files:
            if f not in [self.pdf_list.item(i).text() for i in range(self.pdf_list.count())]:
                self.pdf_list.addItem(f)

    def combine_pdfs(self):
        if self.pdf_list.count() < 2:
            QMessageBox.warning(self, "Warning", "Add at least two PDFs to combine.")
            return
        out_file, _ = QFileDialog.getSaveFileName(
            self, "Save Combined PDF", "", "PDF Files (*.pdf)"
        )
        if not out_file:
            return
        try:
            merger = PdfMerger()
            # Store the full paths in a list
            file_paths = []
            for i in range(self.pdf_list.count()):
                file_paths.append(self.pdf_list.item(i).data(Qt.UserRole))
                merger.append(self.pdf_list.item(i).data(Qt.UserRole))
            merger.write(out_file)
            merger.close()
            QMessageBox.information(self, "Success", "PDFs combined successfully!")
            folder = os.path.dirname(os.path.abspath(out_file))
            system(f'start "" "{folder}"')

        except Exception as e:
            QMessageBox.critical(self, "Error", f"Failed to combine PDFs: {e}")

    def add_pdfs(self):
        files, _ = QFileDialog.getOpenFileNames(
            self, "Select PDF files", "", "PDF Files (*.pdf)"
        )
        existing_paths = [self.pdf_list.item(i).data(Qt.UserRole) for i in range(self.pdf_list.count())]
        for f in files:
            if f not in existing_paths:
                item = QListWidgetItem(os.path.basename(f))
                item.setData(Qt.UserRole, f)
                self.pdf_list.addItem(item)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Delete and self.pdf_list.hasFocus():
            for item in self.pdf_list.selectedItems():
                self.pdf_list.takeItem(self.pdf_list.row(item))
        else:
            super().keyPressEvent(event)

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = PDFHelper()
    window.show()
    sys.exit(app.exec_())