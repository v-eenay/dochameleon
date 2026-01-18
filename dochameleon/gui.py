"""
GUI interface for Dochameleon using QtPy.
"""

import sys
import os
from pathlib import Path
from qtpy.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QFileDialog, QProgressBar,
    QFrame, QMessageBox
)
from qtpy.QtCore import Qt, QThread, Signal
from qtpy.QtGui import QFont, QIcon

from .packages import check_and_install_packages, check_latex_installed
from .pipeline import (
    convert_single_tex_to_pdf,
    convert_single_tex_to_docx,
    convert_single_pdf_to_docx,
    convert_single_docx_to_pdf,
)

# Default output directory
DEFAULT_OUTPUT_DIR = Path(__file__).parent.parent / "output"


class ConversionWorker(QThread):
    """Worker thread for file conversion."""
    finished = Signal(bool, str)
    
    def __init__(self, mode: str, input_file: Path, output_dir: Path):
        super().__init__()
        self.mode = mode
        self.input_file = input_file
        self.output_dir = output_dir
    
    def run(self):
        try:
            if self.mode == 'tex2pdf':
                success = convert_single_tex_to_pdf(self.input_file, self.output_dir)
            elif self.mode == 'tex2docx':
                success = convert_single_tex_to_docx(self.input_file, self.output_dir)
            elif self.mode == 'pdf2docx':
                success = convert_single_pdf_to_docx(self.input_file, self.output_dir)
            elif self.mode == 'docx2pdf':
                success = convert_single_docx_to_pdf(self.input_file, self.output_dir)
            else:
                self.finished.emit(False, "Unknown conversion mode")
                return
            
            if success:
                self.finished.emit(True, "Conversion completed successfully!")
            else:
                self.finished.emit(False, "Conversion failed. Check console for details.")
        except Exception as e:
            self.finished.emit(False, str(e))


class DropZone(QFrame):
    """Drag and drop zone for files."""
    file_dropped = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setMinimumHeight(120)
        self.file_path = None
        self.setup_ui()
    
    def setup_ui(self):
        self.setStyleSheet("""
            DropZone {
                border: 2px dashed #555;
                border-radius: 12px;
                background-color: #2a2a2a;
            }
            DropZone:hover {
                border-color: #7c3aed;
                background-color: #322a3d;
            }
        """)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        
        self.icon_label = QLabel("📄")
        self.icon_label.setFont(QFont("Segoe UI Emoji", 32))
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        self.text_label = QLabel("Drop file here or click to browse")
        self.text_label.setStyleSheet("color: #888; font-size: 14px;")
        self.text_label.setAlignment(Qt.AlignCenter)
        
        self.file_label = QLabel("")
        self.file_label.setStyleSheet("color: #7c3aed; font-size: 12px; font-weight: bold;")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.file_label)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet("""
                DropZone {
                    border: 2px solid #7c3aed;
                    border-radius: 12px;
                    background-color: #3d2a4d;
                }
            """)
    
    def dragLeaveEvent(self, event):
        self.setup_ui()
    
    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.set_file(files[0])
        self.setup_ui()
    
    def mousePressEvent(self, event):
        self.file_dropped.emit("")
    
    def set_file(self, file_path: str):
        self.file_path = file_path
        if file_path:
            name = Path(file_path).name
            self.icon_label.setText("✓")
            self.text_label.setText("File selected:")
            self.file_label.setText(name)
            self.file_dropped.emit(file_path)
        else:
            self.icon_label.setText("📄")
            self.text_label.setText("Drop file here or click to browse")
            self.file_label.setText("")


class MainWindow(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_dir = DEFAULT_OUTPUT_DIR
        self.worker = None
        self.packages = {}
        self.latex_available = False
        
        self.init_checks()
        self.setup_ui()
    
    def init_checks(self):
        """Check for required packages."""
        self.packages = check_and_install_packages()
        self.latex_available = check_latex_installed()
    
    def setup_ui(self):
        self.setWindowTitle("Dochameleon")
        self.setFixedSize(480, 520)
        self.setStyleSheet("""
            QMainWindow {
                background-color: #1e1e1e;
            }
            QLabel {
                color: #e0e0e0;
            }
            QPushButton {
                background-color: #7c3aed;
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 14px;
                font-weight: bold;
            }
            QPushButton:hover {
                background-color: #8b5cf6;
            }
            QPushButton:pressed {
                background-color: #6d28d9;
            }
            QPushButton:disabled {
                background-color: #444;
                color: #888;
            }
            QComboBox {
                background-color: #2a2a2a;
                color: #e0e0e0;
                border: 1px solid #444;
                border-radius: 8px;
                padding: 10px;
                font-size: 14px;
                min-width: 200px;
            }
            QComboBox:hover {
                border-color: #7c3aed;
            }
            QComboBox::drop-down {
                border: none;
                padding-right: 10px;
            }
            QComboBox::down-arrow {
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid #888;
                margin-right: 10px;
            }
            QComboBox QAbstractItemView {
                background-color: #2a2a2a;
                color: #e0e0e0;
                selection-background-color: #7c3aed;
                border: 1px solid #444;
                border-radius: 4px;
            }
            QProgressBar {
                background-color: #2a2a2a;
                border: none;
                border-radius: 4px;
                height: 6px;
            }
            QProgressBar::chunk {
                background-color: #7c3aed;
                border-radius: 4px;
            }
        """)
        
        # Try to set window icon
        icon_path = Path(__file__).parent.parent / "icons" / "dochameleon_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        layout.setContentsMargins(30, 30, 30, 30)
        layout.setSpacing(20)
        
        # Title
        title = QLabel("Dochameleon")
        title.setFont(QFont("Segoe UI", 24, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("color: #7c3aed;")
        layout.addWidget(title)
        
        subtitle = QLabel("Universal Document Converter")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("color: #888; font-size: 12px; margin-bottom: 10px;")
        layout.addWidget(subtitle)
        
        # Drop zone
        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self.on_file_dropped)
        layout.addWidget(self.drop_zone)
        
        # Conversion type
        conv_layout = QHBoxLayout()
        conv_label = QLabel("Convert to:")
        conv_label.setStyleSheet("font-size: 14px;")
        
        self.conv_combo = QComboBox()
        self.conv_combo.addItems(["PDF", "DOCX"])
        self.conv_combo.setCurrentIndex(0)
        
        conv_layout.addWidget(conv_label)
        conv_layout.addWidget(self.conv_combo, 1)
        layout.addLayout(conv_layout)
        
        # Output folder
        out_layout = QHBoxLayout()
        out_label = QLabel("Output:")
        out_label.setStyleSheet("font-size: 14px;")
        
        self.out_path_label = QLabel(str(self.output_dir))
        self.out_path_label.setStyleSheet("""
            background-color: #2a2a2a;
            border: 1px solid #444;
            border-radius: 8px;
            padding: 10px;
            font-size: 12px;
            color: #888;
        """)
        self.out_path_label.setWordWrap(True)
        
        self.browse_out_btn = QPushButton("📁")
        self.browse_out_btn.setFixedSize(44, 44)
        self.browse_out_btn.setStyleSheet("""
            QPushButton {
                background-color: #2a2a2a;
                border: 1px solid #444;
                font-size: 16px;
            }
            QPushButton:hover {
                border-color: #7c3aed;
                background-color: #322a3d;
            }
        """)
        self.browse_out_btn.clicked.connect(self.browse_output)
        
        out_layout.addWidget(out_label)
        out_layout.addWidget(self.out_path_label, 1)
        out_layout.addWidget(self.browse_out_btn)
        layout.addLayout(out_layout)
        
        # Progress bar
        self.progress = QProgressBar()
        self.progress.setTextVisible(False)
        self.progress.setMaximum(0)
        self.progress.hide()
        layout.addWidget(self.progress)
        
        # Status label
        self.status_label = QLabel("")
        self.status_label.setAlignment(Qt.AlignCenter)
        self.status_label.setStyleSheet("font-size: 12px;")
        layout.addWidget(self.status_label)
        
        # Convert button
        self.convert_btn = QPushButton("Convert")
        self.convert_btn.setEnabled(False)
        self.convert_btn.clicked.connect(self.start_conversion)
        layout.addWidget(self.convert_btn)
        
        layout.addStretch()
    
    def on_file_dropped(self, file_path: str):
        if not file_path:
            # User clicked to browse
            file_path, _ = QFileDialog.getOpenFileName(
                self,
                "Select File",
                "",
                "Supported Files (*.tex *.pdf *.docx);;LaTeX Files (*.tex);;PDF Files (*.pdf);;Word Documents (*.docx)"
            )
        
        if file_path:
            self.input_file = Path(file_path)
            self.drop_zone.set_file(file_path)
            self.update_conversion_options()
            self.convert_btn.setEnabled(True)
            self.status_label.setText("")
            self.status_label.setStyleSheet("color: #888; font-size: 12px;")
    
    def update_conversion_options(self):
        """Update combo box based on input file type."""
        if not self.input_file:
            return
        
        ext = self.input_file.suffix.lower()
        self.conv_combo.clear()
        
        if ext == '.tex':
            self.conv_combo.addItems(["PDF", "DOCX"])
        elif ext == '.pdf':
            self.conv_combo.addItems(["DOCX"])
        elif ext == '.docx':
            self.conv_combo.addItems(["PDF"])
    
    def browse_output(self):
        folder = QFileDialog.getExistingDirectory(
            self,
            "Select Output Folder",
            str(self.output_dir)
        )
        if folder:
            self.output_dir = Path(folder)
            self.out_path_label.setText(str(self.output_dir))
    
    def get_conversion_mode(self) -> str:
        """Get conversion mode based on input file and selected output."""
        if not self.input_file:
            return None
        
        ext = self.input_file.suffix.lower()
        target = self.conv_combo.currentText().lower()
        
        mode_map = {
            ('.tex', 'pdf'): 'tex2pdf',
            ('.tex', 'docx'): 'tex2docx',
            ('.pdf', 'docx'): 'pdf2docx',
            ('.docx', 'pdf'): 'docx2pdf',
        }
        
        return mode_map.get((ext, target))
    
    def start_conversion(self):
        mode = self.get_conversion_mode()
        if not mode:
            self.show_error("Invalid conversion")
            return
        
        # Check requirements
        if mode in ['tex2pdf', 'tex2docx'] and not self.latex_available:
            self.show_error("LaTeX (pdflatex) is not installed.\nInstall MiKTeX or TeX Live.")
            return
        
        if mode in ['tex2docx', 'pdf2docx'] and not self.packages.get('pdf2docx'):
            self.show_error("pdf2docx package is not available.")
            return
        
        if mode == 'docx2pdf' and not self.packages.get('docx2pdf'):
            self.show_error("docx2pdf package is not available.")
            return
        
        # Create output directory
        self.output_dir.mkdir(parents=True, exist_ok=True)
        
        # Disable UI during conversion
        self.convert_btn.setEnabled(False)
        self.progress.show()
        self.status_label.setText("Converting...")
        self.status_label.setStyleSheet("color: #7c3aed; font-size: 12px;")
        
        # Start worker thread
        self.worker = ConversionWorker(mode, self.input_file, self.output_dir)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.start()
    
    def on_conversion_finished(self, success: bool, message: str):
        self.progress.hide()
        self.convert_btn.setEnabled(True)
        
        if success:
            self.status_label.setText("✓ " + message)
            self.status_label.setStyleSheet("color: #22c55e; font-size: 12px;")
        else:
            self.status_label.setText("✗ " + message)
            self.status_label.setStyleSheet("color: #ef4444; font-size: 12px;")
    
    def show_error(self, message: str):
        QMessageBox.warning(self, "Error", message)


def run_gui():
    """Launch the GUI application."""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    run_gui()
