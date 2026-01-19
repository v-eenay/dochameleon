"""
GUI interface for Dochameleon using QtPy.
"""

import sys
import os
import subprocess
import platform
from pathlib import Path
from qtpy.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QComboBox, QFileDialog, QProgressBar,
    QFrame, QMessageBox
)
from qtpy.QtCore import Qt, QThread, Signal
from qtpy.QtGui import QFont, QIcon, QPalette, QColor

from .packages import check_and_install_packages, check_latex_installed
from .pipeline import (
    convert_single_tex_to_pdf,
    convert_single_tex_to_docx,
    convert_single_pdf_to_docx,
    convert_single_docx_to_pdf,
)

# Default output directory
DEFAULT_OUTPUT_DIR = Path(__file__).parent.parent / "output"

# Color palette - Light pastel professional
COLORS = {
    'bg': '#f8f9fa',
    'card': '#ffffff',
    'border': '#e2e8f0',
    'border_hover': '#94a3b8',
    'text': '#1e293b',
    'text_muted': '#64748b',
    'primary': '#6366f1',
    'primary_hover': '#4f46e5',
    'primary_pressed': '#4338ca',
    'success': '#10b981',
    'error': '#ef4444',
    'drop_zone': '#f1f5f9',
    'drop_zone_hover': '#e0e7ff',
}


class ConversionWorker(QThread):
    """Worker thread for file conversion."""
    finished = Signal(bool, str, str)  # success, message, output_file_path
    
    def __init__(self, mode: str, input_file: Path, output_dir: Path):
        super().__init__()
        self.mode = mode
        self.input_file = input_file
        self.output_dir = output_dir
        self.output_file = None
    
    def run(self):
        try:
            # Determine output file extension
            ext_map = {
                'tex2pdf': '.pdf',
                'tex2docx': '.docx',
                'pdf2docx': '.docx',
                'docx2pdf': '.pdf',
            }
            output_ext = ext_map.get(self.mode, '')
            self.output_file = self.output_dir / (self.input_file.stem + output_ext)
            
            if self.mode == 'tex2pdf':
                success = convert_single_tex_to_pdf(self.input_file, self.output_dir)
            elif self.mode == 'tex2docx':
                success = convert_single_tex_to_docx(self.input_file, self.output_dir)
            elif self.mode == 'pdf2docx':
                success = convert_single_pdf_to_docx(self.input_file, self.output_dir)
            elif self.mode == 'docx2pdf':
                success = convert_single_docx_to_pdf(self.input_file, self.output_dir)
            else:
                self.finished.emit(False, "Unknown conversion mode", "")
                return
            
            if success:
                self.finished.emit(True, "Conversion completed!", str(self.output_file))
            else:
                self.finished.emit(False, "Conversion failed", "")
        except Exception as e:
            self.finished.emit(False, str(e), "")


class DropZone(QFrame):
    """Drag and drop zone for files."""
    file_dropped = Signal(str)
    
    def __init__(self):
        super().__init__()
        self.setAcceptDrops(True)
        self.setMinimumHeight(130)
        self.file_path = None
        self.setup_ui()
    
    def setup_ui(self):
        self.setStyleSheet(f"""
            DropZone {{
                border: 2px dashed {COLORS['border']};
                border-radius: 12px;
                background-color: {COLORS['drop_zone']};
            }}
            DropZone:hover {{
                border-color: {COLORS['primary']};
                background-color: {COLORS['drop_zone_hover']};
            }}
        """)
        
        layout = QVBoxLayout(self)
        layout.setAlignment(Qt.AlignCenter)
        layout.setSpacing(8)
        
        self.icon_label = QLabel("📄")
        self.icon_label.setFont(QFont("Segoe UI Emoji", 28))
        self.icon_label.setAlignment(Qt.AlignCenter)
        
        self.text_label = QLabel("Drop file here or click to browse")
        self.text_label.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 13px;")
        self.text_label.setAlignment(Qt.AlignCenter)
        
        self.file_label = QLabel("")
        self.file_label.setStyleSheet(f"color: {COLORS['primary']}; font-size: 12px; font-weight: 600;")
        self.file_label.setAlignment(Qt.AlignCenter)
        self.file_label.setWordWrap(True)
        
        layout.addWidget(self.icon_label)
        layout.addWidget(self.text_label)
        layout.addWidget(self.file_label)
    
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            self.setStyleSheet(f"""
                DropZone {{
                    border: 2px solid {COLORS['primary']};
                    border-radius: 12px;
                    background-color: {COLORS['drop_zone_hover']};
                }}
            """)
    
    def dragLeaveEvent(self, event):
        self.setup_ui()
    
    def dropEvent(self, event):
        files = [url.toLocalFile() for url in event.mimeData().urls()]
        if files:
            self.set_file(files[0], emit_signal=True)
        self.setup_ui()
    
    def mousePressEvent(self, event):
        self.file_dropped.emit("")
    
    def set_file(self, file_path: str, emit_signal: bool = False):
        self.file_path = file_path
        if file_path:
            name = Path(file_path).name
            self.icon_label.setText("✓")
            self.icon_label.setStyleSheet(f"color: {COLORS['success']};")
            self.text_label.setText("File selected:")
            self.file_label.setText(name)
            if emit_signal:
                self.file_dropped.emit(file_path)
        else:
            self.icon_label.setText("📄")
            self.icon_label.setStyleSheet("")
            self.text_label.setText("Drop file here or click to browse")
            self.file_label.setText("")


class MainWindow(QMainWindow):
    """Main application window."""
    
    def __init__(self):
        super().__init__()
        self.input_file = None
        self.output_dir = DEFAULT_OUTPUT_DIR
        self.output_file = None
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
        self.setFixedSize(460, 560)
        self.setStyleSheet(f"""
            QMainWindow {{
                background-color: {COLORS['bg']};
            }}
            QLabel {{
                color: {COLORS['text']};
            }}
            QPushButton {{
                background-color: {COLORS['primary']};
                color: white;
                border: none;
                border-radius: 8px;
                padding: 12px 24px;
                font-size: 13px;
                font-weight: 600;
            }}
            QPushButton:hover {{
                background-color: {COLORS['primary_hover']};
            }}
            QPushButton:pressed {{
                background-color: {COLORS['primary_pressed']};
            }}
            QPushButton:disabled {{
                background-color: {COLORS['border']};
                color: {COLORS['text_muted']};
            }}
            QComboBox {{
                background-color: {COLORS['card']};
                color: {COLORS['text']};
                border: 1px solid {COLORS['border']};
                border-radius: 8px;
                padding: 10px 12px;
                font-size: 13px;
                min-width: 180px;
            }}
            QComboBox:hover {{
                border-color: {COLORS['primary']};
            }}
            QComboBox::drop-down {{
                border: none;
                padding-right: 10px;
            }}
            QComboBox::down-arrow {{
                image: none;
                border-left: 5px solid transparent;
                border-right: 5px solid transparent;
                border-top: 6px solid {COLORS['text_muted']};
                margin-right: 10px;
            }}
            QComboBox QAbstractItemView {{
                background-color: {COLORS['card']};
                color: {COLORS['text']};
                selection-background-color: {COLORS['primary']};
                selection-color: white;
                border: 1px solid {COLORS['border']};
                border-radius: 4px;
                padding: 4px;
            }}
            QProgressBar {{
                background-color: {COLORS['border']};
                border: none;
                border-radius: 4px;
                height: 6px;
            }}
            QProgressBar::chunk {{
                background-color: {COLORS['primary']};
                border-radius: 4px;
            }}
        """)
        
        # Try to set window icon
        icon_path = Path(__file__).parent.parent / "icons" / "dochameleon_icon.ico"
        if icon_path.exists():
            self.setWindowIcon(QIcon(str(icon_path)))
        
        # Central widget
        central = QWidget()
        self.setCentralWidget(central)
        
        layout = QVBoxLayout(central)
        layout.setContentsMargins(32, 28, 32, 28)
        layout.setSpacing(16)
        
        # Title
        title = QLabel("Dochameleon")
        title.setFont(QFont("Segoe UI", 22, QFont.Bold))
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet(f"color: {COLORS['primary']};")
        layout.addWidget(title)
        
        subtitle = QLabel("Universal Document Converter")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet(f"color: {COLORS['text_muted']}; font-size: 12px; margin-bottom: 8px;")
        layout.addWidget(subtitle)
        
        # Drop zone
        self.drop_zone = DropZone()
        self.drop_zone.file_dropped.connect(self.on_file_dropped)
        layout.addWidget(self.drop_zone)
        
        # Conversion type
        conv_layout = QHBoxLayout()
        conv_layout.setSpacing(12)
        conv_label = QLabel("Convert to:")
        conv_label.setStyleSheet("font-size: 13px; font-weight: 500;")
        
        self.conv_combo = QComboBox()
        self.conv_combo.addItems(["PDF", "DOCX"])
        self.conv_combo.setCurrentIndex(0)
        
        conv_layout.addWidget(conv_label)
        conv_layout.addWidget(self.conv_combo, 1)
        layout.addLayout(conv_layout)
        
        # Output folder
        out_layout = QHBoxLayout()
        out_layout.setSpacing(8)
        out_label = QLabel("Output:")
        out_label.setStyleSheet("font-size: 13px; font-weight: 500;")
        
        self.out_path_label = QLabel(str(self.output_dir))
        self.out_path_label.setStyleSheet(f"""
            background-color: {COLORS['card']};
            border: 1px solid {COLORS['border']};
            border-radius: 8px;
            padding: 10px 12px;
            font-size: 11px;
            color: {COLORS['text_muted']};
        """)
        self.out_path_label.setWordWrap(True)
        
        self.browse_out_btn = QPushButton("📁")
        self.browse_out_btn.setFixedSize(42, 42)
        self.browse_out_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['card']};
                border: 1px solid {COLORS['border']};
                font-size: 15px;
                color: {COLORS['text']};
            }}
            QPushButton:hover {{
                border-color: {COLORS['primary']};
                background-color: {COLORS['drop_zone_hover']};
            }}
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
        
        # Action buttons container
        self.action_buttons = QWidget()
        action_layout = QHBoxLayout(self.action_buttons)
        action_layout.setContentsMargins(0, 0, 0, 0)
        action_layout.setSpacing(10)
        
        self.open_file_btn = QPushButton("Open File")
        self.open_file_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['success']};
            }}
            QPushButton:hover {{
                background-color: #059669;
            }}
        """)
        self.open_file_btn.clicked.connect(self.open_output_file)
        
        self.open_folder_btn = QPushButton("Show in Folder")
        self.open_folder_btn.setStyleSheet(f"""
            QPushButton {{
                background-color: {COLORS['card']};
                color: {COLORS['text']};
                border: 1px solid {COLORS['border']};
            }}
            QPushButton:hover {{
                background-color: {COLORS['drop_zone']};
                border-color: {COLORS['primary']};
            }}
        """)
        self.open_folder_btn.clicked.connect(self.open_output_folder)
        
        action_layout.addWidget(self.open_file_btn)
        action_layout.addWidget(self.open_folder_btn)
        self.action_buttons.hide()
        layout.addWidget(self.action_buttons)
        
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
            self.action_buttons.hide()
    
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
        self.action_buttons.hide()
        self.progress.show()
        self.status_label.setText("Converting...")
        self.status_label.setStyleSheet(f"color: {COLORS['primary']}; font-size: 12px;")
        
        # Start worker thread
        self.worker = ConversionWorker(mode, self.input_file, self.output_dir)
        self.worker.finished.connect(self.on_conversion_finished)
        self.worker.start()
    
    def on_conversion_finished(self, success: bool, message: str, output_file: str):
        self.progress.hide()
        self.convert_btn.setEnabled(True)
        
        if success:
            self.output_file = Path(output_file) if output_file else None
            self.status_label.setText("✓ " + message)
            self.status_label.setStyleSheet(f"color: {COLORS['success']}; font-size: 12px; font-weight: 500;")
            self.action_buttons.show()
        else:
            self.status_label.setText("✗ " + message)
            self.status_label.setStyleSheet(f"color: {COLORS['error']}; font-size: 12px;")
            self.action_buttons.hide()
    
    def open_output_file(self):
        """Open the converted file with default application."""
        if self.output_file and self.output_file.exists():
            if platform.system() == 'Windows':
                os.startfile(str(self.output_file))
            elif platform.system() == 'Darwin':
                subprocess.run(['open', str(self.output_file)])
            else:
                subprocess.run(['xdg-open', str(self.output_file)])
    
    def open_output_folder(self):
        """Open the output folder in file browser and select the file."""
        if self.output_file and self.output_file.exists():
            if platform.system() == 'Windows':
                subprocess.run(['explorer', '/select,', str(self.output_file)])
            elif platform.system() == 'Darwin':
                subprocess.run(['open', '-R', str(self.output_file)])
            else:
                subprocess.run(['xdg-open', str(self.output_dir)])
        elif self.output_dir.exists():
            if platform.system() == 'Windows':
                os.startfile(str(self.output_dir))
            elif platform.system() == 'Darwin':
                subprocess.run(['open', str(self.output_dir)])
            else:
                subprocess.run(['xdg-open', str(self.output_dir)])
    
    def show_error(self, message: str):
        QMessageBox.warning(self, "Error", message)


def run_gui():
    """Launch the GUI application."""
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    
    # Set light palette
    palette = QPalette()
    palette.setColor(QPalette.Window, QColor(COLORS['bg']))
    palette.setColor(QPalette.WindowText, QColor(COLORS['text']))
    palette.setColor(QPalette.Base, QColor(COLORS['card']))
    palette.setColor(QPalette.Text, QColor(COLORS['text']))
    palette.setColor(QPalette.Button, QColor(COLORS['card']))
    palette.setColor(QPalette.ButtonText, QColor(COLORS['text']))
    palette.setColor(QPalette.Highlight, QColor(COLORS['primary']))
    palette.setColor(QPalette.HighlightedText, QColor('#ffffff'))
    app.setPalette(palette)
    
    window = MainWindow()
    window.show()
    
    sys.exit(app.exec())


if __name__ == "__main__":
    run_gui()
