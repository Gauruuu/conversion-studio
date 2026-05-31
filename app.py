import os
import sys
import subprocess
from PyQt6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, 
                             QHBoxLayout, QPushButton, QListWidget, QFileDialog, 
                             QComboBox, QLabel, QProgressBar, QMessageBox, QTabWidget)
from PyQt6.QtCore import Qt, QThread, pyqtSignal
from PyQt6.QtGui import QIcon
from PIL import Image
import pillow_heif
from pdf2docx import Converter

# Register HEIF decoder with Pillow
pillow_heif.register_heif_opener()

try:
    import comtypes.client
    COM_AVAILABLE = True
except ImportError:
    COM_AVAILABLE = False

# ==============================================================================
# Ultimate Multi-Format Processing Engine (Threaded)
# ==============================================================================
class UltimateWorker(QThread):
    progress = pyqtSignal(int, str)
    finished = pyqtSignal(bool, str)

    def __init__(self, files, conversion_type, output_dir):
        super().__init__()
        self.files = files
        self.conversion_type = conversion_type
        self.output_dir = output_dir

    def run(self):
        try:
            total_files = len(self.files)
            if total_files == 0:
                self.finished.emit(False, "No files selected.")
                return

            # --- SPECIAL HANDLING: MASS IMAGES TO SINGLE PDF ---
            if self.conversion_type == "Mass Images to PDF":
                pdf_path = os.path.join(self.output_dir, "Combined_Images.pdf")
                self.progress.emit(10, "Opening image streams...")
                img_list = []
                for idx, f in enumerate(self.files):
                    img = Image.open(f)
                    if img.mode in ("RGBA", "P"):
                        img = img.convert("RGB")
                    img_list.append(img)
                    self.progress.emit(int(10 + (idx / total_files) * 80), f"Stitching image {idx+1}...")
                img_list[0].save(pdf_path, save_all=True, append_images=img_list[1:], quality=100)
                self.progress.emit(100, "Done!")
                self.finished.emit(True, f"Combined PDF saved to:\n{pdf_path}")
                return

            # --- INDIVIDUAL FILE PROCESSING LOOP ---
            for idx, file_path in enumerate(self.files):
                base_name = os.path.splitext(os.path.basename(file_path))[0]
                self.progress.emit(int((idx / total_files) * 100), f"Processing: {os.path.basename(file_path)}")

                # 1. DOCUMENT TAB ENGINES
                if self.conversion_type == "Word to PDF" and COM_AVAILABLE:
                    out_f = os.path.join(self.output_dir, f"{base_name}.pdf")
                    word = comtypes.client.CreateObject('Word.Application')
                    word.Visible = False
                    doc = word.Documents.Open(os.path.abspath(file_path))
                    doc.SaveAs(os.path.abspath(out_f), FileFormat=17)
                    doc.Close()
                    word.Quit()

                elif self.conversion_type == "PPTX to PDF" and COM_AVAILABLE:
                    out_f = os.path.join(self.output_dir, f"{base_name}.pdf")
                    ppt = comtypes.client.CreateObject('PowerPoint.Application')
                    deck = ppt.Presentations.Open(os.path.abspath(file_path), WithWindow=False)
                    deck.SaveAs(os.path.abspath(out_f), FileFormat=32)
                    deck.Close()
                    ppt.Quit()

                elif self.conversion_type == "PDF to Word":
                    out_f = os.path.join(self.output_dir, f"{base_name}.docx")
                    cv = Converter(file_path)
                    cv.convert(out_f, start=0, end=None)
                    cv.close()

                # 2. IMAGE TAB ENGINES (Updated with Image to ICO)
                elif self.conversion_type == "Image to ICO (App Icon)":
                    out_f = os.path.join(self.output_dir, f"{base_name}.ico")
                    img = Image.open(file_path)
                    # Enforce standard multi-layer Windows desktop app sizing array rules
                    icon_sizes = [(16, 16), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]
                    img.save(out_f, format="ICO", sizes=icon_sizes)

                elif self.conversion_type == "HEIC to JPG":
                    out_f = os.path.join(self.output_dir, f"{base_name}.jpg")
                    Image.open(file_path).save(out_f, "JPEG", quality=100, subsampling=0)

                elif self.conversion_type == "WebP to PNG":
                    out_f = os.path.join(self.output_dir, f"{base_name}.png")
                    Image.open(file_path).save(out_f, "PNG", compress_level=0)

                elif self.conversion_type == "PNG to JPG":
                    out_f = os.path.join(self.output_dir, f"{base_name}.jpg")
                    img = Image.open(file_path).convert("RGB")
                    img.save(out_f, "JPEG", quality=100, subsampling=0)

                elif self.conversion_type == "JPG to PNG":
                    out_f = os.path.join(self.output_dir, f"{base_name}.png")
                    Image.open(file_path).save(out_f, "PNG", compress_level=0)

                # 3. VIDEO TAB ENGINES
                elif self.conversion_type == "MOV to MP4 (DaVinci Optimized)":
                    out_f = os.path.join(self.output_dir, f"{base_name}.mp4")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c:v', 'libx264', '-crf', '14', '-pix_fmt', 'yuv420p', '-c:a', 'aac', '-b:a', '324k', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                elif self.conversion_type == "MP4 to MOV":
                    out_f = os.path.join(self.output_dir, f"{base_name}.mov")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c:v', 'libx264', '-crf', '14', '-c:a', 'pcm_s16le', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                elif self.conversion_type == "MKV to MP4 (Lossless Remux)":
                    out_f = os.path.join(self.output_dir, f"{base_name}.mp4")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c', 'copy', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                elif self.conversion_type == "Video to GIF":
                    out_f = os.path.join(self.output_dir, f"{base_name}.gif")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-vf', 'fps=15,scale=480:-1:flags=lanczos,split[s0][s1];[s0]palettegen[p];[s1][p]paletteuse', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                # 4. AUDIO TAB ENGINES
                elif self.conversion_type == "WAV to MP3 (Primesync 320kbps)":
                    out_f = os.path.join(self.output_dir, f"{base_name}.mp3")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c:a', 'libmp3lame', '-b:a', '320k', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                elif self.conversion_type == "M4A to WAV (Lossless Unpack)":
                    out_f = os.path.join(self.output_dir, f"{base_name}.wav")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c:a', 'pcm_s16le', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

                elif self.conversion_type == "MP3 to WAV":
                    out_f = os.path.join(self.output_dir, f"{base_name}.wav")
                    cmd = ['ffmpeg', '-y', '-i', file_path, '-c:a', 'pcm_s16le', out_f]
                    subprocess.run(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, creationflags=0x08000000)

            self.progress.emit(100, "Done!")
            self.finished.emit(True, f"All processes completed successfully!\nSaved to: {self.output_dir}")
        except Exception as e:
            self.finished.emit(False, str(e))

# ==============================================================================
# Tabbed Studio User Interface
# ==============================================================================
class UltimateStudioApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_files = []
        self.init_ui()

    def init_ui(self):
        self.setWindowTitle("Converter Studio")
        self.resize(900, 650)
        
        # --- APPLICATION ICON INTEGRATION ---
        icon_filename = "app_icon.ico"
        if os.path.exists(icon_filename):
            self.setWindowIcon(QIcon(icon_filename))
        elif os.path.exists(os.path.join(sys._MEIPASS, icon_filename) if hasattr(sys, '_MEIPASS') else icon_filename):
            self.setWindowIcon(QIcon(os.path.join(sys._MEIPASS, icon_filename)))
        
        # Sleek Matrix Dark Aesthetic Styling
        self.setStyleSheet("""
            QMainWindow { background-color: #0F172A; }
            QWidget { color: #F1F5F9; font-family: 'Segoe UI', Arial, sans-serif; font-size: 13px; }
            QTabWidget::pane { border: 1px solid #1E293B; background: #1E293B; border-radius: 6px; }
            QTabBar::tab { background: #0F172A; border: 1px solid #1E293B; padding: 10px 20px; border-top-left-radius: 4px; border-top-right-radius: 4px; }
            QTabBar::tab:selected { background: #1E293B; border-bottom-color: #1E293B; font-weight: bold; color: #38BDF8; }
            QListWidget { background-color: #0B0F19; border: 1px solid #1E293B; border-radius: 6px; padding: 5px; }
            QPushButton { background-color: #2563EB; border: none; border-radius: 4px; padding: 8px 16px; font-weight: bold; }
            QPushButton:hover { background-color: #3B82F6; }
            QPushButton:disabled { background-color: #334155; color: #64748B; }
            QComboBox { background-color: #0B0F19; border: 1px solid #1E293B; border-radius: 4px; padding: 6px; min-width: 260px; color: #F1F5F9; }
            QProgressBar { border: 1px solid #1E293B; border-radius: 4px; text-align: center; background-color: #0B0F19; }
            QProgressBar::chunk { background-color: #0EA5E9; }
            QLabel { font-weight: 500; }
        """)

        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(15)

        # Header Title Area
        header = QLabel("Universal High-Fidelity Media Processor")
        header.setStyleSheet("font-size: 20px; font-weight: bold; color: #38BDF8; padding-bottom: 2px;")
        main_layout.addWidget(header)

        # Processing Task Categories Tabs Layout
        self.tabs = QTabWidget()
        
        self.doc_options = ["Word to PDF", "PPTX to PDF", "PDF to Word"]
        # Added "Image to ICO (App Icon)" option here
        self.img_options = ["Image to ICO (App Icon)", "Mass Images to PDF", "HEIC to JPG", "WebP to PNG", "PNG to JPG", "JPG to PNG"]
        self.vid_options = ["MOV to MP4 (DaVinci Optimized)", "MP4 to MOV", "MKV to MP4 (Lossless Remux)", "Video to GIF"]
        self.aud_options = ["WAV to MP3 (Primesync 320kbps)", "M4A to WAV (Lossless Unpack)", "MP3 to WAV"]

        self.tab_docs = QWidget()
        self.tab_imgs = QWidget()
        self.tab_vids = QWidget()
        self.tab_auds = QWidget()
        
        self.tabs.addTab(self.tab_docs, "📄 Documents")
        self.tabs.addTab(self.tab_imgs, "🖼️ Image Deck")
        self.tabs.addTab(self.tab_vids, "🎬 Video Studio")
        self.tabs.addTab(self.tab_auds, "🎵 Audio Station")
        main_layout.addWidget(self.tabs)

        # Combined Processing Engine Profile Dropdown Selector 
        config_box = QHBoxLayout()
        config_box.addWidget(QLabel("Active Profile Pipeline:"))
        self.combo_engine = QComboBox()
        config_box.addWidget(self.combo_engine)
        config_box.addStretch()
        main_layout.addLayout(config_box)

        # Wire up tab switches to update the dropdown options dynamically
        self.tabs.currentChanged.connect(self.sync_engine_options)
        self.sync_engine_options(0)

        # Files Pipeline Processing Queue Tracker Display Window Layout
        main_layout.addWidget(QLabel("Active Pipeline Operation File Queue:"))
        self.file_list_widget = QListWidget()
        main_layout.addWidget(self.file_list_widget)

        # Add/Clear Selection Operations Layout Links
        file_btn_layout = QHBoxLayout()
        self.btn_add_files = QPushButton("Add Target Files")
        self.btn_add_files.clicked.connect(self.load_files)
        self.btn_clear_queue = QPushButton("Clear Selection")
        self.btn_clear_queue.setStyleSheet("background-color: #EF4444;")
        self.btn_clear_queue.clicked.connect(self.clear_queue)
        file_btn_layout.addWidget(self.btn_add_files)
        file_btn_layout.addWidget(self.btn_clear_queue)
        file_btn_layout.addStretch()
        main_layout.addLayout(file_btn_layout)

        # Processing Metrics Bar Displays
        self.progress_bar = QProgressBar()
        self.progress_bar.setValue(0)
        main_layout.addWidget(self.progress_bar)
        
        self.status_lbl = QLabel("Awaiting initialization sequence instruction inputs...")
        self.status_lbl.setStyleSheet("color: #94A3B8; font-style: italic;")
        main_layout.addWidget(self.status_lbl)

        # Launch Execution Run Commands
        self.btn_run = QPushButton("Execute Operations Sequence")
        self.btn_run.setStyleSheet("background-color: #10B981; font-size: 14px; padding: 12px; color: #FFFFFF;")
        self.btn_run.clicked.connect(self.start_conversion_pipeline)
        main_layout.addWidget(self.btn_run)

    # ==============================================================================
    # Core Controller Logic UI Coordination Engines
    # ==============================================================================
    def sync_engine_options(self, index):
        self.combo_engine.clear()
        if index == 0: self.combo_engine.addItems(self.doc_options)
        elif index == 1: self.combo_engine.addItems(self.img_options)
        elif index == 2: self.combo_engine.addItems(self.vid_options)
        elif index == 3: self.combo_engine.addItems(self.aud_options)

    def load_files(self):
        files, _ = QFileDialog.getOpenFileNames(self, "Load High Fidelity Target Assets", "", "All Supported Files (*.*)")
        if files:
            self.selected_files.extend(files)
            for f in files:
                self.file_list_widget.addItem(os.path.basename(f))

    def clear_queue(self):
        self.selected_files.clear()
        self.file_list_widget.clear()
        self.progress_bar.setValue(0)
        self.status_lbl.setText("Queue cleared. System idle.")

    def start_conversion_pipeline(self):
        if not self.selected_files:
            QMessageBox.warning(self, "Queue Context Empty", "Please append system files to current list queue context map before running.")
            return
            
        output_dir = QFileDialog.getExistingDirectory(self, "Select Asset Extraction Destination Directory Location")
        if not output_dir:
            return

        self.toggle_ui_state(False)
        engine_mode = self.combo_engine.currentText()
        
        self.worker = UltimateWorker(self.selected_files, engine_mode, output_dir)
        self.worker.progress.connect(self.update_progress_ui)
        self.worker.finished.connect(self.processing_finished_callback)
        self.worker.start()

    def toggle_ui_state(self, enabled_state):
        self.btn_run.setEnabled(enabled_state)
        self.btn_add_files.setEnabled(enabled_state)
        self.btn_clear_queue.setEnabled(enabled_state)
        self.tabs.setEnabled(enabled_state)
        self.combo_engine.setEnabled(enabled_state)

    def update_progress_ui(self, value, update_msg):
        self.progress_bar.setValue(value)
        self.status_lbl.setText(update_msg)

    def processing_finished_callback(self, status_success, message_details):
        self.toggle_ui_state(True)
        if status_success:
            QMessageBox.information(self, "Pipeline Execution Complete", message_details)
            self.clear_queue()
        else:
            QMessageBox.critical(self, "Pipeline Internal Critical Halt Error Exception", f"Process halted unexpectedly:\n{message_details}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = UltimateStudioApp()
    window.show()
    sys.exit(app.exec())