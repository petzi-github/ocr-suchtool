# gui.py
import os
import sys
import logging
from PySide6.QtWidgets import (
    QMainWindow, QWidget, QVBoxLayout, QHBoxLayout, QLabel,
    QTextEdit, QPushButton, QComboBox, QCheckBox, QListWidget,
    QListWidgetItem, QMessageBox, QProgressBar, QFileDialog
)
from PySide6.QtCore import Qt, QObject, Signal, QThread
from PySide6.QtGui import QFont

from utils import bildformate, pdfformate
from ocr_engine import starte_ocr, resource_path

# ---------- Logging einrichten ----------
log_dir = os.path.join(os.path.dirname(__file__), "logs")
os.makedirs(log_dir, exist_ok=True)
log_file = os.path.join(log_dir, "ocr_tool.log")

logging.basicConfig(
    level=logging.DEBUG,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(log_file, encoding="utf-8"),
        logging.StreamHandler()
    ]
)

class OCRWorker(QObject):
    finished = Signal(list, str)
    status_signal = Signal(str)
    progress_signal = Signal(int)

    def __init__(self, dateien, suchbegriffe, sprache, optimierung,
                 full_doc, highlight, poppler_path=None, output_dir=None, abbrechen_flag=None):
        super().__init__()
        self.dateien = dateien
        self.suchbegriffe = suchbegriffe
        self.sprache = sprache
        self.optimierung = optimierung
        self.full_doc = full_doc
        self.highlight = highlight
        self.poppler_path = poppler_path
        self.output_dir = output_dir
        self._abbrechen = False
        self.abbrechen_flag = abbrechen_flag if abbrechen_flag else lambda: self._abbrechen

    def run(self):
        try:
            logging.info("OCR-Worker gestartet")
            result, treffer_datei = starte_ocr(
                dateien=self.dateien,
                suchbegriffe=self.suchbegriffe,
                sprache=self.sprache,
                optimierung=self.optimierung,
                full_doc=self.full_doc,
                highlight=self.highlight,
                status_signal=self.status_signal,
                progress_signal=self.progress_signal,
                poppler_path=self.poppler_path,
                output_dir=self.output_dir,
                abbrechen_flag=self.abbrechen_flag
            )
            logging.info("OCR-Worker beendet")
            self.finished.emit(result, treffer_datei)
        except Exception as e:
            logging.exception("Fehler im OCR-Worker")
            self.status_signal.emit(f"[!] Fehler im OCR-Worker: {e}")
            self.finished.emit([], None)

    def abbrechen(self):
        self._abbrechen = True

class OCRApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("OCR Suchtool")
        self.resize(1000, 750)
        self.setAcceptDrops(True)

        self.poppler_path = resource_path(os.path.join("poppler", "Library", "bin"))
        self.dateien = []
        self.output_dir = None

        haupt_widget = QWidget()
        layout = QVBoxLayout(haupt_widget)

        # Suchbegriffe
        layout.addWidget(QLabel("Suchbegriffe (je Zeile ein Wort):"))
        self.keyword_textfeld = QTextEdit()
        self.keyword_textfeld.setFont(QFont("Consolas", 11))
        layout.addWidget(self.keyword_textfeld)

        # Sprache
        sprache_layout = QHBoxLayout()
        sprache_layout.addWidget(QLabel("Sprache:"))
        self.sprache_dropdown = QComboBox()
        self.sprache_dropdown.addItem("Deutsch (modern)", "deu")
        self.sprache_dropdown.addItem("Deutsch (Fraktur)", "deu_frak")
        self.sprache_dropdown.addItem("Deutsch (Fraktur alt)", "frk")
        self.sprache_dropdown.addItem("Deutsch (Mischschrift)", "deu_latf")
        #self.sprache_dropdown.addItem("Deutsch kombiniert (alle)", "deu+deu_frak+frk+deu_latf")
        sprache_layout.addWidget(self.sprache_dropdown)
        layout.addLayout(sprache_layout)

        # Bildoptimierung
        optim_layout = QHBoxLayout()
        optim_layout.addWidget(QLabel("Bildoptimierung:"))
        self.optim_dropdown = QComboBox()
        self.optim_dropdown.addItem("Keine", None)
        self.optim_dropdown.addItem("Pillow: Graustufen, Schärfen, Kontrast", "pillow")
        self.optim_dropdown.addItem("OpenCV: Binarisierung, Rauschfilter", "opencv")
        self.optim_dropdown.addItem("Kombiniert: Pillow + OpenCV", "kombiniert")
        optim_layout.addWidget(self.optim_dropdown)
        layout.addLayout(optim_layout)

        # Checkboxen
        self.doc_checkbox = QCheckBox("Kompletten Scan als .doc speichern")
        self.highlight_checkbox = QCheckBox("Treffer farbig markieren")
        self.highlight_checkbox.setVisible(False)
        layout.addWidget(self.doc_checkbox)
        layout.addWidget(self.highlight_checkbox)
        self.doc_checkbox.stateChanged.connect(self.toggle_highlight_checkbox)

        # Ausgabeordner
        output_layout = QHBoxLayout()
        output_layout.addWidget(QLabel("Ausgabeordner:"))
        self.output_label = QLabel("(noch nicht gewählt)")
        output_layout.addWidget(self.output_label)
        self.output_button = QPushButton("Ordner wählen")
        self.output_button.clicked.connect(self.select_output_folder)
        output_layout.addWidget(self.output_button)
        output_layout.addStretch()
        layout.addLayout(output_layout)

        # Dateiliste
        layout.addWidget(QLabel("Dateien:"))
        self.file_list_widget = QListWidget()
        layout.addWidget(self.file_list_widget)

        # Buttons
        button_layout = QHBoxLayout()
        self.start_button = QPushButton("OCR starten")
        self.clear_button = QPushButton("Liste löschen")
        self.cancel_button = QPushButton("Abbrechen")
        self.cancel_button.setEnabled(False)
        button_layout.addWidget(self.start_button)
        button_layout.addWidget(self.clear_button)
        button_layout.addWidget(self.cancel_button)
        layout.addLayout(button_layout)

        # Status
        layout.addWidget(QLabel("Status:"))
        self.status_label = QLabel("Bereit.")
        layout.addWidget(self.status_label)
        self.progress_bar = QProgressBar()
        self.progress_bar.setRange(0, 100)
        layout.addWidget(self.progress_bar)
        self.status_feld = QTextEdit()
        self.status_feld.setReadOnly(True)
        self.status_feld.setFont(QFont("Consolas", 10))
        layout.addWidget(self.status_feld)

        self.setCentralWidget(haupt_widget)

        # Signale
        self.start_button.clicked.connect(self.start_worker)
        self.clear_button.clicked.connect(self.liste_loeschen)
        self.cancel_button.clicked.connect(self.abbrechen_worker)

        self.worker = None
        self.thread = None

        # Version
        version_label = QLabel("Version 0.1")
        version_label.setAlignment(Qt.AlignRight)
        version_label.setStyleSheet("color: gray; font-size: 10px;")
        layout.addWidget(version_label)

    # ---------- Drag & Drop ----------
    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        for url in event.mimeData().urls():
            pfad = url.toLocalFile()
            if os.path.isdir(pfad):
                for root, _, files in os.walk(pfad):
                    for f in files:
                        self.add_file(os.path.join(root, f))
            else:
                self.add_file(pfad)

    def add_file(self, pfad):
        if pfad.lower().endswith(bildformate + pdfformate):
            if pfad not in self.dateien:
                self.dateien.append(pfad)
                self.file_list_widget.addItem(QListWidgetItem(pfad))

    def liste_loeschen(self):
        self.dateien = []
        self.file_list_widget.clear()
        self.status_feld.clear()
        self.progress_bar.setValue(0)
        self.status_label.setText("Bereit.")

    def toggle_highlight_checkbox(self):
        self.highlight_checkbox.setVisible(self.doc_checkbox.isChecked())

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Ausgabeordner wählen")
        if folder:
            self.output_dir = folder
            self.output_label.setText(folder)

    def start_worker(self):
        try:
            if not self.dateien:
                QMessageBox.warning(self, "Keine Dateien", "Bitte erst Dateien per Drag & Drop hinzufügen.")
                return
            if not self.output_dir:
                QMessageBox.warning(self, "Ausgabeordner fehlt", "Bitte wähle einen Ausgabeordner aus.")
                return

            suchbegriffe = [w.strip().lower() for w in self.keyword_textfeld.toPlainText().splitlines() if w.strip()]
            if not suchbegriffe:
                QMessageBox.warning(self, "Keine Suchbegriffe", "Bitte gib mindestens ein Suchwort ein.")
                return

            sprache = self.sprache_dropdown.currentData()
            optimierung = self.optim_dropdown.currentData()
            self.status_feld.clear()
            self.status_label.setText("Starte OCR...")
            self.progress_bar.setValue(0)

            self.start_button.setEnabled(False)
            self.clear_button.setEnabled(False)
            self.cancel_button.setEnabled(True)

            self.worker = OCRWorker(
                dateien=self.dateien,
                suchbegriffe=suchbegriffe,
                sprache=sprache,
                optimierung=optimierung,
                full_doc=self.doc_checkbox.isChecked(),
                highlight=self.highlight_checkbox.isChecked(),
                poppler_path=self.poppler_path,
                output_dir=self.output_dir
            )
            self.thread = QThread()
            self.worker.moveToThread(self.thread)

            self.worker.status_signal.connect(lambda msg: self.status_feld.append(msg))
            self.worker.progress_signal.connect(self.progress_bar.setValue)
            self.worker.finished.connect(self.ocr_finished)
            self.thread.started.connect(self.worker.run)
            self.thread.start()
        except Exception as e:
            logging.exception("Fehler beim Start des OCR-Workers")
            QMessageBox.critical(self, "Fehler", f"Fehler beim Start des OCR-Workers: {e}")

    def abbrechen_worker(self):
        if self.worker:
            self.worker.abbrechen()
            self.status_feld.append("[!] Abbrechen angefordert...")
            self.cancel_button.setEnabled(False)

    def ocr_finished(self, docs, treffer_datei):
        try:
            self.status_label.setText("Fertig.")
            self.progress_bar.setValue(self.progress_bar.maximum())
            self.start_button.setEnabled(True)
            self.clear_button.setEnabled(True)
            self.cancel_button.setEnabled(False)
            if self.thread:
                self.thread.quit()
                self.thread.wait()

            if not docs:
                QMessageBox.information(self, "OCR beendet", "Keine Dateien wurden erstellt.")
                return

            # Nur Trefferliste öffnen, wenn gewünscht
            if treffer_datei:
                antwort = QMessageBox.question(
                    self,
                    "OCR abgeschlossen",
                    "Der OCR-Prozess ist abgeschlossen.\nMöchten Sie die Trefferliste öffnen?",
                    QMessageBox.Yes | QMessageBox.No,
                    QMessageBox.Yes
                )
                if antwort == QMessageBox.Yes:
                    os.startfile(treffer_datei)

        except Exception as e:
            logging.exception("Fehler beim Beenden des OCR-Prozesses")
            QMessageBox.critical(self, "Fehler", f"Fehler beim Beenden des OCR-Prozesses: {e}")

if __name__ == "__main__":
    try:
        from PySide6.QtWidgets import QApplication
        app = QApplication(sys.argv)
        window = OCRApp()
        window.show()
        sys.exit(app.exec())
    except Exception as e:
        logging.exception("Unerwarteter Fehler in der GUI")
