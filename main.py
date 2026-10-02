# -*- coding: utf-8 -*-
"""مدير ملفات المعلم - Teacher File Manager (PyQt6)."""

import json
import os
import shutil
import subprocess
import sys
import tempfile

from PyQt6.QtCore import QDate, QFile, QSettings, Qt, QThread, QTimer, QUrl, pyqtSignal
from PyQt6.QtGui import QDesktopServices, QFont, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
    QDateEdit,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QFileIconProvider,
    QFormLayout,
    QHBoxLayout,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMenu,
    QMessageBox,
    QProgressDialog,
    QPlainTextEdit,
    QPushButton,
    QScrollArea,
    QSpinBox,
    QSplitter,
    QStackedWidget,
    QTabWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

from teacher_core import (  # المنطق المشترك مع نسخة الأندرويد
    auto_file_name,
    BOOK_FILE,
    BOOK_KEY,
    build_generator_values,
    DEFAULT_KEYWORDS,
    docx,
    extract_for_generator,
    extraction_notes,
    fill_docx,
    fitz,
    fmt_date,
    GUIDE_FILE,
    is_hidden,
    KEYWORDS_FILE,
    list_entries,
    load_keywords,
    MemoError,
    natural_key,
    normalize,
    pdf_page_count,
    placeholder_key,
    section_in_template,
    subfolders,
    TEMPLATE_NAME,
    template_placeholders,
    TEXT_MODES,
    validate_name,
)

APP_NAME = "مدير ملفات المعلم"
MAX_SEARCH_RESULTS = 1000
PATH_ROLE = Qt.ItemDataRole.UserRole
LOADED_ROLE = Qt.ItemDataRole.UserRole + 1



def open_path(path):
    """فتح ملف/مجلد بالبرنامج الافتراضي في النظام."""
    if hasattr(os, "startfile"):
        os.startfile(os.path.normpath(path))  # Windows
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(path))


def reveal_in_explorer(path):
    path = os.path.normpath(path)
    if sys.platform.startswith("win"):
        subprocess.Popen(["explorer", f"/select,{path}"])
    else:
        QDesktopServices.openUrl(QUrl.fromLocalFile(os.path.dirname(path)))


def find_template(root):
    folders = [root]
    if getattr(sys, "frozen", False):
        folders.append(os.path.dirname(sys.executable))  # بجنب ملف exe
        folders.append(getattr(sys, "_MEIPASS", ""))  # النموذج المدمج داخل exe
    else:
        folders.append(os.path.dirname(os.path.abspath(__file__)))
    for folder in folders:
        if not folder:
            continue
        candidate = os.path.join(folder, TEMPLATE_NAME)
        if os.path.isfile(candidate):
            return candidate
    return None


# ---------- تحويل docx إلى PDF ----------
class ConversionError(Exception):
    """فشل تحويل ملف معيّن."""


class NoConverterError(ConversionError):
    """ما فما حتى أداة تحويل تخدم (لا Word ولا LibreOffice)."""


NO_CONVERTER_MESSAGE = (
    "ما نجمتش نحوّل الملف لأنو ما لقيت برنامج يعمل التحويل.\n\n"
    "لازم تثبّت واحد من هاذم:\n"
    "• Microsoft Word (الأفضل، إذا عندك Office)\n"
    "• LibreOffice (مجاني): https://www.libreoffice.org/download\n\n"
    "بعد التثبيت عاود حاول."
)


def find_libreoffice():
    for name in ("soffice", "soffice.exe", "libreoffice"):
        found = shutil.which(name)
        if found:
            return found
    for base in (os.environ.get("ProgramFiles"), os.environ.get("ProgramFiles(x86)")):
        if base:
            candidate = os.path.join(base, "LibreOffice", "program", "soffice.exe")
            if os.path.isfile(candidate):
                return candidate
    return None


def convert_with_word(src, dst):
    import docx2pdf  # يخدم كان Microsoft Word مثبّت

    com = None
    try:  # التحويل يتم في thread، وCOM يحتاج تهيئة في كل thread
        import pythoncom

        pythoncom.CoInitialize()
        com = pythoncom
    except ImportError:
        pass
    try:
        docx2pdf.convert(src, dst)
    finally:
        if com is not None:
            com.CoUninitialize()


def convert_with_libreoffice(src, dst, soffice):
    out_dir = os.path.dirname(dst)
    with tempfile.TemporaryDirectory() as tmp:
        profile = QUrl.fromLocalFile(os.path.join(tmp, "profile")).toString()  # باش ما يتعارضش مع LibreOffice مفتوح
        cmd = [soffice, f"-env:UserInstallation={profile}", "--headless", "--convert-to", "pdf", "--outdir", tmp, src]
        flags = getattr(subprocess, "CREATE_NO_WINDOW", 0)
        result = subprocess.run(cmd, capture_output=True, timeout=300, creationflags=flags)
        produced = os.path.join(tmp, os.path.splitext(os.path.basename(src))[0] + ".pdf")
        if result.returncode != 0 or not os.path.isfile(produced):
            details = result.stderr.decode("utf-8", "replace").strip()
            raise ConversionError(f"LibreOffice فشل في التحويل. {details}".strip())
        shutil.move(produced, dst)
    return out_dir


class PdfConverter:
    """يجرّب Word (docx2pdf) ثم LibreOffice، ويتذكر الطريقة اللي نجحت."""

    def __init__(self):
        self.order = ["word", "libreoffice"]

    def convert(self, src, dst):
        errors = []
        for backend in list(self.order):
            try:
                if backend == "word":
                    convert_with_word(src, dst)
                else:
                    soffice = find_libreoffice()
                    if not soffice:
                        errors.append("LibreOffice: غير مثبّت")
                        continue
                    convert_with_libreoffice(src, dst, soffice)
                if not os.path.isfile(dst):
                    raise ConversionError("ما تكوّن ملف PDF.")
            except Exception as e:  # noqa: BLE001 - كل فشل يجرّب الطريقة الموالية
                label = "Word" if backend == "word" else "LibreOffice"
                errors.append(f"{label}: {type(e).__name__}: {e}".strip())
                continue
            self.order.remove(backend)
            self.order.insert(0, backend)
            return
        if not find_libreoffice():
            error = NoConverterError(NO_CONVERTER_MESSAGE)
        else:
            error = ConversionError("فشل التحويل بكل الطرق المتاحة.")
        error.details = "\n".join(errors)
        raise error


class PdfWorker(QThread):
    """يحوّل قائمة ملفات docx في thread منفصل."""

    progress = pyqtSignal(int, int, str)  # المنجز، المجموع، اسم الملف الحالي
    finished_all = pyqtSignal(list, list, str)  # [pdf], [(ملف, خطأ)], تفاصيل «ما فما أداة تحويل»

    def __init__(self, paths):
        super().__init__()
        self.paths = paths

    def run(self):
        converter = PdfConverter()
        done, failed, no_converter = [], [], ""
        for i, src in enumerate(self.paths):
            if self.isInterruptionRequested():
                break
            self.progress.emit(i, len(self.paths), os.path.basename(src))
            dst = os.path.splitext(src)[0] + ".pdf"
            try:
                converter.convert(src, dst)
                done.append(dst)
            except NoConverterError as e:
                no_converter = e.details or "—"
                break
            except Exception as e:  # noqa: BLE001
                details = getattr(e, "details", "")
                failed.append((src, f"{e}\n{details}".strip()))
        self.finished_all.emit(done, failed, no_converter)


class NewMemoDialog(QDialog):
    def __init__(self, root, initial_section=None, initial_subject=None, parent=None):
        super().__init__(parent)
        self.root = root
        self.name_manual = False
        self.setWindowTitle("مذكرة جديدة")
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setMinimumWidth(420)

        self.section_box = QComboBox()
        self.subject_box = QComboBox()
        self.lesson_edit = QLineEdit()
        self.lesson_edit.setPlaceholderText("مثال: الجمع والطرح")
        self.week_spin = QSpinBox()
        self.week_spin.setRange(1, 60)
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("يتكوّن أوتوماتيك، وتنجم تبدلو")

        form = QFormLayout()
        form.addRow("القسم:", self.section_box)
        form.addRow("المادة:", self.subject_box)
        form.addRow("عنوان الدرس:", self.lesson_edit)
        form.addRow("رقم الأسبوع:", self.week_spin)
        form.addRow("التاريخ:", self.date_edit)
        form.addRow("اسم الملف:", self.name_edit)

        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel
        )
        buttons.button(QDialogButtonBox.StandardButton.Ok).setText("إنشاء")
        buttons.button(QDialogButtonBox.StandardButton.Cancel).setText("إلغاء")
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)

        layout = QVBoxLayout(self)
        layout.addLayout(form)
        layout.addWidget(buttons)

        self.section_box.addItems(subfolders(root))
        self.section_box.currentTextChanged.connect(self.load_subjects)
        if initial_section:
            self.section_box.setCurrentText(initial_section)
        self.load_subjects(self.section_box.currentText())
        if initial_subject:
            self.subject_box.setCurrentText(initial_subject)

        self.subject_box.currentTextChanged.connect(self.update_auto_name)
        self.lesson_edit.textChanged.connect(self.update_auto_name)
        self.week_spin.valueChanged.connect(self.update_auto_name)
        self.name_edit.textEdited.connect(lambda text: setattr(self, "name_manual", bool(text.strip())))
        self.update_auto_name()
        self.lesson_edit.setFocus()

    def load_subjects(self, section):
        self.subject_box.clear()
        if section:
            self.subject_box.addItems(subfolders(os.path.join(self.root, section)))

    def update_auto_name(self):
        if not self.name_manual:
            self.name_edit.setText(
                auto_file_name(self.subject_box.currentText(), self.lesson_edit.text(), self.week_spin.value())
            )

    def accept(self):
        section, subject = self.section_box.currentText(), self.subject_box.currentText()
        if not section:
            QMessageBox.warning(self, APP_NAME, "ما فما حتى قسم في المجلد الرئيسي. أنشئ مجلد قسم أولا.")
            return
        if not subject:
            QMessageBox.warning(self, APP_NAME, "هذا القسم ما فيه حتى مادة. أنشئ مجلد مادة داخلو أولا.")
            return
        name = self.name_edit.text().strip()
        error = validate_name(name)
        if error:
            QMessageBox.warning(self, APP_NAME, error)
            return
        if not name.lower().endswith(".docx"):
            name += ".docx"
        self.target = os.path.join(self.root, section, subject, name)
        if os.path.exists(self.target):
            QMessageBox.warning(self, APP_NAME, "يوجد ملف بنفس الاسم في هذه المادة.")
            return
        d = self.date_edit.date()
        self.values = {
            "القسم": section,
            "المادة": subject,
            "الدرس": self.lesson_edit.text().strip(),
            "الأسبوع": str(self.week_spin.value()),
            "التاريخ": f"{d.day():02d}/{d.month():02d}/{d.year()}",
        }
        super().accept()


def ask_yes_no(parent, title, text, yes_text, no_text, icon=QMessageBox.Icon.Question):
    box = QMessageBox(parent)
    box.setIcon(icon)
    box.setWindowTitle(title)
    box.setText(text)
    box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    yes = box.addButton(yes_text, QMessageBox.ButtonRole.AcceptRole)
    box.addButton(no_text, QMessageBox.ButtonRole.RejectRole)
    box.setDefaultButton(yes)
    box.exec()
    return box.clickedButton() is yes


def rtl_text_edit(read_only=False):
    edit = QPlainTextEdit()
    edit.setReadOnly(read_only)
    edit.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    option = edit.document().defaultTextOption()
    option.setTextDirection(Qt.LayoutDirection.RightToLeft)
    edit.document().setDefaultTextOption(option)
    return edit


class MemoGeneratorDialog(QDialog):
    """مولّد المذكرات (بلا إنترنت): يستخرج نص الدليل والكتاب ويعمّر نموذج Word."""

    def __init__(self, settings, parent=None):
        super().__init__(parent)
        self.settings = settings
        self.library = ""
        self.subject = ""
        self.keywords = {}
        self.page_counts = {}
        self.fields = {}  # اسم الخانة -> (QPlainTextEdit, QLabel تحذير)
        self.created_path = None
        self.setWindowTitle("مولّد المذكرات (بلا إنترنت)")
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.resize(1150, 720)

        self.stack = QStackedWidget()
        self.stack.addWidget(self.build_setup_page())
        self.stack.addWidget(self.build_review_page())
        layout = QVBoxLayout(self)
        layout.addWidget(self.stack)

        saved = settings.value("library_folder", "", type=str)
        if saved and os.path.isdir(saved):
            self.set_library(saved, remember=False)

    # ---------- الصفحة 1: الإعدادات ----------
    def build_setup_page(self):
        page = QWidget()
        self.lib_btn = QPushButton("اختيار مجلد المكتبة")
        self.lib_btn.clicked.connect(self.choose_library)
        self.lib_label = QLabel("لم يتم اختيار مجلد المكتبة")
        self.lib_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.subject_list = QListWidget()
        self.subject_list.currentItemChanged.connect(self.on_subject_changed)

        self.lesson_edit = QLineEdit()
        self.lesson_edit.setPlaceholderText("مثال: الجمع والطرح")
        self.week_spin = QSpinBox()
        self.week_spin.setRange(1, 60)
        self.date_edit = QDateEdit(QDate.currentDate())
        self.date_edit.setCalendarPopup(True)
        self.date_edit.setDisplayFormat("dd/MM/yyyy")
        self.guide_from, self.guide_to = QSpinBox(), QSpinBox()
        self.book_from, self.book_to = QSpinBox(), QSpinBox()
        for spin in (self.guide_from, self.guide_to, self.book_from, self.book_to):
            spin.setRange(1, 1)
        self.guide_from.valueChanged.connect(lambda v: self.guide_to.setValue(max(v, self.guide_to.value())))
        self.book_from.valueChanged.connect(lambda v: self.book_to.setValue(max(v, self.book_to.value())))
        self.mode_box = QComboBox()
        for label, value in TEXT_MODES:
            self.mode_box.addItem(label, value)
        self.info_label = QLabel("")
        self.info_label.setWordWrap(True)

        def pair(a, b):
            row = QHBoxLayout()
            row.addWidget(QLabel("من"))
            row.addWidget(a)
            row.addWidget(QLabel("إلى"))
            row.addWidget(b)
            row.addStretch(1)
            return row

        form = QFormLayout()
        form.addRow("عنوان الدرس:", self.lesson_edit)
        form.addRow("رقم الأسبوع:", self.week_spin)
        form.addRow("التاريخ:", self.date_edit)
        form.addRow("صفحات الدرس في الدليل:", pair(self.guide_from, self.guide_to))
        form.addRow("صفحات الدرس في الكتاب:", pair(self.book_from, self.book_to))
        form.addRow("معالجة اتجاه النص:", self.mode_box)

        kw_btn = QPushButton("تعديل الكلمات المفتاحية")
        kw_btn.clicked.connect(self.edit_keywords)
        kw_reset_btn = QPushButton("الكلمات الافتراضية")
        kw_reset_btn.setToolTip("ترجّع keywords.json متاع المادة للقيم الافتراضية (مراحل الدليل والكتاب)")
        kw_reset_btn.clicked.connect(self.reset_keywords)
        analyze_btn = QPushButton("تحليل ومراجعة")
        analyze_btn.setDefault(True)
        analyze_btn.clicked.connect(self.analyze)
        close_btn = QPushButton("إغلاق")
        close_btn.clicked.connect(self.reject)
        buttons = QHBoxLayout()
        buttons.addWidget(kw_btn)
        buttons.addWidget(kw_reset_btn)
        buttons.addStretch(1)
        buttons.addWidget(analyze_btn)
        buttons.addWidget(close_btn)

        left = QVBoxLayout()
        left.addLayout(form)
        left.addWidget(self.info_label)
        left.addStretch(1)
        subjects = QVBoxLayout()
        subjects.addWidget(QLabel("المواد:"))
        subjects.addWidget(self.subject_list, 1)
        body = QHBoxLayout()
        body.addLayout(subjects, 1)
        body.addLayout(left, 2)

        top = QHBoxLayout()
        top.addWidget(self.lib_btn)
        top.addWidget(self.lib_label, 1)
        root = QVBoxLayout(page)
        root.addLayout(top)
        root.addLayout(body, 1)
        root.addLayout(buttons)
        return page

    def choose_library(self):
        folder = QFileDialog.getExistingDirectory(self, "اختيار مجلد المكتبة", self.library or os.path.expanduser("~"))
        if folder:
            self.set_library(folder)

    def set_library(self, folder, remember=True):
        self.library = os.path.normpath(folder)
        if remember:
            self.settings.setValue("library_folder", self.library)
        self.lib_label.setText(f"المكتبة: {self.library}")
        self.subject_list.clear()
        for name in subfolders(self.library):
            path = os.path.join(self.library, name)
            missing = [f for f in (GUIDE_FILE, BOOK_FILE, TEMPLATE_NAME) if not os.path.isfile(os.path.join(path, f))]
            item = QListWidgetItem(name + (f"   (ناقص: {'، '.join(missing)})" if missing else ""))
            item.setData(PATH_ROLE, name)
            self.subject_list.addItem(item)
        if self.subject_list.count() == 0:
            self.info_label.setText("ما لقيت حتى مجلد مادة داخل المكتبة. كل مادة لازم تكون مجلد فيه guide.pdf وbook.pdf وtemplate.docx.")
        else:
            self.info_label.setText("")

    def subject_dir(self):
        return os.path.join(self.library, self.subject)

    def on_subject_changed(self, item, _previous=None):
        self.subject = item.data(PATH_ROLE) if item else ""
        self.page_counts = {}
        notes = []
        for key, filename, spins in (
            ("guide", GUIDE_FILE, (self.guide_from, self.guide_to)),
            ("book", BOOK_FILE, (self.book_from, self.book_to)),
        ):
            path = os.path.join(self.subject_dir(), filename) if self.subject else ""
            try:
                count = pdf_page_count(path) if path else 0
            except MemoError as e:
                count = 0
                notes.append(str(e))
            self.page_counts[key] = count
            for spin in spins:
                spin.setEnabled(count > 0)
                spin.setRange(1, max(count, 1))
            if count:
                notes.append(f"{filename}: {count} صفحة")
        self.info_label.setText("\n".join(notes))

    def edit_keywords(self):
        if not self.subject:
            QMessageBox.information(self, APP_NAME, "اختر مادة أولا.")
            return
        try:
            load_keywords(self.subject_dir())  # يكوّن الملف بالقيم الافتراضية إذا ما كانش موجود
            open_path(os.path.join(self.subject_dir(), KEYWORDS_FILE))
        except (MemoError, OSError) as e:
            QMessageBox.warning(self, APP_NAME, str(e))
            return
        QMessageBox.information(
            self, APP_NAME,
            f"تم فتح {KEYWORDS_FILE} في محرر النصوص.\nعدّل الكلمات واحفظ الملف، ثم اضغط «تحليل ومراجعة» من جديد.\n"
            "كل قسم فيه الكلمات اللي تدل على عنوانو في «الدليل» وفي «الكتاب».\n"
            "اسم القسم هو الكلمة المعلّمة في النموذج، مثلا {{الوضعية المشكل}}.",
        )

    def reset_keywords(self):
        if not self.subject:
            QMessageBox.information(self, APP_NAME, "اختر مادة أولا.")
            return
        if not ask_yes_no(
            self, "الكلمات الافتراضية",
            f"نرجّع {KEYWORDS_FILE} متاع «{self.subject}» للقيم الافتراضية؟\nالتعديلات اللي عملتها فيه تتمسح.",
            "نعم، رجّعها", "لا", QMessageBox.Icon.Warning,
        ):
            return
        try:
            with open(os.path.join(self.subject_dir(), KEYWORDS_FILE), "w", encoding="utf-8") as f:
                json.dump(DEFAULT_KEYWORDS, f, ensure_ascii=False, indent=2)
        except OSError as e:
            QMessageBox.warning(self, APP_NAME, f"ما نجمتش نكتب {KEYWORDS_FILE}:\n{e}")
            return
        QMessageBox.information(self, APP_NAME, "تم. اضغط «تحليل ومراجعة» من جديد.")

    # ---------- الصفحة 2: المراجعة ----------
    def build_review_page(self):
        page = QWidget()
        self.review_info = QLabel("")
        self.review_info.setWordWrap(True)
        self.fields_layout = QVBoxLayout()
        self.fields_layout.addStretch(1)
        holder = QWidget()
        holder.setLayout(self.fields_layout)
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(holder)

        self.guide_view = rtl_text_edit(read_only=True)
        self.book_view = rtl_text_edit(read_only=True)
        tabs = QTabWidget()
        tabs.addTab(self.guide_view, "النص الأصلي: الدليل")
        tabs.addTab(self.book_view, "النص الأصلي: الكتاب")

        splitter = QSplitter(Qt.Orientation.Horizontal)
        splitter.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        splitter.addWidget(scroll)  # في RTL أول عنصر يكون على اليمين
        splitter.addWidget(tabs)
        splitter.setSizes([600, 500])

        back_btn = QPushButton("رجوع")
        back_btn.clicked.connect(lambda: self.stack.setCurrentIndex(0))
        save_btn = QPushButton("حفظ")
        save_btn.setDefault(True)
        save_btn.clicked.connect(self.save)
        buttons = QHBoxLayout()
        buttons.addWidget(back_btn)
        buttons.addStretch(1)
        buttons.addWidget(save_btn)

        root = QVBoxLayout(page)
        root.addWidget(self.review_info)
        root.addWidget(splitter, 1)
        root.addLayout(buttons)
        return page

    def set_field_state(self, edit, warning):
        empty = not edit.toPlainText().strip()
        warning.setVisible(empty)
        edit.setStyleSheet("QPlainTextEdit { background: #fff1c2; }" if empty else "")

    def rebuild_fields(self, values, found_keys):
        while self.fields_layout.count() > 1:
            item = self.fields_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        self.fields = {}
        for name, text in values.items():
            label = QLabel(f"<b>{name}</b>")
            edit = rtl_text_edit()
            edit.setMinimumHeight(110)
            edit.setPlainText(text)
            if name == BOOK_KEY:
                hint = "ما فما نص للكتاب. عمّرو بيدك أو ارجع وصحّح الصفحات."
            else:
                hint = f"ما لقيتش «{name}» لا في الدليل لا في الكتاب. عمّرو بيدك (أو عدّل الكلمات المفتاحية وعاود التحليل)."
            warning = QLabel("⚠ " + hint)
            warning.setStyleSheet("color: #9a6700;")
            warning.setWordWrap(True)
            edit.textChanged.connect(lambda e=edit, w=warning: self.set_field_state(e, w))
            self.set_field_state(edit, warning)
            index = self.fields_layout.count() - 1
            for widget in (label, edit, warning):
                self.fields_layout.insertWidget(index, widget)
                index += 1
            self.fields[name] = (edit, warning)

    # ---------- التحليل ----------
    def read_range(self, first, last, key, label):
        a, b = first.value(), last.value()
        if a > b:
            raise MemoError(f"صفحات {label}: رقم البداية ({a}) أكبر من رقم النهاية ({b}).")
        if b > self.page_counts.get(key, 0):
            raise MemoError(f"صفحات {label}: الملف فيه {self.page_counts.get(key, 0)} صفحة برك.")
        return a, b

    def analyze(self):
        try:
            self.run_analysis()
        except MemoError as e:
            QApplication.restoreOverrideCursor()
            QMessageBox.warning(self, APP_NAME, str(e))

    def run_analysis(self):
        if not self.library:
            raise MemoError("اختر مجلد المكتبة أولا.")
        if not self.subject:
            raise MemoError("اختر المادة من القائمة.")
        if not self.lesson_edit.text().strip():
            raise MemoError("اكتب عنوان الدرس.")
        error = validate_name(auto_file_name(self.subject, self.lesson_edit.text(), self.week_spin.value()))
        if error:
            raise MemoError(f"اسم المذكرة غير صالح: {error}")
        folder = self.subject_dir()
        for filename in (GUIDE_FILE, TEMPLATE_NAME):
            if not os.path.isfile(os.path.join(folder, filename)):
                raise MemoError(f"ما لقيتش {filename} في مجلد المادة «{self.subject}».")
        if not self.page_counts.get("guide"):
            raise MemoError(f"ما نجمتش نقرا {GUIDE_FILE}. تأكد أنو سليم وفيه صفحات.")
        guide_range = self.read_range(self.guide_from, self.guide_to, "guide", "الدليل")
        has_book = bool(self.page_counts.get("book"))
        book_range = self.read_range(self.book_from, self.book_to, "book", "الكتاب") if has_book else None

        self.keywords = load_keywords(folder)
        mode = self.mode_box.currentData()

        QApplication.setOverrideCursor(Qt.CursorShape.WaitCursor)
        try:
            guide = extract_for_generator(os.path.join(folder, GUIDE_FILE), *guide_range, mode, self.keywords, "guide")
            book = extract_for_generator(os.path.join(folder, BOOK_FILE), *book_range, mode, self.keywords, "book") if has_book else None
        finally:
            QApplication.restoreOverrideCursor()
        guide_lines, book_lines = guide[0], (book[0] if book else [])

        values, self.extra_values = build_generator_values(guide_lines, book_lines, self.keywords)
        self.guide_view.setPlainText("\n".join(guide_lines))
        self.book_view.setPlainText("\n".join(book_lines))
        self.rebuild_fields(values, [k for k, v in values.items() if v])

        missing = [k for k, v in values.items() if k != BOOK_KEY and not v]
        note = f"المادة: {self.subject} — الدرس: {self.lesson_edit.text().strip()} — أسبوع {self.week_spin.value()}."
        if missing:
            note += f"\nأقسام ما تلقاتش لا في الدليل لا في الكتاب (معلّمة بالأصفر): {'، '.join(missing)}."
        if not has_book:
            note += f"\nما فما {BOOK_FILE} في مجلد المادة، خانة الكتاب فارغة."
        for extra in extraction_notes(guide, book, guide_range, book_range):
            note += "\n" + extra
        self.review_info.setText(note)
        self.stack.setCurrentIndex(1)

    # ---------- الحفظ ----------
    def save(self):
        folder = self.subject_dir()
        template = os.path.join(folder, TEMPLATE_NAME)
        name = auto_file_name(self.subject, self.lesson_edit.text(), self.week_spin.value()) + ".docx"
        target = os.path.join(folder, name)
        if os.path.exists(target) and not ask_yes_no(
            self, "الملف موجود", f"يوجد ملف بنفس الاسم:\n{name}\nتحب تستبدلو؟", "نعم، استبدل", "لا", QMessageBox.Icon.Warning
        ):
            return
        values = {
            "المادة": self.subject,
            "الدرس": self.lesson_edit.text().strip(),
            "الأسبوع": str(self.week_spin.value()),
            "التاريخ": fmt_date(self.date_edit.date()),
        }
        values.update(getattr(self, "extra_values", {}))  # {{القسم - الدليل}} و{{القسم - الكتاب}}
        for field, (edit, _warning) in self.fields.items():
            values[field] = edit.toPlainText().strip()
        try:
            present = template_placeholders(template)
            filled = fill_docx(template, target, values)
        except MemoError as e:
            QMessageBox.warning(self, APP_NAME, str(e))
            return
        except Exception as e:  # noqa: BLE001 - نموذج تالف، ملف مفتوح في Word...
            QMessageBox.critical(self, APP_NAME, f"ما نجمتش نحفظ المذكرة (ربما الملف مفتوح في Word):\n{e}")
            return
        self.created_path = target

        known = {placeholder_key(k) for k in values}
        warnings = []
        unknown = [orig for key, orig in present.items() if key not in known]
        if unknown:
            warnings.append("كلمات معلّمة في النموذج ما عندها خانة (بقات كيما هي): " + "، ".join("{{%s}}" % u for u in unknown))
        unused = [f for f, v in values.items()
                  if f in self.fields and f != BOOK_KEY and v and not section_in_template(f, present) and f not in filled]
        if unused and present:  # نموذج بلا {{...}} (يتعمّر بالعناوين): ما نكثروش عليه ملاحظات
            warnings.append("أقسام فيها نص لكن ما فماش كلمة معلّمة ليها في النموذج (ما تدخلتش للمذكرة): " + "، ".join("{{%s}}" % u for u in unused))
        if getattr(filled, "empty", None):
            warnings.append("خانات في نموذجك بقات فارغة خاطر ما لقيتلهمش نص في الدليل ولا الكتاب: "
                            + "، ".join(filled.empty) + ". ثبّت أرقام الصفحات، ولا اكتب النص في خانتو في شاشة المراجعة.")
        if warnings:
            QMessageBox.warning(self, APP_NAME, "تم حفظ المذكرة، لكن:\n\n" + "\n\n".join(warnings))
        if ask_yes_no(self, "تم الحفظ", f"تم حفظ المذكرة:\n{name}\nتحب تفتحها توا؟", "نعم، افتحها", "لا"):
            try:
                open_path(target)
            except OSError as e:
                QMessageBox.warning(self, APP_NAME, f"ما نجمتش نفتح الملف:\n{e}")
        self.accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(900, 650)
        self.settings = QSettings("TeacherFileManager", "TeacherFileManager")
        self.root = ""
        self.items = {}  # مسار -> عنصر الشجرة
        self.icons = QFileIconProvider()
        self.pdf_worker = None
        self.pdf_dialog = None

        self.build_ui()
        self.build_shortcuts()

        saved = self.settings.value("root_folder", "", type=str)
        if saved and os.path.isdir(saved):
            self.set_root(saved, remember=False)
        else:
            self.statusBar().showMessage("اختر المجلد الرئيسي للبدء.")

    # ---------- الواجهة ----------
    def build_ui(self):
        self.choose_btn = QPushButton("اختيار المجلد الرئيسي")
        self.choose_btn.clicked.connect(self.choose_root)
        self.new_btn = QPushButton("مذكرة جديدة")
        self.new_btn.clicked.connect(self.new_memo)
        self.generator_btn = QPushButton("مولّد المذكرات")
        self.generator_btn.clicked.connect(self.open_generator)
        self.refresh_btn = QPushButton("تحديث")
        self.refresh_btn.clicked.connect(self.refresh_all)

        self.search_edit = QLineEdit()
        self.search_edit.setPlaceholderText("بحث بالاسم في كل المجلدات...")
        self.search_edit.setClearButtonEnabled(True)
        self.search_edit.textChanged.connect(lambda: self.search_timer.start())
        self.search_timer = QTimer(self)
        self.search_timer.setSingleShot(True)
        self.search_timer.setInterval(250)
        self.search_timer.timeout.connect(self.run_search)

        top = QHBoxLayout()
        top.addWidget(self.choose_btn)
        top.addWidget(self.new_btn)
        top.addWidget(self.generator_btn)
        top.addWidget(self.search_edit, 1)
        top.addWidget(self.refresh_btn)

        self.root_label = QLabel("لم يتم اختيار مجلد رئيسي")
        self.root_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("الأقسام ← المواد ← الملفات")
        self.tree.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.tree.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.tree_menu)
        self.tree.itemExpanded.connect(self.on_expand)
        self.tree.itemDoubleClicked.connect(lambda item, _c: self.open_if_file(item.data(0, PATH_ROLE)))

        self.results = QListWidget()
        self.results.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.results.setSelectionMode(QAbstractItemView.SelectionMode.ExtendedSelection)
        self.results.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.results.customContextMenuRequested.connect(self.results_menu)
        self.results.itemDoubleClicked.connect(lambda item: self.open_if_file(item.data(PATH_ROLE)))

        self.stack = QStackedWidget()
        self.stack.addWidget(self.tree)
        self.stack.addWidget(self.results)

        central = QWidget()
        layout = QVBoxLayout(central)
        layout.addLayout(top)
        layout.addWidget(self.root_label)
        layout.addWidget(self.stack, 1)
        self.setCentralWidget(central)

    def build_shortcuts(self):
        ctx = Qt.ShortcutContext.WidgetShortcut
        for widget in (self.tree, self.results):
            QShortcut(QKeySequence("F2"), widget, activated=self.rename_current, context=ctx)
            QShortcut(QKeySequence("Delete"), widget, activated=self.delete_current, context=ctx)
            QShortcut(QKeySequence("Return"), widget, activated=self.open_current, context=ctx)
        QShortcut(QKeySequence("F5"), self, activated=self.refresh_all)
        QShortcut(QKeySequence("Ctrl+F"), self, activated=self.search_edit.setFocus)
        QShortcut(QKeySequence("Ctrl+N"), self, activated=self.new_memo)

    # ---------- المجلد الرئيسي ----------
    def choose_root(self):
        folder = QFileDialog.getExistingDirectory(self, "اختيار المجلد الرئيسي", self.root or os.path.expanduser("~"))
        if folder:
            self.set_root(folder)

    def set_root(self, folder, remember=True):
        self.root = os.path.normpath(folder)
        if remember:
            self.settings.setValue("root_folder", self.root)
        self.root_label.setText(f"المجلد الرئيسي: {self.root}")
        self.search_edit.clear()
        self.refresh_all()

    # ---------- الشجرة ----------
    def make_item(self, path, is_dir):
        item = QTreeWidgetItem([os.path.basename(path)])
        item.setData(0, PATH_ROLE, path)
        item.setIcon(0, self.icons.icon(QFileIconProvider.IconType.Folder if is_dir else QFileIconProvider.IconType.File))
        self.items[path] = item
        if is_dir:
            item.setData(0, LOADED_ROLE, False)
            item.addChild(QTreeWidgetItem())  # عنصر وهمي ليظهر السهم
        return item

    def populate(self, parent_item, path):
        """parent_item = None يعني الجذر."""
        if parent_item is not None:
            parent_item.takeChildren()
            parent_item.setData(0, LOADED_ROLE, True)
        dirs, files = list_entries(path)
        entries = [(d, True) for d in dirs] + [(f, False) for f in files]
        nodes = [self.make_item(os.path.join(path, name), is_dir) for name, is_dir in entries]
        if parent_item is None:
            self.tree.addTopLevelItems(nodes)
        else:
            parent_item.addChildren(nodes)

    def on_expand(self, item):
        if not item.data(0, LOADED_ROLE):
            self.populate(item, item.data(0, PATH_ROLE))

    def refresh_tree(self):
        expanded = sorted((p for p, it in self.items.items() if it.isExpanded()), key=len)
        current = self.current_tree_path()
        self.tree.clear()
        self.items = {}
        if not self.root or not os.path.isdir(self.root):
            return
        self.populate(None, self.root)
        for p in expanded:
            item = self.items.get(p)
            if item is not None:
                item.setExpanded(True)
        if current:
            self.select_path(current)

    def select_path(self, path):
        path = os.path.normpath(path)
        rel = os.path.relpath(path, self.root)
        if rel.startswith(".."):
            return
        cur = self.root
        for part in rel.split(os.sep)[:-1]:
            cur = os.path.join(cur, part)
            item = self.items.get(cur)
            if item is None:
                return
            item.setExpanded(True)
        item = self.items.get(path)
        if item is not None:
            self.tree.setCurrentItem(item)
            self.tree.scrollToItem(item)

    def current_tree_path(self):
        item = self.tree.currentItem()
        return item.data(0, PATH_ROLE) if item else None

    def refresh_all(self):
        if self.root and not os.path.isdir(self.root):
            QMessageBox.warning(self, APP_NAME, "المجلد الرئيسي ما عادش موجود. اختر مجلد آخر.")
            self.root = ""
            self.root_label.setText("لم يتم اختيار مجلد رئيسي")
        self.refresh_tree()
        if self.search_edit.text().strip():
            self.run_search()
        elif self.root:
            self.statusBar().showMessage("تم التحديث.", 3000)

    # ---------- البحث ----------
    def run_search(self):
        query = normalize(self.search_edit.text().strip())
        self.results.clear()
        if not query:
            self.stack.setCurrentWidget(self.tree)
            return
        self.stack.setCurrentWidget(self.results)
        if not self.root:
            self.statusBar().showMessage("اختر المجلد الرئيسي أولا.")
            return
        count, truncated = 0, False
        for folder, dirs, files in os.walk(self.root):
            dirs[:] = sorted((d for d in dirs if not is_hidden(d)), key=natural_key)
            names = [(d, True) for d in dirs] + [(f, False) for f in sorted(files, key=natural_key) if not is_hidden(f)]
            for name, is_dir in names:
                if query not in normalize(name):
                    continue
                if count >= MAX_SEARCH_RESULTS:
                    truncated = True
                    break
                path = os.path.join(folder, name)
                where = os.path.relpath(folder, self.root)
                where = "" if where == "." else "   —   " + where.replace(os.sep, " ‹ ")
                item = QListWidgetItem(self.icons.icon(QFileIconProvider.IconType.Folder if is_dir else QFileIconProvider.IconType.File), name + where)
                item.setData(PATH_ROLE, path)
                item.setToolTip(path)
                self.results.addItem(item)
                count += 1
            if truncated:
                break
        if truncated:
            self.statusBar().showMessage(f"أكثر من {MAX_SEARCH_RESULTS} نتيجة، تم عرض الأولى فقط. دقّق البحث.")
        else:
            self.statusBar().showMessage(f"عدد النتائج: {count}")

    # ---------- العمليات على العناصر ----------
    def current_path(self):
        if self.stack.currentWidget() is self.results:
            item = self.results.currentItem()
            return item.data(PATH_ROLE) if item else None
        return self.current_tree_path()

    def open_if_file(self, path):
        if path and os.path.isfile(path):
            self.open_item(path)

    def open_current(self):
        path = self.current_path()
        if path:
            self.open_item(path)

    def open_item(self, path):
        if not os.path.exists(path):
            QMessageBox.warning(self, APP_NAME, "الملف ما عادش موجود. يتم التحديث.")
            self.refresh_all()
            return
        try:
            open_path(path)
        except OSError as e:
            QMessageBox.critical(self, APP_NAME, f"ما نجمتش نفتح الملف:\n{e}")

    def rename_current(self):
        path = self.current_path()
        if path:
            self.rename_item(path)

    def delete_current(self):
        path = self.current_path()
        if path:
            self.delete_item(path)

    def rename_item(self, path):
        is_file = os.path.isfile(path)
        base = os.path.basename(path)
        stem, ext = os.path.splitext(base) if is_file else (base, "")
        dlg = QInputDialog(self)
        dlg.setWindowTitle("إعادة تسمية")
        dlg.setLabelText("الاسم الجديد:")
        dlg.setTextValue(stem)
        dlg.setOkButtonText("موافق")
        dlg.setCancelButtonText("إلغاء")
        dlg.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        if not dlg.exec():
            return
        new_stem = dlg.textValue().strip()
        error = validate_name(new_stem)
        if error:
            QMessageBox.warning(self, APP_NAME, error)
            return
        new_path = os.path.join(os.path.dirname(path), new_stem + ext)
        if new_path == path:
            return
        # على ويندوز الأسماء ما تفرقش بين الحروف الكبيرة والصغيرة
        if os.path.exists(new_path) and os.path.normcase(new_path) != os.path.normcase(path):
            QMessageBox.warning(self, APP_NAME, "يوجد عنصر بنفس الاسم في هذا المكان.")
            return
        try:
            os.rename(path, new_path)
        except OSError as e:
            QMessageBox.critical(self, APP_NAME, f"فشلت إعادة التسمية (ربما الملف مفتوح في برنامج آخر):\n{e}")
            return
        self.refresh_all()
        self.select_path(new_path)

    def delete_item(self, path):
        is_dir = os.path.isdir(path)
        kind = "المجلد" if is_dir else "الملف"
        extra = " وكل محتوياته" if is_dir else ""
        box = QMessageBox(self)
        box.setIcon(QMessageBox.Icon.Warning)
        box.setWindowTitle("تأكيد الحذف")
        box.setText(f"واش تحب تحذف {kind} «{os.path.basename(path)}»{extra}؟")
        box.setInformativeText("سيتم نقله إلى سلة المحذوفات.")
        box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        yes = box.addButton("نعم، احذف", QMessageBox.ButtonRole.DestructiveRole)
        cancel = box.addButton("إلغاء", QMessageBox.ButtonRole.RejectRole)
        box.setDefaultButton(cancel)
        box.exec()
        if box.clickedButton() is not yes:
            return
        res = QFile.moveToTrash(path)
        ok = res[0] if isinstance(res, tuple) else res
        if not ok:
            QMessageBox.critical(self, APP_NAME, "فشل الحذف (ربما الملف مفتوح في برنامج آخر).")
        self.refresh_all()

    # ---------- القوائم (كليك يمين) ----------
    def selected_docx(self):
        """ملفات docx المحددة في القائمة النشطة."""
        if self.stack.currentWidget() is self.results:
            paths = [i.data(PATH_ROLE) for i in self.results.selectedItems()]
        else:
            paths = [i.data(0, PATH_ROLE) for i in self.tree.selectedItems()]
        docs = [p for p in paths if p.lower().endswith(".docx") and os.path.isfile(p)]
        return sorted(docs, key=natural_key)

    def show_menu(self, path, global_pos):
        if not path or not os.path.exists(path):
            return
        is_dir = os.path.isdir(path)
        menu = QMenu(self)
        menu.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        menu.addAction("فتح المجلد" if is_dir else "فتح", lambda: self.open_item(path))
        docs = self.selected_docx()
        if len(docs) > 1 and path in docs:
            menu.addAction(f"تحويل الملفات المحددة إلى PDF ({len(docs)})", lambda: self.convert_to_pdf(docs))
        elif path.lower().endswith(".docx") and os.path.isfile(path):
            menu.addAction("تحويل إلى PDF", lambda: self.convert_to_pdf([path]))
        menu.addAction("إعادة تسمية", lambda: self.rename_item(path))
        menu.addAction("حذف", lambda: self.delete_item(path))
        menu.addSeparator()
        menu.addAction("فتح مكان الملف", lambda: reveal_in_explorer(path))
        menu.exec(global_pos)

    def tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        if item is None:
            return
        if not item.isSelected():  # نحافظو على التحديد المتعدد
            self.tree.setCurrentItem(item)
        self.show_menu(item.data(0, PATH_ROLE), self.tree.viewport().mapToGlobal(pos))

    def results_menu(self, pos):
        item = self.results.itemAt(pos)
        if item is None:
            return
        if not item.isSelected():
            self.results.setCurrentItem(item)
        self.show_menu(item.data(PATH_ROLE), self.results.viewport().mapToGlobal(pos))

    # ---------- تحويل إلى PDF ----------
    def convert_to_pdf(self, paths):
        if self.pdf_worker is not None and self.pdf_worker.isRunning():
            QMessageBox.information(self, APP_NAME, "فما تحويل قيد التنفيذ، استنى يكمل.")
            return
        total = len(paths)
        dlg = QProgressDialog("جاري التحويل إلى PDF...", "إلغاء", 0, 0 if total == 1 else total, self)
        dlg.setWindowTitle("تحويل إلى PDF")
        dlg.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        dlg.setWindowModality(Qt.WindowModality.WindowModal)
        dlg.setMinimumDuration(0)
        dlg.setAutoClose(False)
        dlg.setAutoReset(False)
        worker = PdfWorker(paths)
        worker.progress.connect(self.on_pdf_progress)
        worker.finished_all.connect(self.on_pdf_finished)
        dlg.canceled.connect(worker.requestInterruption)  # يوقف بعد الملف الحالي
        self.pdf_worker, self.pdf_dialog = worker, dlg
        dlg.show()
        worker.start()

    def on_pdf_progress(self, done, total, name):
        dlg = self.pdf_dialog
        if dlg is None:
            return
        dlg.setLabelText(f"جاري تحويل: {name}" + (f"\n({done + 1} من {total})" if total > 1 else ""))
        if total > 1:
            dlg.setValue(done)

    def on_pdf_finished(self, done, failed, no_converter):
        dlg, worker = self.pdf_dialog, self.pdf_worker
        self.pdf_dialog = self.pdf_worker = None
        if dlg is not None:
            dlg.canceled.disconnect()
            dlg.close()
        if worker is not None:
            worker.wait()
        self.refresh_all()
        if no_converter:
            box = QMessageBox(self)
            box.setIcon(QMessageBox.Icon.Warning)
            box.setWindowTitle("تحويل إلى PDF")
            box.setText(NO_CONVERTER_MESSAGE)
            box.setDetailedText(no_converter)
            box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
            box.addButton("موافق", QMessageBox.ButtonRole.AcceptRole)
            box.exec()
            return
        box = QMessageBox(self)
        box.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        box.setWindowTitle("تحويل إلى PDF")
        if failed:
            box.setIcon(QMessageBox.Icon.Warning)
            box.setText(f"تم تحويل {len(done)} ملف، وفشل {len(failed)}.")
            box.setDetailedText("\n\n".join(f"{os.path.basename(src)}:\n{err}" for src, err in failed))
        elif done:
            box.setIcon(QMessageBox.Icon.Information)
            box.setText("تم التحويل بنجاح." if len(done) == 1 else f"تم تحويل {len(done)} ملفات بنجاح.")
        else:
            return  # أُلغي قبل تحويل أي ملف
        open_btn = None
        if len(done) == 1:
            box.setInformativeText(done[0])
            open_btn = box.addButton("فتح PDF", QMessageBox.ButtonRole.AcceptRole)
        box.addButton("إغلاق", QMessageBox.ButtonRole.RejectRole)
        if done:
            self.select_path(done[0])
        box.exec()
        if open_btn is not None and box.clickedButton() is open_btn:
            self.open_item(done[0])

    # ---------- مذكرة جديدة ----------
    def new_memo(self):
        if not self.root:
            QMessageBox.information(self, APP_NAME, "اختر المجلد الرئيسي أولا.")
            return
        template = find_template(self.root)
        if not template:
            QMessageBox.warning(
                self, APP_NAME,
                f"ما لقيتش ملف النموذج {TEMPLATE_NAME}.\nحطو في المجلد الرئيسي أو بجنب main.py.",
            )
            return
        section = subject = None
        sel = self.current_tree_path()
        if sel:
            parts = os.path.relpath(sel, self.root).split(os.sep)
            section = parts[0]
            if len(parts) > 1:
                subject = parts[1]
        dlg = NewMemoDialog(self.root, section, subject, self)
        if not dlg.exec():
            return
        try:
            if docx is None:
                shutil.copyfile(template, dlg.target)
                QMessageBox.warning(
                    self, APP_NAME,
                    "مكتبة python-docx غير مثبتة، تم نسخ النموذج بدون تعمير الكلمات المعلّمة.\n"
                    "نفّذ: pip install python-docx",
                )
            else:
                fill_docx(template, dlg.target, dlg.values)
        except Exception as e:  # noqa: BLE001 - ملف نموذج تالف، ملف مقفول...
            if os.path.exists(dlg.target):
                try:
                    os.remove(dlg.target)
                except OSError:
                    pass
            QMessageBox.critical(self, APP_NAME, f"ما نجمتش نكوّن المذكرة من النموذج:\n{e}")
            return
        self.search_edit.clear()
        self.refresh_all()
        self.select_path(dlg.target)
        self.statusBar().showMessage("تم إنشاء المذكرة.", 5000)
        if self.ask("مذكرة جديدة", "تم إنشاء المذكرة.\nتحب تفتح المذكرة توا؟", "نعم، افتحها", "لا"):
            self.open_item(dlg.target)

    def ask(self, title, text, yes_text, no_text):
        return ask_yes_no(self, title, text, yes_text, no_text)

    def open_generator(self):
        if fitz is None or docx is None:
            QMessageBox.warning(self, APP_NAME, "المكتبات ناقصة. نفّذ: pip install -r requirements.txt")
            return
        MemoGeneratorDialog(self.settings, self).exec()
        self.refresh_all()


def main():
    for name in ("stdout", "stderr"):  # exe بدون نافذة أوامر: sys.stdout يكون None وتنكسر مكتبات تكتب فيه
        if getattr(sys, name) is None:
            setattr(sys, name, open(os.devnull, "w", encoding="utf-8"))
    app = QApplication(sys.argv)
    app.setApplicationName(APP_NAME)
    app.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
    font = QFont()
    font.setFamilies(["Segoe UI", "Tahoma"])
    font.setPointSize(11)
    app.setFont(font)
    window = MainWindow()
    window.show()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()
