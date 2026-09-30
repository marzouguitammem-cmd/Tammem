# -*- coding: utf-8 -*-
"""مدير ملفات المعلم - Teacher File Manager (PyQt6)."""

import os
import re
import shutil
import subprocess
import sys

from PyQt6.QtCore import QFile, QSettings, Qt, QTimer, QUrl
from PyQt6.QtGui import QDesktopServices, QFont, QKeySequence, QShortcut
from PyQt6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QComboBox,
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
    QPushButton,
    QStackedWidget,
    QTreeWidget,
    QTreeWidgetItem,
    QVBoxLayout,
    QWidget,
)

APP_NAME = "مدير ملفات المعلم"
TEMPLATE_NAME = "template.docx"
MAX_SEARCH_RESULTS = 1000
INVALID_CHARS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')
RESERVED_NAMES = {
    "CON", "PRN", "AUX", "NUL",
    *(f"COM{i}" for i in range(1, 10)),
    *(f"LPT{i}" for i in range(1, 10)),
}
IGNORED_FILES = {"thumbs.db", "desktop.ini"}
PATH_ROLE = Qt.ItemDataRole.UserRole
LOADED_ROLE = Qt.ItemDataRole.UserRole + 1

_ARABIC_MAP = str.maketrans({"أ": "ا", "إ": "ا", "آ": "ا", "ة": "ه", "ى": "ي", "ؤ": "و", "ئ": "ي"})
_ARABIC_MARKS = re.compile("[ً-ٰٟـ]")


def normalize(text):
    """تطبيع النص للبحث: حروف صغيرة، بدون تشكيل، وتوحيد الألف والتاء المربوطة."""
    return _ARABIC_MARKS.sub("", text.casefold()).translate(_ARABIC_MAP)


def natural_key(name):
    return [int(p) if p.isdigit() else p for p in re.split(r"(\d+)", normalize(name))]


def is_hidden(name):
    return name.startswith(".") or name.startswith("~$") or name.lower() in IGNORED_FILES


def list_entries(path):
    """يرجع (المجلدات, الملفات) مرتبة، بدون الملفات المخفية."""
    dirs, files = [], []
    try:
        with os.scandir(path) as it:
            for e in it:
                if is_hidden(e.name):
                    continue
                try:
                    (dirs if e.is_dir() else files).append(e.name)
                except OSError:
                    continue
    except OSError:
        pass
    dirs.sort(key=natural_key)
    files.sort(key=natural_key)
    return dirs, files


def subfolders(path):
    return list_entries(path)[0]


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


def validate_name(name):
    """يرجع رسالة خطأ بالعربية أو None إذا الاسم صالح."""
    if not name:
        return "الاسم فارغ."
    if INVALID_CHARS.search(name):
        return 'الاسم فيه رموز ممنوعة: < > : " / \\ | ? *'
    if name != name.rstrip(" ."):
        return "الاسم ما ينجمش ينتهي بنقطة أو فراغ."
    if name.split(".")[0].upper() in RESERVED_NAMES:
        return "هذا الاسم محجوز من النظام."
    return None


def find_template(root):
    here = os.path.dirname(os.path.abspath(__file__))
    for folder in (root, here):
        candidate = os.path.join(folder, TEMPLATE_NAME)
        if os.path.isfile(candidate):
            return candidate
    return None


class NewMemoDialog(QDialog):
    def __init__(self, root, initial_section=None, initial_subject=None, parent=None):
        super().__init__(parent)
        self.root = root
        self.setWindowTitle("مذكرة جديدة")
        self.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.setMinimumWidth(380)

        self.section_box = QComboBox()
        self.subject_box = QComboBox()
        self.name_edit = QLineEdit()
        self.name_edit.setPlaceholderText("مثال: مذكرة الدرس 1")

        form = QFormLayout()
        form.addRow("القسم:", self.section_box)
        form.addRow("المادة:", self.subject_box)
        form.addRow("اسم المذكرة:", self.name_edit)

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
        self.name_edit.setFocus()

    def load_subjects(self, section):
        self.subject_box.clear()
        if section:
            self.subject_box.addItems(subfolders(os.path.join(self.root, section)))

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
        super().accept()


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_NAME)
        self.resize(900, 650)
        self.settings = QSettings("TeacherFileManager", "TeacherFileManager")
        self.root = ""
        self.items = {}  # مسار -> عنصر الشجرة
        self.icons = QFileIconProvider()

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
        top.addWidget(self.search_edit, 1)
        top.addWidget(self.refresh_btn)

        self.root_label = QLabel("لم يتم اختيار مجلد رئيسي")
        self.root_label.setTextInteractionFlags(Qt.TextInteractionFlag.TextSelectableByMouse)

        self.tree = QTreeWidget()
        self.tree.setHeaderLabel("الأقسام ← المواد ← الملفات")
        self.tree.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        self.tree.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.tree.customContextMenuRequested.connect(self.tree_menu)
        self.tree.itemExpanded.connect(self.on_expand)
        self.tree.itemDoubleClicked.connect(lambda item, _c: self.open_if_file(item.data(0, PATH_ROLE)))

        self.results = QListWidget()
        self.results.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
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
    def show_menu(self, path, global_pos):
        if not path or not os.path.exists(path):
            return
        is_dir = os.path.isdir(path)
        menu = QMenu(self)
        menu.setLayoutDirection(Qt.LayoutDirection.RightToLeft)
        menu.addAction("فتح المجلد" if is_dir else "فتح", lambda: self.open_item(path))
        menu.addAction("إعادة تسمية", lambda: self.rename_item(path))
        menu.addAction("حذف", lambda: self.delete_item(path))
        menu.addSeparator()
        menu.addAction("فتح مكان الملف", lambda: reveal_in_explorer(path))
        menu.exec(global_pos)

    def tree_menu(self, pos):
        item = self.tree.itemAt(pos)
        if item is None:
            return
        self.tree.setCurrentItem(item)
        self.show_menu(item.data(0, PATH_ROLE), self.tree.viewport().mapToGlobal(pos))

    def results_menu(self, pos):
        item = self.results.itemAt(pos)
        if item is None:
            return
        self.results.setCurrentItem(item)
        self.show_menu(item.data(PATH_ROLE), self.results.viewport().mapToGlobal(pos))

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
            shutil.copyfile(template, dlg.target)
        except OSError as e:
            QMessageBox.critical(self, APP_NAME, f"ما نجمتش ننسخ النموذج:\n{e}")
            return
        self.search_edit.clear()
        self.refresh_all()
        self.select_path(dlg.target)
        self.statusBar().showMessage("تم إنشاء المذكرة.", 5000)
        self.open_item(dlg.target)


def main():
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
