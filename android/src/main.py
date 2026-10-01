# -*- coding: utf-8 -*-
"""مدير ملفات المعلم - نسخة الأندرويد (Flet).

نفس منطق نسخة ويندوز (teacher_core.py): تصفّح الأقسام والمواد، البحث، مذكرة جديدة
بتعمير القالب، ومولّد المذكرات من الدليل والكتاب. يخدم 100% بلا إنترنت.
"""

import asyncio
import datetime
import json
import mimetypes
import os
import shutil
import sys

import flet as ft

HERE = os.path.dirname(os.path.abspath(__file__))
try:
    import teacher_core as core
except ImportError:  # تشغيل محلي من المستودع: الملف المشترك في جذر المشروع
    sys.path.insert(0, os.path.abspath(os.path.join(HERE, "..", "..")))
    import teacher_core as core

APP_NAME = "مدير ملفات المعلم"
GENERATOR_TITLE = "مولّد المذكرات (بلا إنترنت)"
EMPTY_FIELD_COLOR = ft.Colors.AMBER_100
FILE_ICONS = {
    ".pdf": ft.Icons.PICTURE_AS_PDF,
    ".doc": ft.Icons.DESCRIPTION,
    ".docx": ft.Icons.DESCRIPTION,
    ".xls": ft.Icons.TABLE_CHART,
    ".xlsx": ft.Icons.TABLE_CHART,
    ".ppt": ft.Icons.SLIDESHOW,
    ".pptx": ft.Icons.SLIDESHOW,
    ".jpg": ft.Icons.IMAGE,
    ".jpeg": ft.Icons.IMAGE,
    ".png": ft.Icons.IMAGE,
}
EXTRA_MIME = {
    ".docx": "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    ".xlsx": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    ".pptx": "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    ".doc": "application/msword",
    ".xls": "application/vnd.ms-excel",
    ".ppt": "application/vnd.ms-powerpoint",
}


# ---------- الأندرويد (pyjnius) ----------
def _android():
    """يرجّع (autoclass, activity) أو None خارج الأندرويد."""
    host = os.environ.get("MAIN_ACTIVITY_HOST_CLASS_NAME")
    if not host:
        return None
    from jnius import autoclass  # متوفّر في نسخة الأندرويد فقط

    return autoclass, autoclass(host).mActivity


def mime_type(path):
    ext = os.path.splitext(path)[1].lower()
    return EXTRA_MIME.get(ext) or mimetypes.guess_type(path)[0] or "*/*"


def open_external(path):
    """يفتح الملف بالتطبيق المناسب في التلفون (Word، قارئ PDF...)."""
    android = _android()
    if android is None:  # تجربة على الكمبيوتر
        if hasattr(os, "startfile"):
            os.startfile(path)
        else:
            import subprocess

            subprocess.Popen(["xdg-open", path])
        return
    autoclass, activity = android
    Intent = autoclass("android.content.Intent")
    File = autoclass("java.io.File")
    FileProvider = autoclass("androidx.core.content.FileProvider")
    context = activity.getApplicationContext()
    uri = FileProvider.getUriForFile(context, context.getPackageName() + ".provider", File(path))
    intent = Intent(Intent.ACTION_VIEW)
    intent.setDataAndType(uri, mime_type(path))
    # الأندرويد يفتح التطبيق الافتراضي، أو يعرض قائمة التطبيقات إذا ما فماش افتراضي
    intent.addFlags(Intent.FLAG_GRANT_READ_URI_PERMISSION | Intent.FLAG_GRANT_WRITE_URI_PERMISSION
                    | Intent.FLAG_ACTIVITY_NEW_TASK)
    activity.startActivity(intent)


def has_storage_access():
    android = _android()
    if android is None:
        return True
    autoclass, activity = android
    if autoclass("android.os.Build$VERSION").SDK_INT >= 30:
        return bool(autoclass("android.os.Environment").isExternalStorageManager())
    granted = autoclass("android.content.pm.PackageManager").PERMISSION_GRANTED
    return activity.checkSelfPermission("android.permission.WRITE_EXTERNAL_STORAGE") == granted


def request_storage_access():
    """يفتح شاشة «السماح بإدارة كل الملفات» (أندرويد 11+) أو يطلب الإذن العادي."""
    android = _android()
    if android is None:
        return
    autoclass, activity = android
    if autoclass("android.os.Build$VERSION").SDK_INT >= 30:
        Intent = autoclass("android.content.Intent")
        Settings = autoclass("android.provider.Settings")
        Uri = autoclass("android.net.Uri")
        try:
            intent = Intent(Settings.ACTION_MANAGE_APP_ALL_FILES_ACCESS_PERMISSION,
                            Uri.parse("package:" + activity.getPackageName()))
            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
            activity.startActivity(intent)
        except Exception:  # noqa: BLE001 - بعض التلفونات ما فيهاش الشاشة الخاصة بالتطبيق
            intent = Intent(Settings.ACTION_MANAGE_ALL_FILES_ACCESS_PERMISSION)
            intent.addFlags(Intent.FLAG_ACTIVITY_NEW_TASK)
            activity.startActivity(intent)
    else:
        activity.requestPermissions(
            ["android.permission.READ_EXTERNAL_STORAGE", "android.permission.WRITE_EXTERNAL_STORAGE"], 1
        )


def default_storage_root():
    return "/storage/emulated/0" if _android() else os.path.expanduser("~")


def find_template(root):
    for folder in (root, HERE, os.path.join(HERE, "..", "..")):
        candidate = os.path.join(folder, core.TEMPLATE_NAME) if folder else ""
        if candidate and os.path.isfile(candidate):
            return candidate
    return None


def fmt_date(d):
    return f"{d.day:02d}/{d.month:02d}/{d.year}"


def parse_int(text, label, low=1, high=10000):
    try:
        value = int(str(text).strip())
    except ValueError:
        raise core.MemoError(f"{label}: اكتب رقم صحيح.")
    if not low <= value <= high:
        raise core.MemoError(f"{label}: الرقم لازم يكون بين {low} و{high}.")
    return value


class App:
    def __init__(self, page: ft.Page):
        self.page = page
        self.prefs = ft.SharedPreferences()
        self.picker = ft.FilePicker()
        self.root = ""
        self.library = ""
        self.current = ""
        self.search_token = 0

    # ---------- أدوات الواجهة ----------
    def toast(self, message):
        self.page.show_dialog(ft.SnackBar(ft.Text(message), rtl=True))

    def show_message(self, title, message):
        dialog = ft.AlertDialog(
            rtl=True,
            title=ft.Text(title),
            content=ft.Text(message, selectable=True),
            actions=[ft.TextButton("موافق", on_click=lambda e: self.page.pop_dialog())],
        )
        self.page.show_dialog(dialog)

    async def confirm(self, title, message, yes_text, no_text="لا", danger=False):
        future = asyncio.get_running_loop().create_future()

        def answer(value):
            if not future.done():
                future.set_result(value)
            self.page.pop_dialog()

        dialog = ft.AlertDialog(
            rtl=True,
            modal=True,
            title=ft.Text(title),
            content=ft.Text(message),
            actions=[
                ft.TextButton(no_text, on_click=lambda e: answer(False)),
                ft.FilledButton(
                    yes_text,
                    on_click=lambda e: answer(True),
                    style=ft.ButtonStyle(bgcolor=ft.Colors.RED_700) if danger else None,
                ),
            ],
        )
        self.page.show_dialog(dialog)
        return await future

    def open_file(self, path):
        if not os.path.exists(path):
            self.show_message("غلطة", "الملف ما عادش موجود.")
            self.refresh_files()
            return
        try:
            open_external(path)
        except Exception as e:  # noqa: BLE001 - ما فما حتى تطبيق يفتح هذا النوع...
            self.show_message(
                "ما نجمتش نفتح الملف",
                f"ثبّت تطبيق يفتح هذا النوع من الملفات (مثلا Word أو قارئ PDF).\n\nالتفاصيل: {e}",
            )

    # ---------- البداية ----------
    async def start(self):
        page = self.page
        page.title = APP_NAME
        page.rtl = True
        page.theme = ft.Theme(color_scheme_seed=ft.Colors.TEAL)
        page.locale_configuration = ft.LocaleConfiguration(
            supported_locales=[ft.Locale("ar", "TN"), ft.Locale("ar"), ft.Locale("fr"), ft.Locale("en")],
            current_locale=ft.Locale("ar", "TN"),
        )
        if os.environ.get("TFM_FONT"):  # تجربة في المتصفح بلا إنترنت: خط محلي يدعم العربية
            page.fonts = {"local": os.environ["TFM_FONT"]}
            page.theme = ft.Theme(color_scheme_seed=ft.Colors.TEAL, font_family="local")
        page.padding = 0

        self.root = await self.prefs.get("root_folder") or ""
        self.library = await self.prefs.get("library_folder") or ""
        self.current = self.root

        self.build_files_tab()
        self.build_generator_tab()
        self.build_settings_tab()
        # التبويبات الكل تبقى في الصفحة ونبدّلو الظهور برك: تحديث عناصر موش معروضة يلخبط Flet
        self.body = ft.Column([self.files_tab, self.generator_tab, self.settings_tab], expand=True, spacing=0)
        page.navigation_bar = ft.NavigationBar(
            selected_index=0,
            on_change=self.on_nav,
            destinations=[
                ft.NavigationBarDestination(icon=ft.Icons.FOLDER_OUTLINED, selected_icon=ft.Icons.FOLDER, label="الملفات"),
                ft.NavigationBarDestination(icon=ft.Icons.AUTO_AWESOME_OUTLINED, selected_icon=ft.Icons.AUTO_AWESOME, label="مولّد المذكرات"),
                ft.NavigationBarDestination(icon=ft.Icons.SETTINGS_OUTLINED, selected_icon=ft.Icons.SETTINGS, label="الإعدادات"),
            ],
        )
        page.appbar = ft.AppBar(title=ft.Text(APP_NAME), bgcolor=ft.Colors.TEAL_50)
        page.floating_action_button = self.fab
        page.add(ft.SafeArea(self.body, expand=True))
        self.show_tab(0)
        self.refresh_files()
        self.refresh_subjects()
        self.update_permission_status()
        if not has_storage_access():
            self.show_tab(2)
            self.page.navigation_bar.selected_index = 2
            self.show_message(
                "إذن الوصول للملفات",
                "باش التطبيق يقرا ملفاتك ويكتب فيها، لازم تعطيه إذن «إدارة كل الملفات».\n"
                "اضغط «منح إذن الوصول للملفات» في الإعدادات، فعّل الإذن، وارجع للتطبيق.",
            )
        page.update()

    def on_nav(self, e):
        self.show_tab(e.control.selected_index)

    def show_tab(self, index):
        tabs = [self.files_tab, self.generator_tab, self.settings_tab]
        titles = [APP_NAME, GENERATOR_TITLE, "الإعدادات"]
        for i, tab in enumerate(tabs):
            tab.visible = i == index
        self.page.appbar.title = ft.Text(titles[index])
        self.page.floating_action_button = self.fab if index == 0 else None
        if index == 2:
            self.update_permission_status()
        self.page.update()

    # ---------- تبويب الملفات ----------
    def build_files_tab(self):
        self.search_field = ft.TextField(
            hint_text="بحث بالاسم في كل المجلدات...",
            prefix_icon=ft.Icons.SEARCH,
            on_change=self.on_search_change,
            dense=True,
            rtl=True,
        )
        self.up_button = ft.IconButton(icon=ft.Icons.ARROW_UPWARD, tooltip="المجلد الأعلى", on_click=self.go_up)
        self.path_label = ft.Text("", expand=True, max_lines=2, overflow=ft.TextOverflow.ELLIPSIS)
        self.file_list = ft.ListView(expand=True, spacing=0)
        self.fab = ft.FloatingActionButton(icon=ft.Icons.NOTE_ADD, content=ft.Text("مذكرة جديدة"), on_click=self.new_memo)
        self.files_tab = ft.Column(
            expand=True,
            spacing=4,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            controls=[
                ft.Container(self.search_field, padding=ft.Padding.only(left=12, right=12, top=8)),
                ft.Row([self.up_button, self.path_label], spacing=0),
                ft.Divider(height=1),
                self.file_list,
            ],
        )

    def tile(self, path, is_dir, subtitle=None):
        name = os.path.basename(path)
        icon = ft.Icons.FOLDER if is_dir else FILE_ICONS.get(os.path.splitext(name)[1].lower(), ft.Icons.INSERT_DRIVE_FILE)
        items = [
            ft.PopupMenuItem(content="فتح", icon=ft.Icons.OPEN_IN_NEW, on_click=lambda e: self.activate(path)),
            ft.PopupMenuItem(content="إعادة تسمية", icon=ft.Icons.EDIT, on_click=lambda e: self.rename_item(path)),
            ft.PopupMenuItem(content="حذف", icon=ft.Icons.DELETE, on_click=lambda e: self.page.run_task(self.delete_item, path)),
        ]
        if subtitle is not None:
            items.append(ft.PopupMenuItem(content="فتح مكان الملف", icon=ft.Icons.FOLDER_OPEN,
                                          on_click=lambda e: self.reveal(path)))
        return ft.ListTile(
            leading=ft.Icon(icon, color=ft.Colors.AMBER_700 if is_dir else ft.Colors.TEAL_700),
            title=ft.Text(name),
            subtitle=ft.Text(subtitle) if subtitle else None,
            trailing=ft.PopupMenuButton(icon=ft.Icons.MORE_VERT, items=items),
            on_click=lambda e: self.activate(path),
        )

    def activate(self, path):
        if os.path.isdir(path):
            self.search_field.value = ""
            self.current = path
            self.refresh_files()
        else:
            self.open_file(path)

    def reveal(self, path):
        self.search_field.value = ""
        self.current = os.path.dirname(path)
        self.refresh_files()
        self.toast(f"المجلد: {os.path.basename(self.current)}")

    def go_up(self, e=None):
        if self.current and self.root and os.path.normpath(self.current) != os.path.normpath(self.root):
            self.current = os.path.dirname(self.current)
            self.refresh_files()

    def relative(self, path):
        rel = os.path.relpath(path, self.root) if self.root else path
        return "" if rel == "." else rel.replace(os.sep, " ‹ ")

    def refresh_files(self):
        items = []
        if not self.root:
            self.path_label.value = ""
            self.up_button.visible = False
            items.append(
                ft.Container(
                    padding=24,
                    content=ft.Column(
                        horizontal_alignment=ft.CrossAxisAlignment.CENTER,
                        controls=[
                            ft.Icon(ft.Icons.FOLDER_SPECIAL, size=64, color=ft.Colors.TEAL_300),
                            ft.Text("اختر مجلد الخدمة متاعك (المجلد الرئيسي) من الإعدادات.", text_align=ft.TextAlign.CENTER),
                            ft.FilledButton("الإعدادات", icon=ft.Icons.SETTINGS, on_click=lambda e: self.goto_settings()),
                        ],
                    ),
                )
            )
            self.file_list.controls = items
            self.page.update()
            return
        if self.search_field.value and self.search_field.value.strip():
            return  # النتائج تتعرض في run_search
        if not os.path.isdir(self.current or ""):
            self.current = self.root
        self.up_button.visible = os.path.normpath(self.current) != os.path.normpath(self.root)
        self.path_label.value = os.path.basename(self.root) + ((" ‹ " + self.relative(self.current)) if self.relative(self.current) else "")
        try:
            dirs, files = core.list_entries(self.current)
            if not dirs and not files and not os.access(self.current, os.R_OK):
                raise PermissionError
        except PermissionError:
            dirs, files = [], []
            items.append(ft.Text("ما عنديش إذن نقرا هذا المجلد. أعطي إذن الوصول للملفات من الإعدادات.", color=ft.Colors.RED_700))
        for name in dirs:
            items.append(self.tile(os.path.join(self.current, name), True))
        for name in files:
            items.append(self.tile(os.path.join(self.current, name), False))
        if not dirs and not files:
            items.append(ft.Container(ft.Text("المجلد فارغ.", color=ft.Colors.GREY_600), padding=24))
        self.file_list.controls = items
        self.page.update()

    async def on_search_change(self, e):
        self.search_token += 1
        token = self.search_token
        await asyncio.sleep(0.35)
        if token == self.search_token:
            await self.run_search()

    async def run_search(self):
        query = (self.search_field.value or "").strip()
        if not query:
            self.refresh_files()
            return
        if not self.root:
            self.toast("اختر المجلد الرئيسي أولا.")
            return
        results, truncated = await asyncio.to_thread(core.search_names, self.root, query, 500)
        items = []
        self.up_button.visible = False
        self.path_label.value = f"نتائج البحث: {len(results)}" + (" (أكثر، دقّق البحث)" if truncated else "")
        for path, is_dir in results:
            items.append(self.tile(path, is_dir, subtitle=self.relative(os.path.dirname(path)) or "المجلد الرئيسي"))
        if not results:
            items.append(ft.Container(ft.Text("ما لقيت حتى نتيجة."), padding=24))
        self.file_list.controls = items
        self.page.update()

    def refresh_all(self):
        if (self.search_field.value or "").strip():
            self.page.run_task(self.run_search)
        else:
            self.refresh_files()

    def rename_item(self, path):
        is_file = os.path.isfile(path)
        base = os.path.basename(path)
        stem, ext = os.path.splitext(base) if is_file else (base, "")
        field = ft.TextField(value=stem, label="الاسم الجديد", autofocus=True, rtl=True)

        def do_rename(e):
            new_stem = (field.value or "").strip()
            error = core.validate_name(new_stem)
            if error:
                field.error = error
                self.page.update()
                return
            new_path = os.path.join(os.path.dirname(path), new_stem + ext)
            if new_path != path and os.path.exists(new_path):
                field.error = "يوجد عنصر بنفس الاسم في هذا المكان."
                self.page.update()
                return
            try:
                os.rename(path, new_path)
            except OSError as err:
                field.error = f"فشلت إعادة التسمية: {err}"
                self.page.update()
                return
            self.page.pop_dialog()
            self.refresh_all()

        self.page.show_dialog(ft.AlertDialog(
            rtl=True,
            title=ft.Text("إعادة تسمية"),
            content=field,
            actions=[ft.TextButton("إلغاء", on_click=lambda e: self.page.pop_dialog()),
                     ft.FilledButton("موافق", on_click=do_rename)],
        ))

    async def delete_item(self, path):
        is_dir = os.path.isdir(path)
        kind = "المجلد" if is_dir else "الملف"
        extra = " وكل محتوياته" if is_dir else ""
        ok = await self.confirm(
            "تأكيد الحذف",
            f"واش تحب تحذف {kind} «{os.path.basename(path)}»{extra}؟\nالحذف نهائي (ما فماش سلة محذوفات في الأندرويد).",
            "نعم، احذف", "إلغاء", danger=True,
        )
        if not ok:
            return
        try:
            shutil.rmtree(path) if is_dir else os.remove(path)
        except OSError as e:
            self.show_message("فشل الحذف", str(e))
        self.refresh_all()

    # ---------- مذكرة جديدة ----------
    async def new_memo(self, e=None):
        if not self.root:
            self.show_message("مذكرة جديدة", "اختر المجلد الرئيسي أولا من الإعدادات.")
            return
        template = find_template(self.root)
        if not template:
            self.show_message("مذكرة جديدة", f"ما لقيتش ملف النموذج {core.TEMPLATE_NAME}.\nحطو في المجلد الرئيسي.")
            return
        sections = core.subfolders(self.root)
        if not sections:
            self.show_message("مذكرة جديدة", "ما فما حتى قسم في المجلد الرئيسي. أنشئ مجلد قسم أولا.")
            return
        rel = os.path.relpath(self.current, self.root).split(os.sep) if self.current else ["."]
        init_section = rel[0] if rel[0] in sections else sections[0]
        state = {"date": datetime.date.today(), "manual": False}

        section_dd = ft.Dropdown(label="القسم", value=init_section, expand=True,
                                 options=[ft.DropdownOption(key=s, text=s) for s in sections])
        subject_dd = ft.Dropdown(label="المادة", expand=True)
        lesson = ft.TextField(label="عنوان الدرس", hint_text="مثال: الجمع والطرح", rtl=True)
        week = ft.TextField(label="رقم الأسبوع", value="1", keyboard_type=ft.KeyboardType.NUMBER)
        date_btn = ft.OutlinedButton(fmt_date(state["date"]), icon=ft.Icons.CALENDAR_MONTH)
        name = ft.TextField(label="اسم الملف", helper="يتكوّن أوتوماتيك، وتنجم تبدلو", rtl=True)
        error = ft.Text("", color=ft.Colors.RED_700)

        def load_subjects(initial=None):
            subjects = core.subfolders(os.path.join(self.root, section_dd.value or ""))
            subject_dd.options = [ft.DropdownOption(key=s, text=s) for s in subjects]
            subject_dd.value = initial if initial in subjects else (subjects[0] if subjects else None)

        def auto_name(e=None):
            if not state["manual"]:
                name.value = core.auto_file_name(subject_dd.value or "", lesson.value or "", (week.value or "").strip())
            self.page.update()

        def on_section(e):
            load_subjects()
            auto_name()

        def on_name_change(e):
            state["manual"] = bool((name.value or "").strip())

        def on_date(e):
            if picker.value:
                state["date"] = picker.value
                date_btn.content = fmt_date(picker.value)
                self.page.update()

        picker = ft.DatePicker(value=datetime.datetime.now(), first_date=datetime.datetime(2000, 1, 1),
                               last_date=datetime.datetime(2100, 12, 31), on_change=on_date)
        date_btn.on_click = lambda e: self.page.show_dialog(picker)
        section_dd.on_select = on_section
        subject_dd.on_select = auto_name
        lesson.on_change = auto_name
        week.on_change = auto_name
        name.on_change = on_name_change
        load_subjects(rel[1] if len(rel) > 1 else None)
        auto_name()

        async def create(e):
            try:
                if not subject_dd.value:
                    raise core.MemoError("هذا القسم ما فيه حتى مادة. أنشئ مجلد مادة داخلو أولا.")
                week_no = parse_int(week.value, "رقم الأسبوع", 1, 60)
                file_name = (name.value or "").strip()
                err = core.validate_name(file_name)
                if err:
                    raise core.MemoError(err)
                if not file_name.lower().endswith(".docx"):
                    file_name += ".docx"
                target = os.path.join(self.root, section_dd.value, subject_dd.value, file_name)
                if os.path.exists(target):
                    raise core.MemoError("يوجد ملف بنفس الاسم في هذه المادة.")
                values = {
                    "القسم": section_dd.value, "المادة": subject_dd.value,
                    "الدرس": (lesson.value or "").strip(), "الأسبوع": str(week_no),
                    "التاريخ": fmt_date(state["date"]),
                }
                await asyncio.to_thread(core.fill_docx, template, target, values)
            except core.MemoError as err:
                error.value = str(err)
                self.page.update()
                return
            except Exception as err:  # noqa: BLE001
                error.value = f"ما نجمتش نكوّن المذكرة من النموذج: {err}"
                self.page.update()
                return
            self.page.pop_dialog()
            self.search_field.value = ""
            self.current = os.path.dirname(target)
            self.refresh_files()
            if await self.confirm("مذكرة جديدة", "تم إنشاء المذكرة.\nتحب تفتح المذكرة توا؟", "نعم، افتحها"):
                self.open_file(target)

        self.page.show_dialog(ft.AlertDialog(
            rtl=True,
            title=ft.Text("مذكرة جديدة"),
            scrollable=True,
            content=ft.Column(tight=True, spacing=10, controls=[
                section_dd, subject_dd, lesson, week,
                ft.Row([ft.Text("التاريخ:"), date_btn]), name, error,
            ]),
            actions=[ft.TextButton("إلغاء", on_click=lambda e: self.page.pop_dialog()),
                     ft.FilledButton("إنشاء", on_click=create)],
        ))

    # ---------- تبويب مولّد المذكرات ----------
    def build_generator_tab(self):
        self.gen_subject = ft.Dropdown(label="المادة", expand=True, on_select=self.on_gen_subject)
        self.gen_info = ft.Text("", size=13, color=ft.Colors.GREY_700)
        self.gen_lesson = ft.TextField(label="عنوان الدرس", hint_text="مثال: الجمع والطرح", rtl=True)
        self.gen_week = ft.TextField(label="رقم الأسبوع", value="1", keyboard_type=ft.KeyboardType.NUMBER, expand=True)
        self.gen_date = datetime.date.today()
        self.gen_date_btn = ft.OutlinedButton(fmt_date(self.gen_date), icon=ft.Icons.CALENDAR_MONTH,
                                              on_click=lambda e: self.page.show_dialog(self.gen_picker))
        self.gen_picker = ft.DatePicker(value=datetime.datetime.now(), first_date=datetime.datetime(2000, 1, 1),
                                        last_date=datetime.datetime(2100, 12, 31), on_change=self.on_gen_date)

        def num(label):
            return ft.TextField(label=label, value="1", keyboard_type=ft.KeyboardType.NUMBER, expand=True)

        self.guide_from, self.guide_to = num("من"), num("إلى")
        self.book_from, self.book_to = num("من"), num("إلى")
        self.gen_mode = ft.Dropdown(label="معالجة اتجاه النص", value="auto",
                                    options=[ft.DropdownOption(key=v, text=t) for t, v in core.TEXT_MODES])
        self.gen_busy = ft.ProgressRing(visible=False, width=24, height=24)
        self.gen_form = ft.Column(
            scroll=ft.ScrollMode.AUTO, expand=True, spacing=12,
            horizontal_alignment=ft.CrossAxisAlignment.STRETCH,
            controls=[
                ft.Row([self.gen_subject, ft.IconButton(ft.Icons.TUNE, tooltip="تعديل الكلمات المفتاحية", on_click=self.edit_keywords)]),
                self.gen_info,
                self.gen_lesson,
                ft.Row([self.gen_week, ft.Text("التاريخ:"), self.gen_date_btn]),
                ft.Text("صفحات الدرس في الدليل", weight=ft.FontWeight.BOLD),
                ft.Row([self.guide_from, self.guide_to]),
                ft.Text("صفحات الدرس في الكتاب", weight=ft.FontWeight.BOLD),
                ft.Row([self.book_from, self.book_to]),
                self.gen_mode,
                ft.Row([ft.FilledButton("تحليل ومراجعة", icon=ft.Icons.FACT_CHECK, on_click=self.analyze), self.gen_busy]),
            ],
        )
        # «رجوع» و«حفظ» في شريط ثابت فوق: الصفحة طويلة، والسحب فوق خانة نص كبيرة يحرّك النص موش الصفحة
        self.review_list = ft.Column(scroll=ft.ScrollMode.AUTO, expand=True, spacing=8,
                                     horizontal_alignment=ft.CrossAxisAlignment.STRETCH)
        self.review = ft.Column(expand=True, spacing=6, visible=False, controls=[
            ft.Row([
                ft.OutlinedButton("رجوع", icon=ft.Icons.ARROW_FORWARD, on_click=self.back_to_form),
                ft.FilledButton("حفظ", icon=ft.Icons.SAVE, on_click=self.save_generated),
            ], alignment=ft.MainAxisAlignment.SPACE_BETWEEN),
            ft.Divider(height=1),
            self.review_list,
        ])
        self.generator_tab = ft.Container(ft.Column([self.gen_form, self.review], expand=True), padding=16, expand=True)
        self.page_counts = {}
        self.fields = {}

    def on_gen_date(self, e):
        if self.gen_picker.value:
            self.gen_date = self.gen_picker.value
            self.gen_date_btn.content = fmt_date(self.gen_date)
            self.page.update()

    def subject_dir(self):
        return os.path.join(self.library, self.gen_subject.value or "")

    def refresh_subjects(self):
        if not self.library or not os.path.isdir(self.library):
            self.gen_subject.options = []
            self.gen_subject.value = None
            self.gen_info.value = "اختر مجلد المكتبة من الإعدادات. كل مادة مجلد فيه guide.pdf وbook.pdf وtemplate.docx."
            self.page.update()
            return
        subjects = core.subfolders(self.library)
        self.gen_subject.options = [ft.DropdownOption(key=s, text=s) for s in subjects]
        if self.gen_subject.value not in subjects:
            self.gen_subject.value = subjects[0] if subjects else None
        self.on_gen_subject()

    def on_gen_subject(self, e=None):
        self.page_counts = {}
        notes = []
        if not self.gen_subject.value:
            self.gen_info.value = "ما لقيت حتى مجلد مادة داخل المكتبة."
            self.page.update()
            return
        folder = self.subject_dir()
        missing = [f for f in (core.GUIDE_FILE, core.BOOK_FILE, core.TEMPLATE_NAME) if not os.path.isfile(os.path.join(folder, f))]
        if missing:
            notes.append("ناقص: " + "، ".join(missing))
        for key, filename in (("guide", core.GUIDE_FILE), ("book", core.BOOK_FILE)):
            path = os.path.join(folder, filename)
            if os.path.isfile(path):
                try:
                    self.page_counts[key] = core.pdf_page_count(path)
                    notes.append(f"{filename}: {self.page_counts[key]} صفحة")
                except core.MemoError as err:
                    notes.append(str(err))
        self.gen_info.value = " — ".join(notes)
        self.page.update()

    async def analyze(self, e=None):
        self.gen_busy.visible = True
        self.page.update()
        try:
            await self.run_analysis()
        except core.MemoError as err:
            self.show_message("مولّد المذكرات", str(err))
        except Exception as err:  # noqa: BLE001
            self.show_message("مولّد المذكرات", f"صارت غلطة غير متوقعة:\n{err}")
        finally:
            self.gen_busy.visible = False
            self.page.update()

    async def run_analysis(self):
        if not self.library:
            raise core.MemoError("اختر مجلد المكتبة أولا من الإعدادات.")
        subject = self.gen_subject.value
        if not subject:
            raise core.MemoError("اختر المادة.")
        lesson = (self.gen_lesson.value or "").strip()
        if not lesson:
            raise core.MemoError("اكتب عنوان الدرس.")
        week = parse_int(self.gen_week.value, "رقم الأسبوع", 1, 60)
        error = core.validate_name(core.auto_file_name(subject, lesson, week))
        if error:
            raise core.MemoError(f"اسم المذكرة غير صالح: {error}")
        folder = self.subject_dir()
        for filename in (core.GUIDE_FILE, core.TEMPLATE_NAME):
            if not os.path.isfile(os.path.join(folder, filename)):
                raise core.MemoError(f"ما لقيتش {filename} في مجلد المادة «{subject}».")
        if not self.page_counts.get("guide"):
            self.on_gen_subject()
            if not self.page_counts.get("guide"):
                raise core.MemoError(f"ما نجمتش نقرا {core.GUIDE_FILE}. تأكد أنو سليم وفيه صفحات.")

        def page_range(first, last, key, label):
            a = parse_int(first.value, f"صفحات {label} (من)")
            b = parse_int(last.value, f"صفحات {label} (إلى)")
            if a > b:
                raise core.MemoError(f"صفحات {label}: رقم البداية ({a}) أكبر من رقم النهاية ({b}).")
            if b > self.page_counts.get(key, 0):
                raise core.MemoError(f"صفحات {label}: الملف فيه {self.page_counts.get(key, 0)} صفحة برك.")
            return a, b

        guide_range = page_range(self.guide_from, self.guide_to, "guide", "الدليل")
        has_book = bool(self.page_counts.get("book"))
        book_range = page_range(self.book_from, self.book_to, "book", "الكتاب") if has_book else None
        keywords = core.load_keywords(folder)
        vocab = core.keyword_vocab(keywords)
        mode = self.gen_mode.value or "auto"

        guide_lines = await asyncio.to_thread(core.extract_lines, os.path.join(folder, core.GUIDE_FILE), *guide_range, mode, vocab)
        book_lines = []
        if has_book:
            book_lines = await asyncio.to_thread(core.extract_lines, os.path.join(folder, core.BOOK_FILE), *book_range, mode, vocab)
        values, self.extra_values = core.build_generator_values(guide_lines, book_lines, keywords)
        self.review_meta = {"subject": subject, "lesson": lesson, "week": week, "date": fmt_date(self.gen_date)}
        self.build_review(values, guide_lines, book_lines, has_book)

    def build_review(self, values, guide_lines, book_lines, has_book):
        self.fields = {}
        missing = [k for k, v in values.items() if k != core.BOOK_KEY and not v]
        meta = self.review_meta
        controls = [
            ft.Text(f"{meta['subject']} — {meta['lesson']} — أسبوع {meta['week']}", weight=ft.FontWeight.BOLD, size=16),
        ]
        if missing:
            controls.append(ft.Text(f"أقسام ما تلقاتش لا في الدليل لا في الكتاب (معلّمة بالأصفر): {'، '.join(missing)}.", color=ft.Colors.AMBER_900))
        if not has_book:
            controls.append(ft.Text(f"ما فما {core.BOOK_FILE} في مجلد المادة، خانة الكتاب فارغة.", color=ft.Colors.AMBER_900))
        for name, text in values.items():
            if name == core.BOOK_KEY:
                hint = "ما فما نص للكتاب. عمّرو بيدك أو ارجع وصحّح الصفحات."
            else:
                hint = f"ما لقيتش «{name}» لا في الدليل لا في الكتاب. عمّرو بيدك أو عدّل الكلمات المفتاحية."
            field = ft.TextField(label=name, value=text, multiline=True, min_lines=3, max_lines=14, rtl=True, helper_max_lines=3)

            def paint(e=None, f=field, h=hint):
                empty = not (f.value or "").strip()
                f.bgcolor = EMPTY_FIELD_COLOR if empty else None
                f.filled = empty
                f.helper = ("⚠ " + h) if empty else None
                if e is not None:
                    self.page.update()

            field.on_change = paint
            paint()
            self.fields[name] = field
            controls.append(field)
        controls += [
            ft.ExpansionTile(title=ft.Text("النص الأصلي: الدليل"), controls=[
                ft.Container(ft.Text("\n".join(guide_lines), selectable=True), padding=8)]),
            ft.ExpansionTile(title=ft.Text("النص الأصلي: الكتاب"), controls=[
                ft.Container(ft.Text("\n".join(book_lines) or "—", selectable=True), padding=8)]),
        ]
        self.review_list.controls = controls
        self.gen_form.visible = False
        self.review.visible = True
        self.page.update()

    def back_to_form(self, e=None):
        self.review.visible = False
        self.gen_form.visible = True
        self.page.update()

    async def save_generated(self, e=None):
        meta = self.review_meta
        folder = self.subject_dir()
        template = os.path.join(folder, core.TEMPLATE_NAME)
        name = core.auto_file_name(meta["subject"], meta["lesson"], meta["week"]) + ".docx"
        target = os.path.join(folder, name)
        if os.path.exists(target) and not await self.confirm(
            "الملف موجود", f"يوجد ملف بنفس الاسم:\n{name}\nتحب تستبدلو؟", "نعم، استبدل", danger=True
        ):
            return
        values = {"المادة": meta["subject"], "الدرس": meta["lesson"], "الأسبوع": str(meta["week"]), "التاريخ": meta["date"]}
        values.update(getattr(self, "extra_values", {}))  # {{القسم - الدليل}} و{{القسم - الكتاب}}
        for field_name, field in self.fields.items():
            values[field_name] = (field.value or "").strip()
        try:
            present = await asyncio.to_thread(core.template_placeholders, template)
            await asyncio.to_thread(core.fill_docx, template, target, values)
        except core.MemoError as err:
            self.show_message("الحفظ", str(err))
            return
        except Exception as err:  # noqa: BLE001
            self.show_message("الحفظ", f"ما نجمتش نحفظ المذكرة:\n{err}")
            return
        known = {core.placeholder_key(k) for k in values}
        warnings = []
        unknown = [orig for key, orig in present.items() if key not in known]
        if unknown:
            warnings.append("كلمات معلّمة في النموذج ما عندها خانة (بقات كيما هي): " + "، ".join("{{%s}}" % u for u in unknown))
        unused = [f for f, v in values.items()
                  if f in self.fields and f != core.BOOK_KEY and v and not core.section_in_template(f, present)]
        if unused:
            warnings.append("أقسام فيها نص لكن ما فماش كلمة معلّمة ليها في النموذج: " + "، ".join("{{%s}}" % u for u in unused))
        text = f"تم حفظ المذكرة:\n{name}"
        if warnings:
            text += "\n\nملاحظة:\n" + "\n".join(warnings)
        if await self.confirm("تم الحفظ", text + "\n\nتحب تفتحها توا؟", "نعم، افتحها"):
            self.open_file(target)
        self.back_to_form()

    async def edit_keywords(self, e=None):
        if not self.gen_subject.value:
            self.show_message("الكلمات المفتاحية", "اختر المادة أولا.")
            return
        folder = self.subject_dir()
        try:
            keywords = core.load_keywords(folder)
        except core.MemoError as err:
            keywords = None
            self.toast(str(err))
        path = os.path.join(folder, core.KEYWORDS_FILE)
        try:
            with open(path, encoding="utf-8-sig") as f:
                raw = f.read()
        except OSError:
            raw = json.dumps(keywords or core.DEFAULT_KEYWORDS, ensure_ascii=False, indent=2)
        editor = ft.TextField(value=raw, multiline=True, min_lines=10, max_lines=20, rtl=True, text_size=14)
        error = ft.Text("", color=ft.Colors.RED_700)

        def restore_defaults(e):
            editor.value = json.dumps(core.DEFAULT_KEYWORDS, ensure_ascii=False, indent=2)
            error.value = "القيم الافتراضية متاع الدليل والكتاب. اضغط «حفظ» باش تتسجّل."
            self.page.update()

        def save(e):
            try:
                data = json.loads(editor.value or "")
            except json.JSONDecodeError as err:
                error.value = f"غلطة في الكتابة (سطر {err.lineno}، عمود {err.colno}): {err.msg}"
                self.page.update()
                return
            try:
                core.validate_keywords(data)
                with open(path, "w", encoding="utf-8") as f:
                    json.dump(data, f, ensure_ascii=False, indent=2)
            except (core.MemoError, OSError) as err:
                error.value = str(err)
                self.page.update()
                return
            self.page.pop_dialog()
            self.toast("تم حفظ الكلمات المفتاحية.")

        self.page.show_dialog(ft.AlertDialog(
            rtl=True,
            scrollable=True,
            title=ft.Text(f"الكلمات المفتاحية: {self.gen_subject.value}"),
            content=ft.Column(tight=True, controls=[
                ft.Text("كل قسم: الكلمات اللي تدل على عنوانو في «الدليل» وفي «الكتاب». "
                        "اسم القسم هو الكلمة المعلّمة في النموذج، مثلا {{الوضعية المشكل}}.", size=13),
                editor, error]),
            actions=[ft.TextButton("القيم الافتراضية", on_click=restore_defaults),
                     ft.TextButton("إلغاء", on_click=lambda e: self.page.pop_dialog()),
                     ft.FilledButton("حفظ", on_click=save)],
        ))

    # ---------- تبويب الإعدادات ----------
    def build_settings_tab(self):
        self.root_field = ft.TextField(label="المجلد الرئيسي (مجلد الخدمة)", value=self.root, rtl=True, expand=True)
        self.lib_field = ft.TextField(label="مجلد المكتبة (مولّد المذكرات)", value=self.library, rtl=True, expand=True)
        self.permission_text = ft.Text("")
        self.settings_tab = ft.Column(
            scroll=ft.ScrollMode.AUTO, expand=True, spacing=14,
            controls=[ft.Container(padding=16, content=ft.Column(spacing=14, horizontal_alignment=ft.CrossAxisAlignment.STRETCH, controls=[
                ft.Text("إذن الوصول للملفات", weight=ft.FontWeight.BOLD),
                self.permission_text,
                ft.FilledButton("منح إذن الوصول للملفات", icon=ft.Icons.LOCK_OPEN, on_click=self.ask_permission),
                ft.Divider(),
                ft.Row([self.root_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="اختيار", on_click=self.pick_root)]),
                ft.FilledButton("حفظ المجلد الرئيسي", icon=ft.Icons.SAVE, on_click=self.save_root),
                ft.Divider(),
                ft.Row([self.lib_field, ft.IconButton(ft.Icons.FOLDER_OPEN, tooltip="اختيار", on_click=self.pick_library)]),
                ft.FilledButton("حفظ مجلد المكتبة", icon=ft.Icons.SAVE, on_click=self.save_library),
                ft.Divider(),
                ft.Text("التطبيق يخدم 100% بلا إنترنت. ملفاتك ما تخرجش من التلفون.", color=ft.Colors.GREY_700, size=13),
            ]))],
        )

    def update_permission_status(self):
        try:
            ok = has_storage_access()
        except Exception:  # noqa: BLE001
            ok = False
        self.permission_text.value = "✔ الإذن مفعّل." if ok else "✘ الإذن موش مفعّل: التطبيق ما ينجمش يقرا ملفاتك."
        self.permission_text.color = ft.Colors.GREEN_700 if ok else ft.Colors.RED_700

    def ask_permission(self, e=None):
        try:
            request_storage_access()
        except Exception as err:  # noqa: BLE001
            self.show_message(
                "إذن الوصول للملفات",
                "ما نجمتش نفتح شاشة الإذن. افتحها بيدك: الإعدادات ← التطبيقات ← مدير ملفات المعلم ← "
                f"الأذونات ← الملفات ← «السماح بإدارة كل الملفات».\n\nالتفاصيل: {err}",
            )

    def on_resume(self):
        self.update_permission_status()
        self.refresh_all()
        self.page.update()

    async def pick(self, title, current):
        try:
            return await self.picker.get_directory_path(dialog_title=title, initial_directory=current or None)
        except Exception as err:  # noqa: BLE001 - مثلا في المتصفح
            self.toast(f"ما نجمتش نفتح نافذة الاختيار، اكتب المسار بيدك. ({err})")
            return None

    async def pick_root(self, e=None):
        path = await self.pick("اختيار المجلد الرئيسي", self.root_field.value)
        if path:
            self.root_field.value = path
            await self.save_root()

    async def pick_library(self, e=None):
        path = await self.pick("اختيار مجلد المكتبة", self.lib_field.value)
        if path:
            self.lib_field.value = path
            await self.save_library()

    def check_folder(self, path):
        if not path:
            raise core.MemoError("اكتب مسار المجلد أو اختارو.")
        if not os.path.isdir(path):
            raise core.MemoError(f"المجلد غير موجود أو ما عنديش إذن نقراه:\n{path}")

    async def save_root(self, e=None):
        path = os.path.normpath((self.root_field.value or "").strip()) if (self.root_field.value or "").strip() else ""
        try:
            self.check_folder(path)
        except core.MemoError as err:
            self.show_message("المجلد الرئيسي", str(err))
            return
        self.root = self.current = path
        self.root_field.value = path
        await self.prefs.set("root_folder", path)
        self.search_field.value = ""
        self.refresh_files()
        self.toast("تم حفظ المجلد الرئيسي.")

    async def save_library(self, e=None):
        path = os.path.normpath((self.lib_field.value or "").strip()) if (self.lib_field.value or "").strip() else ""
        try:
            self.check_folder(path)
        except core.MemoError as err:
            self.show_message("مجلد المكتبة", str(err))
            return
        self.library = path
        self.lib_field.value = path
        await self.prefs.set("library_folder", path)
        self.back_to_form()
        self.refresh_subjects()
        self.toast("تم حفظ مجلد المكتبة.")

    def goto_settings(self):
        self.page.navigation_bar.selected_index = 2
        self.show_tab(2)
        if not self.root_field.value:
            self.root_field.value = default_storage_root()
            self.page.update()


async def main(page: ft.Page):
    app = App(page)

    def on_lifecycle(e):
        if getattr(e, "state", None) == ft.AppLifecycleState.RESUME:
            app.on_resume()

    page.on_app_lifecycle_state_change = on_lifecycle
    await app.start()


if __name__ == "__main__":
    if os.environ.get("TFM_WEB_PORT"):
        ft.run(main, view=ft.AppView.WEB_BROWSER, port=int(os.environ["TFM_WEB_PORT"]),
               assets_dir=os.environ.get("TFM_ASSETS", "assets"), no_cdn=True)
    else:
        ft.run(main)
