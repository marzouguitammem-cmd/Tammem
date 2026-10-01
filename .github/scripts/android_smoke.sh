#!/bin/bash
# تجربة آلية على المحاكي: تثبيت الـ APK، تشغيلو، وتجربة التصفح وفتح ملف بالضغط على الأزرار.
set -x
APK=$(ls apk/*x86_64*.apk 2>/dev/null | head -1)
[ -z "$APK" ] && APK=$(ls apk/*.apk | head -1)
echo "APK: $APK"
adb install -r -g "$APK" || exit 1
AAPT=$(ls -d "$ANDROID_HOME"/build-tools/* | tail -1)/aapt
PKG=$("$AAPT" dump badging "$APK" | sed -n "s/package: name='\([^']*\)'.*/\1/p")
echo "PKG: $PKG"
adb shell appops set "$PKG" MANAGE_EXTERNAL_STORAGE allow || true

# ملفات تجريبية (أسماء لاتينية لأن adb input ما يكتبش عربي)
ROOT=/sdcard/Documents/khidma
adb shell mkdir -p "$ROOT/S1/math" "$ROOT/S2"
echo "test" > note.txt
adb push note.txt "$ROOT/S1/math/lesson.txt"
echo "<html><body><h1>page</h1></body></html>" > page.html
adb push page.html "$ROOT/S1/math/page.html"
adb push template.docx "$ROOT/template.docx"
LIB=/sdcard/Documents/lib
adb shell mkdir -p "$LIB/math"
adb push tests/fixtures/android/guide.pdf "$LIB/math/guide.pdf"
adb push tests/fixtures/android/book.pdf "$LIB/math/book.pdf"
adb push tests/fixtures/android/template.docx "$LIB/math/template.docx"

screen() {  # يطبع النصوص الظاهرة
  adb shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
  adb pull /sdcard/ui.xml ui.xml >/dev/null 2>&1
  echo "===== SCREEN: $1 ====="
  python3 - <<'PY'
import re, html
x = open("ui.xml", encoding="utf-8").read()
for n in re.findall(r"<node [^>]*>", x):
    d = html.unescape(re.search(r'content-desc="([^"]*)"', n).group(1))
    t = html.unescape(re.search(r' text="([^"]*)"', n).group(1))
    pkg = re.search(r'package="([^"]*)"', n).group(1)
    cls = re.search(r'class="([^"]*)"', n).group(1)
    h = re.search(r' hint="([^"]*)"', n)
    h = html.unescape(h.group(1)) if h else ""
    if d or t or "EditText" in cls:
        print(f"[{pkg}] ({cls.split('.')[-1]}) {d} {t} {('hint=' + h) if h else ''}".replace("\n", " / "))
PY
}
tap() {  # يضغط على أول عنصر نصّو يطابق regex
  adb shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
  adb pull /sdcard/ui.xml ui.xml >/dev/null 2>&1
  XY=$(python3 - "$1" <<'PY'
import re, sys, html
x = open("ui.xml", encoding="utf-8").read()
for n in re.findall(r"<node [^>]*>", x):
    h = re.search(r' hint="([^"]*)"', n)
    d = html.unescape(re.search(r'content-desc="([^"]*)"', n).group(1)) + " " + html.unescape(re.search(r' text="([^"]*)"', n).group(1)) + " " + (html.unescape(h.group(1)) if h else "")
    if re.search(sys.argv[1], d, re.S):
        b = list(map(int, re.findall(r"\d+", re.search(r'bounds="([^"]*)"', n).group(1))))
        print((b[0] + b[2]) // 2, (b[1] + b[3]) // 2)
        break
PY
)
  echo "tap '$1' -> $XY"
  [ -n "$XY" ] && adb shell input tap $XY
  sleep 3
}

tap_edit() {  # يضغط على الخانة النصية رقم $1 (من 0)
  adb shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
  adb pull /sdcard/ui.xml ui.xml >/dev/null 2>&1
  XY=$(python3 - "$1" <<'PY'
import re, sys
x = open("ui.xml", encoding="utf-8").read()
nodes = [n for n in re.findall(r"<node [^>]*>", x) if "EditText" in n]
if len(nodes) > int(sys.argv[1]):
    b = list(map(int, re.findall(r"\d+", re.search(r'bounds="([^"]*)"', nodes[int(sys.argv[1])]).group(1))))
    print((b[0] + b[2]) // 2, (b[1] + b[3]) // 2)
PY
)
  echo "tap_edit $1 -> $XY"
  [ -n "$XY" ] && adb shell input tap $XY
  sleep 2
}

adb logcat -c
adb shell monkey -p "$PKG" -c android.intent.category.LAUNCHER 1
sleep 75
screen "start"
adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png screen-1-start.png

tap "^الإعدادات.*3 من 3"
screen "settings"
tap_edit 0
adb shell input text "$ROOT"
sleep 2
adb shell input keyevent 111   # إخفاء الكلافيي
sleep 2
tap "حفظ المجلد الرئيسي"
tap "^الملفات.*1 من 3"
screen "files root"
tap "^S1"
tap "^math"
screen "files S1/math"
adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png screen-2-files.png
shot() { adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png "screen-$1.png"; }

# --- فتح ملف (txt: ما فما تطبيق في المحاكي → رسالة واضحة؛ html: يتفتح بـ HTML Viewer) ---
tap "^lesson.txt"
screen "after opening lesson.txt"
tap "^موافق"
tap "^page.html"
sleep 3
screen "after opening page.html (external viewer expected)"
shot 3-open
adb shell input keyevent 4
sleep 3

# --- مذكرة جديدة (python-docx على الأندرويد) ---
tap "^مذكرة جديدة"
screen "new memo dialog"
shot 4-memo
tap "^إنشاء"
sleep 3
screen "after create"
tap "^لا$"
echo "===== FILES IN S1/math ====="
adb shell ls "$ROOT/S1/math"

# --- مولّد المذكرات (PyMuPDF على الأندرويد) ---
tap "^الإعدادات.*3 من 3"
tap_edit 1
adb shell input text "$LIB"
sleep 2
adb shell input keyevent 111
sleep 2
tap "حفظ مجلد المكتبة"
tap "^مولّد المذكرات.*2 من 3"
screen "generator"
tap_edit 1
adb shell input text "Addition"
adb shell input keyevent 111
sleep 2
screen "generator filled"
tap "^تحليل ومراجعة"
sleep 8
screen "review"
shot 5-review
adb shell input swipe 160 500 160 100 300
sleep 1
adb shell input swipe 160 500 160 100 300
sleep 1
adb shell input swipe 160 500 160 100 300
sleep 2
tap "^حفظ$"
sleep 4
screen "after generator save"
tap "^لا$"
echo "===== FILES IN LIB/math ====="
adb shell ls "$LIB/math"

echo "===== LOGCAT (app) ====="
adb logcat -d | grep -i -E "python|flet|traceback|exception|fatal|jnius" | grep -v -i -E "chatty|PeoplePU|gms|Conscrypt|CorpusConfig|MediaScanner" | tail -80 || true
echo "===== END LOGCAT ====="
if adb shell pidof "$PKG"; then echo "APP_RUNNING=yes"; else echo "APP_RUNNING=no"; exit 1; fi
