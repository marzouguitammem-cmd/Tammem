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
    if d or t:
        print(f"[{pkg}] {d} {t}".replace("\n", " / "))
PY
}
tap() {  # يضغط على أول عنصر نصّو يطابق regex
  adb shell uiautomator dump /sdcard/ui.xml >/dev/null 2>&1
  adb pull /sdcard/ui.xml ui.xml >/dev/null 2>&1
  XY=$(python3 - "$1" <<'PY'
import re, sys, html
x = open("ui.xml", encoding="utf-8").read()
for n in re.findall(r"<node [^>]*>", x):
    d = html.unescape(re.search(r'content-desc="([^"]*)"', n).group(1)) + " " + html.unescape(re.search(r' text="([^"]*)"', n).group(1))
    if re.search(sys.argv[1], d):
        b = list(map(int, re.findall(r"\d+", re.search(r'bounds="([^"]*)"', n).group(1))))
        print((b[0] + b[2]) // 2, (b[1] + b[3]) // 2)
        break
PY
)
  echo "tap '$1' -> $XY"
  [ -n "$XY" ] && adb shell input tap $XY
  sleep 3
}

adb logcat -c
adb shell monkey -p "$PKG" -c android.intent.category.LAUNCHER 1
sleep 75
screen "start"
adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png screen-1-start.png

tap "^الإعدادات.*3 من 3"
screen "settings"
tap "المجلد الرئيسي \(مجلد الخدمة\)"
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
tap "^lesson.txt"
sleep 3
screen "after opening lesson.txt (chooser expected)"
adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png screen-3-open.png

echo "===== LOGCAT (app) ====="
adb logcat -d | grep -i -E "python|flet|traceback|exception|fatal|jnius" | grep -v -i -E "chatty|PeoplePU|gms|Conscrypt|CorpusConfig|MediaScanner" | tail -80 || true
echo "===== END LOGCAT ====="
if adb shell pidof "$PKG"; then echo "APP_RUNNING=yes"; else echo "APP_RUNNING=no"; exit 1; fi
