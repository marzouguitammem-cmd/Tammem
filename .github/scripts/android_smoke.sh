#!/bin/bash
# تجربة آلية: تثبيت الـ APK على المحاكي، تشغيلو، وطباعة النصوص الظاهرة على الشاشة.
set -x
APK=$(ls apk/*x86_64*.apk 2>/dev/null | head -1)
[ -z "$APK" ] && APK=$(ls apk/*.apk | head -1)
echo "APK: $APK"
adb install -r -g "$APK" || exit 1
AAPT=$(ls -d "$ANDROID_HOME"/build-tools/* | tail -1)/aapt
PKG=$("$AAPT" dump badging "$APK" | sed -n "s/package: name='\([^']*\)'.*/\1/p")
echo "PKG: $PKG"
adb shell appops set "$PKG" MANAGE_EXTERNAL_STORAGE allow || true
adb shell mkdir -p "/sdcard/Documents/khidma/S1/math"
adb logcat -c
adb shell monkey -p "$PKG" -c android.intent.category.LAUNCHER 1
sleep 75
adb shell uiautomator dump /sdcard/ui.xml || true
echo "===== UI TEXT ====="
adb shell cat /sdcard/ui.xml | grep -o 'content-desc="[^"]*"\|text="[^"]*"' | grep -v '=""' | head -80 || true
echo "===== END UI TEXT ====="
adb shell screencap -p /sdcard/screen.png && adb pull /sdcard/screen.png screen.png || true
echo "===== LOGCAT ====="
adb logcat -d | grep -i -E "python|flet|traceback|exception|fatal" | grep -v -i chatty | tail -150 || true
echo "===== END LOGCAT ====="
if adb shell pidof "$PKG"; then echo "APP_RUNNING=yes"; else echo "APP_RUNNING=no"; exit 1; fi
