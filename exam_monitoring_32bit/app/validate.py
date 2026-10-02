# -*- coding: utf-8 -*-
"""التحقق من الأرقام والرسائل الاحترافية — يعمل بلا واجهة رسومية (قابل للاختبار)."""

MINISTRY_PREFIXES = ("17", "18", "19", "20")          # أمثلة الأوزاري المتداولة
NATIONAL_LEN = 10
PHONE_PREFIXES = ("077", "078", "079", "076")          # إلزامي كما طُلب

MSG_NOT_DIGITS   = "الإدخال يجب أن يحتوي على أرقام فقط."
MSG_BAD_MINISTRY = "الرقم الوزاري غير صحيح: يجب أن يكون رقمًا وزاريًا ministerial مكونًا من 5 إلى 6 خانات يبدأ بـ 17/18/19/20."
MSG_BAD_NATIONAL = "الرقم الوطني غير صحيح: يجب أن يتألف من 10 خانات."
MSG_LOCKED       = "تم تجاوز عدد المحاولات المسموح (3 محاولات) — سيتم إغلاق البرنامج."
MSG_PHONE        = "رقم الهاتف غير صحيح: يجب أن يتألف من 10 خانات ويبدأ بـ 077 أو 078 أو 079 أو 076."
MSG_OK           = "تم الحفظ بنجاح."
MSG_QUIT         = "خرج المستخدم بعد 3 محاولات خاطئة."


def digits_only(text):
    return str(text or "").strip().replace(" ", "").isdigit()


def is_ministry_number(text):
    """رقم وزاري: 5-6 خانات، يبدأ بـ 17/18/19/20 (مثال: 172571)."""
    s = str(text or "").strip()
    if not s.isdigit():
        return False
    if len(s) not in (5, 6):
        return False
    return s[:2] in MINISTRY_PREFIXES


def is_national_number(text):
    """رقم وطني: 10 خانات (مثال: 9881018940)."""
    s = str(text or "").strip()
    return s.isdigit() and len(s) == NATIONAL_LEN


def is_phone(text):
    """هاتف: 10 خانات يبدأ بـ 077/078/079/076 إلزاميًا."""
    s = str(text or "").strip()
    return s.isdigit() and len(s) == 10 and s.startswith(PHONE_PREFIXES)


def login_check(text):
    """يقبل رقمًا وزاريًا أو وطنيًا. يعيد (ok, kind, normalized, message)."""
    s = str(text or "").strip()
    if not digits_only(s):
        return False, None, s, MSG_NOT_DIGITS
    if is_national_number(s):
        return True, "national", s, MSG_OK
    if is_ministry_number(s):
        return True, "ministry", s, MSG_OK
    # تمييز الرسالة الأدق
    if len(s) == NATIONAL_LEN:
        return False, None, s, MSG_BAD_NATIONAL
    return False, None, s, MSG_BAD_MINISTRY
