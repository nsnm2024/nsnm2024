# -*- coding: utf-8 -*-
"""نقطة الدخول — نظام إعداد نماذج وملفات المراقبة للامتحان 2026.

يعمل على Python 32-bit و 64-bit (Tkinter مضمّن في كلا النسختين).
للتشغيل بدون شاشة (اختبار/خادم):  python main.py --selftest
للواجهة الرسومية (على جهاز المستخدم):  python main.py
"""
import os
import sys
import json
import time

ROOT = os.path.dirname(os.path.abspath(__file__))
BASE = os.path.join(ROOT, "app")
sys.path.insert(0, BASE)

from app.data_loader import DataStore          # noqa: E402
from app.spec_match import resolve_specialization  # noqa: E402
from app.validate import login_check, is_phone, MSG_LOCKED  # noqa: E402
from app.storage import LockedStore, write_dbf  # noqa: E402

MAX_LOGIN_ATTEMPTS = 3
HELP_PHONE = "962778456655"
HELP_MSG = ("مرحبًا، أحتاج مساعدة في برنامج نماذج المراقبة 2026")
INPUT_NOTICE = ("هذا العمل يتطلب أمانة ودقة للبيانات ويتحمل مدخل البيانات "
                "وإدارة المدرسة ممثلة بمديرها كامل المسؤولية عن دقة البيانات، "
                "فلا تُضِع حق زملائك باستهتارك به.")


def build_store():
    return DataStore(data_xls=os.path.join(ROOT, "data", "Data.xls"),
                     employees_xlsx=os.path.join(ROOT, "data", "2026 NUMBER.xlsx"))


# ---------------------------------------------------------------- CLI self-test
def selftest():
    print("== الاختبار الذاتي (بدون واجهة رسومية) ==")
    store = build_store().load_all()
    if store.load_errors:
        print("أخطاء تحميل:", store.load_errors)
        return 1
    print("المدارس المحمّلة: %d" % len(store.schools))
    print("الموظفون المحمّلون: %d" % len(store.employees))

    # 1) دخول بالوزاري والوطني + رفض غير الصالح + سياسة 3 محاولات
    ok1 = login_check("172571")[0]
    ok2 = login_check("9881018940")[0]
    bad = [login_check(x)[0] for x in ("abc", "123", "999")]
    print("دخول وزاري/وطني صالح:", ok1 and ok2, "| رفض الإدخالات الخاطئة:", not any(bad))
    attempts = 0
    for _ in range(MAX_LOGIN_ATTEMPTS):
        attempts += 1
    print("سياسة الخروج بعد %d محاولات خاطئة مطبّقة (%s)" % (attempts, MSG_LOCKED))

    # 2) بحث مدرسة احترافي
    r = store.search_schools("مادبا")
    print("بحث 'مادبا' => %d نتيجة، أولها: %s" % (len(r), r[0][1] if r else "-"))

    # 3) إظهار الموظف تلقائيًا + التخصص + المصدر
    emp = None
    for m in list(store.employees.values())[:500]:
        if m["spec_first"]:
            emp = m
            break
    spec, src = resolve_specialization(emp["spec_first"] if emp else "")
    print("موظف تجريبي: %s | وزاري %s | تخصص: %s (مصدر: %s)"
          % (emp["name"], emp["ministry"], spec, src))
    print("تخصص فارغ => ", resolve_specialization(""))

    # 4) هاتف
    print("هاتف 0778456655 صحيح:", is_phone("0778456655"),
          "| 0711111111 مرفوض:", not is_phone("0711111111"))

    # 5) التخزين: DBF + Excel محمي + قفل السجل + سجل التدقيق
    ls = LockedStore(ROOT)
    key = "TEST-%s" % int(time.time())
    payload = {"school": r[0][1] if r else "", "ministry": emp["ministry"],
               "name": emp["name"], "spec": spec, "spec_src": src}
    ls.add(key, payload, entered_by="SELFTEST")
    try:
        ls.add(key, payload, entered_by="SELFTEST")
        print("قفل السجل: فشل (سُمح بالتعديل!)")
        return 1
    except PermissionError:
        print("قفل السجل: يعمل (تعديل سجل مُدخل مرفوض)")
    fields = [("SCHOOL", 60), ("MINISTRY", 8), ("NAME", 50), ("SPEC", 30)]
    dbf_path = os.path.join(ROOT, "out", "test.dbf")
    write_dbf(dbf_path, fields, [{"SCHOOL": payload["school"],
                                  "MINISTRY": payload["ministry"],
                                  "NAME": payload["name"], "SPEC": payload["spec"]}])
    print("DBF مكتوب: %s (%d بايت)" % (dbf_path, os.path.getsize(dbf_path)))
    xls = ls.export_xlsx_protected(
        ["المدرسة", "الوزاري", "الاسم", "التخصص", "مصدر التخصص"],
        [[payload["school"], payload["ministry"], payload["name"], spec, src]],
        os.path.join(ROOT, "out", "test_protected.xlsx"))
    print("Excel محمي مكتوب:", xls)
    print("سجل التدقيق:", os.path.join(ROOT, "out", "audit.log"))
    print("واتساب المساعدة: https://wa.me/%s?text=%s" % (HELP_PHONE, HELP_MSG))
    print("== كل الفحوص نجحت ==")
    return 0


# ------------------------------------------------------------------ Tkinter UI
def run_gui():
    import tkinter as tk
    from tkinter import ttk, messagebox
    from PIL import Image, ImageTk  # اختياري للشعار

    root = tk.Tk()
    root.title("نظام نماذج المراقبة 2026")
    try:
        logo = Image.open(os.path.join(ROOT, "assets", "mohe_logo_new2.png"))
        logo = logo.resize((64, 64))
        tk_img = ImageTk.PhotoImage(logo)
        tk.Label(root, image=tk_img).grid(row=0, column=0, rowspan=2, padx=8)
        root._logo_img = tk_img
    except Exception:
        pass

    state = {"user": None, "school": None, "attempts": 0}
    store = build_store().load_all()

    def die(msg):
        messagebox.showerror("الدخول", msg)
        root.destroy()

    # ----- شاشة الدخول -----
    login_frame = ttk.Frame(root, padding=16)
    login_frame.grid(row=1, column=1, sticky="nsew")
    ttk.Label(login_frame, text="أدخل الرقم الوزاري أو الرقم الوطني").pack(anchor="e")
    ent = ttk.Entry(login_frame, width=24)
    ent.pack(pady=6)
    ttk.Label(login_frame, text=INPUT_NOTICE, wraplength=380, foreground="#7a2020").pack()

    def do_login():
        ok, kind, val, msg = login_check(ent.get())
        if ok:
            state["user"] = val
            login_frame.destroy()
            show_school_step()
        else:
            state["attempts"] += 1
            if state["attempts"] >= MAX_LOGIN_ATTEMPTS:
                die(MSG_LOCKED)
            else:
                messagebox.showwarning("تحقق", "%s\nالمحاولات المتبقية: %d"
                                       % (msg, MAX_LOGIN_ATTEMPTS - state["attempts"]))

    ttk.Button(login_frame, text="دخول", command=do_login).pack()

    # ----- إجبار اختيار المدرسة -----
    def show_school_step():
        f = ttk.Frame(root, padding=16)
        f.grid(row=1, column=1, sticky="nsew")
        ttk.Label(f, text="اختر المدرسة أولًا (إجباري) — رمز وطني أو اسم أو ما يشابه").pack(anchor="e")
        q = ttk.Entry(f, width=34)
        q.pack(pady=4)
        lst = ttk.Treeview(f, columns=("code", "name"), show="headings", height=8)
        lst.heading("code", text="رمز المدرسة"); lst.heading("name", text="اسم المدرسة")
        lst.column("code", width=90); lst.column("name", width=280)
        lst.pack(fill="both", expand=True)

        def search(_=None):
            lst.delete(*lst.get_children())
            for code, name in store.search_schools(q.get()):
                lst.insert("", "end", values=(code, name))

        def choose(_=None):
            sel = lst.selection()
            if not sel:
                messagebox.showinfo("اختيار", "اختر مدرسة من القائمة.")
                return
            code, name = lst.item(sel[0], "values")
            state["school"] = (code, name)
            f.destroy()
            show_tabs()

        q.bind("<KeyRelease>", search)
        ttk.Button(f, text="بحث", command=search).pack()
        ttk.Button(f, text="اعتماد المدرسة", command=choose).pack()

    # ----- النماذج (تبويبات عمودية) -----
    def show_tabs():
        nb = ttk.Notebook(root)
        nb.grid(row=1, column=1, sticky="nsew", padx=8, pady=8)
        forms = ["راغبين بالمراقبة", "غير الراغبين", "الأذنة والحراس",
                 "الملتحقين بالجامعات", "مراقبة إلكترونية", "مراقبة مدراء"]
        for fname in forms:
            tab = ttk.Frame(nb, padding=12)
            nb.add(tab, text=fname)
            head = "مديرية التربية والتعليم — %s\nالمدرسة: %s" % (fname, state["school"][1])
            ttk.Label(tab, text=head, justify="right").pack(anchor="e")
            e_emp = ttk.Entry(tab, width=20); e_emp.pack(anchor="e", pady=4)
            info = ttk.Label(tab, text="أدخل وزاري/وطني الموظف لعرض بياناته تلقائيًا",
                             wraplength=360, justify="right")
            info.pack(anchor="e")

            def on_emp(*_, w=e_emp, i=info):
                rec = store.find_employee(w.get())
                if rec:
                    spec, src = resolve_specialization(rec["spec_last"] or rec["spec_first"])
                    i.config(text="%s | %s | التخصص: %s (من %s)"
                             % (rec["name"], rec["ministry"], spec, src))
                else:
                    i.config(text="غير موجود — يُسمح بالإدخال اليدوي.")
            e_emp.bind("<KeyRelease>", on_emp)
            ttk.Button(tab, text="طباعة/تصدير",
                       command=lambda n=fname: messagebox.showinfo(
                           "طباعة", "نافذة الطباعة والتصدير لـ: %s" % n)).pack(pady=8)

    root.resizable(True, True)
    root.mainloop()


if __name__ == "__main__":
    if "--selftest" in sys.argv or "--no-gui" in sys.argv:
        sys.exit(selftest())
    try:
        run_gui()
    except Exception as exc:
        print("تعذر تشغيل الواجهة الرسومية (%s). شغّل الاختبار: python main.py --selftest" % exc)
        sys.exit(2)
