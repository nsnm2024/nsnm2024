# -*- coding: utf-8 -*-
"""التخزين المحلي: DBF (صيغة dBASE III كما الأصل) + Excel محمي + قفل السجل + سجل التدقيق.

منفذ هنا بدون مكتبات خارجية للـ DBF (ترميز cp1256 للعربية)، وبopenpyxl لحماية Excel.
"""
import json
import os
import struct
import time

try:
    import openpyxl
    from openpyxl.worksheet.protection import SheetProtection
except ImportError:  # pragma: no cover
    openpyxl = None

DBF_TYPE_CHAR = b"03"           # dBASE III بدون memo
CP = "cp1256"                   # العربية في ملفات الأصل


def _dbf_header(fields, nrec):
    """يبني ترويسة DBF لعدد حقول text بطول موحد (dBASE III)."""
    header_len = 32 + 32 * len(fields) + 1
    record_len = 1 + sum(size for _, size in fields)   # بايت الحذف + الحقول
    yy = time.localtime().tm_year - 1900
    mm = time.localtime().tm_mon
    dd = time.localtime().tm_mday
    out = bytearray(32)          # الترويسة الأساسية 32 بايت
    out[0:1] = b"\x03"                       # [0] نوع الملف: dBASE III بدون memo
    out[1:4] = bytes([yy % 256, mm, dd])     # [1..3] تاريخ آخر تحديث (y,m,d)
    out[4:8] = struct.pack("<I", nrec)         # [4..7] عدد السجلات
    out[8:10] = struct.pack("<H", header_len)  # [8..9] طول الترويسة
    out[10:12] = struct.pack("<H", record_len) # [10..11] طول السجل
    # [12..31] محجوز = أصفاف
    for name, size in fields:
        fd = bytearray(32)
        fd[0:11] = name.encode(CP)[:10].ljust(11, b"\x00")
        fd[11:12] = b"C"                      # نوع الحقل: Character (حرف واحد خام)
        fd[16:17] = bytes([size])              # طول الحقل (بايت واحد خام)
        out += fd
    out += b"\x0d"
    return bytes(out)


def write_dbf(path, fields, records):
    """fields=[('NAME',10),...] records=[{...}] — يكتب أرقامًا كنص كما في الأصل."""
    data = bytearray(_dbf_header(fields, len(records)))
    widths = dict(fields)
    for rec in records:
        data += b" "  # غير محذوف
        for name, size in fields:
            val = str(rec.get(name, "") or "").encode(CP, errors="replace")
            data += val.ljust(size, b" ")[:size]
    data += b"\x1a"
    with open(path, "wb") as f:
        f.write(bytes(data))


def read_dbf(path):
    """قراءة بسيطة لملفات هذا الموديول (نص فقط)."""
    with open(path, "rb") as f:
        raw = f.read()
    nrec = struct.unpack("<I", raw[4:8])[0]
    hlen = struct.unpack("<H", raw[8:10])[0]
    rlen = struct.unpack("<H", raw[10:12])[0]
    fields = []
    p = 32
    while p < hlen - 1 and raw[p] != 0x0D:
        name = raw[p:p + 11].split(b"\x00")[0].decode(CP, errors="replace")
        size = raw[p + 16]
        fields.append((name, size))
        p += 32
    recs = []
    start = hlen
    for i in range(nrec):
        off = start + i * rlen
        row, q = {}, off + 1
        for name, size in fields:
            row[name] = raw[q:q + size].decode(CP, errors="replace").strip()
            q += size
        recs.append(row)
    return fields, recs


class LockedStore:
    """مخزن JSON للسجلات المقفلة + سجل التدقيق (من أدخل ومتى)."""

    def __init__(self, base_dir):
        self.dir = os.path.join(base_dir, "out")
        os.makedirs(self.dir, exist_ok=True)
        self.records_file = os.path.join(self.dir, "records.json")
        self.audit_file = os.path.join(self.dir, "audit.log")
        self.records = {}
        if os.path.exists(self.records_file):
            try:
                with open(self.records_file, "r", encoding="utf-8") as f:
                    self.records = json.load(f)
            except Exception:
                self.records = {}

    def add(self, key, payload, entered_by):
        if key in self.records:
            raise PermissionError("السجل مقفل ولا يمكن تعديله بعد الإدخال.")
        self.records[key] = payload
        self._save()
        self._audit(entered_by, key)
        return payload

    def get(self, key):
        return self.records.get(key)

    def all(self):
        return dict(self.records)

    def export_xlsx_protected(self, headers, rows, path, password=""):
        """حفظ Excel محمي من التعديل (كما طُلب للنموذج الإلكتروني)."""
        if openpyxl is None:
            raise RuntimeError("openpyxl غير مثبت.")
        wb = openpyxl.Workbook()
        ws = wb.active
        ws.title = "البيانات"
        ws.append(headers)
        for r in rows:
            ws.append(r)
        ws.protection = SheetProtection(sheet=True, formatCells=False,
                                        insertRows=False, deleteRows=False)
        wb.save(path)
        return path

    def _save(self):
        with open(self.records_file, "w", encoding="utf-8") as f:
            json.dump(self.records, f, ensure_ascii=False, indent=1)

    def _audit(self, entered_by, key):
        stamp = time.strftime("%Y-%m-%d %H:%M:%S")
        with open(self.audit_file, "a", encoding="utf-8") as f:
            f.write("%s | مدخل البيانات: %s | السجل: %s\n" % (stamp, entered_by, key))
