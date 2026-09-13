# -*- coding: utf-8 -*-
"""VBA 版の「正解」を作る。

genba_norm.py の関数（parse_date / canon / parse_num / classify）をそのまま使い、
VBA 版と同じ列の表を CSV に出す。compare.py がこの CSV と VBA の出力を1セルずつ突き合わせる。

  python vba/oracle.py samples/raw/点検記録_2026-09.xlsx

ponytail: 行のループは genba_norm.norm_inspection() と同じ規則を写してある（十数行）。
          norm_inspection() を行単位の関数に割ったら、ここはそれを呼ぶだけにする。
"""
import csv
import io
import os
import re
import sys

HERE = os.path.dirname(os.path.abspath(__file__))
sys.path.insert(0, os.path.join(HERE, ".."))
import genba_norm as G  # noqa: E402

CHECK_KEYS = (u"掻き取り", u"異音", u"発泡", u"色")


def table(path):
    rows, _, reader = G.read_xlsx(path)
    base = os.path.splitext(os.path.basename(path))[0]
    m = re.search(r"(\d{4})-(\d{2})", base)
    dy, dm = (int(m.group(1)), int(m.group(2))) if m else (None, None)

    head1, head2 = rows[0], rows[1]
    cols = [(G.z2h(a).strip() + u" " + G.z2h(b).strip()).strip()
            for a, b in zip(head1, list(head2) + [u""] * len(head1))]
    check = [i for i, c in enumerate(cols) if any(k in c for k in CHECK_KEYS)]

    out = [[u"行", u"日付", u"担当", u"部署", u"汚泥厚"] + [cols[i] for i in check]]
    for r, row in enumerate(rows[2:], start=3):
        if not any((u"%s" % x).strip() for x in row):
            continue
        d = G.parse_date(row[0], dy, dm)
        num = G.parse_num(row[2])
        out.append([u"%d" % r,
                    d.isoformat() if d else u"読めない",
                    G.canon(u"担当", row[1]),
                    G.canon(u"部署", row[7]) if len(row) > 7 else u"",
                    u"%.1f" % num if num is not None else u"読めない"]
                   + [G.classify(row[i] if i < len(row) else u"")[0] for i in check])
    return out, reader


def main():
    src = sys.argv[1] if len(sys.argv) > 1 else os.path.join(
        HERE, "..", "samples", "raw", u"点検記録_2026-09.xlsx")
    out, reader = table(src)
    dst = os.path.splitext(src)[0] + u"_正解.csv"
    with io.open(dst, "w", encoding="utf-8-sig", newline="") as f:
        csv.writer(f).writerows(out)
    dates = sum(1 for r in out[1:] if r[1] != u"読めない")
    print(u"%s  %d行（日付確定 %d）  読み取り: %s" % (dst, len(out) - 1, dates, reader))


if __name__ == "__main__":
    main()
