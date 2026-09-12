# -*- coding: utf-8 -*-
u"""make_samples.py — 現場の帳票の「汚れ方」を再現したサンプルを作る

    python samples/make_samples.py

出るもの:
    samples/raw/点検記録_2026-09.xlsx   加圧浮上槽の日常点検表（1シート・30行）
    samples/raw/日報/*.txt              日報 20件

**実物は使えない。**前職の点検記録は社外秘。
そこで「実物がどう汚れているか」だけを写したサンプルを作る。
汚し方そのものが、現場を見たことがあるかどうかの証拠になる。

再現した汚れ（すべて実務で見かける形）:

  1. 日付の表記が1つの列で混ざる
     2026/9/1 ・ 2026-09-02 ・ 令和8年9月3日 ・ 9/4 ・ R8.9.5 ・ 26.9.6
     ＋ Excel が勝手にシリアル値へ変えてしまった行

  2. 日付セルが縦に結合されている
     1日ぶんの点検が3行あると、日付は先頭行にしか入らない。
     素朴に読むと3行のうち2行が「日付なし」になる

  3. ヘッダが2段で、上段が横に結合されている
     「加圧浮上槽」が3列にまたがり、その下に「汚泥厚」「発泡」「色」が並ぶ

  4. 人名・部署名・設備名の表記ゆれ
     山田／山田 太郎／山田太郎／ヤマダ
     第1製造部／第一製造部／製造1部／Ｍ１
     加圧浮上槽／浮上槽／加圧浮上

  5. 数値に単位と全角が混ざる
     １２．３ ・ " 12.3 " ・ 12.3℃ ・ 約12

  6. **点検結果の欄が「書いた」のか「書き忘れた」のか区別できない**
     異常なし／特になし／OK／レ／✓／良  ← 記入あり
     異常あり…／NG                     ← 異常
     要観察／経過観察／軽微             ← **条件付き許容。異常ではない**
     計画停止中／休転／対象外           ← **対象外。未記入ではない**
     （空欄）                          ← 未記入
     - ／ ー ／ ／ ／ N/A               ← **どちらとも取れる**

**6 がいちばん危ない。**空欄を黙って落とすと、残るのは「異常なし」だけになり、
「9月に異常は無かった」と読めてしまう。実際は点検されていない日がある。

依存: 標準ライブラリのみ（xlsx は zip + XML なので zipfile と ElementTree で書ける）
"""
import io
import os
import random
import zipfile

HERE = os.path.dirname(os.path.abspath(__file__))
RAW = os.path.join(HERE, "raw")

# --- 汚れの材料 -------------------------------------------------------------

TANTOU = [u"山田", u"山田 太郎", u"山田太郎", u"ヤマダ",
          u"佐藤", u"佐藤 花子", u"さとう"]
BUSHO = [u"第1製造部", u"第一製造部", u"製造1部", u"Ｍ１"]

# 記入あり・正常とみなせるもの
OK_MARKS = [u"異常なし", u"特になし", u"OK", u"レ", u"✓", u"良"]
# 「該当なし」なのか「未チェック」なのか決められないもの
AMBIG = [u"-", u"ー", u"／", u"N/A"]

IJOU = [u"異常あり：掻き取り機に異音", u"異常あり　汚泥が厚い", u"NG"]
# 稼働は継続しているが、そのままにはしない状態。**異常ではない**
COND = [u"要観察（発泡多め）", u"経過観察", u"軽微な付着あり"]
# 設備が止まっている。**空欄でも未記入ではない**
EXCL = [u"計画停止中", u"休転", u"対象外"]


def dirty_date(y, m, d, i):
    u"""同じ日付を、その行ごとに違う書き方で返す。1列の中で6通りが混ざる。"""
    forms = [
        u"%d/%d/%d" % (y, m, d),
        u"%04d-%02d-%02d" % (y, m, d),
        u"令和%d年%d月%d日" % (y - 2018, m, d),
        u"%d/%d" % (m, d),
        u"R%d.%d.%d" % (y - 2018, m, d),
        u"%02d.%d.%d" % (y % 100, m, d),
    ]
    return forms[i % len(forms)]


def excel_serial(y, m, d):
    u"""1900年日付システムのシリアル値。1900年うるう年バグぶんの +1 を含む。"""
    import datetime
    return (datetime.date(y, m, d) - datetime.date(1899, 12, 30)).days


def dirty_num(v, i):
    z = u"０１２３４５６７８９．"
    if i % 4 == 0:
        return u"".join(z[u"0123456789.".index(c)] if c in u"0123456789." else c
                        for c in u"%.1f" % v)          # 全角
    if i % 4 == 1:
        return u"  %.1f " % v                          # 前後に空白
    if i % 4 == 2:
        return u"%.1f℃" % v                            # 単位つき
    return u"約%d" % int(v)                            # 概数


# --- 最小の xlsx 書き出し ---------------------------------------------------
#
# openpyxl を入れれば数行で済む。入れないのは、この repo が
# 「依存を足さない」方針で作られているため（PILLAR_2_DEV.md §8）。
# 読む側（genba_norm.py）も同じ理由で標準ライブラリだけで書いてある。

def _esc(s):
    return (s.replace(u"&", u"&amp;").replace(u"<", u"&lt;").replace(u">", u"&gt;"))


def col_name(n):
    s = u""
    while n >= 0:
        s = chr(ord("A") + n % 26) + s
        n = n // 26 - 1
    return s


def write_xlsx(path, rows, merges):
    u"""rows: [[cell, ...], ...]  cell は文字列か ('n', 値, スタイル番号)
    merges: ["A3:A5", ...]
    """
    shared, order = {}, []

    def sid(s):
        if s not in shared:
            shared[s] = len(order)
            order.append(s)
        return shared[s]

    body = []
    for r, row in enumerate(rows, start=1):
        cells = []
        for c, v in enumerate(row):
            if v is None or v == u"":
                continue                        # 空欄はセルごと書かない（実物もこうなる）
            ref = u"%s%d" % (col_name(c), r)
            if isinstance(v, tuple):            # 数値
                _, num, style = v
                st = u' s="%d"' % style if style else u""
                cells.append(u'<c r="%s"%s><v>%s</v></c>' % (ref, st, num))
            else:
                cells.append(u'<c r="%s" t="s"><v>%d</v></c>' % (ref, sid(v)))
        body.append(u'<row r="%d">%s</row>' % (r, u"".join(cells)))

    merge_xml = u""
    if merges:
        merge_xml = (u'<mergeCells count="%d">%s</mergeCells>' %
                     (len(merges),
                      u"".join(u'<mergeCell ref="%s"/>' % m for m in merges)))

    sheet = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
             u'<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
             u'<sheetData>%s</sheetData>%s</worksheet>' % (u"".join(body), merge_xml))

    sst = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
           u'<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
           u'count="%d" uniqueCount="%d">%s</sst>'
           % (len(order), len(order),
              u"".join(u"<si><t xml:space=\"preserve\">%s</t></si>" % _esc(s) for s in order)))

    # s="1" を「日付」書式にする。これが無いとシリアル値がただの数に見える
    styles = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              u'<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
              u'<numFmts count="0"/>'
              u'<fonts count="1"><font><sz val="11"/><name val="Calibri"/></font></fonts>'
              u'<fills count="1"><fill><patternFill patternType="none"/></fill></fills>'
              u'<borders count="1"><border/></borders>'
              u'<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
              u'<cellXfs count="2">'
              u'<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
              u'<xf numFmtId="14" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>'
              u'</cellXfs>'
              u'<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>'
              u'<dxfs count="0"/>'
              u'</styleSheet>')

    ct = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          u'<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          u'<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          u'<Default Extension="xml" ContentType="application/xml"/>'
          u'<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          u'<Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>'
          u'<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>'
          u'<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          u'</Types>')

    rels = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
            u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
            u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
            u'</Relationships>')

    wb = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          u'<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" '
          u'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
          u'<sheets><sheet name="日常点検" sheetId="1" r:id="rId1"/></sheets></workbook>')

    wbrels = (u'<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
              u'<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
              u'<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet1.xml"/>'
              u'<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>'
              u'<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>'
              u'</Relationships>')

    with zipfile.ZipFile(path, "w", zipfile.ZIP_DEFLATED) as z:
        z.writestr("[Content_Types].xml", ct.encode("utf-8"))
        z.writestr("_rels/.rels", rels.encode("utf-8"))
        z.writestr("xl/workbook.xml", wb.encode("utf-8"))
        z.writestr("xl/_rels/workbook.xml.rels", wbrels.encode("utf-8"))
        z.writestr("xl/sharedStrings.xml", sst.encode("utf-8"))
        z.writestr("xl/styles.xml", styles.encode("utf-8"))
        z.writestr("xl/worksheets/sheet1.xml", sheet.encode("utf-8"))


# --- 点検記録を組み立てる ---------------------------------------------------

def build_inspection():
    rng = random.Random(20260912)          # 毎回同じものが出るように固定する

    rows = [
        # 1行目: 上段ヘッダ。「加圧浮上槽」が C:E に、「活性汚泥槽」が F:G にまたがる
        [u"日付", u"担当", u"加圧浮上槽", u"", u"", u"活性汚泥槽", u"", u"部署", u"備考"],
        # 2行目: 下段ヘッダ
        [u"", u"", u"汚泥厚(mm)", u"掻き取り", u"異音", u"発泡", u"色", u"", u""],
    ]
    merges = ["A1:A2", "B1:B2", "C1:E1", "F1:G1", "H1:H2", "I1:I2"]

    r = 3
    day = 1
    i = 0
    log = []                                # 答え合わせ用（正解を控えておく）

    while day <= 12:
        n = rng.choice([1, 2, 3])           # 1日に1〜3回点検する
        top = r
        for j in range(n):
            # 日付は先頭行にしか入れない。残りは結合セルで空になる（実物と同じ）
            if j == 0:
                if i % 7 == 6:
                    cell = ("n", excel_serial(2026, 9, day), 1)   # シリアル値の行
                else:
                    cell = dirty_date(2026, 9, day, i)
            else:
                cell = u""

            atsu = dirty_num(10 + rng.random() * 6, i + j)

            # 掻き取り・異音・発泡・色 の4項目。ここに「未記入」と「判別不能」を混ぜる
            vals = []
            for col in range(4):
                p = rng.random()
                if p < 0.10:
                    vals.append(u"")                       # 未記入
                elif p < 0.20:
                    vals.append(rng.choice(AMBIG))         # どちらとも取れる
                elif p < 0.26:
                    vals.append(rng.choice(IJOU))          # 異常
                elif p < 0.34:
                    vals.append(rng.choice(COND))          # 条件付き許容（異常ではない）
                elif p < 0.40:
                    vals.append(rng.choice(EXCL))          # 対象外（未記入ではない）
                else:
                    vals.append(rng.choice(OK_MARKS))      # 記入あり・正常

            rows.append([cell, rng.choice(TANTOU), atsu,
                         vals[0], vals[1], vals[2], vals[3],
                         rng.choice(BUSHO),
                         u"" if rng.random() < 0.8 else u"次回、掻き取り機の増し締め"])
            log.append((day, vals))
            i += 1
            r += 1
        if n > 1:
            merges.append("A%d:A%d" % (top, top + n - 1))
        day += 1

    return rows, merges, log


NIPPOU = u"""{date} 日報
部署: {busho}
記入: {tantou}

■ 設備
{setsubi} の点検を実施。{kekka}

■ 生産
{seihin} を {kazu} 個 加工。段取り替え {dan} 回。

■ 連絡
{renraku}
"""

SETSUBI = [u"加圧浮上槽", u"浮上槽", u"加圧浮上", u"活性汚泥槽", u"汚泥槽"]
RENRAKU = [u"特になし", u"", u"-",
           u"ラインが止まった場合は、まず保全係へ内線。不在なら製造課長へ。",
           u"掻き取り機の増し締めを次回までに実施。"]


def build_nippou():
    rng = random.Random(777)
    out = []
    for k in range(20):
        day = k + 1
        d = dirty_date(2026, 9, day, k)
        out.append((u"日報_%02d.txt" % day, NIPPOU.format(
            date=d,
            busho=rng.choice(BUSHO),
            tantou=rng.choice(TANTOU),
            setsubi=rng.choice(SETSUBI),
            kekka=rng.choice([u"異常なし。", u"特になし", u"", u"発泡が多い。要観察。"]),
            seihin=rng.choice([u"A-100", u"Ａ－１００", u"A100"]),
            kazu=rng.randint(80, 400),
            dan=rng.randint(0, 3),
            renraku=rng.choice(RENRAKU),
        )))
    return out


def main():
    os.makedirs(os.path.join(RAW, u"日報"), exist_ok=True)

    rows, merges, log = build_inspection()
    xlsx = os.path.join(RAW, u"点検記録_2026-09.xlsx")
    write_xlsx(xlsx, rows, merges)

    # 答え合わせ用の内訳を数えておく（正規化がここに一致すれば正しい）
    flat = [x for _, v in log for x in v]
    n_blank = sum(1 for x in flat if x == u"")
    n_amb = sum(1 for x in flat if x in AMBIG)
    n_ijou = sum(1 for x in flat if x in IJOU)
    n_cond = sum(1 for x in flat if x in COND)
    n_excl = sum(1 for x in flat if x in EXCL)
    n_ok = sum(1 for x in flat if x in OK_MARKS)

    for name, text in build_nippou():
        io.open(os.path.join(RAW, u"日報", name), "w", encoding="utf-8").write(text)

    print(u"点検記録: %s（%d行・結合%d箇所）" % (xlsx, len(rows), len(merges)))
    print(u"  点検欄の内訳: 記入あり %d / 異常 %d / 条件付き許容 %d / 対象外 %d / 未記入 %d / 判別不能 %d"
          % (n_ok, n_ijou, n_cond, n_excl, n_blank, n_amb))
    print(u"日報: %s/日報/*.txt  20件" % RAW)
    print(u"")
    print(u"**未記入 %d と判別不能 %d を黙って落とすと、「異常なし」だけが残る。**" % (n_blank, n_amb))


if __name__ == "__main__":
    main()
