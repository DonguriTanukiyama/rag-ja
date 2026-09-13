Attribute VB_Name = "GenbaNorm"
Option Explicit
'
' 現場の点検表（Excel）を、集計できる表に整える。
' rag-ja/genba_norm.py（Python）と同じ規則で書いてある。
' 結果は vba/compare.py で、Python 版の出力と1セルずつ突き合わせる。
'
' 使い方
'   1. VBA エディタ（Alt+F11）→ ファイル → ファイルのインポート → このファイル
'   2. まず SelfTest を実行（Alt+F8）。「SelfTest OK」と出ることを確かめる
'   3. Normalize を実行 → 点検表の xlsx を選ぶ
'   4. 同じフォルダに「<元の名前>_VBA.csv」ができる
'
' 元のファイルは読み取り専用で開く。書き換えない。
'
' ============================================================
' 書くときに避けた罠
' ============================================================
'
' ・StrConv(s, vbNarrow) は使わない。カタカナまで半角カナになる（ポンプ → ﾎﾟﾝﾌﾟ）。
'   英数記号（U+FF01 から U+FF5E）と全角スペースだけを半角にする（Z2H）。
' ・VBScript.RegExp は使わない。Microsoft が VBScript の廃止を発表している。
'   日付と数値は1文字ずつ読む。
' ・Scripting.Dictionary も使わない。対応表は配列で持つ。
' ・このファイルは Shift-JIS（cp932）で保存する。VBA エディタは UTF-8 を読めない。
'   cp932 に無い文字（チェックマーク・マイナス記号・全角ダッシュ）は ChrW で作る。
' ・CSV に書く前にセルを文字列書式にする。しないと Excel が "2026-09-01" を
'   日付に変え、"2026/09/01" で書き出す。突き合わせが全部ずれる。
' ・波ダッシュ（U+301C）は cp932 を往復すると全角チルダ（U+FF5E）に化ける。
'   このファイルでは使わない（範囲は「から」と書く）。
' ・Trim$ は全角スペースを落とさない。Python の strip() と揃えるため StripW を使う。
'
' ============================================================
' 既知の限界
' ============================================================
'
' ・1904年日付システム（Mac 由来）には未対応。Python 版と同じ。
' ・Z2H は英数記号だけ。Python の NFKC が変える他の文字（半角カナ・丸数字など）は変えない。
' ・年が100未満の日付（0001/1/1 のような入力）は読まない。Python 版は読む。
' ・汚泥厚の丸めは Format$。小数第2位が5ちょうどの値で Python と食い違うことがある。

' --- 対応表（表記ゆれ）-------------------------------------------------------
' 各要素は Array(代表名, 別名1, 別名2, ...)。
' **現場ごとに違う。必ず書き換えて使う。**（genba_norm.py の ALIASES と同じ中身）

Private Function AliasTable(ByVal kind As String) As Variant
    Select Case kind
    Case "設備"
        AliasTable = Array( _
            Array("加圧浮上槽", "浮上槽", "加圧浮上", "加圧浮上そう"), _
            Array("活性汚泥槽", "汚泥槽", "活性汚泥"))
    Case "部署"
        AliasTable = Array( _
            Array("第1製造部", "第一製造部", "製造1部", "製造一部", "M1", "Ｍ１"))
    Case "担当"
        AliasTable = Array( _
            Array("山田太郎", "山田", "山田 太郎", "ヤマダ", "やまだ"), _
            Array("佐藤花子", "佐藤", "佐藤 花子", "さとう", "サトウ"))
    Case Else
        AliasTable = Array()
    End Select
End Function

' --- 点検欄の判定に使う語 -----------------------------------------------------

Private Function OkMarks() As Variant
    OkMarks = Array("異常なし", "異常無し", "特になし", "特に無し", _
                    "OK", "ok", "レ", ChrW(&H2713), "良", "可", "○", "〇")
End Function

Private Function AmbigMarks() As Variant
    AmbigMarks = Array("-", "ー", ChrW(&H2212), ChrW(&H2014), "/", "／", "N/A", "n/a", "NA")
End Function

Private Function ExcludedHead() As Variant
    ExcludedHead = Array("対象外", "点検対象外", "計画停止", "停止中", "休転", "非稼働", "停止")
End Function

Private Function CondHead() As Variant
    CondHead = Array("要観察", "経過観察", "継続監視", "軽微", "稼働可", "要確認", "要注意")
End Function

Private Function AbnormalHead() As Variant
    AbnormalHead = Array("異常", "NG", "不良", "要修理", "故障")
End Function

' --- 入口 -------------------------------------------------------------------

Public Sub Normalize()
    Dim f As Variant
    f = Application.GetOpenFilename("Excel ファイル (*.xlsx;*.xlsm),*.xlsx;*.xlsm", , "点検表を選んでください")
    If VarType(f) = vbBoolean Then Exit Sub
    NormalizeFile CStr(f)
End Sub

Public Sub NormalizeFile(ByVal path As String)
    Dim wbSrc As Workbook, wbOut As Workbook, ws As Worksheet
    Dim lastRow As Long, lastCol As Long, r As Long, c As Long, k As Long, n As Long
    Dim dy As Long, dm As Long, stem As String, outPath As String
    Dim checkCols() As Long, nCheck As Long, heads() As String
    Dim res() As Variant, dateStr As String, ok As Boolean, num As Double
    Dim nDate As Long, tally(0 To 5) As Long, kind As String
    Dim alerts As Boolean, msg As String

    alerts = Application.DisplayAlerts
    On Error GoTo Fail

    Set wbSrc = Workbooks.Open(Filename:=path, ReadOnly:=True, UpdateLinks:=0)
    Set ws = wbSrc.Worksheets(1)
    With ws.UsedRange
        lastRow = .Row + .Rows.Count - 1
        lastCol = .Column + .Columns.Count - 1
    End With

    ' 見出しは2段。上段と下段をつなげて1つの名前にする（結合セルは展開して読む）
    ReDim heads(1 To lastCol)
    ReDim checkCols(1 To lastCol)
    For c = 1 To lastCol
        heads(c) = StripW(StripW(Z2H(CStr(CellVal(ws, 1, c)))) & " " & _
                          StripW(Z2H(CStr(CellVal(ws, 2, c)))))
        If InStr(heads(c), "掻き取り") > 0 Or InStr(heads(c), "異音") > 0 Or _
           InStr(heads(c), "発泡") > 0 Or InStr(heads(c), "色") > 0 Then
            nCheck = nCheck + 1
            checkCols(nCheck) = c
        End If
    Next

    ' ファイル名の「2026-09」から、年の無い日付（9/4 など）の年を補う
    stem = Mid$(path, InStrRev(path, "\") + 1)
    If InStrRev(stem, ".") > 0 Then stem = Left$(stem, InStrRev(stem, ".") - 1)
    YearMonth stem, dy, dm

    ReDim res(0 To lastRow, 1 To 5 + nCheck)
    res(0, 1) = "行": res(0, 2) = "日付": res(0, 3) = "担当": res(0, 4) = "部署": res(0, 5) = "汚泥厚"
    For k = 1 To nCheck
        res(0, 5 + k) = heads(checkCols(k))
    Next

    For r = 3 To lastRow
        If RowHasValue(ws, r, lastCol) Then
            n = n + 1
            res(n, 1) = CStr(r)

            dateStr = ParseDate(CellVal(ws, r, 1), dy, dm)
            If Len(dateStr) > 0 Then
                res(n, 2) = dateStr
                nDate = nDate + 1
            Else
                res(n, 2) = "読めない"
            End If

            res(n, 3) = Canon("担当", CellVal(ws, r, 2))
            If lastCol >= 8 Then
                res(n, 4) = Canon("部署", CellVal(ws, r, 8))
            Else
                res(n, 4) = ""
            End If

            num = ParseNum(CellVal(ws, r, 3), ok)
            If ok Then
                res(n, 5) = Format$(num, "0.0")
            Else
                res(n, 5) = "読めない"
            End If

            For k = 1 To nCheck
                kind = Classify(CellVal(ws, r, checkCols(k)))
                res(n, 5 + k) = kind
                tally(KindIndex(kind)) = tally(KindIndex(kind)) + 1
            Next
        End If
    Next

    ' 書き出す。**先に文字列書式にする**（日付や数値に化けさせない）
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    With wbOut.Worksheets(1).Range("A1").Resize(n + 1, 5 + nCheck)
        .NumberFormat = "@"
        .Value = res
    End With

    outPath = Left$(path, InStrRev(path, ".") - 1) & "_VBA.csv"
    Application.DisplayAlerts = False
    wbOut.SaveAs Filename:=outPath, FileFormat:=62   ' 62 = CSV UTF-8
    wbOut.Close SaveChanges:=False
    Set wbOut = Nothing
    Application.DisplayAlerts = alerts
    wbSrc.Close SaveChanges:=False
    Set wbSrc = Nothing

    MsgBox "整形しました。" & vbCrLf & vbCrLf & _
           "行数: " & n & vbCrLf & _
           "日付が確定した行: " & nDate & " / " & n & vbCrLf & _
           "点検欄: 記入あり " & tally(0) & " / 異常 " & tally(1) & _
           " / 条件付き許容 " & tally(2) & " / 対象外 " & tally(3) & _
           " / 未記入 " & tally(4) & " / 判別不能 " & tally(5) & vbCrLf & vbCrLf & _
           "出力: " & outPath, vbInformation, "GenbaNorm"
    Exit Sub

Fail:
    msg = Err.Number & ": " & Err.Description
    Application.DisplayAlerts = alerts
    On Error Resume Next
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=False
    If Not wbSrc Is Nothing Then wbSrc.Close SaveChanges:=False
    MsgBox "止まりました。" & vbCrLf & msg, vbExclamation, "GenbaNorm"
End Sub

' --- 自己確認 -----------------------------------------------------------------
' genba_norm.py --demo と同じ観点。**Excel の地域設定が違っても同じ答えを返すか**もここで分かる。

Public Sub SelfTest()
    Dim fails As String, ok As Boolean

    Check fails, "令和", ParseDate("令和8年9月3日", 0, 0), "2026-09-03"
    Check fails, "R8.9.5", ParseDate("R8.9.5", 0, 0), "2026-09-05"
    Check fails, "26.9.6", ParseDate("26.9.6", 0, 0), "2026-09-06"
    Check fails, "2026-09-02", ParseDate("2026-09-02", 0, 0), "2026-09-02"
    Check fails, "9/4 年を補う", ParseDate("9/4", 2026, 9), "2026-09-04"
    Check fails, "9/4 年が無い", ParseDate("9/4", 0, 0), ""
    Check fails, "シリアル値", ParseDate(46266#, 0, 0), "2026-09-01"
    Check fails, "シリアル文字", ParseDate("46266", 0, 0), "2026-09-01"
    Check fails, "2月30日", ParseDate("2026/2/30", 0, 0), ""
    Check fails, "12.5 は日付でない", ParseDate("12.5", 0, 0), ""

    Check fails, "異常なし", Classify("異常なし"), "記入あり"
    Check fails, "異常あり", Classify("異常あり：異音"), "異常"
    Check fails, "要観察", Classify("要観察（発泡多め）"), "条件付き許容"
    Check fails, "計画停止中", Classify("計画停止中"), "対象外"
    Check fails, "全角スラッシュ", Classify("／"), "判別不能"
    Check fails, "空欄", Classify(""), "未記入"
    Check fails, "全角空白だけ", Classify(ChrW(&H3000)), "未記入"
    Check fails, "NG", Classify("NG"), "異常"
    Check fails, "チェックマーク", Classify(ChrW(&H2713)), "記入あり"

    Check fails, "全角数値", Format$(ParseNum("１２．３", ok), "0.0"), "12.3"
    Check fails, "約12", Format$(ParseNum("約12", ok), "0.0"), "12.0"
    Check fails, "単位つき", Format$(ParseNum("12.3℃", ok), "0.0"), "12.3"
    Call ParseNum("なし", ok)
    Check fails, "数値なし", IIf(ok, "T", "F"), "F"

    Check fails, "カタカナを半角にしない", Z2H("ポンプＭ１"), "ポンプM1"
    Check fails, "部署の別名", Canon("部署", "Ｍ１"), "第1製造部"
    Check fails, "担当の別名", Canon("担当", "さとう"), "佐藤花子"
    Check fails, "対応表に無い", Canon("担当", "鈴木"), "鈴木"

    If Len(fails) = 0 Then
        MsgBox "SelfTest OK", vbInformation, "GenbaNorm"
    Else
        MsgBox "SelfTest NG" & vbCrLf & fails, vbExclamation, "GenbaNorm"
    End If
End Sub

Private Sub Check(ByRef fails As String, ByVal label As String, ByVal got As String, ByVal want As String)
    If got <> want Then fails = fails & label & ": 期待 [" & want & "] / 実際 [" & got & "]" & vbCrLf
End Sub

' --- 判定 -------------------------------------------------------------------

Private Function Classify(ByVal v As Variant) As String
    Dim s As String
    s = StripW(CStr(v))
    If Len(s) = 0 Then Classify = "未記入": Exit Function
    If InListZ(Z2H(s), AmbigMarks(), False) Then Classify = "判別不能": Exit Function
    ' 正常の語を先に見る。「異常なし」は「異常」で始まるので、後に置くと異常に化ける
    If InListZ(Z2H(RStripChar(s, "。")), OkMarks(), True) Then Classify = "記入あり": Exit Function
    If StartsAny(s, ExcludedHead()) Then Classify = "対象外": Exit Function
    ' 条件付き許容も異常より先に見る
    If StartsAny(s, CondHead()) Then Classify = "条件付き許容": Exit Function
    If StartsAny(s, AbnormalHead()) Then Classify = "異常": Exit Function
    Classify = "記入あり"
End Function

Private Function KindIndex(ByVal kind As String) As Long
    Select Case kind
    Case "記入あり": KindIndex = 0
    Case "異常": KindIndex = 1
    Case "条件付き許容": KindIndex = 2
    Case "対象外": KindIndex = 3
    Case "未記入": KindIndex = 4
    Case Else: KindIndex = 5
    End Select
End Function

Private Function Canon(ByVal kind As String, ByVal v As Variant) As String
    Dim s As String, e As Variant, i As Long
    s = Z2H(StripW(CStr(v)))
    For Each e In AliasTable(kind)
        If s = e(0) Or s = Z2H(CStr(e(0))) Then Canon = e(0): Exit Function
        For i = 1 To UBound(e)
            If s = e(i) Then Canon = e(0): Exit Function
        Next
    Next
    Canon = s
End Function

' --- 日付 -------------------------------------------------------------------
' 受ける形: 2026/9/1  2026-09-01  令和8年9月3日  R8.9.5  26.9.6  9/4  シリアル値
' 読めなければ "" を返す（捨てない。呼び出し側が「読めない」と書く）

Private Function ParseDate(ByVal v As Variant, ByVal dy As Long, ByVal dm As Long) As String
    Dim s As String, t As String, a As String, b As String, c As String, n As Double

    If VarType(v) = vbDouble Or VarType(v) = vbDate Then
        n = CDbl(v)
        If n >= 10000# And n < 100000# Then ParseDate = Format$(CDate(Int(n)), "yyyy-mm-dd")
        Exit Function
    End If

    s = StripW(Z2H(CStr(v)))
    If Len(s) = 0 Then Exit Function

    If IsSerialText(s) Then
        ParseDate = Format$(CDate(Int(Val(s))), "yyyy-mm-dd")
        Exit Function
    End If

    ' 令和8年9月3日 / R8.9.5
    t = ""
    If Left$(s, 2) = "令和" Then
        t = LStripW(Mid$(s, 3))
    ElseIf Left$(s, 1) = "R" Or Left$(s, 1) = "r" Then
        t = LStripW(Mid$(s, 2))
    End If
    If Len(t) > 0 Then
        If Read3(t, ".年/-", ".月/-", True, a, b, c) Then
            If Len(a) <= 2 And Len(b) <= 2 And Len(c) <= 2 Then
                ParseDate = MkDate(2018 + CLng(a), CLng(b), CLng(c))
                Exit Function
            End If
        End If
    End If

    ' 2026/9/1  2026-09-01  2026年9月1日
    If Read3(s, "/-.年", "/-.月", True, a, b, c) Then
        If Len(a) = 4 And Len(b) <= 2 And Len(c) <= 2 Then
            ParseDate = MkDate(CLng(a), CLng(b), CLng(c))
            Exit Function
        End If
    End If

    ' 26.9.6（2桁の年）。「12.5」のような2つ区切りは日付と見ない
    If Read3(s, ".-/", ".-/", False, a, b, c) Then
        If Len(a) = 2 And Len(b) <= 2 And Len(c) <= 2 Then
            ParseDate = MkDate(2000 + CLng(a), CLng(b), CLng(c))
            Exit Function
        End If
    End If

    ' 9/4（年が無い）。ファイル名から年を補えるときだけ
    If dy > 0 Then
        If Read2(s, "/-月", a, b) Then
            If Len(a) <= 2 And Len(b) <= 2 Then
                ParseDate = MkDate(dy, CLng(a), CLng(b))
                Exit Function
            End If
        End If
    End If
End Function

Private Function MkDate(ByVal y As Long, ByVal m As Long, ByVal d As Long) As String
    Dim dt As Date
    If y < 100 Or y > 9999 Or m < 1 Or m > 12 Or d < 1 Or d > 31 Then Exit Function
    dt = DateSerial(y, m, d)
    ' DateSerial は 2月30日を 3月2日に繰り上げる。繰り上がったら日付ではない
    If Year(dt) = y And Month(dt) = m And Day(dt) = d Then MkDate = Format$(dt, "yyyy-mm-dd")
End Function

Private Function IsSerialText(ByVal s As String) As Boolean
    Dim p As Long, a As String, b As String
    p = 1
    a = Digits(s, p)
    If Len(a) <> 5 Then Exit Function
    If p > Len(s) Then IsSerialText = True: Exit Function
    If Mid$(s, p, 1) <> "." Then Exit Function
    p = p + 1
    b = Digits(s, p)
    IsSerialText = (Len(b) > 0 And p > Len(s))
End Function

' 数字・区切り・数字・区切り・数字・[日] を先頭から末尾まで読む
Private Function Read3(ByVal s As String, ByVal sep1 As String, ByVal sep2 As String, _
                       ByVal nichi As Boolean, ByRef a As String, ByRef b As String, _
                       ByRef c As String) As Boolean
    Dim p As Long
    p = 1
    a = Digits(s, p)
    If Len(a) = 0 Or p > Len(s) Then Exit Function
    If InStr(1, sep1, Mid$(s, p, 1), vbBinaryCompare) = 0 Then Exit Function
    p = p + 1
    b = Digits(s, p)
    If Len(b) = 0 Or p > Len(s) Then Exit Function
    If InStr(1, sep2, Mid$(s, p, 1), vbBinaryCompare) = 0 Then Exit Function
    p = p + 1
    c = Digits(s, p)
    If Len(c) = 0 Then Exit Function
    If p > Len(s) Then Read3 = True: Exit Function
    If nichi And Mid$(s, p) = "日" Then Read3 = True
End Function

' 数字・区切り・数字・[日]
Private Function Read2(ByVal s As String, ByVal sep As String, _
                       ByRef a As String, ByRef b As String) As Boolean
    Dim p As Long
    p = 1
    a = Digits(s, p)
    If Len(a) = 0 Or p > Len(s) Then Exit Function
    If InStr(1, sep, Mid$(s, p, 1), vbBinaryCompare) = 0 Then Exit Function
    p = p + 1
    b = Digits(s, p)
    If Len(b) = 0 Then Exit Function
    If p > Len(s) Then Read2 = True: Exit Function
    If Mid$(s, p) = "日" Then Read2 = True
End Function

Private Sub YearMonth(ByVal stem As String, ByRef dy As Long, ByRef dm As Long)
    Dim i As Long
    dy = 0: dm = 0
    For i = 5 To Len(stem) - 2
        If Mid$(stem, i, 1) = "-" Then
            If AllDigits(Mid$(stem, i - 4, 4)) And AllDigits(Mid$(stem, i + 1, 2)) Then
                dy = CLng(Mid$(stem, i - 4, 4))
                dm = CLng(Mid$(stem, i + 1, 2))
                Exit Sub
            End If
        End If
    Next
End Sub

' --- 数値 -------------------------------------------------------------------
' '１２．３'  ' 12.3 '  '12.3℃'  '約12' → 12.3 / 12。最初に出てくる数を読む

Private Function ParseNum(ByVal v As Variant, ByRef ok As Boolean) As Double
    Dim s As String, p As Long, st As Long, ch As String, t As String
    ok = False
    If VarType(v) = vbDouble Then ParseNum = CDbl(v): ok = True: Exit Function

    s = StripW(Z2H(CStr(v)))
    For p = 1 To Len(s)
        ch = Mid$(s, p, 1)
        If IsDigitCh(ch) Then st = p: Exit For
        If ch = "-" And p < Len(s) Then
            If IsDigitCh(Mid$(s, p + 1, 1)) Then st = p: Exit For
        End If
    Next
    If st = 0 Then Exit Function

    p = st
    If Mid$(s, p, 1) = "-" Then t = "-": p = p + 1
    t = t & Digits(s, p)
    If p < Len(s) Then
        If Mid$(s, p, 1) = "." And IsDigitCh(Mid$(s, p + 1, 1)) Then
            p = p + 1
            t = t & "." & Digits(s, p)
        End If
    End If
    ParseNum = Val(t)
    ok = True
End Function

' --- 文字の下ごしらえ ---------------------------------------------------------

' 全角の英数記号（U+FF01 から U+FF5E）と全角スペースだけを半角にする。カタカナは触らない
Private Function Z2H(ByVal s As String) As String
    Dim i As Long, code As Long
    For i = 1 To Len(s)
        code = AscW(Mid$(s, i, 1)) And &HFFFF&
        If code >= &HFF01& And code <= &HFF5E& Then
            Mid$(s, i, 1) = ChrW(code - &HFEE0&)
        ElseIf code = &H3000& Then
            Mid$(s, i, 1) = " "
        End If
    Next
    Z2H = s
End Function

Private Function IsWs(ByVal ch As String) As Boolean
    Select Case AscW(ch) And &HFFFF&
    Case 9, 10, 11, 12, 13, 32, &HA0&, &H3000&
        IsWs = True
    End Select
End Function

Private Function StripW(ByVal s As String) As String
    Dim a As Long, b As Long
    a = 1
    b = Len(s)
    Do While a <= b
        If Not IsWs(Mid$(s, a, 1)) Then Exit Do
        a = a + 1
    Loop
    Do While b >= a
        If Not IsWs(Mid$(s, b, 1)) Then Exit Do
        b = b - 1
    Loop
    If b >= a Then StripW = Mid$(s, a, b - a + 1)
End Function

Private Function LStripW(ByVal s As String) As String
    Dim a As Long
    a = 1
    Do While a <= Len(s)
        If Not IsWs(Mid$(s, a, 1)) Then Exit Do
        a = a + 1
    Loop
    LStripW = Mid$(s, a)
End Function

Private Function RStripChar(ByVal s As String, ByVal ch As String) As String
    Dim b As Long
    b = Len(s)
    Do While b > 0
        If Mid$(s, b, 1) <> ch Then Exit Do
        b = b - 1
    Loop
    RStripChar = Left$(s, b)
End Function

Private Function IsDigitCh(ByVal ch As String) As Boolean
    If Len(ch) <> 1 Then Exit Function
    IsDigitCh = (ch >= "0" And ch <= "9")
End Function

Private Function Digits(ByVal s As String, ByRef p As Long) As String
    Dim st As Long
    st = p
    Do While p <= Len(s)
        If Not IsDigitCh(Mid$(s, p, 1)) Then Exit Do
        p = p + 1
    Loop
    Digits = Mid$(s, st, p - st)
End Function

Private Function AllDigits(ByVal s As String) As Boolean
    Dim i As Long
    If Len(s) = 0 Then Exit Function
    For i = 1 To Len(s)
        If Not IsDigitCh(Mid$(s, i, 1)) Then Exit Function
    Next
    AllDigits = True
End Function

Private Function InListZ(ByVal z As String, ByVal list As Variant, ByVal stripMaru As Boolean) As Boolean
    Dim e As Variant, x As String
    For Each e In list
        x = CStr(e)
        If stripMaru Then x = RStripChar(x, "。")
        If z = Z2H(x) Then InListZ = True: Exit Function
    Next
End Function

Private Function StartsAny(ByVal s As String, ByVal heads As Variant) As Boolean
    Dim e As Variant
    For Each e In heads
        If Left$(s, Len(e)) = e Then StartsAny = True: Exit Function
    Next
End Function

' --- セルを読む ---------------------------------------------------------------

' 結合セルは左上の値を返す。エラー値は空として扱う
Private Function CellVal(ByVal ws As Worksheet, ByVal r As Long, ByVal c As Long) As Variant
    Dim cell As Range
    Set cell = ws.Cells(r, c)
    If cell.MergeCells Then Set cell = cell.MergeArea.Cells(1, 1)
    If IsError(cell.Value2) Then
        CellVal = ""
    Else
        CellVal = cell.Value2
    End If
End Function

Private Function RowHasValue(ByVal ws As Worksheet, ByVal r As Long, ByVal lastCol As Long) As Boolean
    Dim c As Long
    For c = 1 To lastCol
        If Len(StripW(CStr(CellVal(ws, r, c)))) > 0 Then RowHasValue = True: Exit Function
    Next
End Function
