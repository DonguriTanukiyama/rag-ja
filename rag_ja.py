# -*- coding: utf-8 -*-
u"""rag_ja.py — 日本語文書の検索つき回答（RAG）

  python rag_ja.py index <ディレクトリ>     文書を読んで索引を作る
  python rag_ja.py ask "質問"               検索して答える
  python rag_ja.py eval <質問セット.json>   検索の当たり具合を測る
  python rag_ja.py --demo                   自己確認

設計の方針:

  1. 出典の無い文を答えに出さない。
     答えの各文に、根拠にした文書と位置を必ず付ける。
     出典を付けられない部分は「資料に記載がありません」と言う。

  2. 埋め込みは差し替えられる。
     Ollama に埋め込みモデルがあればそれを使い、無ければ Chroma の既定に落ちる。
     どちらを使ったかは必ず表示する。黙って質の低いほうに落ちない。

  3. 測れる。
     eval に質問と正解文書の対を渡すと、top-k に正解が入った割合を出す。
     埋め込みを替えたときの前後比較に使う。

依存: chromadb, requests
"""
import argparse
import io
import json
import os
import re
import sys

OLLAMA = os.environ.get("OLLAMA_HOST", "http://127.0.0.1:11434")
DB_DIR = os.path.join(os.path.dirname(os.path.abspath(__file__)), "_index")
COLLECTION = "docs"

# Ollama に入っていれば使う埋め込みモデル。上から順に探す
EMBED_CANDIDATES = ["bge-m3", "mxbai-embed-large", "nomic-embed-text"]

SUFFIXES = (".md", ".txt", ".rst")


# --- 埋め込み ---------------------------------------------------------------

def _ollama_models():
    try:
        import requests
        r = requests.get(OLLAMA + "/api/tags", timeout=5)
        return [m["name"].split(":")[0] for m in r.json().get("models", [])]
    except Exception:
        return []


def pick_embedder():
    u"""使える埋め込みを1つ選ぶ。何を選んだかを呼び出し側に返す。"""
    have = _ollama_models()
    for name in EMBED_CANDIDATES:
        if name in have:
            return ("ollama", name)
    return ("chroma-default", "all-MiniLM-L6-v2")


def embed(texts, backend, model):
    if backend == "ollama":
        import requests
        out = []
        for t in texts:                       # 1件ずつ。まとめると落ちる実装がある
            r = requests.post(OLLAMA + "/api/embed",
                              json={"model": model, "input": t}, timeout=120)
            r.raise_for_status()
            d = r.json()
            v = d.get("embeddings") or d.get("embedding")
            out.append(v[0] if isinstance(v[0], list) else v)
        return out
    return None                               # Chroma の既定に任せる


# --- 文書を読む -------------------------------------------------------------

def split(text, size=600, overlap=120):
    u"""段落の切れ目を優先して分ける。文の途中で切らない。

    日本語は空白で単語が切れないので、文字数で数える。
    段落（空行）→ 句点 の順で切れ目を探し、どちらも無ければ位置で切る。
    """
    text = re.sub(r"\r\n?", "\n", text).strip()
    if not text:
        return []
    chunks, i = [], 0
    while i < len(text):
        end = min(i + size, len(text))
        if end < len(text):
            window = text[i:end]
            cut = max(window.rfind("\n\n"), window.rfind("。\n"))
            if cut < size // 3:
                cut = window.rfind("。")
            if cut >= size // 3:
                end = i + cut + 1
        chunk = text[i:end].strip()
        if chunk:
            chunks.append((chunk, i))          # 本文と、元文書での開始位置
        if end >= len(text) or end <= i:
            break                              # 末尾に達したら重ねない（重複の元）
        i = max(end - overlap, i + 1)
    return chunks


def iter_docs(root):
    for dirpath, dirnames, filenames in os.walk(root):
        dirnames[:] = [d for d in dirnames
                       if d not in (".git", "__pycache__", "_index", "node_modules")]
        for fn in sorted(filenames):
            if fn.lower().endswith(SUFFIXES):
                yield os.path.join(dirpath, fn)


# --- 索引 -------------------------------------------------------------------

def get_collection(reset=False):
    import chromadb
    client = chromadb.PersistentClient(path=DB_DIR)
    if reset:
        try:
            client.delete_collection(COLLECTION)
        except Exception:
            pass
    return client.get_or_create_collection(COLLECTION)


def cmd_index(root, reset=True):
    backend, model = pick_embedder()
    say(u"埋め込み: %s / %s" % (backend, model))
    if backend == "chroma-default":
        say(u"  注意: 日本語の検索精度は期待できない。"
            u"`ollama pull bge-m3` で差し替えられる（eval で差が測れる）")

    col = get_collection(reset=reset)
    ids, docs, metas, n_files = [], [], [], 0

    for path in iter_docs(root):
        try:
            text = io.open(path, encoding="utf-8").read()
        except Exception as e:
            say(u"  読めない: %s (%s)" % (path, type(e).__name__))
            continue
        rel = os.path.relpath(path, root).replace("\\", "/")
        pieces = split(text)
        if not pieces:
            continue
        n_files += 1
        for k, (chunk, pos) in enumerate(pieces):
            ids.append("%s#%d" % (rel, k))
            docs.append(chunk)
            metas.append({"file": rel, "chunk": k, "pos": pos,
                          "line": text[:pos].count("\n") + 1})

    if not docs:
        say(u"索引に入れるものが無い（.md .txt .rst を探した）")
        return 1

    B = 64
    for s in range(0, len(docs), B):
        part = docs[s:s + B]
        vecs = embed(part, backend, model)
        kw = {"ids": ids[s:s + B], "documents": part, "metadatas": metas[s:s + B]}
        if vecs is not None:
            kw["embeddings"] = vecs
        col.upsert(**kw)
        say(u"  %d/%d" % (min(s + B, len(docs)), len(docs)))

    io.open(os.path.join(DB_DIR, "backend.json"), "w", encoding="utf-8").write(
        json.dumps({"backend": backend, "model": model, "root": os.path.abspath(root)},
                   ensure_ascii=False))
    say(u"索引: %d文書 → %d片" % (n_files, len(docs)))
    return 0


def _saved_backend():
    try:
        d = json.load(io.open(os.path.join(DB_DIR, "backend.json"), encoding="utf-8"))
        return d["backend"], d["model"]
    except Exception:
        return pick_embedder()


def index_is_empty():
    return get_collection().count() == 0


def search(question, k=5):
    backend, model = _saved_backend()
    col = get_collection()
    kw = {"n_results": k}
    vecs = embed([question], backend, model)
    if vecs is not None:
        kw["query_embeddings"] = vecs
    else:
        kw["query_texts"] = [question]
    r = col.query(**kw)
    hits = []
    for i in range(len(r["ids"][0])):
        hits.append({
            "id": r["ids"][0][i],
            "text": r["documents"][0][i],
            "file": r["metadatas"][0][i]["file"],
            "line": r["metadatas"][0][i]["line"],
            "dist": r["distances"][0][i] if r.get("distances") else None,
        })
    return hits, backend, model


# --- 回答 -------------------------------------------------------------------

PROMPT = u"""あなたは社内資料の検索窓口です。以下の【資料】だけを根拠に、質問に答えてください。

守ること:
- 【資料】に書かれていないことは書かない。推測で補わない。
- 分からない場合は「資料に記載がありません」とだけ答える。
- 答えの各文の末尾に、根拠にした資料の番号を [1] のように付ける。
- 数値を出すときは、必ず根拠の番号を付ける。

【資料】
%s

【質問】
%s

【答え】"""


GEN_MODEL = "qwen2.5:7b"   # 埋め込み(1.2GB)と同居させる。12B は VRAM 10GB で溢れて遅い


def cmd_ask(question, k=5, model=GEN_MODEL):
    import requests
    if index_is_empty():
        say(u"索引が空。先に `python rag_ja.py index <ディレクトリ>` を実行する")
        return 1
    hits, backend, emodel = search(question, k)
    say(u"埋め込み: %s / %s   検索: 上位%d件" % (backend, emodel, len(hits)))
    if not hits:
        say(u"該当なし")
        return 1

    src = u"\n\n".join(u"[%d] %s:%d\n%s" % (i + 1, h["file"], h["line"], h["text"])
                       for i, h in enumerate(hits))
    r = requests.post(OLLAMA + "/api/generate",
                      json={"model": model, "prompt": PROMPT % (src, question),
                            "stream": False, "options": {"temperature": 0}},
                      timeout=600)
    r.raise_for_status()
    say(u"")
    say(r.json().get("response", "").strip())
    say(u"")
    say(u"--- 根拠 ---")
    for i, h in enumerate(hits):
        say(u"[%d] %s:%d  %s" % (i + 1, h["file"], h["line"],
                                 h["text"][:70].replace("\n", " ")))
    return 0


# --- 測る -------------------------------------------------------------------

def cmd_eval(path, k=5):
    u"""質問セットで検索の当たり具合を測る。

    形式: [{"q": "質問", "file": "正解の文書（相対パス。部分一致）"}, ...]
    出すもの: Recall@k と、1位に正解が来た割合
    """
    if index_is_empty():
        say(u"索引が空。先に `python rag_ja.py index <ディレクトリ>` を実行する")
        return 1
    cases = json.load(io.open(path, encoding="utf-8"))
    hit_k = hit_1 = 0
    rows = []
    backend = emodel = "?"
    for c in cases:
        hits, backend, emodel = search(c["q"], k)
        files = [h["file"].replace("\\", "/") for h in hits]
        want = c["file"].replace("\\", "/")
        in_k = any(want in f for f in files)
        in_1 = bool(files) and want in files[0]
        hit_k += in_k
        hit_1 += in_1
        rows.append((c["q"], want, files[0] if files else "-", in_k, in_1))

    n = len(cases)
    say(u"埋め込み: %s / %s" % (backend, emodel))
    say(u"件数 %d   Recall@%d %d/%d (%.0f%%)   1位正解 %d/%d (%.0f%%)"
        % (n, k, hit_k, n, 100.0 * hit_k / n, hit_1, n, 100.0 * hit_1 / n))
    say(u"")
    for q, want, got, in_k, in_1 in rows:
        mark = u"OK " if in_1 else (u"k  " if in_k else u"NG ")
        say(u"  %s %-30s 期待=%s / 1位=%s" % (mark, q[:30], want, got))
    return 0


# --- 自己確認 ---------------------------------------------------------------

def demo():
    NL = u"\n"

    # 1) ふつうの文書は、段落か句点で切れる
    def para(n):
        return u"第%d段落。" % n + u"これは説明の文です。" * 12 + NL + NL

    t = u"".join(para(n) for n in range(1, 8))
    ps = split(t, size=600, overlap=120)
    assert len(ps) >= 2, len(ps)
    assert all(isinstance(p, tuple) and len(p) == 2 for p in ps)
    assert all(p[0].endswith(u"。") for p in ps), repr([p[0][-8:] for p in ps])
    assert all(ps[i][1] < ps[i + 1][1] for i in range(len(ps) - 1))

    # 2) 句点の無い長文でも止まらず進み、文字を落とさない
    runon = u"あ" * 2000
    rs = split(runon, size=600, overlap=120)
    assert len(rs) >= 3, len(rs)
    assert all(rs[i][1] < rs[i + 1][1] for i in range(len(rs) - 1))
    covered = set()
    for chunk, pos in rs:
        covered.update(range(pos, pos + len(chunk)))
    assert covered == set(range(2000)), len(covered)   # 穴が無い

    # 3) 端
    assert split(u"") == []
    assert split(u"   " + NL + NL + u"  ") == []
    assert len(split(u"短い文書。")) == 1
    assert split(u"改行だけ\r\n混在。")[0][0].find(u"\r") < 0    # CRLF を潰している

    # 4) 重なりがある（前の片の末尾が次の片に入る）
    assert ps[1][1] < ps[0][1] + len(ps[0][0])

    b, m = pick_embedder()
    say(u"demo OK  通常 %d片 / 句点なし %d片 / 埋め込み %s:%s" % (len(ps), len(rs), b, m))


# --- 入口 -------------------------------------------------------------------

def say(s):
    enc = getattr(sys.stdout, "encoding", None) or "utf-8"
    try:
        print(s)
    except UnicodeEncodeError:
        print(s.encode(enc, "replace").decode(enc, "replace"))


def main():
    ap = argparse.ArgumentParser(description=u"日本語文書のRAG")
    ap.add_argument("cmd", nargs="?", choices=["index", "ask", "eval"])
    ap.add_argument("arg", nargs="?")
    ap.add_argument("-k", type=int, default=5)
    ap.add_argument("--model", default=GEN_MODEL)
    ap.add_argument("--demo", action="store_true")
    ap.add_argument("--add", action="store_true", help=u"索引を作り直さず追加する")
    a = ap.parse_args()

    if a.demo:
        return demo()
    if not a.cmd or not a.arg:
        ap.print_help()
        return 1

    import requests
    try:
        if a.cmd == "index":
            return cmd_index(a.arg, reset=not a.add)
        if a.cmd == "ask":
            return cmd_ask(a.arg, a.k, a.model)
        if a.cmd == "eval":
            return cmd_eval(a.arg, a.k)
    except requests.exceptions.RequestException as e:
        # スタックトレースを出さない。何を直せばいいかだけ出す
        say(u"Ollama (%s) に届かない。`ollama serve` が動いているか確認する" % OLLAMA)
        say(u"  %s: %s" % (type(e).__name__, str(e)[:120]))
        return 2


if __name__ == "__main__":
    sys.exit(main() or 0)
