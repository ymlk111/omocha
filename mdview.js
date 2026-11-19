// =============================================================================
//  MarkdownPreview.js - サクラエディタ用 Markdown プレビューマクロ (WSH / JScript)
// -----------------------------------------------------------------------------
//  編集中の内容（未保存の変更も含む）を HTML に変換し、ブラウザで表示します。
//
//   - 変換はブラウザ側の marked.js（GFM：表・タスクリスト・取り消し線・自動リンク）
//   - ```mermaid のコードブロックは図として描画（含まれるときだけ mermaid.js を読込）
//   - 言語指定付きコードブロックは highlight.js で色付け（含まれるときだけ読込）
//   - 相対パスの画像・リンクは .md ファイルのあるフォルダを基準に解決
//   - 見出しに GitHub 風の id を付与（目次の [x](#見出し) リンクが機能）
//   - 先頭の YAML Front Matter は折りたたみ表示
//   - OS のダークモード設定に追従
//
//  出力先: %TEMP%\sakura_md_preview\<ファイル名>_<パスのハッシュ>.html
//
//  【登録方法】
//   1. このファイルを「マクロフォルダー」（共通設定 → マクロ タブで確認）に置く
//   2. 共通設定 → マクロ タブ: 空き番号を選び、名前と File を指定して「設定」
//   3. 共通設定 → キー割り当て タブ: 種別「外部マクロ」から選び、好きなキーを割り当て
//   ※ 試しに動かすだけなら「名前を指定してマクロ実行」でこのファイルを選べば OK
//
//  【重要】このファイルは「BOM 付き UTF-8」または「Shift_JIS」で保存してください。
//          サクラエディタは BOM なし UTF-8 のマクロを Shift_JIS として読むため、
//          日本語が文字化けしてエラーになります。
// =============================================================================

var CONFIG = {
    // 開くブラウザの exe。空なら .html に関連付けられた既定のブラウザで開く
    //   例: "C:\\Program Files (x86)\\Microsoft\\Edge\\Application\\msedge.exe"
    browser: "",

    // プレビュー HTML の出力先フォルダ。空なら %TEMP%\sakura_md_preview
    outDir: "",

    // true: 改行をそのまま <br> にする / false: GitHub の .md 表示と同じ（空行で段落）
    breaks: false,

    // 使用ライブラリ（URL またはローカルの絶対パス）。
    // ネットに出られない環境では、各ファイルをダウンロードしてローカルパスを指定する。
    //   例: markedJs: "C:\\tools\\mdpreview\\marked.umd.js"
    markedJs:    "https://cdn.jsdelivr.net/npm/marked@18/lib/marked.umd.js",
    mermaidJs:   "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js",              // "" で無効
    hljsJs:      "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/highlight.min.js", // "" で無効
    hljsCss:     "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/styles/github.min.css",
    hljsCssDark: "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/styles/github-dark.min.css"
};

// =============================================================================
//  ここから下は通常変更不要
// =============================================================================

(function () {
    try {
        main();
    } catch (e) {
        Editor.ErrorMsg("Markdown プレビューの作成に失敗しました。\n\n" + (e.message || e.description || String(e)));
    }
})();

function main() {
    var fso   = new ActiveXObject("Scripting.FileSystemObject");
    var shell = new ActiveXObject("WScript.Shell");

    // 1) 編集中のテキスト（未保存の変更も含む）
    var md = getEditorText();

    // 2) ファイル情報（無題なら path は ""）
    var path  = String(Editor.GetFilename() || "");
    var title = path ? fso.GetFileName(path) : "(無題)";
    var baseHref = "";
    if (path) {
        baseHref = toFileUrl(fso.GetParentFolderName(path).replace(/\\+$/, "") + "\\");
    }

    // 3) 出力先
    var outDir = CONFIG.outDir || (fso.GetSpecialFolder(2).Path + "\\sakura_md_preview");
    ensureFolder(fso, outDir);
    var outFile = outDir.replace(/\\+$/, "") + "\\" +
                  (path ? fso.GetBaseName(path).replace(/[#%&;'`^~{}\[\] ]/g, "_") : "untitled") +
                  "_" + hash(path.toLowerCase()) + ".html";

    // 4) HTML を生成して UTF-8 で保存
    writeUtf8(outFile, buildHtml(md, title, path, baseHref));

    // 5) ブラウザで開く
    if (CONFIG.browser) {
        shell.Run(quote(CONFIG.browser) + " " + quote(outFile), 1, false);
    } else {
        shell.Run(quote(outFile), 1, false);
    }
}

// 全行を取得（GetLineStr は改行コード付きで返るので除去して LF で連結）
function getEditorText() {
    var n = Editor.GetLineCount(0);
    var buf = [];
    for (var i = 1; i <= n; i++) {
        buf.push(String(Editor.GetLineStr(i)).replace(/(\r\n|[\r\n\u0085\u2028\u2029])$/, ""));
    }
    return buf.join("\n");
}

function buildHtml(md, title, path, baseHref) {
    var opt = {
        md:          md,
        breaks:      !!CONFIG.breaks,
        mermaidJs:   toUrl(CONFIG.mermaidJs),
        hljsJs:      toUrl(CONFIG.hljsJs),
        hljsCss:     toUrl(CONFIG.hljsCss),
        hljsCssDark: toUrl(CONFIG.hljsCssDark)
    };
    var h = [];
    h.push('<!DOCTYPE html>');
    h.push('<html lang="ja">');
    h.push('<head>');
    h.push('<meta charset="utf-8">');
    h.push('<meta name="viewport" content="width=device-width, initial-scale=1">');
    if (baseHref) {
        h.push('<base href="' + escapeHtml(baseHref) + '">');
    }
    h.push('<title>' + escapeHtml(title) + ' - Markdown Preview</title>');
    h.push('<style>' + heredoc(CSS) + '</style>');
    h.push('</head>');
    h.push('<body>');
    h.push('<div id="bar"><span class="path" title="' + escapeHtml(path || title) + '">' +
           escapeHtml(path || title) + '</span><span id="msg"></span><span>' +
           escapeHtml(formatDate(new Date())) + ' 生成</span></div>');
    h.push('<article id="content" class="markdown-body"><p class="note">変換中...</p></article>');
    h.push('<script src="' + escapeHtml(toUrl(CONFIG.markedJs)) + '"></script>');
    h.push('<script>');
    h.push('(' + String(browserMain) + ')(' + toJs(opt) + ');');
    h.push('</script>');
    h.push('</body>');
    h.push('</html>');
    return h.join("\n");
}

// -----------------------------------------------------------------------------
//  ブラウザ側で実行される処理（関数のソースをそのまま HTML に埋め込む）
//  ※ マクロ実行時に JScript (ES3) でも構文解析されるため、ES3 の構文で書くこと
// -----------------------------------------------------------------------------
function browserMain(opt) {
    var out = document.getElementById("content");
    var src = opt.md;

    function esc(s) {
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
                        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }
    function warn(text) {
        var m = document.getElementById("msg");
        m.className = "warn";
        m.textContent = (m.textContent ? m.textContent + " / " : "") + text;
    }
    function loadScript(url, onload) {
        var s = document.createElement("script");
        s.src = url;
        s.onload = onload;
        s.onerror = function () { warn(url.replace(/^.*\//, "") + " を読み込めませんでした"); };
        document.head.appendChild(s);
    }
    function addCss(url, media) {
        if (!url) { return; }
        var l = document.createElement("link");
        l.rel = "stylesheet";
        l.href = url;
        if (media) { l.media = media; }
        document.head.appendChild(l);
    }

    // marked.js が読めなかった場合は原文表示
    if (typeof window.marked === "undefined") {
        out.innerHTML = '<p class="warn">marked.js を読み込めませんでした。' +
            'ネットに出られない環境では、マクロの CONFIG.markedJs にローカルファイルを指定してください。</p>';
        var raw = document.createElement("pre");
        raw.textContent = src;
        out.appendChild(raw);
        return;
    }

    // YAML Front Matter
    var frontMatter = null;
    var fm = /^---[ \t]*\n([\s\S]*?)\n(?:---|\.\.\.)[ \t]*(?:\n|$)/.exec(src);
    if (fm) {
        frontMatter = fm[1];
        src = src.slice(fm[0].length);
    }

    // GitHub 風の見出し id
    var used = {};
    var tmp = document.createElement("div");
    function slugify(html) {
        tmp.innerHTML = html;
        var s = (tmp.textContent || "").toLowerCase();
        try {
            s = s.replace(new RegExp("[^\\p{L}\\p{M}\\p{N}\\s_-]", "gu"), "");
        } catch (e) {
            s = s.replace(/[!-,.\/:-@\[-\^`{-~]/g, "");
        }
        s = s.replace(/\s/g, "-");
        var n = used[s] || 0;
        used[s] = n + 1;
        return n ? s + "-" + n : s;
    }

    marked.use({
        gfm: true,
        breaks: !!opt.breaks,
        renderer: {
            heading: function (token) {
                var inner = this.parser.parseInline(token.tokens);
                return "<h" + token.depth + ' id="' + esc(slugify(inner)) + '">' +
                       inner + "</h" + token.depth + ">\n";
            },
            code: function (token) {
                var lang = String(token.lang || "").replace(/^\s+/, "").split(/\s/)[0].toLowerCase();
                if (lang === "mermaid") {
                    return '<pre class="mermaid">' + esc(token.text) + "</pre>\n";
                }
                return false; // それ以外は marked 標準の描画
            }
        }
    });

    out.innerHTML = marked.parse(src);

    if (frontMatter !== null) {
        var d = document.createElement("details");
        d.className = "front-matter";
        d.innerHTML = "<summary>Front Matter</summary>";
        var p = document.createElement("pre");
        p.textContent = frontMatter;
        d.appendChild(p);
        out.insertBefore(d, out.firstChild);
    }

    // <base> を付けているため、ページ内リンク (#...) はスクリプトでスクロールさせる
    document.addEventListener("click", function (e) {
        var a = (e.target && e.target.closest) ? e.target.closest("a[href^='#']") : null;
        if (!a) { return; }
        var id = a.getAttribute("href").slice(1);
        try { id = decodeURIComponent(id); } catch (err) { }
        var el = id ? document.getElementById(id) : document.body;
        e.preventDefault();
        if (el) { el.scrollIntoView(); }
    });

    // コードのシンタックスハイライト（言語指定ありのブロックのみ）
    var codes = out.querySelectorAll("pre code[class*='language-']");
    if (opt.hljsJs && codes.length) {
        addCss(opt.hljsCss, opt.hljsCssDark ? "(prefers-color-scheme: light)" : "");
        addCss(opt.hljsCssDark, "(prefers-color-scheme: dark)");
        loadScript(opt.hljsJs, function () {
            for (var i = 0; i < codes.length; i++) {
                try { window.hljs.highlightElement(codes[i]); } catch (err) { }
            }
        });
    }

    // Mermaid
    var diagrams = out.querySelectorAll("pre.mermaid");
    if (opt.mermaidJs && diagrams.length) {
        loadScript(opt.mermaidJs, function () {
            var dark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
            window.mermaid.initialize({ startOnLoad: false, theme: dark ? "dark" : "default" });
            window.mermaid.run({ nodes: diagrams, suppressErrors: true });
        });
    }
}

// -----------------------------------------------------------------------------
//  スタイル（GitHub 風）。heredoc で取り出すため、この中に CSS コメントは書かないこと
// -----------------------------------------------------------------------------
function CSS() {/*
:root {
  color-scheme: light dark;
  --fg: #1f2328; --bg: #ffffff; --muted: #59636e; --border: #d1d9e0;
  --code-bg: #f6f8fa; --link: #0969da; --bar-bg: #f6f8fa; --warn: #d1242f;
}
@media (prefers-color-scheme: dark) {
  :root {
    --fg: #f0f6fc; --bg: #0d1117; --muted: #9198a1; --border: #3d444d;
    --code-bg: #151b23; --link: #4493f8; --bar-bg: #151b23; --warn: #f85149;
  }
}
html, body { margin: 0; background: var(--bg); }
body {
  color: var(--fg);
  font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", "Noto Sans JP", "Hiragino Sans", Meiryo, sans-serif;
  font-size: 16px; line-height: 1.7;
}
#bar {
  position: sticky; top: 0; z-index: 10;
  display: flex; gap: 16px; align-items: center;
  padding: 6px 16px; font-size: 12px; color: var(--muted);
  background: var(--bar-bg); border-bottom: 1px solid var(--border);
}
#bar .path { flex: 1; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; }
.warn { color: var(--warn) !important; }
.note { color: var(--muted); }
.markdown-body {
  max-width: 980px; margin: 0 auto; padding: 32px 40px 96px;
  box-sizing: border-box; overflow-wrap: break-word;
}
.markdown-body > :first-child { margin-top: 0 !important; }
.markdown-body [id] { scroll-margin-top: 48px; }
.markdown-body h1, .markdown-body h2, .markdown-body h3,
.markdown-body h4, .markdown-body h5, .markdown-body h6 {
  margin: 24px 0 16px; font-weight: 600; line-height: 1.25;
}
.markdown-body h1 { font-size: 2em; padding-bottom: .3em; border-bottom: 1px solid var(--border); }
.markdown-body h2 { font-size: 1.5em; padding-bottom: .3em; border-bottom: 1px solid var(--border); }
.markdown-body h3 { font-size: 1.25em; }
.markdown-body h4 { font-size: 1em; }
.markdown-body h5 { font-size: .875em; }
.markdown-body h6 { font-size: .85em; color: var(--muted); }
.markdown-body p, .markdown-body blockquote, .markdown-body ul, .markdown-body ol,
.markdown-body dl, .markdown-body table, .markdown-body pre, .markdown-body details {
  margin-top: 0; margin-bottom: 16px;
}
.markdown-body a { color: var(--link); text-decoration: none; }
.markdown-body a:hover { text-decoration: underline; }
.markdown-body ul, .markdown-body ol { padding-left: 2em; }
.markdown-body li + li { margin-top: .25em; }
.markdown-body li > ul, .markdown-body li > ol { margin-bottom: 0; }
.markdown-body ul:has(> li > input[type="checkbox"]) { list-style: none; padding-left: 1.2em; }
.markdown-body li > input[type="checkbox"] { margin: 0 .4em 0 -1.2em; vertical-align: middle; }
.markdown-body blockquote { margin-left: 0; margin-right: 0; padding: 0 1em; color: var(--muted); border-left: .25em solid var(--border); }
.markdown-body code {
  font-family: ui-monospace, SFMono-Regular, Consolas, "BIZ UDGothic", "MS Gothic", monospace;
  font-size: 85%; padding: .2em .4em; background: var(--code-bg); border-radius: 6px;
}
.markdown-body pre {
  padding: 16px; overflow: auto; font-size: 85%; line-height: 1.45;
  background: var(--code-bg); border-radius: 6px;
}
.markdown-body pre code, .markdown-body pre code.hljs {
  padding: 0; font-size: 100%; background: transparent; white-space: pre; display: inline;
}
.markdown-body pre.mermaid { background: transparent; text-align: center; }
.markdown-body table { border-collapse: collapse; border-spacing: 0; display: block; width: max-content; max-width: 100%; overflow: auto; }
.markdown-body th, .markdown-body td { padding: 6px 13px; border: 1px solid var(--border); }
.markdown-body th { font-weight: 600; }
.markdown-body tr:nth-child(2n) { background: var(--code-bg); }
.markdown-body img { max-width: 100%; box-sizing: content-box; }
.markdown-body hr { height: .25em; padding: 0; margin: 24px 0; background: var(--border); border: 0; }
.markdown-body kbd {
  display: inline-block; padding: 3px 5px; font-size: 11px; line-height: 10px;
  border: 1px solid var(--border); border-bottom-width: 2px; border-radius: 6px; background: var(--code-bg);
}
.front-matter summary { cursor: pointer; color: var(--muted); font-size: .85em; }
@media print {
  #bar { display: none; }
  .markdown-body { max-width: none; padding: 0; }
}
*/}

// -----------------------------------------------------------------------------
//  ユーティリティ（JScript / ES3）
// -----------------------------------------------------------------------------
function heredoc(fn) {
    var s = String(fn);
    return s.substring(s.indexOf("/*") + 2, s.lastIndexOf("*/"));
}

function escapeHtml(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
                    .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
}

// JS の文字列リテラル化（</script> 対策で < > & も \uXXXX にする）
function jsStr(s) {
    return '"' + String(s).replace(/[\\"\x00-\x1f<>&\u2028\u2029]/g, function (c) {
        var hex = c.charCodeAt(0).toString(16);
        return "\\u" + "0000".substr(hex.length) + hex;
    }) + '"';
}

function toJs(v) {
    if (v === null || v === undefined) { return "null"; }
    if (typeof v === "boolean" || typeof v === "number") { return String(v); }
    if (typeof v === "string") { return jsStr(v); }
    var parts = [];
    for (var k in v) {
        if (v.hasOwnProperty(k)) { parts.push(jsStr(k) + ":" + toJs(v[k])); }
    }
    return "{" + parts.join(",") + "}";
}

// Windows パス → file:// URL（UNC 対応）
function toFileUrl(p) {
    var s = String(p).replace(/\\/g, "/");
    s = (s.substr(0, 2) === "//") ? "file:" + s : "file:///" + s;
    return encodeURI(s).replace(/#/g, "%23").replace(/\?/g, "%3F");
}

// CONFIG の値（URL またはローカルパス）→ URL
function toUrl(v) {
    if (!v) { return ""; }
    if (/^[A-Za-z]:[\\\/]/.test(v) || /^\\\\/.test(v)) { return toFileUrl(v); }
    return v;
}

function ensureFolder(fso, dir) {
    if (!dir || fso.FolderExists(dir)) { return; }
    var parent = fso.GetParentFolderName(dir);
    if (parent && parent !== dir) { ensureFolder(fso, parent); }
    fso.CreateFolder(dir);
}

function writeUtf8(file, text) {
    var st = new ActiveXObject("ADODB.Stream");
    st.Type = 2;            // adTypeText
    st.Charset = "UTF-8";
    st.Open();
    st.WriteText(text);
    st.SaveToFile(file, 2); // adSaveCreateOverWrite
    st.Close();
}

function hash(s) {
    var h = 5381;
    for (var i = 0; i < s.length; i++) {
        h = (h * 33 + s.charCodeAt(i)) % 4294967296;
    }
    var hex = h.toString(16);
    return "00000000".substr(hex.length) + hex;
}

function quote(s) {
    return '"' + s + '"';
}

function formatDate(d) {
    function p(n) { return (n < 10 ? "0" : "") + n; }
    return d.getFullYear() + "/" + p(d.getMonth() + 1) + "/" + p(d.getDate()) + " " +
           p(d.getHours()) + ":" + p(d.getMinutes()) + ":" + p(d.getSeconds());
}
