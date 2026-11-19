// =============================================================================
//  MarkdownPreview.js - サクラエディタ用 Markdown プレビュー＆編集マクロ (WSH / JScript)
// -----------------------------------------------------------------------------
//  編集中の内容（未保存の変更も含む）を HTML に変換し、ブラウザで表示します。
//
//  ■ 表示
//   - 変換はブラウザ側の marked.js（GFM：表・タスクリスト・取り消し線・自動リンク）
//   - ```mermaid のコードブロックは図として描画（含まれるときだけ mermaid.js を読込）
//   - 言語指定付きコードブロックは highlight.js で色付け（含まれるときだけ読込）
//   - 相対パスの画像・リンクは .md ファイルのあるフォルダを基準に解決
//   - 見出しに GitHub 風の id を付与／先頭の YAML Front Matter は折りたたみ表示
//   - OS のダークモード設定に追従
//
//  ■ 編集（画面上部のボタン）
//   - [プレビュー] [コード] で表示を切り替え。どちらでも直接編集できます
//   - プレビューで編集した場合、触っていないブロックは元の Markdown をそのまま残し、
//     編集したブロックだけを Markdown に変換し直します
//   - [コピー] で編集後の Markdown をコピー → サクラで Ctrl+A → Ctrl+V で反映
//     （改行コードはサクラで開いている文書に合わせます）
//   - プレビュー編集中のリンクは Ctrl+クリックで開きます
//   - Front Matter と mermaid の図はコード表示で編集してください
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

    // 使用ライブラリ（URL／ローカルの絶対パス／相対パス）。
    // 相対パスは「このマクロファイルがあるフォルダー」が基準。
    // ネットに出られない環境では、各ファイルをダウンロードしてローカルに置く。
    //   例: markedJs: "lib/marked.umd.js"              （マクロと同じ階層の lib フォルダー）
    //       markedJs: "C:\\tools\\mdpreview\\marked.umd.js"
    markedJs:    "https://cdn.jsdelivr.net/npm/marked@18/lib/marked.umd.js",
    turndownJs:  "https://cdn.jsdelivr.net/npm/turndown@7/lib/turndown.browser.umd.js",  // プレビュー編集用
    mermaidJs:   "https://cdn.jsdelivr.net/npm/mermaid@11/dist/mermaid.min.js",              // "" で無効
    hljsJs:      "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/highlight.min.js", // "" で無効
    hljsCss:     "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/styles/github.min.css",
    hljsCssDark: "https://cdn.jsdelivr.net/npm/@highlightjs/cdn-assets@11/styles/github-dark.min.css"
};

// =============================================================================
//  ここから下は通常変更不要
// =============================================================================

var MACRO_DIR = ""; // 実行中のマクロファイルのフォルダー（相対パスの基準）

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

    // 0) このマクロファイルのフォルダー（$M = 実行中のマクロファイルのフルパス）
    try {
        var macroPath = String(Editor.ExpandParameter("$M") || "");
        if (/^([A-Za-z]:\\|\\\\)/.test(macroPath)) { MACRO_DIR = fso.GetParentFolderName(macroPath); }
    } catch (e) { }

    // 1) 編集中のテキスト（未保存の変更も含む）と改行コード
    var md = getEditorText();
    var eol = "\r\n";
    try {
        eol = ["\r\n", "\r", "\n", "\u0085", "\u2028", "\u2029"][Editor.GetLineCode()] || "\r\n";
    } catch (e) { }

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
    writeUtf8(outFile, buildHtml(md, eol, title, path, baseHref));

    // 5) ブラウザで開く
    if (CONFIG.browser) {
        shell.Run(quote(CONFIG.browser) + " " + quote(outFile), 1, false);
    } else {
        shell.Run(quote(outFile), 1, false);
    }
}

// 全行を取得（GetLineStr は改行コード付き）。末尾の改行の有無も保ったまま LF に統一
function getEditorText() {
    var n = Editor.GetLineCount(0);
    var buf = [];
    for (var i = 1; i <= n; i++) {
        buf.push(String(Editor.GetLineStr(i)));
    }
    return buf.join("").replace(/\r\n|[\r\u0085\u2028\u2029]/g, "\n");
}

function buildHtml(md, eol, title, path, baseHref) {
    var opt = {
        md:          md,
        eol:         eol,
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
    h.push('<div id="bar">' +
           '<span class="seg">' +
           '<button type="button" id="btn-preview" class="on" title="プレビューでも直接編集できます（リンクは Ctrl+クリックで開く）">プレビュー</button>' +
           '<button type="button" id="btn-code" title="Markdown のコードを表示・編集">コード</button>' +
           '</span>' +
           '<button type="button" id="btn-copy" title="編集後の Markdown をクリップボードにコピー">コピー</button>' +
           '<span id="msg"></span><span id="warn" class="warn"></span>' +
           '<span class="path" title="' + escapeHtml(path || title) + '">' + escapeHtml(path || title) + '</span>' +
           '<span>' + escapeHtml(formatDate(new Date())) + ' 生成</span>' +
           '</div>');
    h.push('<article id="content" class="markdown-body"></article>');
    h.push('<div id="code-wrap" hidden><textarea id="code" spellcheck="false"></textarea></div>');
    h.push('<script src="' + escapeHtml(toUrl(CONFIG.markedJs)) + '"></script>');
    h.push('<script src="' + escapeHtml(toUrl(CONFIG.turndownJs)) + '"></script>');
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
//     （アロー関数・let/const・正規表現の u フラグや後読みのリテラルは不可）
// -----------------------------------------------------------------------------
function browserMain(opt) {
    var out = document.getElementById("content");
    var codeWrap = document.getElementById("code-wrap");
    var code = document.getElementById("code");
    var btnPreview = document.getElementById("btn-preview");
    var btnCode = document.getElementById("btn-code");
    var btnCopy = document.getElementById("btn-copy");
    var msg = document.getElementById("msg");

    var FM_RE = /^---[ \t]*\n([\s\S]*?)\n(?:---|\.\.\.)[ \t]*(?:\n|$)/;
    var BLOCK_TAGS = /^(P|DIV|H[1-6]|UL|OL|LI|PRE|BLOCKQUOTE|TABLE|HR|DETAILS|DL|FIGURE|SECTION|ARTICLE)$/;
    var hasMarked = typeof window.marked !== "undefined";
    var td = (typeof window.TurndownService !== "undefined") ? createTurndown() : null;

    var mode = "preview";
    var current = opt.md;      // 最新の Markdown（プレビュー上の未反映編集を除く）
    var previewDirty = false;  // プレビューで編集され、current に未反映
    var changed = false;       // 最後のコピー以降に編集があったか
    var blocks = [];           // プレビューの各ブロック { raw, dirty, count }
    var fmRaw = "", lead = "", srcTail = "";
    var used = {};
    var slugTmp = document.createElement("div");
    var tmp = document.createElement("div");
    var libs = {};
    var cssAdded = false;
    var mermaidReady = false;
    var observer = null;
    var OBS = { subtree: true, childList: true, characterData: true, attributes: true, attributeFilter: ["checked"] };

    function esc(s) {
        return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
                        .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
    }
    function setMsg(text, cls) {
        msg.textContent = text;
        msg.className = cls || "";
    }
    function warn(text) {
        var w = document.getElementById("warn");
        w.textContent = (w.textContent ? w.textContent + " / " : "") + text;
    }
    function markChanged() {
        changed = true;
        setMsg("● 未コピーの変更あり", "dirty");
    }
    function loadLib(url, name, cb) {
        if (window[name]) { cb(); return; }
        if (libs[url]) { libs[url].push(cb); return; }
        libs[url] = [cb];
        var s = document.createElement("script");
        s.src = url;
        s.onload = function () {
            var q = libs[url];
            for (var i = 0; i < q.length; i++) { q[i](); }
        };
        s.onerror = function () { warn(url.replace(/^.*\//, "") + " を読み込めませんでした"); };
        document.head.appendChild(s);
    }
    function addCssOnce() {
        if (cssAdded) { return; }
        cssAdded = true;
        var list = [[opt.hljsCss, opt.hljsCssDark ? "(prefers-color-scheme: light)" : ""],
                    [opt.hljsCssDark, "(prefers-color-scheme: dark)"]];
        for (var i = 0; i < list.length; i++) {
            if (!list[i][0]) { continue; }
            var l = document.createElement("link");
            l.rel = "stylesheet";
            l.href = list[i][0];
            if (list[i][1]) { l.media = list[i][1]; }
            document.head.appendChild(l);
        }
    }

    // ---- GitHub 風の見出し id ------------------------------------------------
    function slugify(html) {
        slugTmp.innerHTML = html;
        var s = (slugTmp.textContent || "").toLowerCase();
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

    // ---- HTML → Markdown（プレビューで編集したブロック用） -------------------
    function createTurndown() {
        var s = new window.TurndownService({
            headingStyle: "atx", hr: "---", bulletListMarker: "-", codeBlockStyle: "fenced",
            fence: "```", emDelimiter: "*", strongDelimiter: "**"
        });
        s.keep(["kbd", "sub", "sup", "u", "mark", "ins", "details", "summary"]);
        s.escape = escapeMd;

        function cleanUrl(u) {
            try {
                var d = decodeURI(u);
                if (!/[\s()<>]/.test(d)) { return d; }
            } catch (e) { }
            return u;
        }
        function titlePart(n) {
            var t = n.getAttribute("title");
            return t ? ' "' + t.replace(/"/g, '\\"') + '"' : "";
        }
        function cellMd(cell) {
            return s.turndown(cell.innerHTML).replace(/\n+/g, "<br>")
                    .replace(/\|/g, "\\|").replace(/^\s+|\s+$/g, "");
        }

        s.addRule("strike", {
            filter: ["del", "s", "strike"],
            replacement: function (c) { return "~~" + c + "~~"; }
        });
        s.addRule("taskbox", {
            filter: function (n) { return n.nodeName === "INPUT" && n.type === "checkbox"; },
            replacement: function (c, n) { return n.checked ? "[x] " : "[ ] "; }
        });
        s.addRule("listItem", {
            filter: "li",
            replacement: function (content, node, o) {
                var prefix = o.bulletListMarker + " ";
                var parent = node.parentNode;
                if (parent.nodeName === "OL") {
                    var start = parent.getAttribute("start");
                    var idx = Array.prototype.indexOf.call(parent.children, node);
                    prefix = (start ? Number(start) + idx : idx + 1) + ". ";
                }
                var isPara = /\n$/.test(content);
                content = content.replace(/^\n+/, "").replace(/\n+$/, "") + (isPara ? "\n" : "");
                content = content.replace(/^\[( |x)\]\s+/, "[$1] ");
                content = content.replace(/\n/g, "\n" + new Array(prefix.length + 1).join(" "));
                return prefix + content + (node.nextSibling ? "\n" : "");
            }
        });
        s.addRule("mermaid", {
            filter: function (n) { return n.nodeName === "PRE" && /(^|\s)mermaid(\s|$)/.test(n.className); },
            replacement: function (c, n) {
                return "\n\n```mermaid\n" + (n.getAttribute("data-src") || "") + "\n```\n\n";
            }
        });
        s.addRule("link", {
            filter: function (n) { return n.nodeName === "A" && n.getAttribute("href"); },
            replacement: function (content, n) {
                var raw = n.getAttribute("href");
                var href = cleanUrl(raw);
                var text = content.replace(/\\([!-\/:-@\[-`{-~])/g, "$1");
                if (!n.getAttribute("title") && (text === href || text === raw) && /^(https?:|mailto:)/i.test(href)) {
                    return href; // 自動リンク
                }
                return "[" + content + "](" + href + titlePart(n) + ")";
            }
        });
        s.addRule("image", {
            filter: "img",
            replacement: function (c, n) {
                var src = cleanUrl(n.getAttribute("src") || "");
                var alt = (n.getAttribute("alt") || "").replace(/([\[\]\\])/g, "\\$1");
                return src ? "![" + alt + "](" + src + titlePart(n) + ")" : "";
            }
        });
        s.addRule("table", {
            filter: "table",
            replacement: function (c, node) {
                var rows = node.rows, lines = [], cols = 0, i, j;
                if (!rows || !rows.length) { return ""; }
                for (i = 0; i < rows.length; i++) { cols = Math.max(cols, rows[i].cells.length); }
                for (i = 0; i < rows.length; i++) {
                    var cells = [];
                    for (j = 0; j < cols; j++) {
                        cells.push(rows[i].cells[j] ? cellMd(rows[i].cells[j]) : "");
                    }
                    lines.push("| " + cells.join(" | ") + " |");
                    if (i === 0) {
                        var seps = [];
                        for (j = 0; j < cols; j++) {
                            var cell = rows[0].cells[j];
                            var a = cell ? (cell.getAttribute("align") || cell.style.textAlign || "") : "";
                            seps.push(a === "center" ? ":---:" : a === "right" ? "---:" : a === "left" ? ":---" : "---");
                        }
                        lines.push("| " + seps.join(" | ") + " |");
                    }
                }
                return "\n\n" + lines.join("\n") + "\n\n";
            }
        });
        return s;
    }

    // Markdown として意味を持つ文字だけを最小限エスケープ（turndown 標準より控えめ）
    var US_RE = null;
    try {
        US_RE = new RegExp("(^|[^\\p{L}\\p{N}])_|_(?=[^\\p{L}\\p{N}]|$)", "gu");
    } catch (e) {
        US_RE = /(^|\W)_|_(?=\W|$)/g;
    }
    function escapeMd(str) {
        return str
            .replace(/\\(?=[!-\/:-@\[-`{-~])/g, "\\\\")
            .replace(/\*/g, function (m, i, s) {
                return (/\s/.test(s.charAt(i - 1) || " ") && /\s/.test(s.charAt(i + 1) || " ")) ? "*" : "\\*";
            })
            .replace(US_RE, "$1\\_")
            .replace(/`/g, "\\`")
            .replace(/[\[\]]/g, "\\$&")
            .replace(/~~/g, "\\~\\~")
            .replace(/<(?=[A-Za-z\/!?])/g, "\\<")
            .replace(/&(?=#?\w+;)/g, "\\&")
            .replace(/^(\s*)(#{1,6})(?=[ \t])/gm, "$1\\$2")
            .replace(/^(\s*)>/gm, "$1\\>")
            .replace(/^(\s*)([-+*])(?=[ \t])/gm, "$1\\$2")
            .replace(/^(\s*)(\d+)([.)])(?=[ \t])/gm, "$1$2\\$3")
            .replace(/^(=+|-+)([ \t]*)$/gm, "\\$1$2");
    }

    // ---- Markdown → プレビュー ------------------------------------------------
    function render(md) {
        if (observer) { observer.disconnect(); }
        used = {};
        blocks = [];
        fmRaw = "";
        lead = "";
        srcTail = (/\n*$/.exec(md) || [""])[0];
        out.innerHTML = "";

        var body = md;
        var fm = FM_RE.exec(md);
        if (fm) {
            fmRaw = fm[0];
            body = md.slice(fm[0].length);
            var d = document.createElement("details");
            d.className = "front-matter";
            d.setAttribute("data-fm", "");
            d.setAttribute("contenteditable", "false");
            d.innerHTML = "<summary>Front Matter（編集はコード表示で）</summary>";
            var p = document.createElement("pre");
            p.textContent = fm[1];
            d.appendChild(p);
            out.appendChild(d);
        }

        // トップレベルのブロックごとに描画し、元の Markdown（raw）と対応付ける
        var tokens = marked.lexer(body);
        for (var i = 0; i < tokens.length; i++) {
            var t = tokens[i];
            if (t.type === "space") {
                if (blocks.length) { blocks[blocks.length - 1].raw += t.raw; } else { lead += t.raw; }
                continue;
            }
            var id = blocks.length;
            var b = { raw: t.raw, dirty: false, count: 0 };
            blocks.push(b);

            tmp.innerHTML = marked.parser([t]);
            var nodes = [], hasText = false, j, n;
            for (j = 0; j < tmp.childNodes.length; j++) {
                n = tmp.childNodes[j];
                if (n.nodeType === 1) { nodes.push(n); }
                else if (n.nodeType === 3 && /\S/.test(n.nodeValue)) { hasText = true; }
            }
            if (hasText) {
                var w = document.createElement("div");
                while (tmp.firstChild) { w.appendChild(tmp.firstChild); }
                nodes = [w];
            }
            if (!nodes.length) {
                // リンク定義・HTML コメントなど表示されないブロック
                var hid = document.createElement("div");
                hid.className = "md-hidden";
                hid.setAttribute("contenteditable", "false");
                nodes = [hid];
            }
            for (j = 0; j < nodes.length; j++) {
                nodes[j].setAttribute("data-md", String(id));
                out.appendChild(nodes[j]);
            }
            b.count = nodes.length;
        }
        tmp.innerHTML = "";

        if (td) {
            var cbs = out.querySelectorAll("input[type='checkbox']");
            for (var k = 0; k < cbs.length; k++) { cbs[k].removeAttribute("disabled"); }
        }
        previewDirty = false;
        decorate();
        if (observer) { observer.observe(out, OBS); }
    }

    function decorate() {
        var codes = out.querySelectorAll("pre code[class*='language-']");
        if (opt.hljsJs && codes.length) {
            addCssOnce();
            loadLib(opt.hljsJs, "hljs", function () {
                for (var i = 0; i < codes.length; i++) {
                    try { window.hljs.highlightElement(codes[i]); } catch (e) { }
                }
                if (observer) { observer.takeRecords(); } // 色付けは編集として扱わない
            });
        }
        var diagrams = out.querySelectorAll("pre.mermaid");
        if (opt.mermaidJs && diagrams.length) {
            loadLib(opt.mermaidJs, "mermaid", function () {
                if (!mermaidReady) {
                    var dark = window.matchMedia && window.matchMedia("(prefers-color-scheme: dark)").matches;
                    window.mermaid.initialize({ startOnLoad: false, theme: dark ? "dark" : "default" });
                    mermaidReady = true;
                }
                window.mermaid.run({ nodes: diagrams, suppressErrors: true });
            });
        }
    }

    // ---- プレビュー → Markdown -----------------------------------------------
    function blockId(el) {
        return (el.nodeType === 1 && el.hasAttribute("data-md")) ? el.getAttribute("data-md") : null;
    }
    function previewToMarkdown() {
        var kids = out.childNodes, present = {}, i, el, id;
        for (i = 0; i < kids.length; i++) {
            id = blockId(kids[i]);
            if (id !== null) { present[id] = (present[id] || 0) + 1; }
        }
        var res = fmRaw + lead;
        var head = res;
        var lastRawId = -1;   // 直前に原文のまま出力したブロック
        var lastWasRaw = true;
        var inline = "";      // ブロックに入っていないテキスト・インライン要素

        function sep() {
            if (res === "" || res === head) {
                if (res && !/\n$/.test(res)) { res += "\n"; }
                return;
            }
            if (!/\n\n$/.test(res)) { res += /\n$/.test(res) ? "\n" : "\n\n"; }
        }
        function emitConverted(html) {
            var m = td.turndown(html).replace(/^\n+|\n+$/g, "");
            if (!m) { return; }
            sep();
            res += m + "\n\n";
            lastWasRaw = false;
            lastRawId = -1;
        }
        function flushInline() {
            if (/\S/.test(inline)) { emitConverted("<p>" + inline + "</p>"); }
            inline = "";
        }

        for (i = 0; i < kids.length; i++) {
            el = kids[i];
            id = blockId(el);
            if (el.nodeType === 3) { inline += esc(el.nodeValue); continue; }
            if (el.nodeType !== 1 || el.hasAttribute("data-fm")) { continue; }
            if (id === null && !BLOCK_TAGS.test(el.nodeName)) { inline += el.outerHTML; continue; }
            flushInline();

            var b = id !== null ? blocks[+id] : null;
            if (b && (/(^|\s)md-hidden(\s|$)/.test(el.className) || (!b.dirty && present[id] === b.count))) {
                if (lastRawId === +id) { continue; }           // 同じブロックの 2 要素目以降
                if (!(lastWasRaw && lastRawId === +id - 1)) { sep(); }
                res += b.raw;
                lastRawId = +id;
                lastWasRaw = true;
            } else {
                emitConverted(el.outerHTML);
            }
        }
        flushInline();
        if (!lastWasRaw) {
            res = res.replace(/\n+$/, "") + srcTail;
        }
        return res;
    }

    function onMutate(list) {
        var hit = false;
        for (var i = 0; i < list.length; i++) {
            var m = list[i];
            var n = m.target.nodeType === 1 ? m.target : m.target.parentNode;
            if (n && n.closest && n.closest("pre.mermaid")) { continue; } // 図の描画は無視
            hit = true;
            var top = m.target;
            while (top && top.parentNode !== out) { top = top.parentNode; }
            if (top) { markBlock(top); }
            if (m.type === "childList") {
                var j;
                for (j = 0; j < m.addedNodes.length; j++) { markBlock(m.addedNodes[j]); }
                for (j = 0; j < m.removedNodes.length; j++) { markBlock(m.removedNodes[j]); }
            }
        }
        if (hit) {
            previewDirty = true;
            markChanged();
        }
    }
    function markBlock(el) {
        var id = blockId(el);
        if (id !== null && blocks[+id]) { blocks[+id].dirty = true; }
    }

    // ---- モード切替・コピー --------------------------------------------------
    function getMarkdown() {
        if (mode === "code") { return code.value; }
        if (previewDirty && td) {
            current = previewToMarkdown();
            previewDirty = false;
        }
        return current;
    }

    function setMode(m) {
        if (m === mode) { return; }
        var de = document.documentElement;
        var ratio;
        if (m === "code") {
            ratio = de.scrollHeight > window.innerHeight ? window.scrollY / (de.scrollHeight - window.innerHeight) : 0;
            code.value = getMarkdown();
            out.hidden = true;
            codeWrap.hidden = false;
            mode = "code";
            code.scrollTop = ratio * (code.scrollHeight - code.clientHeight);
        } else {
            if (!hasMarked) { return; }
            ratio = code.scrollHeight > code.clientHeight ? code.scrollTop / (code.scrollHeight - code.clientHeight) : 0;
            if (code.value !== current) {
                current = code.value;
                render(current);
            }
            codeWrap.hidden = true;
            out.hidden = false;
            mode = "preview";
            window.scrollTo(0, ratio * (de.scrollHeight - window.innerHeight));
        }
        btnPreview.className = mode === "preview" ? "on" : "";
        btnCode.className = mode === "code" ? "on" : "";
    }

    function fallbackCopy(text) {
        var ta = document.createElement("textarea");
        ta.value = text;
        ta.style.position = "fixed";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.select();
        var ok = false;
        try { ok = document.execCommand("copy"); } catch (e) { }
        document.body.removeChild(ta);
        return ok;
    }

    function copy() {
        var text = getMarkdown().replace(/\n/g, opt.eol || "\r\n");
        function done(ok) {
            if (ok) {
                changed = false;
                setMsg("コピーしました → サクラで Ctrl+A → Ctrl+V", "ok");
            } else {
                setMsg("コピーに失敗しました", "warn");
            }
        }
        if (navigator.clipboard && navigator.clipboard.writeText) {
            navigator.clipboard.writeText(text).then(
                function () { done(true); },
                function () { done(fallbackCopy(text)); });
        } else {
            done(fallbackCopy(text));
        }
    }

    // ---- イベント ------------------------------------------------------------
    btnPreview.onclick = function () { setMode("preview"); };
    btnCode.onclick = function () { setMode("code"); };
    btnCopy.onclick = copy;
    code.addEventListener("input", markChanged);

    // タスクリストのチェック状態を属性に反映（Markdown 変換で拾うため）
    out.addEventListener("change", function (e) {
        var t = e.target;
        if (t && t.type === "checkbox") {
            if (t.checked) { t.setAttribute("checked", ""); } else { t.removeAttribute("checked"); }
        }
    });

    // 貼り付けは書式なしテキストとして
    out.addEventListener("paste", function (e) {
        if (!out.isContentEditable || !e.clipboardData) { return; }
        e.preventDefault();
        document.execCommand("insertText", false, e.clipboardData.getData("text/plain"));
    });

    // リンク: 編集中は Ctrl+クリックで開く。<base> を付けているため #... はスクリプトでスクロール
    out.addEventListener("click", function (e) {
        var a = (e.target && e.target.closest) ? e.target.closest("a[href]") : null;
        if (!a) { return; }
        var editable = out.isContentEditable;
        if (editable && !(e.ctrlKey || e.metaKey)) { return; }
        var href = a.getAttribute("href");
        if (href.charAt(0) === "#") {
            var id = href.slice(1);
            try { id = decodeURIComponent(id); } catch (err) { }
            var el = id ? document.getElementById(id) : document.body;
            e.preventDefault();
            if (el) { el.scrollIntoView(); }
        } else if (editable) {
            e.preventDefault();
            window.open(a.href, "_blank");
        }
    });

    window.addEventListener("beforeunload", function (e) {
        if (changed) {
            e.preventDefault();
            e.returnValue = "";
        }
    });

    // ---- 初期化 --------------------------------------------------------------
    if (hasMarked) {
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
                        return '<pre class="mermaid" contenteditable="false" data-src="' + esc(token.text) + '">' +
                               esc(token.text) + "</pre>\n";
                    }
                    return false; // それ以外は marked 標準の描画
                }
            }
        });
        if (td) {
            out.setAttribute("contenteditable", "true");
            out.spellcheck = false;
            try { document.execCommand("defaultParagraphSeparator", false, "p"); } catch (e) { }
            observer = new MutationObserver(onMutate);
        } else {
            warn("turndown.js を読み込めなかったため、プレビューでの編集は無効です（コード表示では編集できます）");
        }
        render(current);
    } else {
        warn("marked.js を読み込めませんでした。ネットに出られない環境では CONFIG.markedJs にローカルファイルを指定してください");
        btnPreview.disabled = true;
        setMode("code");
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
  --ok: #1a7f37; --attn: #9a6700;
}
@media (prefers-color-scheme: dark) {
  :root {
    --fg: #f0f6fc; --bg: #0d1117; --muted: #9198a1; --border: #3d444d;
    --code-bg: #151b23; --link: #4493f8; --bar-bg: #151b23; --warn: #f85149;
    --ok: #3fb950; --attn: #d29922;
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
  display: flex; gap: 12px; align-items: center;
  padding: 6px 16px; font-size: 12px; color: var(--muted);
  background: var(--bar-bg); border-bottom: 1px solid var(--border);
}
#bar .path { flex: 1; min-width: 0; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; text-align: right; }
#bar button {
  font: inherit; font-size: 12px; line-height: 1.6; padding: 2px 12px; cursor: pointer;
  color: var(--fg); background: var(--bg); border: 1px solid var(--border); border-radius: 6px;
}
#bar button:hover { border-color: var(--muted); }
#bar button:disabled { opacity: .5; cursor: default; }
#bar .seg { display: inline-flex; }
#bar .seg button:first-child { border-radius: 6px 0 0 6px; }
#bar .seg button + button { border-radius: 0 6px 6px 0; border-left: 0; }
#bar .seg button.on { color: #ffffff; background: var(--link); border-color: var(--link); }
#btn-copy { font-weight: 600; }
#msg { white-space: nowrap; }
#msg.dirty { color: var(--attn); }
#msg.ok { color: var(--ok); }
.warn { color: var(--warn) !important; }
.md-hidden { display: none !important; }
[hidden] { display: none !important; }
#content { outline: none; caret-color: var(--link); }
#code-wrap { max-width: 980px; margin: 0 auto; padding: 16px 40px; box-sizing: border-box; }
#code {
  display: block; width: 100%; height: calc(100vh - 72px); box-sizing: border-box;
  padding: 16px; resize: none; outline: none; tab-size: 4;
  font-family: ui-monospace, SFMono-Regular, Consolas, "BIZ UDGothic", "MS Gothic", monospace;
  font-size: 14px; line-height: 1.6;
  color: var(--fg); background: var(--code-bg); border: 1px solid var(--border); border-radius: 6px;
}
#code:focus { border-color: var(--link); }
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
  #bar, #code-wrap { display: none; }
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

// CONFIG の値（URL／絶対パス／相対パス）→ URL
function toUrl(v) {
    if (!v) { return ""; }
    if (/^[A-Za-z]:[\\\/]/.test(v) || /^\\\\/.test(v)) { return toFileUrl(v); }   // 絶対パス
    if (/^[A-Za-z][A-Za-z0-9+.\-]*:/.test(v) || /^\/\//.test(v)) { return v; }     // URL
    if (MACRO_DIR) {                                                               // 相対パス
        return toFileUrl(MACRO_DIR.replace(/\\+$/, "") + "\\" + v.replace(/\//g, "\\"));
    }
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
