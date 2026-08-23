// make-pptx.mjs — 可重複使用的簡報產生器（pptxgenjs）
// v4 —「編輯式顧問簡報」設計系統（McKinsey / a16z 語言）
//   * 扁平、克制：無陰影、無圓角裝飾、無 ghost 數字
//   * 嚴格網格：0.65" 邊距、統一模組、hairline 分隔線 + 強調色段
//   * Action title：標題即完整洞察；每頁一個 takeaway
//   * 雙字體：拉丁/數字用 Segoe UI（小寫距大寫標籤），中文用微軟正黑體
//   * 深色 takeaway panel（McKinsey insight box）+ 金色論點線
//   * v4 新增版型：statRow（統計列 / number band）、table（機構對比表）、
//     quote（引句頁 pull quote）；沿用既有版型，spec 格式向下相容
// 用法:
//   node make-pptx.mjs                 -> 用內建示範 spec，產出 demo-DSH快速指南.pptx
//   node make-pptx.mjs spec.json        -> 用自訂 spec 產出（spec 內 output 欄位指定檔名）
//   node make-pptx.mjs spec.json out.pptx
//
// spec.json 格式:
// {
//   "title", "author", "date", "footer", "output",
//   "palette": { ... 可覆蓋任一色票 ... },
//   "slides": [
//     { "layout": "cover",          "kicker", "title", "subtitle" },
//     { "layout": "section",        "title", "subtitle" },
//     { "layout": "content",        "title", "bullets", "notes" },
//     { "layout": "statement",      "kicker", "title", "support", "bullets" },
//     { "layout": "statementSplit", "kicker", "title", "support", "bullets",
//       "card": { "label", "title", "desc" } },
//     { "layout": "statementCards", "kicker", "title", "support", "bullets" },
//     { "layout": "statRow",        "kicker", "title", "support",
//       "stats": [ { "value", "label", "note" } ] },
//     { "layout": "table",          "kicker", "title", "support",
//       "columns": [ ... ], "rows": [ [ ... ] | { col: val } ] },
//     { "layout": "quote",          "kicker", "title", "support", "attribution" },
//     { "layout": "closing",        "kicker", "title", "support" }
//   ]
// }
// 卡片化規則：bullet 含「：/→/－/—」時自動拆成「主詞 + 說明」（主詞 ≤14 字）。

import fs from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import pptxgen from "pptxgenjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const [, , specArg, outArg] = process.argv;

// ---- v3 色票：深墨藍 + 鋼藍強調 + 金（僅論點線）+ 中性灰 ----
const C = {
  ink: "0B1B33",        // 深墨藍（背景/標題）
  inkSoft: "122B52",    // 深藍面板內部輔色
  accent: "1D4ED8",     // 鋼藍（主強調）
  accentBright: "3B82F6", // 亮藍（面板內標籤/細強調）
  gold: "C9A227",       // 金（論點線，點綴）
  gray: "5B6B7C",       // 次要文字
  grayLight: "8A97A8",  // 頁尾/註腳
  line: "E2E8F0",       // hairline
  panel: "F4F7FA",      // 淺卡片底
  white: "FFFFFF",
  panelText: "C7D2E0",  // 深面板內淺色文字
};
const F_CJK = "Microsoft JhengHei"; // 微軟正黑體（zh-TW 系統內建）
const F_LATIN = "Segoe UI";         // 拉丁/數字（Windows 內建，Helvetica 系）
const W = 13.33;
const H = 7.5;
const M = 0.65; // 統一左右邊距

// 依內容自動選字體：含 CJK → 正黑體，純拉丁/數字 → Segoe UI
function pickFont(t) {
  return /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/.test(t) ? F_CJK : F_LATIN;
}

const DEFAULT_SPEC = {
  title: "DeepSeek Harness 快速指南",
  author: "DeepSeek Harness",
  date: "2026-08-18",
  footer: "DEEPSEEK HARNESS | GUIDE BRIEF",
  output: "demo-DSH快速指南.pptx",
  slides: [
    {
      layout: "cover",
      kicker: "出差協作利器",
      title: "DeepSeek Harness 快速指南",
      subtitle: "啟動方式 × Codex 協作 SOP × 資料紀律",
      notes: "歡迎使用 DeepSeek Harness：帶著筆電，雙擊桌面圖示即可開工。",
    },
    { layout: "section", title: "01　啟動方式", subtitle: "開機 → 雙擊 → 開工" },
    {
      layout: "content",
      title: "桌面與檔案位置",
      bullets: [
        "桌面：啟動 DSH.bat（雙擊即開，瀏覽器自動連到 127.0.0.1:3080）",
        "07_腳本：重啟 DSH.bat / 停止 DSH.bat / 開機自動啟動.bat",
        "對話紀錄：C:/Users/grays/.dsh/sessions（停止服務不會刪除）",
        "啟動指令：npx --yes @deepseek-ai/dsh web",
      ],
      notes: "開機自啟設定好之後，只需要雙擊 啟動 DSH.bat 一次。",
    },
    {
      layout: "content",
      title: "與 Codex 協作 SOP",
      bullets: [
        "產物生成 → Codex：xlsx、程式碼、artifact 管線",
        "稽核接手 → DeepSeek：讀懂專案、Shadow Audit、Handoff",
        "中文文書 → DeepSeek：LINE 草稿、說明文件、回覆客戶",
        "交接橋樑 → AGENTS.md + Git + 明確標示",
      ],
      notes: "Codex 在台北主力產出，DeepSeek 在出差時稽核與文書，雙方互當 reviewer。",
    },
    {
      layout: "content",
      title: "資料紀律",
      bullets: [
        "真實客戶資料與模擬資料分開存放",
        "模擬一律標示 DEMO / SIMULATION",
        "未知值標「未提供」，不自行補全",
        "無法判斷時如實標示：SCORE = UNKNOWN",
      ],
    },
    {
      layout: "content",
      title: "常用任務",
      bullets: [
        "做 PPT / 圖表 / 流程圖（本工具，支援自訂 spec）",
        "稽核 Codex 留下的專案",
        "寫回覆客戶的 LINE 草稿",
        "整理試算表與資料",
      ],
      notes: "有任何需求直接對 DeepSeek Harness 說即可；非安全事項大部分可自主執行。",
    },
  ],
};

function loadSpec() {
  if (specArg) {
    const p = path.resolve(__dirname, specArg);
    if (!fs.existsSync(p)) {
      console.error(`[ERROR] 找不到 spec 檔案: ${p}`);
      process.exit(1);
    }
    return JSON.parse(fs.readFileSync(p, "utf8"));
  }
  return DEFAULT_SPEC;
}

// ---------- 小工具 ----------

function splitLead(b) {
  const m = /^([^：:→－—｜|]{1,14})[：:→－—｜|]\s*(.+)$/.exec(b);
  return m ? { t: m[1].trim(), d: m[2].trim() } : { t: b, d: "" };
}

function runs(text, baseOpts = {}) {
  const parts = String(text).split("\n");
  return parts.map((p, i) => ({
    text: p,
    options: { ...baseOpts, breakLine: i < parts.length - 1 },
  }));
}

// ---------- 共用模組 ----------

// 頁首：kicker（小寫距大寫）→ action title → 強調色段 + hairline
function addHeader(slide, PAL, s) {
  if (s.kicker) {
    slide.addText(String(s.kicker).toUpperCase(), {
      x: M, y: 0.38, w: 12.0, h: 0.3, fontSize: 10, bold: true,
      color: PAL.accent, charSpacing: 2.5, fontFace: pickFont(s.kicker),
    });
  }
  slide.addText(runs(s.title, { bold: true }), {
    x: M - 0.02, y: 0.74, w: 12.1, h: 1.05, fontSize: 22,
    color: PAL.ink, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.08,
  });
  slide.addShape("rect", { x: M, y: 1.93, w: 0.62, h: 0.045, fill: { color: PAL.accent } });
  slide.addShape("rect", { x: M, y: 1.955, w: 12.03, h: 0.012, fill: { color: PAL.line } });
}

// 頁尾：hairline + 品牌標 + 頁碼
function addFooter(slide, deck, PAL, n, total) {
  slide.addShape("rect", { x: M, y: 7.06, w: 12.03, h: 0.01, fill: { color: PAL.line } });
  if (deck.footer) {
    slide.addShape("rect", { x: M, y: 7.19, w: 0.1, h: 0.1, fill: { color: PAL.accent } });
    slide.addText(deck.footer, { x: M + 0.18, y: 7.13, w: 9.5, h: 0.22, fontSize: 8, color: PAL.grayLight, charSpacing: 1.5, fontFace: F_LATIN });
  }
  slide.addText(String(n).padStart(2, "0"), { x: 11.95, y: 7.09, w: 0.73, h: 0.22, fontSize: 9, color: PAL.gray, align: "right", fontFace: F_LATIN });
}

// ---------- 版型 ----------

// 封面：深墨藍 + 大標題 + 金色論點線（a16z 編輯感）
function layoutCover(slide, deck, s, PAL) {
  slide.background = { color: PAL.ink };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: PAL.accent } });
  if (s.kicker) {
    slide.addText(String(s.kicker).toUpperCase(), { x: 0.9, y: 1.72, w: 11, h: 0.32, fontSize: 11, bold: true, color: PAL.accentBright, charSpacing: 3, fontFace: pickFont(s.kicker) });
  }
  slide.addText(runs(s.title, { bold: true }), { x: 0.88, y: 2.22, w: 11.4, h: 1.8, fontSize: 38, color: PAL.white, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.12 });
  slide.addShape("rect", { x: 0.95, y: 4.12, w: 1.5, h: 0.05, fill: { color: PAL.gold } });
  if (s.subtitle) slide.addText(s.subtitle, { x: 0.9, y: 4.42, w: 11, h: 0.7, fontSize: 15, color: PAL.panelText, fontFace: F_CJK });
  slide.addShape("rect", { x: 0.9, y: 6.58, w: 0.12, h: 0.12, fill: { color: PAL.accent } });
  slide.addText(`${deck.author || ""}　${deck.date || ""}`, { x: 1.14, y: 6.5, w: 9, h: 0.3, fontSize: 10.5, color: PAL.panelText, fontFace: F_CJK });
}

// 章節頁：深墨藍 + 大數字標題
function layoutSection(slide, deck, s, PAL) {
  slide.background = { color: PAL.ink };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: PAL.accent } });
  slide.addText(runs(s.title, { bold: true }), { x: 1.0, y: 2.85, w: 11.3, h: 1.0, fontSize: 34, color: PAL.white, fontFace: F_CJK });
  slide.addShape("rect", { x: 1.05, y: 3.95, w: 1.3, h: 0.045, fill: { color: PAL.gold } });
  if (s.subtitle) slide.addText(s.subtitle, { x: 1.0, y: 4.2, w: 11.3, h: 0.5, fontSize: 14, color: PAL.panelText, fontFace: F_CJK });
}

// 論點頁（通用）
function layoutStatement(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, s);
  let y = 2.35;
  if (s.support) {
    slide.addText(s.support, { x: M, y, w: 11.9, h: 0.5, fontSize: 13.5, color: PAL.gray, fontFace: F_CJK });
    y += 0.58;
  }
  if (s.bullets && s.bullets.length) {
    slide.addText(
      s.bullets.map((b) => ({ text: b, options: { bullet: { code: "25AA", indent: 14 }, color: "3B4754", fontSize: 14, paraSpaceAfter: 14, lineSpacingMultiple: 1.25 } })),
      { x: M, y, w: 11.9, h: 6.85 - y, valign: "top", fontFace: F_CJK }
    );
  }
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 左文右 takeaway panel（McKinsey insight box）
function layoutStatementSplit(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, s);
  let y = 2.35;
  if (s.support) {
    slide.addText(s.support, { x: M, y, w: 7.5, h: 0.55, fontSize: 13, color: PAL.gray, fontFace: F_CJK });
    y += 0.62;
  }
  if (s.bullets && s.bullets.length) {
    slide.addText(
      s.bullets.map((b) => ({ text: b, options: { bullet: { code: "25AA", indent: 14 }, color: "3B4754", fontSize: 14, paraSpaceAfter: 15, lineSpacingMultiple: 1.25 } })),
      { x: M, y, w: 7.5, h: 6.85 - y, valign: "top", fontFace: F_CJK }
    );
  }
  const card = s.card;
  if (card) {
    const cx = 8.6, cw = 4.08, cy = 2.35, ch = 4.2;
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: ch, fill: { color: PAL.ink } });
    slide.addShape("rect", { x: cx, y: cy, w: cw, h: 0.07, fill: { color: PAL.accent } });
    if (card.label) slide.addText(String(card.label).toUpperCase(), { x: cx + 0.4, y: cy + 0.5, w: cw - 0.8, h: 0.28, fontSize: 10, bold: true, color: PAL.accentBright, charSpacing: 2, fontFace: F_LATIN });
    if (card.title) slide.addText(runs(card.title, { bold: true }), { x: cx + 0.4, y: cy + 0.92, w: cw - 0.8, h: 1.05, fontSize: 17, color: PAL.white, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.15 });
    slide.addShape("rect", { x: cx + 0.4, y: cy + 2.05, w: cw - 0.8, h: 0.012, fill: { color: "24405F" } });
    if (card.desc) slide.addText(card.desc, { x: cx + 0.4, y: cy + 2.25, w: cw - 0.8, h: 1.7, fontSize: 12.5, color: PAL.panelText, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.3 });
  }
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 卡片網格：扁平淺底面板 + 強調色邊條（無圓角、無陰影）
function layoutStatementCards(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, s);
  let top = 2.35;
  if (s.support) {
    slide.addText(s.support, { x: M, y: top, w: 11.9, h: 0.45, fontSize: 13, color: PAL.gray, fontFace: F_CJK });
    top += 0.55;
  }
  const items = (s.bullets || []).map(splitLead);
  if (!items.length) { addFooter(slide, deck, PAL, n, deck.slides.length); return; }
  const gap = 0.26;
  const ch = 1.9;
  const cols = items.length <= 3 ? items.length : 2;
  const rows = Math.ceil(items.length / cols);
  const areaW = 12.03;
  const cw = (areaW - gap * (cols - 1)) / cols;
  const areaBottom = 6.9;
  const gridH = rows * ch + (rows - 1) * gap;
  const startY = top + Math.max(0, (areaBottom - top - gridH) / 2);
  items.forEach((it, i) => {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = M + c * (cw + gap);
    const y = startY + r * (ch + gap);
    slide.addShape("rect", { x, y, w: cw, h: ch, fill: { color: PAL.panel } });
    slide.addShape("rect", { x, y, w: 0.07, h: ch, fill: { color: PAL.accent } });
    if (it.d) {
      slide.addText(it.t, { x: x + 0.32, y: y + 0.28, w: cw - 0.62, h: 0.42, fontSize: 14.5, bold: true, color: PAL.ink, fontFace: F_CJK });
      slide.addText(it.d, { x: x + 0.32, y: y + 0.75, w: cw - 0.62, h: ch - 1.0, fontSize: 12.5, color: PAL.gray, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.25 });
    } else {
      slide.addText(it.t, { x: x + 0.32, y: y + 0.2, w: cw - 0.62, h: ch - 0.4, fontSize: 14, bold: true, color: PAL.ink, fontFace: F_CJK, valign: "middle", lineSpacingMultiple: 1.2 });
    }
  });
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 單頁總覽（one-pager）：深墨藍底 + 左文右卡片，一頁裝完
function layoutOnePager(slide, deck, s, PAL) {
  slide.background = { color: PAL.ink };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: PAL.gold } });
  // 左欄：kicker + 大標 + support
  if (s.kicker) slide.addText(String(s.kicker).toUpperCase(), { x: 0.9, y: 0.85, w: 11.5, h: 0.3, fontSize: 10.5, bold: true, color: PAL.accentBright, charSpacing: 3, fontFace: F_LATIN });
  slide.addText(runs(s.title, { bold: true }), { x: 0.88, y: 1.3, w: 11.5, h: 1.5, fontSize: 24, color: PAL.white, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.18 });
  if (s.support) slide.addText(s.support, { x: 0.9, y: 2.9, w: 11.5, h: 0.55, fontSize: 13, color: PAL.panelText, fontFace: F_CJK });
  // 卡片網格（2 欄；最多 6 張，3 排，用較矮卡片）
  const items = (s.bullets || []).map(splitLead);
  const cols = 2;
  const gap = 0.22;
  const cw = (11.53 - gap) / cols;
  const ch = 0.86;
  const x0 = 0.9;
  const y0 = 3.62;
  items.forEach((it, i) => {
    const r = Math.floor(i / cols);
    const c = i % cols;
    const x = x0 + c * (cw + gap);
    const y = y0 + r * (ch + gap);
    slide.addShape("rect", { x, y, w: cw, h: ch, fill: { color: PAL.inkSoft } });
    slide.addShape("rect", { x, y, w: 0.06, h: ch, fill: { color: PAL.gold } });
    slide.addText(it.t, { x: x + 0.28, y: y + 0.12, w: cw - 0.5, h: 0.32, fontSize: 12, bold: true, color: PAL.white, fontFace: F_CJK });
    if (it.d) slide.addText(it.d, { x: x + 0.28, y: y + 0.44, w: cw - 0.5, h: ch - 0.5, fontSize: 9.5, color: PAL.panelText, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.12 });
  });
  // 底部：作者 + 聯絡
  slide.addShape("rect", { x: 0.9, y: 6.72, w: 0.12, h: 0.12, fill: { color: PAL.accent } });
  slide.addText(`${deck.author || ""}　${deck.footer || ""}`, { x: 1.14, y: 6.64, w: 11, h: 0.3, fontSize: 10.5, color: PAL.panelText, fontFace: F_CJK });
}

// 結尾：深墨藍 + 大論點 + 金線
function layoutClosing(slide, deck, s, PAL) {  slide.background = { color: PAL.ink };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: PAL.accent } });
  if (s.kicker) slide.addText(String(s.kicker).toUpperCase(), { x: 0.9, y: 1.6, w: 10, h: 0.3, fontSize: 10.5, bold: true, color: PAL.accentBright, charSpacing: 3, fontFace: F_LATIN });
  slide.addText(runs(s.title, { bold: true }), { x: 0.88, y: 2.15, w: 11.4, h: 1.8, fontSize: 27, color: PAL.white, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.22 });
  slide.addShape("rect", { x: 0.95, y: 4.1, w: 1.5, h: 0.05, fill: { color: PAL.gold } });
  if (s.support) slide.addText(s.support, { x: 0.9, y: 4.4, w: 10.5, h: 0.6, fontSize: 14, color: PAL.panelText, fontFace: F_CJK });
  slide.addShape("rect", { x: 0.9, y: 6.58, w: 0.12, h: 0.12, fill: { color: PAL.accent } });
  slide.addText(`${deck.author || ""}　${deck.date || ""}`, { x: 1.14, y: 6.5, w: 9, h: 0.3, fontSize: 10.5, color: PAL.panelText, fontFace: F_CJK });
  slide.addText("FIN", { x: 11.8, y: 6.48, w: 0.9, h: 0.3, fontSize: 10, bold: true, color: PAL.accentBright, align: "right", charSpacing: 3, fontFace: F_LATIN });
}

// 內容頁（標題 + 項目符號）
function layoutContent(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, { ...s, kicker: s.kicker || "" });
  const bullets = s.bullets || [];
  slide.addText(
    bullets.map((b) => ({ text: b, options: { bullet: { code: "25CF" }, color: "3B4754", fontSize: 15, paraSpaceAfter: 13 } })),
    { x: M, y: 2.3, w: 11.9, h: 4.6, valign: "top", lineSpacingMultiple: 1.25, fontFace: F_CJK }
  );
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 統計列（McKinsey number band）：一列大數字 + 標籤，克制單色
function layoutStatRow(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, s);
  let y = 2.35;
  if (s.support) {
    slide.addText(s.support, { x: M, y, w: 11.9, h: 0.45, fontSize: 13, color: PAL.gray, fontFace: F_CJK });
    y += 0.6;
  }
  const stats = s.stats || [];
  if (!stats.length) { addFooter(slide, deck, PAL, n, deck.slides.length); return; }
  const gap = 0.28;
  const cw = (12.03 - gap * (stats.length - 1)) / stats.length;
  const ch = 2.7;
  const startY = y + Math.max(0, (6.7 - y - ch) / 2);
  stats.forEach((st, i) => {
    const x = M + i * (cw + gap);
    slide.addShape("rect", { x, y: startY, w: 0.05, h: ch, fill: { color: PAL.accent } });
    slide.addText(String(st.value), { x: x + 0.35, y: startY + 0.3, w: cw - 0.6, h: 1.1, fontSize: 42, bold: true, color: PAL.ink, fontFace: pickFont(String(st.value)), valign: "top" });
    slide.addShape("rect", { x: x + 0.35, y: startY + 1.55, w: cw - 0.6, h: 0.014, fill: { color: PAL.line } });
    slide.addText(String(st.label), { x: x + 0.35, y: startY + 1.7, w: cw - 0.6, h: 0.55, fontSize: 12.5, color: PAL.gray, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.2 });
    if (st.note) slide.addText(String(st.note), { x: x + 0.35, y: startY + 2.25, w: cw - 0.6, h: 0.4, fontSize: 9.5, color: PAL.grayLight, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.15 });
  });
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 機構對比表：hairline 分隔、表頭小寫距大寫、狀態欄強調色、正文 15pt、行高自適應
// cell 值可為 string 或 [主文, 附註]：附註以 11.5pt muted gray 印在主文下方（claim boundary 用）
function estLines(text, widthIn, sizePt) {
  // 粗估換行：CJK 字寬 = 1em，拉丁 ≈ 0.55em（Segoe UI 平均字寬）
  let used = 0, lines = 1;
  for (const ch of String(text)) {
    const w = /[\u2E80-\u9FFF\uF900-\uFAFF\uFF00-\uFFEF\u3000-\u303F]/.test(ch) ? sizePt / 72 : (0.55 * sizePt) / 72;
    if (used + w > widthIn) { lines++; used = 0; }
    used += w;
  }
  return lines;
}
function layoutTable(slide, deck, s, PAL, n) {
  slide.background = { color: PAL.white };
  addHeader(slide, PAL, s);
  let y = 2.35;
  if (s.support) {
    slide.addText(s.support, { x: M, y, w: 11.9, h: 0.45, fontSize: 13, color: PAL.gray, fontFace: F_CJK });
    y += 0.55;
  }
  const cols = s.columns || [];
  const rows = s.rows || [];
  const x0 = M, w = 12.03;
  const hh = 0.42;
  const CELL = 15;   // 正文 ≥14pt（14–16 範圍取 15）
  const NOTE = 11.5; // 附註（11–12pt muted gray）
  const pad = 0.15;
  cols.forEach((c, i) => {
    const cw = w / cols.length;
    slide.addText(String(c).toUpperCase(), { x: x0 + i * cw + pad, y, w: cw - pad * 2, h: hh, fontSize: 11, bold: true, color: PAL.accent, charSpacing: 1.5, fontFace: F_LATIN, valign: "middle" });
  });
  slide.addShape("rect", { x: x0, y: y + hh, w, h: 0.016, fill: { color: PAL.ink } });
  y += hh + 0.3;
  rows.forEach((r, ri) => {
    const cells = Array.isArray(r) ? r : cols.map((c) => (r[c] != null ? String(r[c]) : ""));
    const cw = w / cols.length;
    const innerW = cw - pad * 2;
    // 行高：主文行數（15pt）+ 附註行數（11.5pt）
    const mainLines = Math.max(...cells.map((c) => estLines(Array.isArray(c) ? c[0] : c, innerW, CELL)));
    const noteLines = Math.max(...cells.map((c) => (Array.isArray(c) && c[1] != null ? estLines(c[1], innerW, NOTE) : 0)));
    const rowH = Math.max(0.7, mainLines * (CELL / 72) * 1.4 + noteLines * (NOTE / 72) * 1.35 + 0.34);
    if (ri > 0) slide.addShape("rect", { x: x0, y: y - 0.14, w, h: 0.008, fill: { color: PAL.line } });
    cells.forEach((cell, ci) => {
      const main = Array.isArray(cell) ? String(cell[0]) : String(cell);
      const note = Array.isArray(cell) && cell[1] != null ? String(cell[1]) : null;
      const isStatus = /^(TESTED|PENDING|VERIFIED|NOT YET PROVEN)/.test(main);
      const baseOpts = { x: x0 + ci * cw + pad, w: innerW, fontSize: CELL, color: isStatus ? PAL.accent : PAL.ink, bold: isStatus, fontFace: pickFont(main), lineSpacingMultiple: 1.2 };
      if (note) {
        // 主文上緣對齊、附註跟隨其下
        const mainH = estLines(main, innerW, CELL) * (CELL / 72) * 1.4;
        slide.addText(main, { ...baseOpts, y: y - 0.08, h: mainH, valign: "top" });
        slide.addText(note, { x: x0 + ci * cw + pad, y: y - 0.08 + mainH + 0.04, w: innerW, h: rowH - mainH - 0.1, fontSize: NOTE, color: PAL.grayLight, fontFace: F_LATIN, valign: "top", lineSpacingMultiple: 1.2 });
      } else {
        slide.addText(main, { ...baseOpts, y: y - 0.1, h: rowH + 0.1, valign: "middle" });
      }
    });
    y += rowH;
  });
  addFooter(slide, deck, PAL, n, deck.slides.length);
}

// 引句頁（pull quote）：深墨藍 + 單一論點 + 金線
function layoutQuote(slide, deck, s, PAL) {
  slide.background = { color: PAL.ink };
  slide.addShape("rect", { x: 0, y: 0, w: W, h: 0.06, fill: { color: PAL.accent } });
  if (s.kicker) slide.addText(String(s.kicker).toUpperCase(), { x: 0.9, y: 1.55, w: 11, h: 0.3, fontSize: 10.5, bold: true, color: PAL.accentBright, charSpacing: 3, fontFace: F_LATIN });
  slide.addText(runs(s.title, { bold: true }), { x: 0.88, y: 2.1, w: 11.5, h: 2.2, fontSize: 30, color: PAL.white, fontFace: F_CJK, valign: "top", lineSpacingMultiple: 1.22 });
  slide.addShape("rect", { x: 0.95, y: 4.5, w: 1.2, h: 0.045, fill: { color: PAL.gold } });
  if (s.support) slide.addText(s.support, { x: 0.9, y: 4.85, w: 11, h: 0.5, fontSize: 13.5, color: PAL.panelText, fontFace: F_CJK });
  if (s.attribution) slide.addText(s.attribution, { x: 0.9, y: 6.3, w: 11, h: 0.4, fontSize: 10.5, color: PAL.grayLight, fontFace: F_CJK });
}

// ---------- 組裝 ----------

function build(deck) {
  const pres = new pptxgen();
  pres.defineLayout({ name: "WIDE", width: W, height: H });
  pres.layout = "WIDE";
  pres.author = deck.author || "";
  pres.company = deck.company || "";
  pres.title = deck.title || "";
  pres.subject = deck.subject || "";

  const PAL = { ...C, ...(deck.palette || {}) };
  const total = deck.slides.length;

  deck.slides.forEach((s, i) => {
    const n = i + 1;
    const slide = pres.addSlide();
    const L = s.layout || "content";
    if (L === "cover" || L === "coverLight") layoutCover(slide, deck, s, PAL);
    else if (L === "section") layoutSection(slide, deck, s, PAL);
    else if (L === "statement") layoutStatement(slide, deck, s, PAL, n);
    else if (L === "statementSplit") layoutStatementSplit(slide, deck, s, PAL, n);
    else if (L === "statementCards") layoutStatementCards(slide, deck, s, PAL, n);
    else if (L === "onePager") layoutOnePager(slide, deck, s, PAL);
    else if (L === "statRow") layoutStatRow(slide, deck, s, PAL, n);
    else if (L === "table") layoutTable(slide, deck, s, PAL, n);
    else if (L === "quote") layoutQuote(slide, deck, s, PAL);
    else if (L === "closing") layoutClosing(slide, deck, s, PAL);
    else layoutContent(slide, deck, s, PAL, n);

    if (L !== "cover" && L !== "section" && L !== "closing") {
      addFooter(slide, deck, PAL, n, total);
    }
    if (s.notes) slide.addNotes(s.notes);
  });

  return pres;
}

function main() {
  const spec = loadSpec();
  const outName = outArg || spec.output || "demo.pptx";
  const outPath = path.resolve(__dirname, outName);
  const pres = build(spec);
  pres
    .writeFile({ fileName: outPath })
    .then(() => {
      const size = fs.statSync(outPath).size;
      console.log(`[OK] 已產生: ${outPath}`);
      console.log(`     ${spec.slides.length} 頁，${(size / 1024).toFixed(1)} KB`);
    })
    .catch((e) => {
      console.error("[ERROR]", e);
      process.exit(1);
    });
}

main();
