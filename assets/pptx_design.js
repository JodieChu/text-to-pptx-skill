// pptx_design.js — text-to-pptx 可复用版式库（已在「耀途×光合」对比 PPT 实战验证）
// 用法：
//   const pptxgen = require("pptxgenjs");
//   const D = require("./pptx_design");
//   const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE";
//   const C = D.palette("executive");           // 选/改调色板
//   D.titleSlide(pres, C, {...});               // 封面
//   const s = pres.addSlide(); D.bgPaper(s, C); D.header(pres, s, C, "KICKER", "标题", C.accentA);
//   D.cardGrid(pres, s, C, items, {...});       // 内容
//   D.footer(pres, s, C, "页脚文字"); D.pageNum(s, C, 3);
//   await pres.writeFile({ fileName: "x.pptx" });
//
// 画布：LAYOUT_WIDE = 13.33in × 7.5in。所有坐标按此设计。
// 字体：默认 "PingFang SC"（Mac 原生清晰黑体；Windows PowerPoint 会替换为雅黑，亦正常）。
//
// ⚠️ 视觉 QA 必做：生成后用 scripts/qa_render.sh 出逐页图，肉眼检查溢出/碰撞。
//    LibreOffice 渲染中文会用「楷体」替代（它没 PingFang），那是渲染假象，PowerPoint 里是黑体。只看版式。

const W = 13.33, H = 7.5;
const FONT = "PingFang SC";

// ---------- 调色板 ----------
// 每个调色板：navy(深主色) / navy2(深色块) / ink/body/muted(文字三档) / paper(浅底) /
//             card/line(卡片) / accentA+aBg / accentB+bBg(两个强调色，可代表两个对比对象)
// 每套都带一句中文定位 + 触发词。flat:true = 极简观感（无投影、靠发丝线分隔）。
const PALETTES = {
  // 投资/金融/严肃汇报：海军蓝 + 金 + 绿（「耀途×光合」用的就是它）
  executive: {
    _desc:"投资/金融/严肃汇报 · 海军蓝+金+绿", _triggers:["商务","金融","投资","严肃","汇报","executive"],
    navy:"13294B", navy2:"1F3A5F", ink:"1A1A2E", body:"39414D", muted:"7A8699",
    paper:"F6F5F1", card:"FFFFFF", line:"E2E1DB",
    accentA:"C8901C", aBg:"FBF1DC", accentB:"2A8C7C", bBg:"E1F0EC",
    white:"FFFFFF", onDarkMuted:"C7D4E6", onDarkFaint:"9FB3CF",
  },
  // 科技/产品：深石板 + 青 + 珊瑚
  tech: {
    _desc:"科技/产品 · 深石板+青+珊瑚", _triggers:["科技","产品","互联网","tech"],
    navy:"0F172A", navy2:"1E293B", ink:"0F172A", body:"334155", muted:"94A3B8",
    paper:"F8FAFC", card:"FFFFFF", line:"E2E8F0",
    accentA:"0E7490", aBg:"CFFAFE", accentB:"EA580C", bBg:"FFEDD5",
    white:"FFFFFF", onDarkMuted:"CBD5E1", onDarkFaint:"94A3B8",
  },
  // 高科技/未来感：近黑靛蓝 + 电光紫 + 青
  electric: {
    _desc:"高科技/未来/AI · 靛蓝+电光紫+青", _triggers:["未来","高科技","炫","AI","酷","electric","赛博"],
    navy:"0E1530", navy2:"1A2347", ink:"141A2E", body:"36405C", muted:"8A93AC",
    paper:"F7F8FC", card:"FFFFFF", line:"E3E7F0",
    accentA:"4F46E5", aBg:"E8E7FB", accentB:"06B6D4", bBg:"D6F5FB",
    white:"FFFFFF", onDarkMuted:"C5CCE6", onDarkFaint:"8A93AC",
  },
  // 极简：纯白大留白 + 发丝线（无投影）+ 克制石墨蓝
  minimal: {
    _desc:"极简/留白 · 纯白+发丝线+石墨蓝", _triggers:["简约","极简","留白","干净","简洁","minimal"],
    navy:"242629", navy2:"31343A", ink:"17181A", body:"45484D", muted:"9AA0A6",
    paper:"FFFFFF", card:"FFFFFF", line:"E4E4E6",
    accentA:"2B4C7E", aBg:"EAEFF6", accentB:"6B7280", bBg:"F0F1F2",
    white:"FFFFFF", onDarkMuted:"D2D5DA", onDarkFaint:"9AA0A6", flat:true,
  },
  // 高智感/编辑感：深墨紫 + 香槟金 + 暖米白（知性、精致）
  premium: {
    _desc:"高智感/知性/编辑感 · 深墨+香槟金+米白", _triggers:["高智","高级感","知性","精致","编辑","质感","premium","高端"],
    navy:"211D2B", navy2:"2F2A3C", ink:"201C29", body:"4B4656", muted:"8B8596",
    paper:"F4F1EA", card:"FBF9F4", line:"E6E0D4",
    accentA:"A8842F", aBg:"F0E7CF", accentB:"4F6D7A", bBg:"E3EAEC",
    white:"FFFFFF", onDarkMuted:"D9D3C6", onDarkFaint:"A79E8D",
  },
  // 高级灰单色：炭灰阶 + 唯一一抹暖铜（最克制、最耐看）
  mono: {
    _desc:"高级灰/单色 · 炭灰阶+一抹暖铜", _triggers:["高级灰","单色","黑白","莫兰迪","性冷淡","mono","灰"],
    navy:"2B2B2B", navy2:"3A3A3A", ink:"1C1C1C", body:"4D4D4D", muted:"9A9A9A",
    paper:"F7F7F6", card:"FFFFFF", line:"E5E5E3",
    accentA:"33373B", aBg:"ECECEA", accentB:"9C7A3C", bBg:"F0E9DC",
    white:"FFFFFF", onDarkMuted:"D4D4D4", onDarkFaint:"9A9A9A", flat:true,
  },
  // 学术/研究：墨绿 + 赭 + 苔
  scholar: {
    _desc:"学术/研究/政策 · 墨绿+赭+苔", _triggers:["学术","研究","政策","报告","scholar","稳重"],
    navy:"1B3A2F", navy2:"294D40", ink:"1A2E26", body:"3A4A42", muted:"7E8C84",
    paper:"F5F4EE", card:"FFFFFF", line:"E0DFD6",
    accentA:"B07A2E", aBg:"F3E8D5", accentB:"4F7A52", bBg:"E3EEE2",
    white:"FFFFFF", onDarkMuted:"CBD8CF", onDarkFaint:"9FB0A4",
  },
};
function palette(name){ return JSON.parse(JSON.stringify(PALETTES[name] || PALETTES.executive)); }
// 列出所有风格（名称→定位+触发词），供 SKILL/调试用
function listStyles(){ return Object.entries(PALETTES).map(([k,v])=>({ name:k, desc:v._desc, triggers:v._triggers })); }

const shadow = () => ({ type:"outer", color:"000000", blur:9, offset:3, angle:135, opacity:0.12 });
// flat 风格（minimal/mono）不用投影，靠发丝线分隔
function sh(C){ return (C && C.flat) ? undefined : shadow(); }

// ---------- 基础元件 ----------
function bgPaper(s, C){ s.background = { color: C.paper }; }
function bgDark(s, C){ s.background = { color: C.navy }; }

function pageNum(s, C, n){
  s.addText(String(n).padStart(2,"0"), { x: W-1.1, y: H-0.62, w:0.8, h:0.4,
    fontFace:FONT, fontSize:11, color:C.muted, align:"right" });
}
function footer(pres, s, C, text){
  s.addText(text, { x:0.6, y:H-0.62, w:9.5, h:0.4, fontFace:FONT, fontSize:10, color:C.muted, align:"left" });
}
// 内容页页眉：竖色条 + 英文小标签 + 中文大标题
function header(pres, s, C, kicker, title, accent){
  accent = accent || C.navy;
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.62, w:0.16, h:0.62, fill:{color:accent} });
  s.addText(kicker, { x:0.9, y:0.58, w:11, h:0.3, fontFace:FONT, fontSize:12.5, color:accent, bold:true, charSpacing:2, margin:0 });
  s.addText(title, { x:0.88, y:0.86, w:11.8, h:0.62, fontFace:FONT, fontSize:27, color:C.ink, bold:true, margin:0 });
}
// 深色页页眉（标题白字）
function headerDark(pres, s, C, kicker, title, accent){
  accent = accent || C.accentA;
  s.addShape(pres.shapes.RECTANGLE, { x:0.6, y:0.62, w:0.16, h:0.62, fill:{color:accent} });
  s.addText(kicker, { x:0.9, y:0.58, w:11, h:0.3, fontFace:FONT, fontSize:12.5, color:accent, bold:true, charSpacing:2, margin:0 });
  s.addText(title, { x:0.88, y:0.86, w:11.8, h:0.6, fontFace:FONT, fontSize:27, color:C.white, bold:true, margin:0 });
}
function chip(pres, s, C, x, y, w, txt, fg, bg){
  s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x, y, w, h:0.36, fill:{color:bg}, rectRadius:0.18 });
  s.addText(txt, { x, y, w, h:0.36, fontFace:FONT, fontSize:12, color:fg, bold:true, align:"center", valign:"middle" });
}

// ---------- 封面 ----------
// opts: { kicker, titleParts:[{text,color,bold}], subtitle, descLines:[..], meta }
function titleSlide(pres, C, opts){
  const s = pres.addSlide(); bgDark(s, C);
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.28, h:H, fill:{color:C.accentA} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.28, y:0, w:0.16, h:H, fill:{color:C.accentB} });
  if (opts.kicker) s.addText(opts.kicker, { x:1.1, y:1.35, w:10, h:0.4, fontFace:FONT, fontSize:15, color:C.onDarkFaint, bold:true, charSpacing:3, margin:0 });
  s.addText(opts.titleParts, { x:1.05, y:1.95, w:11.5, h:1.2, fontFace:FONT, fontSize:50, bold:true, margin:0 });
  if (opts.subtitle) s.addText(opts.subtitle, { x:1.1, y:3.25, w:11, h:0.7, fontFace:FONT, fontSize:25, color:C.white, bold:true, margin:0 });
  s.addShape(pres.shapes.LINE, { x:1.12, y:4.2, w:4.2, h:0, line:{color:"3C5680", width:1.5} });
  if (opts.descLines) s.addText(opts.descLines.map((t,i)=>({ text:t, options:{ breakLine:true, color: i===0?C.onDarkMuted:C.onDarkFaint } })),
    { x:1.12, y:4.45, w:10.8, h:1.1, fontFace:FONT, fontSize:14.5, lineSpacingMultiple:1.25, margin:0 });
  if (opts.meta) s.addText(opts.meta, { x:1.12, y:6.5, w:7, h:0.4, fontFace:FONT, fontSize:12.5, color:"7A8FB0", margin:0 });
  return s;
}

// ---------- 2 列内容卡片网格（每卡：大序号 + 标题 + 正文）----------
// items: [{ no, title, body }]  最多 4 张（2×2）。accent 用于序号色。
function cardGrid(pres, s, C, items, opts){
  opts = opts || {};
  const accent = opts.accent || C.accentA, accentAlt = opts.accentAlt;
  const x0=0.65, y0=1.95, gw=6.0, gh=2.18, gx=0.32, gy=0.28;
  items.slice(0,4).forEach((it,i)=>{
    const cx = x0 + (i%2)*(gw+gx), cy = y0 + Math.floor(i/2)*(gh+gy);
    const ac = (accentAlt && i%2) ? accentAlt : accent;
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y:cy, w:gw, h:gh, fill:{color:C.card}, line:{color:C.line, width:1}, rectRadius:0.06, shadow:sh(C) });
    if (opts.bar) s.addShape(pres.shapes.RECTANGLE, { x:cx, y:cy, w:0.1, h:gh, fill:{color:ac} });
    s.addText(it.no, { x:cx+0.28, y:cy+0.22, w:1.0, h:0.8, fontFace:FONT, fontSize:30, color:ac, bold:true, margin:0 });
    s.addText(it.title, { x:cx+1.25, y:cy+0.26, w:gw-1.5, h:0.75, fontFace:FONT, fontSize:15, color:C.ink, bold:true, lineSpacingMultiple:1.1, valign:"top", margin:0 });
    s.addText(it.body, { x:cx+1.25, y:cy+1.02, w:gw-1.5, h:1.0, fontFace:FONT, fontSize:12, color:C.body, lineSpacingMultiple:1.22, valign:"top", margin:0 });
  });
}

// ---------- 深色页：4 个大数字 callout ----------
// stats: [{ big, label, desc, color }]  建议 4 个
function statCallouts(pres, s, C, stats, footnote){
  const x0=0.65, y=2.1, cw=2.95, cg=0.18, ch=3.7;
  stats.slice(0,4).forEach((st,i)=>{
    const cx = x0 + i*(cw+cg), col = st.color || (i%2 ? C.accentB : C.accentA);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y, w:cw, h:ch, fill:{color:C.navy2}, rectRadius:0.07 });
    s.addShape(pres.shapes.RECTANGLE, { x:cx, y, w:cw, h:0.1, fill:{color:col} });
    s.addText(st.big, { x:cx, y:y+0.55, w:cw, h:1.0, fontFace:FONT, fontSize:33, color:C.white, bold:true, align:"center", margin:0 });
    s.addText(st.label, { x:cx, y:y+1.75, w:cw, h:0.4, fontFace:FONT, fontSize:15, color:col, bold:true, align:"center", margin:0 });
    s.addText(st.desc, { x:cx+0.3, y:y+2.35, w:cw-0.6, h:1.1, fontFace:FONT, fontSize:12, color:C.onDarkMuted, align:"center", lineSpacingMultiple:1.25, valign:"top", margin:0 });
  });
  if (footnote) s.addText(footnote, { x:0.65, y:6.25, w:12.0, h:0.6, fontFace:FONT, fontSize:13.5, color:C.onDarkFaint, italic:true, align:"center", margin:0 });
}

// ---------- 4 步流程条（带编号圆 + 箭头）----------
// steps: [{ title, body }]  建议 4 个
function processFlow(pres, s, C, steps, accent, y){
  accent = accent || C.accentB; y = y || 2.0;
  const x0=0.65, sw=2.92, sg=0.2, shh=1.95;
  steps.slice(0,4).forEach((st,i)=>{
    const cx = x0 + i*(sw+sg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:cx, y, w:sw, h:shh, fill:{color:C.card}, line:{color:C.line, width:1}, rectRadius:0.06, shadow:sh(C) });
    s.addShape(pres.shapes.OVAL, { x:cx+0.28, y:y+0.28, w:0.62, h:0.62, fill:{color:accent} });
    s.addText(String(i+1), { x:cx+0.28, y:y+0.28, w:0.62, h:0.62, fontFace:FONT, fontSize:20, color:C.white, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(st.title, { x:cx+1.05, y:y+0.34, w:sw-1.2, h:0.5, fontFace:FONT, fontSize:15.5, color:C.ink, bold:true, valign:"middle", margin:0 });
    s.addText(st.body, { x:cx+0.3, y:y+1.05, w:sw-0.55, h:0.8, fontFace:FONT, fontSize:11.5, color:C.body, lineSpacingMultiple:1.18, valign:"top", margin:0 });
    if (i<steps.length-1 && i<3) s.addText("›", { x:cx+sw-0.02, y:y+0.55, w:0.24, h:0.6, fontFace:FONT, fontSize:30, color:accent, bold:true, align:"center", valign:"middle", margin:0 });
  });
}

// ---------- 深色页：编号要点行（左标题 + 右详述）----------
// rows: [{ title, detail }]  建议 ≤5
function takeawayRows(pres, s, C, rows, accent){
  accent = accent || C.accentA;
  const y0=1.95, rh=1.0, rg=0.1;
  rows.slice(0,5).forEach((it,i)=>{
    const ry = y0 + i*(rh+rg);
    s.addShape(pres.shapes.ROUNDED_RECTANGLE, { x:0.65, y:ry, w:12.0, h:rh, fill:{color:C.navy2}, rectRadius:0.05 });
    s.addShape(pres.shapes.OVAL, { x:0.9, y:ry+0.24, w:0.52, h:0.52, fill:{color:accent} });
    s.addText(String(i+1), { x:0.9, y:ry+0.24, w:0.52, h:0.52, fontFace:FONT, fontSize:18, color:C.navy, bold:true, align:"center", valign:"middle", margin:0 });
    s.addText(it.title, { x:1.65, y:ry+0.13, w:3.7, h:0.75, fontFace:FONT, fontSize:15, color:C.white, bold:true, valign:"middle", lineSpacingMultiple:1.05, margin:0 });
    s.addText(it.detail, { x:5.45, y:ry+0.12, w:7.0, h:0.78, fontFace:FONT, fontSize:11.5, color:C.onDarkMuted, valign:"middle", lineSpacingMultiple:1.18, margin:0 });
  });
}

// ---------- 结尾页 ----------
// opts: { title, body, srcHeader, srcLines:[..] }
function closingSlide(pres, C, opts){
  const s = pres.addSlide(); bgDark(s, C);
  s.addShape(pres.shapes.RECTANGLE, { x:0, y:0, w:0.28, h:H, fill:{color:C.accentA} });
  s.addShape(pres.shapes.RECTANGLE, { x:0.28, y:0, w:0.16, h:H, fill:{color:C.accentB} });
  s.addText(opts.title, { x:1.1, y:2.4, w:11, h:0.9, fontFace:FONT, fontSize:38, color:C.white, bold:true, margin:0 });
  if (opts.body) s.addText(opts.body, { x:1.12, y:3.5, w:10.6, h:1.2, fontFace:FONT, fontSize:15, color:C.onDarkMuted, lineSpacingMultiple:1.35, margin:0 });
  s.addShape(pres.shapes.LINE, { x:1.12, y:5.1, w:4.0, h:0, line:{color:"3C5680", width:1.5} });
  const src = [];
  if (opts.srcHeader) src.push({ text:opts.srcHeader, options:{ bold:true, color:C.accentA, breakLine:true, paraSpaceAfter:4 } });
  (opts.srcLines||[]).forEach((t,i)=> src.push({ text:t, options:{ color: i===0?C.onDarkFaint:"7A8FB0", breakLine:true } }));
  if (src.length) s.addText(src, { x:1.12, y:5.35, w:11, h:1.3, fontFace:FONT, fontSize:12, lineSpacingMultiple:1.25, margin:0 });
  return s;
}

// ---------- 数据表格（表头着色 + 斑马纹）----------
// opts: { x, y, w, headers:[], rows:[[]], colW:[], fontSize, headerFontSize, accent, zebra, colAlign:[], rowH, align }
// 行数多时把 fontSize 调到 9.5–10。colW 之和需等于 w。
function dataTable(pres, s, C, opts){
  const x = (opts.x != null ? opts.x : 0.65), y = (opts.y != null ? opts.y : 1.9), w = (opts.w != null ? opts.w : 12.0);
  const fs = opts.fontSize || 11, hfs = opts.headerFontSize || fs;
  const accent = opts.accent || C.navy, zebra = opts.zebra || "F2F2EF";
  const cellAlign = (ci) => (opts.colAlign && opts.colAlign[ci]) || opts.align || "left";
  const headerRow = opts.headers.map((h, ci) => {
    return { text:String(h), options:{ fill:{color:accent}, color:C.white, bold:true, fontSize:hfs, valign:"middle", align:cellAlign(ci) } };
  });
  const bodyRows = opts.rows.map((r, ri) => {
    return r.map((cell, ci) => {
      return { text:(cell==null?"":String(cell)), options:{ fill:{color: ri%2 ? zebra : C.card}, color:C.body, fontSize:fs, valign:"middle", align:cellAlign(ci) } };
    });
  });
  s.addTable([headerRow, ...bodyRows], {
    x, y, w, colW:opts.colW, border:{type:"solid", pt:0.5, color:C.line},
    fontFace:FONT, valign:"middle", margin:opts.margin!=null?opts.margin:4,
    rowH:opts.rowH, autoPage:false });
}

module.exports = { W, H, FONT, palette, listStyles, shadow, sh, bgPaper, bgDark, pageNum, footer,
  header, headerDark, chip, titleSlide, cardGrid, statCallouts, processFlow, takeawayRows, dataTable, closingSlide };
