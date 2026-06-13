---
name: text-to-pptx
description: 把文字内容（研究报告/纪要/总结/要点）做成专业的中文 PPT。当 Jodie 说「做一份 PPT」「把 X 转成 PPT/幻灯片」「根据这篇分析生成 PPT」「做个汇报片」时使用。内置已实战验证的版式库 + 零试错的视觉 QA 渲染链路（本机无 PowerPoint/LibreOffice，已固化用缓存 dmg 出图）。
---

# text-to-pptx · 文字转专业中文 PPT

把一段文字（行研报告、会议纪要、项目分析、要点清单）转成一份**版式专业、配色得体、中文渲染干净**的 PPT，并**必做逐页视觉 QA** 再交付。

> 本 skill 把「耀途×光合」对比 PPT 的成功做法固化下来：可复用版式库（`assets/pptx_design.js`）+ 零试错的 QA 出图脚本（`scripts/qa_render.sh`）。下次直接填内容、跑 QA，不再重写版式、不再为渲染工具反复试错。

---

## 一、何时用本 skill

- Jodie 说「做 PPT / 幻灯片 / 汇报片 / 上会材料的片子」，或「把这篇 X 转成 PPT」。
- 输入通常是 Vault 里某篇 md（研究/纪要/分析）或 Jodie 给的一段文字/要点。
- 如果只是要 .pptx 的**读取/解析**或基于现成模板编辑，用官方 `document-skills:pptx`；本 skill 专注「从文字**新建**一份设计良好的中文 deck」。

---

## 二、标准流程（5 步，照做即可）

### 步骤 1 · 读源、定大纲
- 读全源文（md 用 Read；PDF/DOCX/PPTX 先 `~/scripts/file2md.py` 转 md）。
- 把内容收敛成 **10–14 页**的大纲：封面 → 背景/为什么 → 几个核心论点（每点 1 页）→ 数据页 → 结论/要点 → 结尾来源。
- 一页只讲一件事；每页正文精炼（中文每行 ~15–18 字，卡片正文别超 4 行）。

### 步骤 2 · 准备工程目录
```bash
mkdir -p ~/ai_projects/ppt-<主题> && cd ~/ai_projects/ppt-<主题>
npm init -y >/dev/null 2>&1
npm install pptxgenjs >/dev/null 2>&1
cp ~/.claude/skills/text-to-pptx/assets/pptx_design.js .
cp ~/.claude/skills/text-to-pptx/assets/image_search.js .   # 需要自动配图时
```
（项目一律放 `~/ai_projects/`，见全局 CLAUDE.md。）

### 步骤 3 · 写 build.js，用版式库拼页
```js
const pptxgen = require("pptxgenjs");
const D = require("./pptx_design");
const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE";
const C = D.palette("executive");   // 7 套风格，见下「三、风格与触发」

// 封面
D.titleSlide(pres, C, {
  kicker:"一手团队访谈对比研究",
  titleParts:[{text:"耀途资本",options:{color:C.accentA,bold:true}},
              {text:"  ×  ",options:{color:"6E83A6"}},
              {text:"光合创投",options:{color:"7FD4C4",bold:true}}],
  subtitle:"两家硬科技早期基金的打法对比　与对复旦科创的启示",
  descLines:["基于两份一手团队访谈纪要…","把打法掰开揉碎，提炼可借鉴经验"],
  meta:"研究日期　2026-06-10",
});

// 内容页（浅底）
let s = pres.addSlide(); D.bgPaper(s, C);
D.header(pres, s, C, "STRATEGY · 投资思路", "从撒胡椒面到敢重仓 + 复利", C.navy);
D.cardGrid(pres, s, C, [
  {no:"01", title:"…", body:"…"}, {no:"02", title:"…", body:"…"},
  {no:"03", title:"…", body:"…"}, {no:"04", title:"…", body:"…"},
]);
D.footer(pres, s, C, "页脚一句话"); D.pageNum(s, C, 3);

// 深色数据页 / 结论页用 D.bgDark + D.headerDark + D.statCallouts / D.takeawayRows
// 流程页用 D.processFlow；结尾用 D.closingSlide
await pres.writeFile({ fileName:"主题.pptx" });
```
版式库提供的积木（详见 `assets/pptx_design.js` 顶部注释与每个函数签名）：
- `titleSlide` 封面 · `closingSlide` 结尾来源页
- `header`/`headerDark` 页眉 · `footer`+`pageNum` 页脚页码
- `cardGrid` 2×2 编号卡片 · `statCallouts` 4 个大数字 · `processFlow` 4 步流程 · `takeawayRows` 编号要点行 · `dataTable` 数据表格 · `chip` 小标签
- **配图**（借鉴 ppt-master，配合 `image_search.js`）：`titleSlide({bgImage})` 封面背景图 · `heroSplit` 半图大图页 · `imageRow` 图卡行 · `imageBox` 任意位置放图（含占位兜底）· `imageCredits` 来源合规标注
- 7 套风格调色板（见「三、风格与触发」），`accentA / accentB` 两个强调色天然适合「两个对象对比」。

**自动配图用法**（按关键词抓免费图片，失败自动降级占位框）：
```js
const IS = require("./image_search");
const img = await IS.searchImage("人工智能 芯片", { orient:"landscape" });  // 返回 {path,source,license,...} 或 null
D.titleSlide(pres, C, { bgImage: img, /* ... */ });           // 封面铺底
D.heroSplit(pres, C, { img, side:"right", title:"…", bullets:[…] });
D.imageRow(pres, s, C, [{img, title:"…", caption:"…"}, …]);
```
- 默认图源 **Wikimedia Commons**（关键词、免 key、直连、CC/公共领域、带署名）——你的网络实测直连可达。
- 想要更漂亮的图库照片：设环境变量 `PEXELS_KEY`（或 `PIXABAY_KEY`），并传 `{source:"best"}`，会优先用 Pexels/Pixabay。
- Openverse 源需走代理（模块会自动用 `127.0.0.1:7890` 或 `TTP_PROXY`/`HTTPS_PROXY`）。
- 图片缓存在 `~/ai_projects/.tools-cache/img-cache/`，重复关键词不重抓。CC 图记得用 `imageCredits` 标注来源。

**模板填充（用你自己的 .pptx 模板，借鉴 ppt-master）** — `assets/template_fill.py`（python-pptx）
当 Jodie 要「按机构/公司标准模板出片」时用：把内容填进一份**真实 .pptx 模板**，外观（母版/Logo/配色/字体）完全沿用模板，只换内容。
- 模板里标占位符（标一次，长期复用）：
  - 文本：任意文本框/单元格写 `{{key}}` → 被 `content["key"]` 替换（**保留原格式**）。
  - 重复表格行：模板表格的「模板行」首格写 `{{#each LISTKEY}}`，其余格用列名 `{{name}}`/`{{amount}}` → 按 `content["LISTKEY"]`（dict 列表）逐条克隆填充。
  - 图片：给模板图片设「替代文字 alt text」为 `{{img:KEY}}` → 用 `content["img"][KEY]`（本地图片路径，可来自 image_search）替换，保持原位置尺寸。
- 用法：`/usr/bin/python3 ~/.claude/skills/text-to-pptx/assets/template_fill.py 模板.pptx 内容.json -o 输出.pptx`
- 校验：填充后 grep `{{` 应无残留；表格行数 = 1 表头 + N 数据。再用 `qa_render.sh` 出图核对版式未被破坏。
- 何时用「模板填充」vs「7 套风格代码生成」：有指定模板/品牌要求 → 模板填充；从零做、要好看排版 → 代码生成（`pptx_design.js`）。两者可混用（代码生成的图当 `{{img}}` 填进模板）。
- 示例模板见 `examples/sample_template.pptx` + `examples/content.json`。

设计纪律（避免 AI 味）：标题 27pt 加粗、正文 12–13pt；**不在标题下加装饰线**；卡片留白足；一种强调色主导。中文字体用 `PingFang SC`（库已内置）。

### 步骤 4 · 【必做】视觉 QA — 零试错出图
```bash
bash ~/.claude/skills/text-to-pptx/scripts/qa_render.sh "主题.pptx"
```
脚本自动：挂载缓存的 LibreOffice dmg → 转 PDF → `pdftoppm` 出 `qa/slide-01.jpg…` → 卸载。
然后**用 Read 逐页看图**（可派 subagent 看，但务必把结论要回来），重点查：
- 文字溢出卡片/页面边缘、正文压到下一个元素
- **页脚/页码与上方内容碰撞**（最常见，这次第6页表格就踩过）
- 卡片不等高、间距忽大忽小、低对比文字
- 中文断行出现孤字
发现问题改 build.js → 重新生成 → 重跑 QA，直到一遍扫下来无新问题。

> ⚠️ LibreOffice 没 PingFang，渲染出的中文正文是「楷体」替代字——**那是渲染假象**，PowerPoint/Keynote 里是正常黑体（封面可用 `qlmanage -t -s 1600 x.pptx -o /tmp/ql` 原生渲染验证一页）。QA 只判版式，别管字体样式。

### 步骤 5 · 交付
- 成片存到合适位置（研究类 deck → Vault `04-Research/`，与源文同目录；项目类 → `03-Projects/`；遵循 Vault CLAUDE.md 的研究产出归置规则）。
- `build.js` 和 `qa/` 留在 `~/ai_projects/ppt-<主题>/` 备查。
- 按「与 Jodie 协作铁律」：5 行内告知完成状态、文件路径；提示可按需调整页数/配色。

---

## 三、风格与触发（7 套）

每套都跟着同一套版式引擎走，只换配色与质感。`flat` 标记的两套用「发丝线 + 无投影」做真极简观感。

| 风格名 | 中文定位 | 配色 | Jodie 这样说就用它（触发词） |
|--------|---------|------|------------------------------|
| `executive` | 投资/金融/严肃汇报（默认） | 海军蓝 + 金 + 绿 | 商务 / 金融 / 投资 / 严肃 / 汇报 |
| `tech` | 科技/产品 | 深石板 + 青 + 珊瑚 | 科技 / 产品 / 互联网 |
| `electric` | 高科技/未来/AI | 靛蓝 + 电光紫 + 青 | 未来 / 高科技 / 炫 / AI / 酷 / 赛博 |
| `minimal` ⟂flat | 极简/留白 | 纯白 + 发丝线 + 石墨蓝 | 简约 / 极简 / 留白 / 干净 / 简洁 |
| `premium` | 高智感/知性/编辑感 | 深墨 + 香槟金 + 米白 | 高智 / 高级感 / 知性 / 精致 / 质感 / 高端 |
| `mono` ⟂flat | 高级灰/单色 | 炭灰阶 + 一抹暖铜 | 高级灰 / 单色 / 黑白 / 莫兰迪 / 性冷淡 |
| `scholar` | 学术/研究/政策 | 墨绿 + 赭 + 苔 | 学术 / 研究 / 政策 / 稳重 |

**触发方式**：Jodie 在请求里带上风格词即可，例如：
- 「用**科技风**做这份 PPT」→ `tech`；「**未来感/AI 风**」→ `electric`
- 「**简约/极简**点」→ `minimal`；「要**高智感/高级感**」→ `premium`；「**高级灰/性冷淡**风」→ `mono`
- 不指定 → 默认 `executive`（投资/汇报最稳）。可视主题自动选：投资项目→executive、AI/科技→electric/tech、学术研究→scholar。

代码里就一行：`const C = D.palette("电脑里的风格名");`。运行 `node -e "require('<skill>/assets/pptx_design.js').listStyles()"` 可随时列出全部风格与触发词。

---

## 四、本机渲染环境（重要，已固化、勿再试错）

本机（Apple Silicon Mac）**无 PowerPoint、无 LibreOffice、Keynote 在 launchd 下 AppleScript 导出会超时**，`brew` 装 LibreOffice 的 documentfoundation 下载对本机网络一直 404。**唯一可靠路径**已写进 `qa_render.sh`：
- LibreOffice 25.8.7 的 dmg **缓存在** `~/ai_projects/.tools-cache/LibreOffice-25.8.7.dmg`（清华 TUNA 镜像，286MB，已下好；脚本检测缺失会自动重下）。
- 用 `hdiutil attach` 按需挂载、跑完 `hdiutil detach`，**不做完整安装**（省 ~700MB 磁盘）。
- `pdftoppm` 来自 `brew install poppler`（已装）。

详见记忆 [[reference-pptx-qa-rendering]]。

---

## 五、依赖自检
- `node -e "require('pptxgenjs')"` —— 在工程目录本地装（全局 nvm 没有）。
- `command -v pdftoppm` —— 没有则 `brew install poppler`。
- `ls ~/ai_projects/.tools-cache/LibreOffice-*.dmg` —— 没有则 qa_render.sh 会自动下。
