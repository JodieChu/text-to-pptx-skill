# text-to-pptx · 文字转专业中文 PPT 技能

一个 Claude Code 技能（Skill）：把一段文字（行研报告、会议纪要、项目分析、要点清单）转成**版式专业、配色得体、中文渲染干净**的 PPT，并内置**逐页视觉 QA**。

## 它能做什么

- **一行换风格**：内置 7 套精心调过的配色 / 质感，用一句话触发——
  `executive`（商务汇报）、`tech`（科技）、`electric`（高科技/未来/AI）、`minimal`（极简，无阴影发丝线）、`premium`（高智感/编辑感）、`mono`（高级灰）、`scholar`（学术）。
- **可复用版式积木**（`assets/pptx_design.js`）：封面 / 结尾、页眉页脚、2×2 编号卡片、大数字 callout、四步流程条、编号要点行、**数据表格 `dataTable`**（表头着色 + 斑马纹 + 数字右对齐）、小标签 chip。
- **自动配图**（`assets/image_search.js`，借鉴 [ppt-master](https://github.com/hugohe3/ppt-master) 的 image_search）：按关键词抓免费图片并排版——封面背景图 `titleSlide({bgImage})`、半图大图 `heroSplit`、图卡行 `imageRow`，抓不到自动降级占位框，CC 图自动署名。默认图源 Wikimedia（免 key 直连），可选 Pexels/Pixabay（设 `PEXELS_KEY`/`PIXABAY_KEY`）/Openverse。
- **零试错视觉 QA**（`scripts/qa_render.sh`）：把 pptx 转成逐页 JPG 供肉眼/子 agent 检查溢出与碰撞。

## 文件

| 文件 | 作用 |
|------|------|
| `SKILL.md` | 技能说明书：触发场景、5 步标准流程、7 套风格与触发词、渲染环境 |
| `assets/pptx_design.js` | 可复用版式库（调色板 + 所有版式函数，含 `dataTable` 表格与配图函数） |
| `assets/image_search.js` | 按关键词抓免费图片（Wikimedia/Pexels/Pixabay/Openverse/Picsum），仅用 Node 内置 + curl |
| `scripts/qa_render.sh` | pptx → 逐页图，做视觉 QA |

## 安装（Claude Code）

```bash
git clone <this-repo> ~/.claude/skills/text-to-pptx
# 或 clone 到别处后软链：ln -s /path/to/repo ~/.claude/skills/text-to-pptx
```
之后在 Claude Code 里说「做一份 PPT / 把 X 转成幻灯片 / 用科技风做个汇报片」即可触发。

## 依赖

- Node + `pptxgenjs`（在每个 PPT 工程目录本地 `npm install pptxgenjs`）。
- 视觉 QA 需要 LibreOffice（`qa_render.sh` 会按需从镜像下载 dmg、挂载使用、不做完整安装）+ poppler 的 `pdftoppm`（`brew install poppler`）。
- 中文字体默认 `PingFang SC`（macOS 原生；Windows PowerPoint 会替换为雅黑，亦正常）。

## 用法速览

```js
const pptxgen = require("pptxgenjs");
const D = require("./pptx_design");
const pres = new pptxgen(); pres.layout = "LAYOUT_WIDE";
const C = D.palette("tech");                 // 选风格
D.titleSlide(pres, C, {...});                // 封面
let s = pres.addSlide(); D.bgPaper(s, C);
D.header(pres, s, C, "KICKER", "标题", C.navy);
D.cardGrid(pres, s, C, [...]);               // 卡片
D.dataTable(pres, s, C, {headers, rows, colW});  // 表格
D.footer(pres, s, C, "页脚"); D.pageNum(s, C, 2);
await pres.writeFile({ fileName: "x.pptx" });
```

完整 API 与设计纪律见 `SKILL.md` 与 `assets/pptx_design.js` 顶部注释。

---
🤖 由 Claude Code 生成与维护
