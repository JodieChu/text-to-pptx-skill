#!/bin/bash
# qa_render.sh — 把 .pptx 转成逐页 JPG，供视觉 QA。
# 用法:  qa_render.sh <文件.pptx> [输出目录=./qa] [DPI=110]
#
# 渲染链路（本机已固化，零试错）：
#   缓存的 LibreOffice dmg → 挂载 → soffice 转 PDF → pdftoppm 转逐页 JPG → 卸载 dmg
# 依赖：① ~/ai_projects/.tools-cache/LibreOffice-*.dmg（缺失则自动从清华 TUNA 镜像重下）
#       ② poppler 的 pdftoppm（缺失则 brew install poppler）
#
# 注意：LibreOffice 没 PingFang，中文正文会渲染成「楷体」替代字 —— 那是渲染假象，
#       PowerPoint/Keynote 打开是正常黑体。QA 只看版式/溢出/碰撞，别被字体样式误导。
set -uo pipefail

PPTX="${1:?用法: qa_render.sh <文件.pptx> [输出目录] [DPI]}"
OUTDIR="${2:-./qa}"
DPI="${3:-110}"
CACHE_DIR="$HOME/ai_projects/.tools-cache"
LO_VER="25.8.7"
DMG="$CACHE_DIR/LibreOffice-$LO_VER.dmg"
TUNA="https://mirrors.tuna.tsinghua.edu.cn/libreoffice/libreoffice/stable/$LO_VER/mac/aarch64/LibreOffice_${LO_VER}_MacOS_aarch64.dmg"

[ -f "$PPTX" ] || { echo "✗ 找不到文件: $PPTX"; exit 1; }
PPTX_ABS="$(cd "$(dirname "$PPTX")" && pwd)/$(basename "$PPTX")"

# 1) 确保缓存 dmg 存在
if [ ! -f "$DMG" ] || [ "$(stat -f%z "$DMG" 2>/dev/null || echo 0)" -lt 100000000 ]; then
  echo "→ 缓存 dmg 缺失，从清华 TUNA 镜像下载（约 286MB，一次性）…"
  mkdir -p "$CACHE_DIR"
  ( unset HTTP_PROXY HTTPS_PROXY http_proxy https_proxy ALL_PROXY all_proxy
    curl -L -s -o "$DMG" "$TUNA" -w "  下载完成 HTTP %{http_code}  %{size_download} bytes\n" )
fi

# 2) 确保 pdftoppm 存在
if ! command -v pdftoppm >/dev/null 2>&1; then
  echo "→ 安装 poppler（提供 pdftoppm）…"; brew install poppler >/dev/null 2>&1
fi
PDFTOPPM="$(command -v pdftoppm)"

# 3) 挂载 dmg（按需），找 soffice
MOUNT_POINT="/Volumes/LibreOffice"
DETACH=0
if [ ! -x "$MOUNT_POINT/LibreOffice.app/Contents/MacOS/soffice" ]; then
  echo "→ 挂载 LibreOffice dmg…"
  hdiutil attach "$DMG" -nobrowse -quiet
  DETACH=1
  for i in $(seq 1 15); do
    [ -x "$MOUNT_POINT/LibreOffice.app/Contents/MacOS/soffice" ] && break; sleep 1
  done
fi
SOFFICE="$MOUNT_POINT/LibreOffice.app/Contents/MacOS/soffice"
[ -x "$SOFFICE" ] || { echo "✗ 挂载后仍找不到 soffice"; exit 1; }

cleanup(){ [ "$DETACH" = "1" ] && hdiutil detach "$MOUNT_POINT" -quiet 2>/dev/null || true; }
trap cleanup EXIT

# 4) 转 PDF（临时 profile 目录避免权限问题）
WORK="$(dirname "$PPTX_ABS")"
echo "→ soffice 转 PDF…"
"$SOFFICE" --headless -env:UserInstallation=file:///tmp/lo_qa_profile \
  --convert-to pdf --outdir "$WORK" "$PPTX_ABS" >/dev/null 2>&1
PDF="${PPTX_ABS%.pptx}.pdf"
[ -f "$PDF" ] || { echo "✗ PDF 生成失败"; exit 1; }

# 5) 逐页转 JPG
mkdir -p "$OUTDIR"
"$PDFTOPPM" -jpeg -r "$DPI" "$PDF" "$OUTDIR/slide" >/dev/null 2>&1
N=$(ls "$OUTDIR"/slide-*.jpg 2>/dev/null | wc -l | tr -d ' ')
rm -f "$PDF"   # 清理中间 PDF
echo "✓ 已生成 $N 页图到 $OUTDIR/  (slide-01.jpg …)"
echo "  下一步：用 Read 逐页看图做视觉 QA（重点查溢出/碰撞/页脚压内容）。"
