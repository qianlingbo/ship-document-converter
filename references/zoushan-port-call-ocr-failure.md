# PORT CALL LIST 扫描件处理备忘

## 本次教训（ZHOUSHAN 船舶 PORT CALL LIST，2026-06-06）

### 文件特征
- 文件：`6.PORT CALL LIST(1).pdf`
- 来源：NAPS2 扫描仪（PDFsharp 1.50.4000 产生）
- 页数：1页 A4
- 关键问题：**无文字层**，纯图像 PDF

### 处理失败路径（不要重复）
| 方案 | 失败原因 |
|------|----------|
| `pdfplumber` / `pymupdf` 直接读 text | PDF 无文字层，返回空 |
| `pdftotext` | 同样无文字 |
| `marker_single` (venv) | 超时（~180s+）|
| `tesseract` + `chi_sim` | **训练数据损坏**：`/opt/homebrew/share/tessdata/chi_sim.traineddata` 只有 8KB，正常应 40MB+ |
| `tesseract` + `eng` | 能运行但输出全乱码符号 |

### 成功路径
```bash
# Step 1: 用 pypdfium2 渲染 PDF 为图像（venv 中有）
cd ~/.hermes/hermes-agent
./venv/bin/python3 -c "
import pypdfium2 as pdfium
pdf = pdfium.PdfDocument('port_call.pdf')
page = pdf[0]
pil_img = page.render(scale=2).to_pil()
pil_img.save('port_call_render.png')
"

# Step 2: 图像已生成，但 tesseract 中文 OCR 损坏
# 替代方案待探索：
# - 重新下载 chi_sim.traineddata
# - 使用在线 OCR API（如 OCR.space，需 API key）
# - 让用户手动复制/截图
```

### chi_sim.traineddata 修复
```bash
# 检查大小（正常 ~40MB）
ls -lh /opt/homebrew/share/tessdata/chi_sim.traineddata
# 8KB = 损坏，需要重新下载
# 来源：https://github.com/tesseract-ocr/tessdata/raw/main/chi_sim.traineddata
```

### 经验总结
- **先检查是否有文字层**：用 `pdfplumber` 或 `pymupdf` 读一页，有 text = 正常 PDF，无 text = 扫描件
- **扫描件用 pypdfium2 渲染**：`page.render(scale=2).to_pil()` 比 `convert` (ImageMagick) 更可靠
- **中文 OCR 是难点**：Tesseract 中文数据在 macOS Homebrew 上容易损坏，需定期验证
- **扫描件优先问用户要文字版**：避免浪费精力在 OCR 上

### 下次遇到扫描件 PODD FILE
1. 用 pypdfium2 渲染确认可读
2. 检查 Tesseract chi_sim 状态：`ls -lh $(brew --prefix)/share/tessdata/chi_sim.traineddata`
3. 若损坏，尝试修复或换用其他 OCR 方案
4. 若无法 OCR，在 skill 中记录，并向用户说明需要文字版或手动提供数据
