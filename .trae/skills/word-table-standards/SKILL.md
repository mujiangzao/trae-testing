---
name: word-table-standards
description: 制作或修改 Word 文件（.doc/.docx）时的排版与表格规范，尤其适用于含表格的文件。当用户要求生成、编辑、排版 Word 文档或表格时使用。规范要点：对外文件使用简体中文、A4 直式、上下左右边距 2.54 厘米、行距 1.5 倍、中文用新细明体、英文用 Times New Roman、表头加粗、框线一律黑色、单元格不填色彩、表头水平居中、内文靠左、垂直方向一律居中、总页数超过 1 页时右下角标注阿拉伯数字页码。
---

# Word 表格制作规范

## 触发条件
凡需要制作或修改对外 Word 文件（`.doc` / `.docx`），尤其是包含表格的文件时，必须套用以下全部规范。

## 语言规范
- 所有对外文件一律使用**简体中文**。

## 页面设置
- 纸张：A4 **直式**（纵向）。
- 页边距：上、下、左、右均为 **2.54 厘米**。
- 行距：**1.5 倍**行距。

## 字体规范
- 中文字体：**新细明体**（PMingLiU）。
- 英文字体：**Times New Roman**。
- 中英文混排时，中文套用新细明体，英文套用 Times New Roman。

## 表格规范
1. 表格中的**标题文字（表头）一律加粗**。
2. 表格框线一律采用**黑色**。
3. 单元格内**不要填上任何色彩**（无底纹、无填充色）。
4. 对齐方式：
   - 标题（表头）文字：水平**居中**。
   - 一般内文：水平**靠左**。
   - 垂直方向：所有单元格一律**居中**。

## 页码规范
- 若文件总页数**超过 1 页**，须在页面**右下角**标注页码，使用**阿拉伯数字**。
- 总页数仅 1 页时，不需要页码。

## 实现提示（python-docx）
生成文件时按下述方式落实规范：

- 页面：`section.page_width = Cm(21)`、`section.page_height = Cm(29.7)`（A4 纵向）；上下左右边距均设 `Cm(2.54)`。
- 行距：`paragraph.paragraph_format.line_spacing = 1.5`。
- 字体：`run.font.name = "Times New Roman"`，并另行设置东亚字体：
  ```python
  from docx.oxml.ns import qn
  run._element.rPr.rFonts.set(qn("w:eastAsia"), "新细明体")
  ```
- 表头加粗：表头单元格内所有 `run` 设 `font.bold = True`。
- 黑色框线：在 `w:tblBorders`（或单元格 `w:tcBorders`）中将各边框的 `w:color` 设为 `000000`；不要设置任何 `w:shd` 底纹填充。
- 水平对齐：表头用 `WD_ALIGN_PARAGRAPH.CENTER`，内文用 `WD_ALIGN_PARAGRAPH.LEFT`；垂直对齐：`cell.vertical_alignment = WD_CELL_VERTICAL_ALIGNMENT.CENTER`。
- 页码：预估总页数超过 1 页时，在页脚插入右对齐的 `PAGE` 域（阿拉伯数字）：
  ```python
  from docx.oxml import OxmlElement
  from docx.enum.text import WD_ALIGN_PARAGRAPH

  footer = section.footer
  p = footer.paragraphs[0]
  p.alignment = WD_ALIGN_PARAGRAPH.RIGHT
  run = p.add_run()
  fld_begin = OxmlElement("w:fldChar"); fld_begin.set(qn("w:fldCharType"), "begin")
  instr = OxmlElement("w:instrText"); instr.set(qn("xml:space"), "preserve"); instr.text = "PAGE"
  fld_end = OxmlElement("w:fldChar"); fld_end.set(qn("w:fldCharType"), "end")
  run._r.append(fld_begin); run._r.append(instr); run._r.append(fld_end)
  ```

## 产出前检查清单
逐项确认后再交付文件：
- [ ] 内容为简体中文
- [ ] A4 直式、四边距 2.54 厘米、1.5 倍行距
- [ ] 中文新细明体、英文 Times New Roman
- [ ] 表头加粗且水平居中，一般内文靠左，垂直方向全部居中
- [ ] 框线一律黑色，单元格无任何底色
- [ ] 总页数超过 1 页时，右下角有阿拉伯数字页码
