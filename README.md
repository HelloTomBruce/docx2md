# docx2md

样式驱动的 DOCX → Markdown 转换器（Rust）。

与启发式纯文本方案（先抽文本再猜标题）不同，本库直接利用 DOCX 自带的结构化元数据：

- **标题**：`w:pStyle` → `styles.xml` 中 `heading N` 的真实映射 + `w:outlineLvl` 直接大纲级别兜底，自动跳过 TOC 目录段落
- **列表（ul/ol）**：完整解析 `numbering.xml` 三层引用（段落 `numPr(numId, ilvl)` → `w:num` → `w:abstractNum` 的 `lvl`），准确还原嵌套层级与编号
- **自动编号**：计数器状态机渲染 `lvlText` 中的 `%1-%9`，支持多级编号（`1.1.2`）、`(1)`、`①`、罗马数字、中文数字等 `numFmt`
- **表格**：`w:tbl` → Markdown 表格，支持合并单元格（`gridSpan` 占列 / `vMerge` 纵合）、单元格多段落（`<br>`）、单元格内图片与链接
- **图片**：按文档顺序内联提取（DrawingML `a:blip r:embed` / 旧式 VML `v:imagedata r:id`），
  经 `document.xml.rels` 映射到 `word/media/*`，导出文件并生成 `![alt](prefix/file.png)`；
  alt 取 `wp:docPr` 的 descr/name；同一文件多次引用自动去重；`mc:AlternateContent` 只取 Choice 分支避免重复
- **超链接**：`w:hyperlink`（外部 URL / 内部书签锚点）、`w:fldSimple` 与复杂域代码（`fldChar`/`instrText`）三种形式 → `[text](url)`
- **行内格式**：加粗 `**`、斜体 `*`、删除线 `~~`，相邻同格式 run 自动合并
- **脚注/尾注**：`w:footnoteReference`/`w:endnoteReference` → `[^N]` / `[^eN]`，定义按首次引用顺序追加文末；分隔符类条目自动跳过
- **数学公式**：OMML → LaTeX（`m:f` 分数、上下标、根式、n-ary 运算符、定界符、函数、矩阵），输出 `$...$`
- **文档元数据**：`docProps/core.xml` → YAML front matter（`--front-matter`，含 title/author/created/modified）
- **修订**：track changes 默认接受插入（`w:ins`）、拒绝删除（`w:del`）
- **边界情况**：段落级 numPr 优先、样式级 numPr 兜底、`numId=0` 表示取消编号

## 安装

```bash
# 作为 CLI 全局安装
cargo install --path /Users/zhangbei/code/docx2md

# 或作为库依赖（Cargo.toml）
docx2md = { path = "/Users/zhangbei/code/docx2md" }
# 或 git 依赖
docx2md = { git = "ssh://git@your-git-server/docx2md.git" }
```

## CLI 用法

```bash
# 输出到 stdout
 docx2md 需求文档.docx

# 输出到文件（图片默认导出到 output.media/，md 中自动引用）
docx2md 需求文档.docx -o output.md

# 自定义图片目录 / 链接前缀
docx2md 需求文档.docx -o output.md --images-dir assets/img --image-prefix "assets/img/"

# 不导出图片（md 中也不生成图片链接）
docx2md 需求文档.docx -o output.md --no-images

# 输出 YAML front matter（文档属性元数据）
docx2md 需求文档.docx --front-matter -o output.md

# 保留 TOC 目录（默认跳过）
docx2md 需求文档.docx --keep-toc -o output.md
```

## 库用法

```rust
use docx2md::{convert_file, ConvertOptions};

let options = ConvertOptions {
    image_link_prefix: "output.media/".into(),
    ..ConvertOptions::new()
};
let result = convert_file("需求文档.docx", &options)?;
std::fs::write("output.md", &result.markdown)?;
result.write_images("output.media")?;  // 导出图片文件
```

## 已知限制

- 页眉/页脚中的图片不提取（它们有自己的 rels 文件）
- 纯文本 URL 不自动转为链接（GFM 渲染器通常会自动识别）
- 表格内的列表标记不还原（单元格多段落以 `<br>` 连接）
- `lvlOverride`（局部覆盖某级编号格式）未处理
- 样式级行内格式（如 Hyperlink 字符样式的下划线）不渲染

## 开发

```bash
cargo test          # 单元测试
cargo build --release
```
