# docx2md

样式驱动的 DOCX → Markdown 转换器（Rust）。

与启发式纯文本方案（先抽文本再猜标题）不同，本库直接利用 DOCX 自带的结构化元数据：

- **标题**：`w:pStyle` → `styles.xml` 中 `heading N` 的真实映射，自动跳过 TOC 目录段落
- **列表（ul/ol）**：完整解析 `numbering.xml` 三层引用（段落 `numPr(numId, ilvl)` → `w:num` → `w:abstractNum` 的 `lvl`），准确还原嵌套层级与编号
- **自动编号**：计数器状态机渲染 `lvlText` 中的 `%1-%9`，支持多级编号（`1.1.2`）、`(1)`、`①`、罗马数字、中文数字等 `numFmt`
- **表格**：`w:tbl` → Markdown 表格
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

# 输出到文件
docx2md 需求文档.docx -o output.md

# 保留 TOC 目录（默认跳过）
docx2md 需求文档.docx --keep-toc -o output.md
```

## 库用法

```rust
use docx2md::{convert_file, ConvertOptions};

let md = convert_file("需求文档.docx", &ConvertOptions::default())?;
println!("{md}");
```

## 已知限制

- 图片不导出（`word/media/*` 未处理）
- 表格单元格内的列表标记会丢失
- `lvlOverride`（局部覆盖某级编号格式）未处理
- 合并单元格会产生空列

## 开发

```bash
cargo test          # 单元测试
cargo build --release
```
