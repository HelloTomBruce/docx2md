//! 样式驱动的 DOCX → Markdown 转换器。
//!
//! 核心思路（与启发式纯文本方案的本质区别）：
//! - 标题层级来自 `w:pStyle` → `styles.xml` 中 `heading N` 的真实映射，而非猜测文本特征
//! - 列表（ul/ol）完整解析 `numbering.xml` 的三层间接引用：
//!   段落 `numPr(numId, ilvl)` → `w:num` → `w:abstractNum` 的 `lvl(numFmt, lvlText, start)`
//! - 编号来源三级解析：段落级 numPr 优先 → 样式级 numPr 兜底 → numId=0 表示取消编号
//! - 计数器状态机：按 numId 维护各级 ilvl 计数，推进当前层、重置更深层、上层补 start 值；
//!   `lvlText` 中的 `%1-%9` 按对应层的 numFmt 渲染替换（支持多级编号如 "1.1.2"、"(1)、"①"）

use anyhow::{bail, Context, Result};
use roxmltree::{Document, Node};
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

// ---------------------------------------------------------------------------
// 数据结构
// ---------------------------------------------------------------------------

#[derive(Debug, Clone)]
struct LvlDef {
    fmt: String,   // numFmt: bullet / decimal / lowerLetter / ...
    text: String,  // lvlText: "%1." / "(%1)" / "%1.%2.%3"
    start: u32,    // 起始计数
}

impl Default for LvlDef {
    fn default() -> Self {
        Self { fmt: "bullet".into(), text: String::new(), start: 1 }
    }
}

#[derive(Debug, Default, Clone)]
struct StyleInfo {
    num_id: Option<String>,
    ilvl: u32,
}

#[derive(Debug, Default)]
struct Numbering {
    num_to_abs: HashMap<String, String>,
    abstracts: HashMap<String, HashMap<u32, LvlDef>>,
}

#[derive(Debug, Default)]
struct Styles {
    /// styleId → 标题级别（1-9）
    headings: HashMap<String, u32>,
    /// styleId → 信息
    info: HashMap<String, StyleInfo>,
    /// TOC 目录样式（跳过）
    toc: Vec<String>,
}

// ---------------------------------------------------------------------------
// XML 辅助
// ---------------------------------------------------------------------------

/// 按本地名取属性（忽略命名空间前缀，docx 属性均为 w:xxx）
fn attr<'a>(node: &Node<'a, 'a>, local: &str) -> Option<&'a str> {
    node.attributes()
        .find(|a| a.name().rsplit(':').next() == Some(local) || a.name() == local)
        .map(|a| a.value())
}

/// 节点的本地标签名
fn tag<'a>(node: &Node<'a, 'a>) -> &'a str {
    node.tag_name().name()
}

/// 在直接子节点中按路径找（如 ["pPr", "numPr", "numId"]）
fn find_path<'a, 'b>(node: &'b Node<'a, 'a>, path: &[&str]) -> Option<Node<'a, 'a>> {
    let mut cur = *node;
    for name in path {
        cur = cur
            .children()
            .find(|c| c.is_element() && tag(c) == *name)?;
    }
    Some(cur)
}

// ---------------------------------------------------------------------------
// 解析 numbering.xml / styles.xml
// ---------------------------------------------------------------------------

fn parse_numbering(xml: &str) -> Numbering {
    let mut numbering = Numbering::default();
    let Ok(doc) = Document::parse(xml) else { return numbering };

    for num in doc.descendants().filter(|n| tag(n) == "num") {
        if let (Some(id), Some(abs)) = (
            attr(&num, "numId"),
            find_path(&num, &["abstractNumId"]).and_then(|n| attr(&n, "val")),
        ) {
            numbering.num_to_abs.insert(id.to_string(), abs.to_string());
        }
    }
    for abs in doc.descendants().filter(|n| tag(n) == "abstractNum") {
        let Some(abs_id) = attr(&abs, "abstractNumId") else { continue };
        let mut lvls = HashMap::new();
        for lvl in abs.children().filter(|n| tag(n) == "lvl") {
            let ilvl: u32 = attr(&lvl, "ilvl").and_then(|v| v.parse().ok()).unwrap_or(0);
            lvls.insert(ilvl, LvlDef {
                fmt: find_path(&lvl, &["numFmt"]).and_then(|n| attr(&n, "val")).unwrap_or("decimal").to_string(),
                text: find_path(&lvl, &["lvlText"]).and_then(|n| attr(&n, "val")).unwrap_or("%1.").to_string(),
                start: find_path(&lvl, &["start"]).and_then(|n| attr(&n, "val")).and_then(|v| v.parse().ok()).unwrap_or(1),
            });
        }
        numbering.abstracts.insert(abs_id.to_string(), lvls);
    }
    numbering
}

fn parse_styles(xml: &str) -> Styles {
    let mut styles = Styles::default();
    let Ok(doc) = Document::parse(xml) else { return styles };

    for style in doc.descendants().filter(|n| tag(n) == "style") {
        let Some(sid) = attr(&style, "styleId") else { continue };
        let name = find_path(&style, &["name"])
            .and_then(|n| attr(&n, "val"))
            .unwrap_or("")
            .to_lowercase();

        let mut info = StyleInfo::default();
        if let Some(numpr) = find_path(&style, &["pPr", "numPr"]) {
            info.num_id = find_path(&numpr, &["numId"]).and_then(|n| attr(&n, "val")).map(String::from);
            info.ilvl = find_path(&numpr, &["ilvl"]).and_then(|n| attr(&n, "val")).and_then(|v| v.parse().ok()).unwrap_or(0);
        }

        if let Some(level) = name.strip_prefix("heading ").and_then(|d| d.parse::<u32>().ok()) {
            styles.headings.insert(sid.to_string(), level);
        }
        if name.starts_with("toc") {
            styles.toc.push(sid.to_string());
        }
        styles.info.insert(sid.to_string(), info);
    }
    styles
}

// ---------------------------------------------------------------------------
// 计数器状态机
// ---------------------------------------------------------------------------

struct CounterState<'a> {
    numbering: &'a Numbering,
    /// numId → (ilvl → 当前值)
    counters: HashMap<String, HashMap<u32, u32>>,
}

impl<'a> CounterState<'a> {
    fn new(numbering: &'a Numbering) -> Self {
        Self { numbering, counters: HashMap::new() }
    }

    /// 推进 (numId, ilvl) 的计数器，返回 (是否 bullet, 渲染出的编号文本)
    fn resolve(&mut self, num_id: &str, ilvl: u32) -> (bool, String) {
        let Some(abs_id) = self.numbering.num_to_abs.get(num_id) else {
            return (true, String::new());
        };
        let empty = HashMap::new();
        let lvls = self.numbering.abstracts.get(abs_id).unwrap_or(&empty);
        let lvl = lvls.get(&ilvl).cloned().unwrap_or_default();

        let cnt = self.counters.entry(num_id.to_string()).or_default();
        // 推进本层
        *cnt.entry(ilvl).or_insert(lvl.start.saturating_sub(1)) += 1;
        // 重置更深层
        cnt.retain(|k, _| *k <= ilvl);
        // 未走过的上层补 start
        for k in 0..ilvl {
            cnt.entry(k).or_insert(lvls.get(&k).map(|l| l.start).unwrap_or(1));
        }

        if lvl.fmt == "bullet" {
            return (true, String::new());
        }
        // 替换 lvlText 中的 %1-%9
        let mut text = lvl.text.clone();
        for k in (1..=9u32).rev() {
            let placeholder = format!("%{k}");
            if text.contains(&placeholder) {
                let lv = lvls.get(&(k - 1)).cloned().unwrap_or_default();
                let val = cnt.get(&(k - 1)).copied().unwrap_or(lv.start);
                text = text.replace(&placeholder, &format_counter(&lv.fmt, val));
            }
        }
        (false, text)
    }
}

/// 按 numFmt 渲染计数器值
fn format_counter(fmt: &str, n: u32) -> String {
    match fmt {
        "decimal" | "decimalZero" => n.to_string(),
        "lowerLetter" => char::from_u32(u32::from(b'a') + n.saturating_sub(1))
            .filter(|_| (1..=26).contains(&n))
            .map(|c| c.to_string())
            .unwrap_or_else(|| n.to_string()),
        "upperLetter" => char::from_u32(u32::from(b'A') + n.saturating_sub(1))
            .filter(|_| (1..=26).contains(&n))
            .map(|c| c.to_string())
            .unwrap_or_else(|| n.to_string()),
        "lowerRoman" => to_roman(n).to_lowercase(),
        "upperRoman" => to_roman(n),
        "decimalEnclosedCircleChinese" | "decimalEnclosedCircle" => {
            const CIRCLED: [char; 20] = ['①','②','③','④','⑤','⑥','⑦','⑧','⑨','⑩',
                                         '⑪','⑫','⑬','⑭','⑮','⑯','⑰','⑱','⑲','⑳'];
            CIRCLED.get(n.saturating_sub(1) as usize)
                .filter(|_| (1..=20).contains(&n))
                .map(|c| c.to_string())
                .unwrap_or_else(|| n.to_string())
        }
        "chineseCounting" | "chineseCountingThousand" | "chineseLegalSimplified" => {
            const DIGITS: [char; 10] = ['零','一','二','三','四','五','六','七','八','九'];
            DIGITS.get(n as usize)
                .filter(|_| n <= 9)
                .map(|c| c.to_string())
                .unwrap_or_else(|| n.to_string())
        }
        _ => n.to_string(),
    }
}

fn to_roman(mut n: u32) -> String {
    let table = [(1000, "M"), (900, "CM"), (500, "D"), (400, "CD"), (100, "C"),
                 (90, "XC"), (50, "L"), (40, "XL"), (10, "X"), (9, "IX"),
                 (5, "V"), (4, "IV"), (1, "I")];
    let mut out = String::new();
    for (v, s) in table {
        while n >= v { out.push_str(s); n -= v; }
    }
    if out.is_empty() { "I".into() } else { out }
}

// ---------------------------------------------------------------------------
// 段落 / 表格提取
// ---------------------------------------------------------------------------

/// 段落纯文本：拼接 w:t，tab/br 视为空格
fn para_text(p: &Node) -> String {
    let mut parts = String::new();
    for node in p.descendants() {
        match tag(&node) {
            "t" => parts.push_str(&node.text().unwrap_or("")),
            "tab" | "br" => parts.push(' '),
            _ => {}
        }
    }
    parts.trim().to_string()
}

/// 解析段落的编号来源：段落级 numPr 优先，样式级 numPr 兜底。
/// 返回 (numId, ilvl, styleId)。numId 为 "0" 表示取消编号（视为 None）。
fn para_numpr<'a>(p: &Node<'a, 'a>, styles: &Styles) -> (Option<String>, u32, Option<String>) {
    let sid = find_path(p, &["pPr", "pStyle"]).and_then(|n| attr(&n, "val")).map(String::from);

    if let Some(numpr) = find_path(p, &["pPr", "numPr"]) {
        if let Some(num_id) = find_path(&numpr, &["numId"]).and_then(|n| attr(&n, "val")) {
            if num_id != "0" {
                let ilvl = find_path(&numpr, &["ilvl"]).and_then(|n| attr(&n, "val")).and_then(|v| v.parse().ok()).unwrap_or(0);
                return (Some(num_id.to_string()), ilvl, sid);
            }
            return (None, 0, sid);
        }
    }
    if let Some(id) = &sid {
        if let Some(info) = styles.info.get(id) {
            if let Some(num_id) = &info.num_id {
                if num_id != "0" {
                    return (Some(num_id.clone()), info.ilvl, sid);
                }
            }
        }
    }
    (None, 0, sid)
}

/// 表格 → Markdown 表格
fn table_md(tbl: &Node) -> String {
    let mut rows: Vec<Vec<String>> = Vec::new();
    for tr in tbl.children().filter(|n| tag(n) == "tr") {
        let mut cells = Vec::new();
        for tc in tr.children().filter(|n| tag(n) == "tc") {
            // 单元格取第一个段落的文本；合并单元格可能产生空列
            let text = tc.children()
                .find(|n| tag(n) == "p")
                .map(|p| para_text(&p).replace('|', "\\|"))
                .unwrap_or_default();
            cells.push(text);
        }
        rows.push(cells);
    }
    if rows.is_empty() {
        return String::new();
    }
    let width = rows.iter().map(Vec::len).max().unwrap_or(0);
    let mut out = String::new();
    for (i, row) in rows.iter().enumerate() {
        out.push_str("| ");
        let mut cells: Vec<&str> = row.iter().map(String::as_str).collect();
        cells.resize(width, "");
        out.push_str(&cells.join(" | "));
        out.push_str(" |\n");
        if i == 0 {
            out.push_str(&"|---".repeat(width));
            out.push_str("|\n");
        }
    }
    out.trim_end().to_string()
}

// ---------------------------------------------------------------------------
// 主转换入口
// ---------------------------------------------------------------------------

/// 转换选项
#[derive(Debug, Clone, Default)]
pub struct ConvertOptions {
    /// 是否保留 TOC 目录段落（默认跳过）
    pub keep_toc: bool,
}

/// 从 DOCX 字节流转换为 Markdown
pub fn convert_bytes(bytes: &[u8], options: &ConvertOptions) -> Result<String> {
    let cursor = std::io::Cursor::new(bytes);
    let mut archive = zip::ZipArchive::new(cursor).context("无法作为 zip 打开 DOCX")?;

    let mut read_entry = |name: &str| -> Result<String> {
        let mut f = archive.by_name(name).with_context(|| format!("DOCX 中缺少 {name}"))?;
        let mut buf = String::new();
        f.read_to_string(&mut buf)?;
        Ok(buf)
    };

    let document_xml = read_entry("word/document.xml")?;
    let numbering_xml = read_entry("word/numbering.xml").unwrap_or_default();
    let styles_xml = read_entry("word/styles.xml").unwrap_or_default();

    convert_xml(&document_xml, &styles_xml, &numbering_xml, options)
}

/// 从 DOCX 文件路径转换为 Markdown
pub fn convert_file(path: impl AsRef<Path>, options: &ConvertOptions) -> Result<String> {
    let path = path.as_ref();
    let mut f = File::open(path).with_context(|| format!("无法打开 {}", path.display()))?;
    let mut bytes = Vec::new();
    f.read_to_end(&mut bytes)?;
    convert_bytes(&bytes, options)
}

/// 从解析好的 XML 字符串转换（便于测试）
pub fn convert_xml(document_xml: &str, styles_xml: &str, numbering_xml: &str, options: &ConvertOptions) -> Result<String> {
    let doc = Document::parse(document_xml).context("document.xml 解析失败")?;
    let styles = parse_styles(styles_xml);
    let numbering = parse_numbering(numbering_xml);
    let mut counters = CounterState::new(&numbering);

    let Some(body) = doc.descendants().find(|n| tag(n) == "body") else {
        bail!("document.xml 中没有 body");
    };

    let mut md = String::new();
    for el in body.children().filter(|n| n.is_element()) {
        match tag(&el) {
            "tbl" => {
                let t = table_md(&el);
                if !t.is_empty() {
                    md.push_str(&t);
                    md.push_str("\n\n");
                }
            }
            "p" => {
                let text = para_text(&el);
                if text.is_empty() {
                    continue;
                }
                let (num_id, ilvl, sid) = para_numpr(&el, &styles);

                if !options.keep_toc {
                    if let Some(id) = &sid {
                        if styles.toc.contains(id) {
                            continue;
                        }
                    }
                }

                if let Some(id) = &sid {
                    if let Some(&level) = styles.headings.get(id) {
                        // 标题：可能带自动编号（多级号）
                        let prefix = match &num_id {
                            Some(nid) => {
                                let (is_bullet, numtext) = counters.resolve(nid, ilvl);
                                if !is_bullet && !numtext.is_empty() { format!("{numtext} ") } else { String::new() }
                            }
                            None => String::new(),
                        };
                        let hashes = "#".repeat(level.min(6) as usize);
                        md.push_str(&format!("{hashes} {prefix}{text}\n\n"));
                        continue;
                    }
                }

                match num_id {
                    Some(nid) => {
                        let (is_bullet, numtext) = counters.resolve(&nid, ilvl);
                        let indent = "  ".repeat(ilvl as usize);
                        let marker = if is_bullet { "- ".to_string() } else { format!("{numtext} ") };
                        md.push_str(&format!("{indent}{marker}{text}\n"));
                    }
                    None => md.push_str(&format!("{text}\n\n")),
                }
            }
            _ => {}
        }
    }
    Ok(md)
}

// ---------------------------------------------------------------------------
// 测试
// ---------------------------------------------------------------------------

#[cfg(test)]
mod tests {
    use super::*;

    const STYLES: &str = r#"<w:styles xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:style w:styleId="3"><w:name w:val="heading 1"/><w:pPr><w:numPr><w:ilvl w:val="0"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
        <w:style w:styleId="4"><w:name w:val="heading 2"/><w:pPr><w:numPr><w:ilvl w:val="1"/><w:numId w:val="1"/></w:numPr></w:pPr></w:style>
        <w:style w:styleId="19"><w:name w:val="toc 1"/></w:style>
    </w:styles>"#;

    const NUMBERING: &str = r#"<w:numbering xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">
        <w:abstractNum w:abstractNumId="0">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1"/></w:lvl>
            <w:lvl w:ilvl="1"><w:start w:val="1"/><w:numFmt w:val="decimal"/><w:lvlText w:val="%1.%2"/></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="1">
            <w:lvl w:ilvl="0"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
            <w:lvl w:ilvl="1"><w:numFmt w:val="bullet"/><w:lvlText w:val=""/></w:lvl>
        </w:abstractNum>
        <w:abstractNum w:abstractNumId="2">
            <w:lvl w:ilvl="0"><w:start w:val="1"/><w:numFmt w:val="decimalEnclosedCircleChinese"/><w:lvlText w:val="%1"/></w:lvl>
        </w:abstractNum>
        <w:num w:numId="1"><w:abstractNumId w:val="0"/></w:num>
        <w:num w:numId="2"><w:abstractNumId w:val="1"/></w:num>
        <w:num w:numId="3"><w:abstractNumId w:val="2"/></w:num>
    </w:numbering>"#;

    fn doc(body: &str) -> String {
        format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:body>{body}</w:body></w:document>"#)
    }

    fn p(text: &str) -> String {
        format!(r#"<w:p><w:r><w:t>{text}</w:t></w:r></w:p>"#)
    }

    fn p_num(text: &str, num_id: u32, ilvl: u32) -> String {
        format!(r#"<w:p><w:pPr><w:numPr><w:ilvl w:val="{ilvl}"/><w:numId w:val="{num_id}"/></w:numPr></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
    }

    fn p_style(text: &str, sid: u32) -> String {
        format!(r#"<w:p><w:pPr><w:pStyle w:val="{sid}"/></w:pPr><w:r><w:t>{text}</w:t></w:r></w:p>"#)
    }

    #[test]
    fn headings_get_multilevel_numbers() {
        let body = format!("{}{}{}{}", p_style("简介", 3), p_style("文档目的", 4), p_style("概述", 3), p_style("背景", 4));
        let md = convert_xml(&doc(&body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(md.contains("# 1 简介"), "{md}");
        assert!(md.contains("## 1.1 文档目的"), "{md}");
        assert!(md.contains("# 2 概述"), "{md}");
        assert!(md.contains("## 2.1 背景"), "{md}");
    }

    #[test]
    fn bullet_list_with_indent() {
        let body = format!("{}{}{}", p_num("一级A", 2, 0), p_num("二级A", 2, 1), p_num("一级B", 2, 0));
        let md = convert_xml(&doc(&body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(md.contains("- 一级A\n  - 二级A\n- 一级B"), "{md}");
    }

    #[test]
    fn enclosed_circle_ordered_list() {
        let body = format!("{}{}", p_num("地图展示", 3, 0), p_num("数据采集", 3, 0));
        let md = convert_xml(&doc(&body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(md.contains("① 地图展示"), "{md}");
        assert!(md.contains("② 数据采集"), "{md}");
    }

    #[test]
    fn numid_zero_cancels_numbering() {
        let body = format!(r#"<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>普通段落</w:t></w:r></w:p>"#);
        let md = convert_xml(&doc(&body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(md.contains("普通段落\n"), "{md}");
        assert!(!md.contains("- 普通段落"), "{md}");
    }

    #[test]
    fn toc_skipped() {
        let body = format!("{}{}", p_style("1 简介\t1", 19), p("正文"));
        let md = convert_xml(&doc(&body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(!md.contains("简介"), "{md}");
        assert!(md.contains("正文"), "{md}");
    }

    #[test]
    fn table_rendered() {
        let body = r#"<w:tbl>
            <w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>"#;
        let md = convert_xml(&doc(body), STYLES, NUMBERING, &ConvertOptions::default()).unwrap();
        assert!(md.contains("| 名称 | 值 |"), "{md}");
        assert!(md.contains("|---|---|"), "{md}");
        assert!(md.contains("| a | b |"), "{md}");
    }
}
