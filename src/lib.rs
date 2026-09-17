//! 样式驱动的 DOCX → Markdown 转换器。
//!
//! 核心思路（与启发式纯文本方案的本质区别）：
//! - 标题层级来自 `w:pStyle` → `styles.xml` 中 `heading N` 的真实映射，而非猜测文本特征
//! - 列表（ul/ol）完整解析 `numbering.xml` 的三层间接引用：
//!   段落 `numPr(numId, ilvl)` → `w:num` → `w:abstractNum` 的 `lvl(numFmt, lvlText, start)`
//! - 编号来源三级解析：段落级 numPr 优先 → 样式级 numPr 兜底 → numId=0 表示取消编号
//! - 计数器状态机：按 numId 维护各级 ilvl 计数，推进当前层、重置更深层、上层补 start 值；
//!   `lvlText` 中的 `%1-%9` 按对应层的 numFmt 渲染替换（支持多级编号如 "1.1.2"、"(1)、"①"）
//! - 图片按文档顺序内联提取：`<a:blip r:embed>` / `<v:imagedata r:id>`
//!   → `word/_rels/document.xml.rels` → `word/media/*`，输出 `![alt](prefix/file_name)`

use anyhow::{bail, Context, Result};
use roxmltree::{Document, Node};
use std::collections::HashMap;
use std::fs::File;
use std::io::Read;
use std::path::Path;

// ---------------------------------------------------------------------------
// 公开类型
// ---------------------------------------------------------------------------

/// 转换选项
#[derive(Debug, Clone, Default)]
pub struct ConvertOptions {
    /// 是否保留 TOC 目录段落（默认跳过）
    pub keep_toc: bool,
    /// Markdown 图片链接前缀，如 "out.media/"。空字符串表示只用文件名
    pub image_link_prefix: String,
    /// 是否提取图片（false 时 Markdown 中不生成图片链接，默认 true）
    pub extract_images: bool,
    /// 标题是否保留 Word 自动章节号（如 "# 1.1 文档目的"，默认 true；false 输出 "# 文档目的"）
    pub heading_numbers: bool,
}

impl ConvertOptions {
    pub fn new() -> Self {
        Self { keep_toc: false, image_link_prefix: String::new(), extract_images: true, heading_numbers: true }
    }
}

/// 从 DOCX 中提取的图片
#[derive(Debug, Clone)]
pub struct ExtractedImage {
    /// 输出文件名（取自 word/media/ 原始文件名，已去重）
    pub file_name: String,
    /// 图片二进制数据（`convert_xml` 路径下为空，仅 `convert_bytes`/`convert_file` 填充）
    pub data: Vec<u8>,
}

/// 转换结果
#[derive(Debug, Clone)]
pub struct ConversionResult {
    pub markdown: String,
    /// 文档正文中实际引用到的图片，按首次出现顺序排列
    pub images: Vec<ExtractedImage>,
}

impl ConversionResult {
    /// 将所有图片写入指定目录
    pub fn write_images(&self, dir: impl AsRef<Path>) -> Result<()> {
        let dir = dir.as_ref();
        if !self.images.is_empty() {
            std::fs::create_dir_all(dir)
                .with_context(|| format!("无法创建图片目录 {}", dir.display()))?;
            for img in &self.images {
                std::fs::write(dir.join(&img.file_name), &img.data)
                    .with_context(|| format!("无法写入图片 {}", img.file_name))?;
            }
        }
        Ok(())
    }
}

// ---------------------------------------------------------------------------
// 内部数据结构
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

/// rId → word/media/ 下的原始文件名（不含目录）
type ImageRels = HashMap<String, String>;

/// rId → 超链接 URL
type LinkRels = HashMap<String, String>;

/// 段落内联内容（文本与图片按文档顺序交错）
#[derive(Debug)]
enum Inline {
    Text(String),
    Image { r_id: String, alt: String },
    /// 行内格式：加粗/斜体/删除线
    Styled { bold: bool, italic: bool, strike: bool, children: Vec<Inline> },
    /// 超链接：外部 URL 或文档内书签锚点
    Link { target: Option<String>, anchor: Option<String>, children: Vec<Inline> },
}

// ---------------------------------------------------------------------------
// XML 辅助
// ---------------------------------------------------------------------------

/// 按本地名取属性（忽略命名空间前缀，docx 属性均为 w:xxx / r:xxx）
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
// 解析 numbering.xml / styles.xml / document.xml.rels
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

/// 解析 word/_rels/document.xml.rels。
/// 返回 (rId → 原始文件名, 原始文件名 → zip 内路径, rId → 超链接 URL)。
fn parse_rels(xml: &str) -> (ImageRels, HashMap<String, String>, LinkRels) {
    let mut r_id_to_file: ImageRels = HashMap::new();
    let mut file_to_zip: HashMap<String, String> = HashMap::new();
    let mut links: LinkRels = HashMap::new();
    let Ok(doc) = Document::parse(xml) else { return (r_id_to_file, file_to_zip, links) };

    // 同一 target 被多个 rId 引用时复用同一文件名
    let mut target_to_file: HashMap<String, String> = HashMap::new();
    let mut used_names: Vec<String> = Vec::new();

    for rel in doc.descendants().filter(|n| tag(n) == "Relationship") {
        let rel_type = attr(&rel, "Type").unwrap_or("");
        let (Some(id), Some(target)) = (attr(&rel, "Id"), attr(&rel, "Target")) else { continue };

        if rel_type.ends_with("/hyperlink") {
            links.insert(id.to_string(), target.to_string());
            continue;
        }
        if !rel_type.ends_with("/image") {
            continue;
        }

        // 归一化为 zip 内路径："/word/media/x.png" 或相对 document.xml 的 "media/x.png"
        let zip_path = if let Some(stripped) = target.strip_prefix('/') {
            stripped.to_string()
        } else {
            format!("word/{target}")
        };

        let file_name = target_to_file.entry(target.to_string()).or_insert_with(|| {
            let base = target.rsplit('/').next().unwrap_or("image");
            let mut name = base.to_string();
            // 防止不同目录下同名文件冲突
            let mut i = 1;
            while used_names.contains(&name) {
                name = format!("{}_{}", i, base);
                i += 1;
            }
            used_names.push(name.clone());
            name
        });
        file_to_zip.insert(file_name.clone(), zip_path);
        r_id_to_file.insert(id.to_string(), file_name.clone());
    }
    (r_id_to_file, file_to_zip, links)
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
// 段落 / 表格提取（含图片、超链接、行内格式）
// ---------------------------------------------------------------------------

/// 转换上下文：渲染时需要的只读引用集合
struct Ctx<'a> {
    images: &'a ImageRels,
    links: &'a LinkRels,
    image_prefix: &'a str,
    extract_images: bool,
    used_images: Vec<String>,
}

/// rPr 中的开/关属性（w:b、w:i 等）：存在且 val 不是 false/0/off 即为开
fn is_on(el: Option<Node>) -> bool {
    match el {
        None => false,
        Some(n) => !matches!(attr(&n, "val"), Some("false") | Some("0") | Some("off")),
    }
}

/// 从 drawing/pict 容器中提取图片（docPr 的 descr/name 作为 alt）
fn image_from_container(container: &Node) -> Option<Inline> {
    let mut alt = String::new();
    let mut r_id: Option<String> = None;
    for node in container.descendants() {
        match tag(&node) {
            "inline" | "anchor" => alt.clear(),
            "docPr" => {
                alt = attr(&node, "descr").filter(|s| !s.is_empty())
                    .or_else(|| attr(&node, "name"))
                    .unwrap_or("").to_string();
            }
            "blip" => {
                if r_id.is_none() {
                    r_id = attr(&node, "embed").or_else(|| attr(&node, "link")).map(String::from);
                }
            }
            "imagedata" => {
                if r_id.is_none() {
                    r_id = attr(&node, "id").map(String::from);
                }
            }
            _ => {}
        }
    }
    r_id.map(|r_id| Inline::Image { r_id, alt })
}

/// 提取 run（w:r）内容：解析 rPr 行内格式 + 子内容
fn run_inline(r: &Node) -> Vec<Inline> {
    let mut bold = false;
    let mut italic = false;
    let mut strike = false;
    let mut children: Vec<Inline> = Vec::new();

    for child in r.children().filter(|n| n.is_element()) {
        match tag(&child) {
            "rPr" => {
                bold = is_on(find_path(&child, &["b"]));
                italic = is_on(find_path(&child, &["i"]));
                strike = is_on(find_path(&child, &["strike"])) || is_on(find_path(&child, &["dstrike"]));
            }
            "t" => {
                let text = child.text().unwrap_or("");
                if !text.is_empty() {
                    children.push(Inline::Text(text.to_string()));
                }
            }
            "tab" | "br" | "cr" => children.push(Inline::Text(" ".into())),
            "drawing" | "pict" => {
                if let Some(img) = image_from_container(&child) {
                    children.push(img);
                }
            }
            _ => {}
        }
    }

    if (bold || italic || strike) && !children.is_empty() {
        vec![Inline::Styled { bold, italic, strike, children }]
    } else {
        children
    }
}

/// 解析 fldSimple 的 HYPERLINK 域指令，提取 URL
fn fld_simple_url(instr: &str) -> Option<String> {
    if !instr.contains("HYPERLINK") {
        return None;
    }
    // 形如 ` HYPERLINK "https://example.com" `
    instr.split('"').nth(1).map(String::from)
}

/// 推入内联节点：相邻且格式相同的 Styled 自动合并，避免产生 `****` 这类断裂标记
fn push_inline(out: &mut Vec<Inline>, item: Inline) {
    if let Inline::Styled { bold, italic, strike, .. } = &item {
        let (b, i, s) = (*bold, *italic, *strike);
        if let Some(Inline::Styled { bold: pb, italic: pi, strike: ps, children: pc }) = out.last_mut() {
            if *pb == b && *pi == i && *ps == s {
                if let Inline::Styled { children, .. } = item {
                    pc.extend(children);
                }
                return;
            }
        }
    }
    out.push(item);
}

/// 段落内联内容：递归遍历，支持超链接嵌套、行内格式、修订（接受插入/拒绝删除）、
/// 复杂域代码超链接（fldChar/instrText）、mc:AlternateContent 只取 Choice 分支。
fn walk_inlines<'a>(node: &Node<'a, 'a>, out: &mut Vec<Inline>) {
    let mut fld_instr = String::new();          // 域指令累积（begin → separate 之间）
    let mut fld_collecting = false;
    let mut fld_display: Option<Vec<Inline>> = None; // separate → end 之间的显示内容

    for child in node.children().filter(|n| n.is_element()) {
        match tag(&child) {
            "r" => {
                // 复杂域代码状态机：fldChar begin/separate/end + instrText
                let fld = find_path(&child, &["fldChar"]).and_then(|n| attr(&n, "fldCharType"));
                match fld {
                    Some("begin") => { fld_instr.clear(); fld_collecting = true; continue; }
                    Some("separate") => { fld_collecting = false; fld_display = Some(Vec::new()); continue; }
                    Some("end") => {
                        if let Some(inner) = fld_display.take() {
                            match fld_simple_url(&fld_instr) {
                                Some(url) if !inner.is_empty() =>
                                    push_inline(out, Inline::Link { target: Some(url), anchor: None, children: inner }),
                                _ => { for it in inner { push_inline(out, it); } }
                            }
                        }
                        continue;
                    }
                    _ => {}
                }
                if fld_collecting {
                    if let Some(instr) = find_path(&child, &["instrText"]) {
                        fld_instr.push_str(instr.text().unwrap_or(""));
                    }
                    continue;
                }
                let items = run_inline(&child);
                if let Some(disp) = &mut fld_display {
                    for it in items { push_inline(disp, it); }
                } else {
                    for it in items { push_inline(out, it); }
                }
            }
            "hyperlink" => {
                let mut inner = Vec::new();
                walk_inlines(&child, &mut inner);
                if !inner.is_empty() {
                    let r_id = attr(&child, "id").map(String::from);
                    let anchor = attr(&child, "anchor").map(String::from);
                    push_inline(out, Inline::Link { target: r_id, anchor, children: inner });
                }
            }
            "fldSimple" => {
                // 旧式域代码超链接：<w:fldSimple w:instr=" HYPERLINK \"url\" ">
                let url = attr(&child, "instr").and_then(fld_simple_url);
                let mut inner = Vec::new();
                walk_inlines(&child, &mut inner);
                match url {
                    Some(u) if !inner.is_empty() =>
                        push_inline(out, Inline::Link { target: Some(u), anchor: None, children: inner }),
                    _ => { for it in inner { push_inline(out, it); } }
                }
            }
            // 修订：接受插入，拒绝删除
            "ins" => walk_inlines(&child, out),
            "del" | "moveFrom" => {}
            // 结构化文档标签 / 智能标签：递归取内容
            "sdt" | "sdtContent" | "smartTag" => walk_inlines(&child, out),
            // AlternateContent：只取 Choice，跳过 Fallback（否则图片重复）
            "AlternateContent" => {
                if let Some(choice) = child.children().find(|n| tag(n) == "Choice") {
                    walk_inlines(&choice, out);
                }
            }
            "Fallback" => {}
            // 图片容器直接出现在段落层（少数情况）
            "drawing" | "pict" => {
                if let Some(img) = image_from_container(&child) {
                    push_inline(out, img);
                }
            }
            _ => walk_inlines(&child, out),
        }
    }
}

/// 段落内联内容入口
fn para_inlines(p: &Node) -> Vec<Inline> {
    let mut out = Vec::new();
    walk_inlines(p, &mut out);
    out
}

/// 渲染内联内容为 Markdown 文本。
/// 图片 → ![alt](prefix/file_name)；超链接 → [text](url)；格式 → **/* /~~。
fn render_inlines(inlines: &[Inline], ctx: &mut Ctx) -> String {
    let mut text = String::new();
    for item in inlines {
        match item {
            Inline::Text(t) => text.push_str(t),
            Inline::Image { r_id, alt } if ctx.extract_images => {
                if let Some(file_name) = ctx.images.get(r_id) {
                    if !ctx.used_images.contains(file_name) {
                        ctx.used_images.push(file_name.clone());
                    }
                    text.push_str(&format!("![{alt}]({}{file_name})", ctx.image_prefix));
                }
            }
            Inline::Image { .. } => {}
            Inline::Styled { bold, italic, strike, children } => {
                let inner = render_inlines(children, ctx);
                if inner.trim().is_empty() {
                    continue;
                }
                let open = format!("{}{}{}", if *strike { "~~" } else { "" }, if *bold { "**" } else { "" }, if *italic { "*" } else { "" });
                let close = format!("{}{}{}", if *italic { "*" } else { "" }, if *bold { "**" } else { "" }, if *strike { "~~" } else { "" });
                text.push_str(&format!("{open}{inner}{close}"));
            }
            Inline::Link { target, anchor, children } => {
                let inner = render_inlines(children, ctx);
                if inner.trim().is_empty() {
                    continue;
                }
                let url = target.as_ref().and_then(|id| ctx.links.get(id)).cloned()
                    .or_else(|| target.clone().filter(|t| t.starts_with("http")))  // fldSimple 直接存 URL
                    .map(String::from)
                    .or_else(|| anchor.as_ref().map(|a| format!("#{a}")));
                match url {
                    Some(u) => text.push_str(&format!("[{inner}]({u})")),
                    None => text.push_str(&inner),
                }
            }
        }
    }
    text
}

/// 段落渲染为 Markdown 文本（trim）
fn para_text(p: &Node, ctx: &mut Ctx) -> String {
    render_inlines(&para_inlines(p), ctx).trim().to_string()
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

/// 段落的直接大纲级别（w:outlineLvl，0 索引）。无标题样式但有大纲级别的段落可兜底为标题。
fn para_outline_level(p: &Node) -> Option<u32> {
    find_path(p, &["pPr", "outlineLvl"])
        .and_then(|n| attr(&n, "val"))
        .and_then(|v| v.parse().ok())
        .filter(|&v| v <= 5) // 0-5 → 一至六级标题；9 表示正文
}

/// 表格 → Markdown 表格。
/// 处理合并单元格（gridSpan 橫向占列补空、vMerge 纵向合并续行留空）、
/// 单元格多段落（<br> 连接）、单元格内图片与超链接。
fn table_md(tbl: &Node, ctx: &mut Ctx) -> String {
    let mut rows: Vec<Vec<String>> = Vec::new();
    for tr in tbl.children().filter(|n| tag(n) == "tr") {
        let mut cells: Vec<String> = Vec::new();
        for tc in tr.children().filter(|n| tag(n) == "tc") {
            let tcpr = tc.children().find(|n| tag(n) == "tcPr");
            let grid_span: usize = tcpr
                .and_then(|p| find_path(&p, &["gridSpan"]))
                .and_then(|n| attr(&n, "val"))
                .and_then(|v| v.parse().ok())
                .unwrap_or(1);
            // vMerge：无 val 或 val="continue" 为续行（内容留空），"restart" 为合并起始
            let is_vmerge_continue = tcpr
                .and_then(|p| find_path(&p, &["vMerge"]))
                .map(|n| !matches!(attr(&n, "val"), Some("restart")))
                .unwrap_or(false);

            // 单元格内所有段落，用 <br> 连接
            let text = if is_vmerge_continue {
                String::new()
            } else {
                tc.children()
                    .filter(|n| tag(n) == "p")
                    .map(|p| para_text(&p, ctx))
                    .filter(|t| !t.is_empty())
                    .collect::<Vec<_>>()
                    .join("<br>")
            };
            cells.push(text.replace('|', "\\|"));
            // 橫向合并占列：补空单元格
            for _ in 1..grid_span {
                cells.push(String::new());
            }
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

/// 核心转换：从解析好的 XML 字符串生成 Markdown。
/// 返回 (markdown, 引用到的图片文件名列表)。
fn convert_doc(
    document_xml: &str,
    styles_xml: &str,
    numbering_xml: &str,
    rels_xml: &str,
    options: &ConvertOptions,
) -> Result<(String, Vec<String>, ImageRels)> {
    let doc = Document::parse(document_xml).context("document.xml 解析失败")?;
    let styles = parse_styles(styles_xml);
    let numbering = parse_numbering(numbering_xml);
    let (rels, _, links) = parse_rels(rels_xml);
    let mut counters = CounterState::new(&numbering);
    let mut ctx = Ctx {
        images: &rels,
        links: &links,
        image_prefix: &options.image_link_prefix,
        extract_images: options.extract_images,
        used_images: Vec::new(),
    };

    let Some(body) = doc.descendants().find(|n| tag(n) == "body") else {
        bail!("document.xml 中没有 body");
    };

    let mut md = String::new();
    for el in body.children().filter(|n| n.is_element()) {
        match tag(&el) {
            "tbl" => {
                let t = table_md(&el, &mut ctx);
                if !t.is_empty() {
                    md.push_str(&t);
                    md.push_str("\n\n");
                }
            }
            "p" => {
                let text = para_text(&el, &mut ctx);
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

                // 标题判定：标题样式优先，w:outlineLvl 直接大纲级别兜底
                let heading_level = sid.as_ref()
                    .and_then(|id| styles.headings.get(id).copied())
                    .or_else(|| para_outline_level(&el).map(|v| v + 1));

                if let Some(level) = heading_level {
                    // 标题：可能带自动编号（多级号）；计数器始终推进，是否展示由 heading_numbers 控制
                    let prefix_num = match &num_id {
                        Some(nid) => {
                            let (is_bullet, numtext) = counters.resolve(nid, ilvl);
                            if options.heading_numbers && !is_bullet && !numtext.is_empty() {
                                format!("{numtext} ")
                            } else {
                                String::new()
                            }
                        }
                        None => String::new(),
                    };
                    let hashes = "#".repeat(level.min(6) as usize);
                    md.push_str(&format!("{hashes} {prefix_num}{text}\n\n"));
                    continue;
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
    Ok((md, ctx.used_images.clone(), rels))
}

/// 从 DOCX 字节流转换为 Markdown，并提取正文引用的图片数据
pub fn convert_bytes(bytes: &[u8], options: &ConvertOptions) -> Result<ConversionResult> {
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
    let rels_xml = read_entry("word/_rels/document.xml.rels").unwrap_or_default();

    let (_, file_to_zip, _) = parse_rels(&rels_xml);
    let (markdown, used_files, _) = convert_doc(&document_xml, &styles_xml, &numbering_xml, &rels_xml, options)?;

    // 只提取正文中实际引用的图片数据
    let mut images = Vec::new();
    for file_name in &used_files {
        if let Some(zip_path) = file_to_zip.get(file_name) {
            match archive.by_name(zip_path) {
                Ok(mut f) => {
                    let mut data = Vec::new();
                    f.read_to_end(&mut data)?;
                    images.push(ExtractedImage { file_name: file_name.clone(), data });
                }
                Err(_) => eprintln!("警告: zip 中找不到图片 {zip_path}，已跳过"),
            }
        }
    }
    Ok(ConversionResult { markdown, images })
}

/// 从 DOCX 文件路径转换为 Markdown，并提取正文引用的图片数据
pub fn convert_file(path: impl AsRef<Path>, options: &ConvertOptions) -> Result<ConversionResult> {
    let path = path.as_ref();
    let mut f = File::open(path).with_context(|| format!("无法打开 {}", path.display()))?;
    let mut bytes = Vec::new();
    f.read_to_end(&mut bytes)?;
    convert_bytes(&bytes, options)
}

/// 从解析好的 XML 字符串转换（便于测试；图片数据为空）
pub fn convert_xml(document_xml: &str, styles_xml: &str, numbering_xml: &str, rels_xml: &str, options: &ConvertOptions) -> Result<ConversionResult> {
    let (markdown, used_files, _) = convert_doc(document_xml, styles_xml, numbering_xml, rels_xml, options)?;
    let images = used_files.into_iter()
        .map(|file_name| ExtractedImage { file_name, data: Vec::new() })
        .collect();
    Ok(ConversionResult { markdown, images })
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

    const RELS: &str = r#"<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
        <Relationship Id="rId5" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image1.png"/>
        <Relationship Id="rId6" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/image2.jpg"/>
        <Relationship Id="rId7" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="/word/media/image3.gif"/>
        <Relationship Id="rId8" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>
        <Relationship Id="rId10" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Target="https://modao.cc/proto/xxx" TargetMode="External"/>
    </Relationships>"#;

    fn doc(body: &str) -> String {
        format!(r#"<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:v="urn:schemas-microsoft-com:vml"><w:body>{body}</w:body></w:document>"#)
    }

    fn convert(body: &str) -> ConversionResult {
        convert_xml(&doc(body), STYLES, NUMBERING, RELS, &ConvertOptions {
            image_link_prefix: "img/".into(),
            ..ConvertOptions::new()
        }).unwrap()
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

    fn p_drawing(alt: &str, r_id: &str) -> String {
        format!(r#"<w:p><w:r><w:drawing><wp:inline><wp:docPr name="{alt}" descr=""/><a:graphic><a:graphicData><a:blip r:embed="{r_id}"/></a:graphicData></a:graphic></wp:inline></w:drawing></w:r></w:p>"#)
    }

    #[test]
    fn headings_get_multilevel_numbers() {
        let body = format!("{}{}{}{}", p_style("简介", 3), p_style("文档目的", 4), p_style("概述", 3), p_style("背景", 4));
        let r = convert(&body);
        assert!(r.markdown.contains("# 1 简介"), "{}", r.markdown);
        assert!(r.markdown.contains("## 1.1 文档目的"), "{}", r.markdown);
        assert!(r.markdown.contains("# 2 概述"), "{}", r.markdown);
        assert!(r.markdown.contains("## 2.1 背景"), "{}", r.markdown);
    }

    #[test]
    fn bullet_list_with_indent() {
        let body = format!("{}{}{}", p_num("一级A", 2, 0), p_num("二级A", 2, 1), p_num("一级B", 2, 0));
        let r = convert(&body);
        assert!(r.markdown.contains("- 一级A\n  - 二级A\n- 一级B"), "{}", r.markdown);
    }

    #[test]
    fn enclosed_circle_ordered_list() {
        let body = format!("{}{}", p_num("地图展示", 3, 0), p_num("数据采集", 3, 0));
        let r = convert(&body);
        assert!(r.markdown.contains("① 地图展示"), "{}", r.markdown);
        assert!(r.markdown.contains("② 数据采集"), "{}", r.markdown);
    }

    #[test]
    fn numid_zero_cancels_numbering() {
        let body = r#"<w:p><w:pPr><w:numPr><w:numId w:val="0"/></w:numPr></w:pPr><w:r><w:t>普通段落</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("普通段落\n"), "{}", r.markdown);
        assert!(!r.markdown.contains("- 普通段落"), "{}", r.markdown);
    }

    #[test]
    fn toc_skipped() {
        let body = format!("{}{}", p_style("1 简介\t1", 19), p("正文"));
        let r = convert(&body);
        assert!(!r.markdown.contains("简介"), "{}", r.markdown);
        assert!(r.markdown.contains("正文"), "{}", r.markdown);
    }

    #[test]
    fn table_rendered() {
        let body = r#"<w:tbl>
            <w:tr><w:tc><w:p><w:r><w:t>名称</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>值</w:t></w:r></w:p></w:tc></w:tr>
            <w:tr><w:tc><w:p><w:r><w:t>a</w:t></w:r></w:p></w:tc><w:tc><w:p><w:r><w:t>b</w:t></w:r></w:p></w:tc></w:tr>
        </w:tbl>"#;
        let r = convert(body);
        assert!(r.markdown.contains("| 名称 | 值 |"), "{}", r.markdown);
        assert!(r.markdown.contains("|---|---|"), "{}", r.markdown);
        assert!(r.markdown.contains("| a | b |"), "{}", r.markdown);
    }

    #[test]
    fn drawing_image_rendered_with_alt_and_prefix() {
        let body = format!("{}{}", p("架构如下图所示："), p_drawing("系统架构图", "rId5"));
        let r = convert(&body);
        assert!(r.markdown.contains("![系统架构图](img/image1.png)"), "{}", r.markdown);
        assert_eq!(r.images.len(), 1);
        assert_eq!(r.images[0].file_name, "image1.png");
    }

    #[test]
    fn vml_imagedata_supported() {
        let body = format!(r#"<w:p><w:r><w:pict><v:shape><v:imagedata r:id="rId6"/></v:shape></w:pict></w:r></w:p>"#);
        let r = convert(&body);
        assert!(r.markdown.contains("![](img/image2.jpg)"), "{}", r.markdown);
    }

    #[test]
    fn absolute_target_normalized() {
        let body = p_drawing("", "rId7");
        let r = convert(&body);
        assert!(r.markdown.contains("img/image3.gif"), "{}", r.markdown);
    }

    #[test]
    fn unreferenced_images_not_extracted() {
        let body = p("纯文本，无图片");
        let r = convert(&body);
        assert!(r.images.is_empty(), "{:?}", r.images);
    }

    #[test]
    fn outline_lvl_fallback_heading() {
        // 无标题样式但有 w:outlineLvl 的段落兜底为标题
        let body = r#"<w:p><w:pPr><w:outlineLvl w:val="1"/></w:pPr><w:r><w:t>直接格式化的节标题</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("## 直接格式化的节标题"), "{}", r.markdown);
    }

    #[test]
    fn outline_lvl_body_text_not_heading() {
        // outlineLvl=9 是正文，不应当成标题
        let body = r#"<w:p><w:pPr><w:outlineLvl w:val="9"/></w:pPr><w:r><w:t>正文段落</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(!r.markdown.contains('#'), "{}", r.markdown);
    }

    #[test]
    fn heading_numbers_disabled() {
        let body = format!("{}{}", p_style("简介", 3), p_style("文档目的", 4));
        let r = convert_xml(&doc(&body), STYLES, NUMBERING, RELS, &ConvertOptions {
            heading_numbers: false,
            ..ConvertOptions::new()
        }).unwrap();
        assert!(r.markdown.contains("# 简介\n"), "{}", r.markdown);
        assert!(r.markdown.contains("## 文档目的\n"), "{}", r.markdown);
        assert!(!r.markdown.contains("# 1 简介"), "{}", r.markdown);
    }

    #[test]
    fn hyperlink_external() {
        let body = r#"<w:p><w:r><w:t>详见</w:t></w:r><w:hyperlink r:id="rId10"><w:r><w:t>原型地址</w:t></w:r></w:hyperlink><w:r><w:t>。</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("详见[原型地址](https://modao.cc/proto/xxx)。"), "{}", r.markdown);
    }

    #[test]
    fn hyperlink_internal_anchor() {
        let body = r#"<w:p><w:hyperlink w:anchor="chapter1"><w:r><w:t>跳转到第一章</w:t></w:r></w:hyperlink></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("[跳转到第一章](#chapter1)"), "{}", r.markdown);
    }

    #[test]
    fn hyperlink_fld_simple() {
        let body = r#"<w:p><w:fldSimple w:instr=" HYPERLINK &quot;https://example.com&quot; "><w:r><w:t>示例站点</w:t></w:r></w:fldSimple></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("[示例站点](https://example.com)"), "{}", r.markdown);
    }

    #[test]
    fn inline_formatting() {
        let body = r#"<w:p>
            <w:r><w:rPr><w:b/></w:rPr><w:t>加粗</w:t></w:r>
            <w:r><w:rPr><w:i/></w:rPr><w:t>斜体</w:t></w:r>
            <w:r><w:rPr><w:b/><w:i/></w:rPr><w:t>加粗斜体</w:t></w:r>
            <w:r><w:rPr><w:strike/></w:rPr><w:t>删除线</w:t></w:r>
            <w:r><w:rPr><w:b w:val="false"/></w:rPr><w:t>不加粗</w:t></w:r>
        </w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("**加粗**"), "{}", r.markdown);
        assert!(r.markdown.contains("*斜体*"), "{}", r.markdown);
        assert!(r.markdown.contains("***加粗斜体***"), "{}", r.markdown);
        assert!(r.markdown.contains("~~删除线~~"), "{}", r.markdown);
        assert!(r.markdown.contains("不加粗"), "{}", r.markdown);
        assert!(!r.markdown.contains("**不加粗**"), "{}", r.markdown);
    }

    #[test]
    fn track_changes_accept_insert_reject_delete() {
        let body = r#"<w:p><w:r><w:t>前</w:t></w:r><w:ins><w:r><w:t>插入</w:t></w:r></w:ins><w:del><w:r><w:delText>删除</w:delText></w:r></w:del><w:r><w:t>后</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("前插入后"), "{}", r.markdown);
        assert!(!r.markdown.contains("删除"), "{}", r.markdown);
    }

    #[test]
    fn alternate_content_no_duplicate_images() {
        let body = r#"<w:p><mc:AlternateContent xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006">
            <mc:Choice><w:drawing><wp:inline><wp:docPr name="图"/><a:blip r:embed="rId5"/></wp:inline></w:drawing></mc:Choice>
            <mc:Fallback><w:pict><v:imagedata r:id="rId5"/></w:pict></mc:Fallback>
        </mc:AlternateContent></w:p>"#;
        let r = convert(body);
        assert_eq!(r.markdown.matches("image1.png").count(), 1, "{}", r.markdown);
    }

    #[test]
    fn table_grid_span_and_vmerge() {
        let body = r#"<w:tbl>
            <w:tr>
                <w:tc><w:tcPr><gridSpan xmlns=""/><w:gridSpan w:val="2"/></w:tcPr><w:p><w:r><w:t>跨两列</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:vMerge w:val="restart"/></w:tcPr><w:p><w:r><w:t>纵合起始</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>A</w:t></w:r></w:p></w:tc>
            </w:tr>
            <w:tr>
                <w:tc><w:tcPr><w:vMerge/></w:tcPr><w:p><w:r><w:t>被吞掉</w:t></w:r></w:p></w:tc>
                <w:tc><w:p><w:r><w:t>B</w:t></w:r></w:p></w:tc>
            </w:tr>
        </w:tbl>"#;
        let r = convert(body);
        assert!(r.markdown.contains("| 跨两列 |  |"), "{}", r.markdown);
        assert!(r.markdown.contains("| 纵合起始 | A |"), "{}", r.markdown);
        assert!(r.markdown.contains("|  | B |"), "{}", r.markdown);
        assert!(!r.markdown.contains("被吞掉"), "{}", r.markdown);
    }

    #[test]
    fn table_multi_paragraph_cell() {
        let body = r#"<w:tbl><w:tr><w:tc>
            <w:p><w:r><w:t>第一段</w:t></w:r></w:p>
            <w:p><w:r><w:t>第二段</w:t></w:r></w:p>
        </w:tc></w:tr></w:tbl>"#;
        let r = convert(body);
        assert!(r.markdown.contains("第一段<br>第二段"), "{}", r.markdown);
    }

    #[test]
    fn adjacent_same_style_merged() {
        // 相邻同格式 run 合并，不产生 `****` 断裂
        let body = r#"<w:p><w:r><w:rPr><w:b/></w:rPr><w:t>杭州海络</w:t></w:r><w:r><w:rPr><w:b/></w:rPr><w:t>信息技术</w:t></w:r></w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("**杭州海络信息技术**"), "{}", r.markdown);
        assert!(!r.markdown.contains("****"), "{}", r.markdown);
    }

    #[test]
    fn complex_field_hyperlink() {
        // fldChar begin/instrText/separate/end 序列 → [text](url)
        let body = r#"<w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText> HYPERLINK &quot;https://example.com&quot; </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>显示文本</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("[显示文本](https://example.com)"), "{}", r.markdown);
    }

    #[test]
    fn complex_field_internal_link_plain_text() {
        // HYPERLINK \l _Toc 内部锚点：无 URL，保留显示文本不生成链接
        let body = r#"<w:p>
            <w:r><w:fldChar w:fldCharType="begin"/></w:r>
            <w:r><w:instrText> HYPERLINK \l _Toc123 </w:instrText></w:r>
            <w:r><w:fldChar w:fldCharType="separate"/></w:r>
            <w:r><w:t>章节引用</w:t></w:r>
            <w:r><w:fldChar w:fldCharType="end"/></w:r>
        </w:p>"#;
        let r = convert(body);
        assert!(r.markdown.contains("章节引用"), "{}", r.markdown);
        assert!(!r.markdown.contains("]("), "{}", r.markdown);
    }

    #[test]
    fn image_in_table_cell() {
        let body = format!(r#"<w:tbl>
            <w:tr><w:tc><w:p><w:r><w:t>截图</w:t></w:r></w:p></w:tc>
            <w:tc>{}</w:tc></w:tr>
        </w:tbl>"#, p_drawing("界面", "rId5"));
        let r = convert(&body);
        assert!(r.markdown.contains("![界面](img/image1.png)"), "{}", r.markdown);
        assert_eq!(r.images.len(), 1);
    }
}
