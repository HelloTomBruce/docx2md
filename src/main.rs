use anyhow::{Context, Result};
use clap::Parser;
use std::path::PathBuf;

/// 样式驱动的 DOCX → Markdown 转换器
#[derive(Parser)]
#[command(name = "docx2md", version, about)]
struct Cli {
    /// 输入 DOCX 文件路径
    input: PathBuf,

    /// 输出 Markdown 文件路径（默认输出到 stdout）
    #[arg(short, long)]
    output: Option<PathBuf>,

    /// 保留 TOC 目录段落（默认跳过）
    #[arg(long)]
    keep_toc: bool,
}

fn main() -> Result<()> {
    let cli = Cli::parse();
    let options = docx2md::ConvertOptions { keep_toc: cli.keep_toc };
    let md = docx2md::convert_file(&cli.input, &options)
        .with_context(|| format!("转换失败: {}", cli.input.display()))?;

    match cli.output {
        Some(path) => {
            std::fs::write(&path, &md).with_context(|| format!("无法写入 {}", path.display()))?;
            eprintln!("已写入 {} ({} 字符)", path.display(), md.len());
        }
        None => print!("{md}"),
    }
    Ok(())
}
