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

    /// 图片导出目录（默认：<输出文件名去掉扩展名>.media/，stdout 模式下为 <输入名>.media/）
    #[arg(long)]
    images_dir: Option<PathBuf>,

    /// Markdown 中图片链接前缀（默认取图片目录名 + "/"）
    #[arg(long)]
    image_prefix: Option<String>,

    /// 不导出图片（Markdown 中也不生成图片链接）
    #[arg(long)]
    no_images: bool,

    /// 标题不保留 Word 自动章节号（如 "1.1"）
    #[arg(long)]
    no_heading_numbers: bool,

    /// 保留 TOC 目录段落（默认跳过）
    #[arg(long)]
    keep_toc: bool,
}

fn main() -> Result<()> {
    let cli = Cli::parse();

    // 确定图片导出目录
    let images_dir = if cli.no_images {
        None
    } else {
        Some(cli.images_dir.clone().unwrap_or_else(|| {
            let base = cli.output.as_ref().unwrap_or(&cli.input);
            let stem = base.file_stem().map(|s| s.to_string_lossy().into_owned()).unwrap_or_else(|| "docx".into());
            base.parent().unwrap_or_else(|| std::path::Path::new(".")).join(format!("{stem}.media"))
        }))
    };

    // 链接前缀：默认用目录名（相对路径），保证 md 与图片目录同级放置时可直接打开
    let image_link_prefix = if cli.no_images {
        String::new()
    } else {
        cli.image_prefix.clone().unwrap_or_else(|| {
            let dir = images_dir.as_ref().unwrap();
            let name = dir.file_name().map(|s| s.to_string_lossy().into_owned()).unwrap_or_default();
            format!("{name}/")
        })
    };

    let options = docx2md::ConvertOptions {
        keep_toc: cli.keep_toc,
        image_link_prefix,
        extract_images: !cli.no_images,
        heading_numbers: !cli.no_heading_numbers,
    };
    let result = docx2md::convert_file(&cli.input, &options)
        .with_context(|| format!("转换失败: {}", cli.input.display()))?;

    // 导出图片
    if !cli.no_images {
        if let Some(dir) = &images_dir {
            result.write_images(dir)?;
            if !result.images.is_empty() {
                eprintln!("已导出 {} 张图片到 {}", result.images.len(), dir.display());
            }
        }
    }

    match cli.output {
        Some(path) => {
            std::fs::write(&path, &result.markdown).with_context(|| format!("无法写入 {}", path.display()))?;
            eprintln!("已写入 {} ({} 字符)", path.display(), result.markdown.len());
        }
        None => print!("{}", result.markdown),
    }
    Ok(())
}
