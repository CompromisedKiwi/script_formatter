use std::fs;
use std::path::{Path, PathBuf};

use docx_rs::*;
use anyhow::Result;
use lazy_regex::{Lazy, regex, Regex};
use rfd::FileDialog;

const APP_NAME: &str = concat!("🧙‍♂️Script Formatter v", env!("CARGO_PKG_VERSION"));

fn main() {
    // Pick files
    let files_path = FileDialog::new()
        .add_filter("Word / WPS 文档", &["docx", "doc"])
        .set_title(&format!(
            "{APP_NAME} - 请选择要格式化的剧本文件（可选多个）"
        ))
        .pick_files();

    // Process files
    if let Some(paths) = files_path {
        match process_files(&paths) {
            Ok(output_paths) => {
                create_dialog(&format!("格式化完成！已保存至:\n{}", output_paths)).show();
            }
            Err(e) => {
                create_dialog(&format!("处理失败，请截图上报Bug🐞:\n{e:?}")).show();
            }
        }
    }
}

fn create_dialog(content: &str) -> rfd::MessageDialog {
    rfd::MessageDialog::new()
        .set_title(APP_NAME)
        .set_description(content)
        .set_buttons(rfd::MessageButtons::Ok)
}

fn process_files(paths: &[PathBuf]) -> Result<String> {
    const OUTPUT_DIR_NAME: &str = "已格式化";

    if paths.len() < 1 {
        return Ok(String::new());
    }
    let first_dir = paths[0]
        .parent()
        .ok_or_else(|| anyhow::anyhow!("无法获取文件目录"))?;
    let out_dir = first_dir.join(OUTPUT_DIR_NAME);
    fs::create_dir_all(&out_dir)?;

    let mut output_paths = Vec::new();
    for path in paths {
        let p = formatting(path, &out_dir)?;
        output_paths.push(p);
    }
    Ok(output_paths.join("\n"))
}

fn get_text_with_breaks(para: &Paragraph) -> String {
    // Treat soft enter as spliter of paragraph too.
    // Hard enter which copyed from web could be transformed to soft enter, occasionally.
    let mut full_text = String::new();

    for child in &para.children {
        if let ParagraphChild::Run(run) = child {
            for run_child in &run.children {
                match run_child {
                    RunChild::Text(t) => {
                        full_text.push_str(&t.text);
                    }
                    RunChild::Break(_) => {
                        full_text.push('\n');
                    }
                    _ => {}
                }
            }
        }
    }
    full_text
}

fn formatting(input_path: &Path, out_dir: &Path) -> Result<String> {
    // Load doc
    let file_bytes = fs::read(input_path)?;
    let mut reader = read_docx(&file_bytes)?;

    // New doc
    // Set default font to 宋体（正文）, size 小四（12磅）
    let mut new_doc = Docx::new()
        .default_fonts(RunFonts::new().east_asia("宋体").ascii("Times New Roman"))
        .default_size(24); // Unit: half point

    // Extract lines
    for child in reader.document.children.drain(..) {
        // Process paragraph only
        let DocumentChild::Paragraph(para) = child else {
            continue;
        };
        let raw = get_text_with_breaks(&para);
        let lines = raw.split('\n').collect::<Vec<&str>>();
        for line in lines {
            if line.is_empty() {
                continue;
            }
            let new_para = process_line(line.to_string());
            new_doc = new_doc.add_paragraph(new_para);
        }
    }

    // Save new doc
    let file_name = input_path
        .file_name()
        .ok_or_else(|| anyhow::anyhow!("无法获取文件名"))?
        .to_str()
        .ok_or_else(|| anyhow::anyhow!("文件名包含无效UTF-8字符"))?;
    let output_path = out_dir.join(file_name);
    let out_file = fs::File::create(&output_path)?;
    new_doc.build().pack(out_file)?;

    Ok(output_path
        .to_str()
        .ok_or_else(|| anyhow::anyhow!("文件名包含无效UTF-8字符"))?
        .to_string())
}

fn process_line(mut text: String) -> Paragraph {
    // Replace @ with △ in the beginning
    text = text.replace('@', "△");

    // Create new paragraph
    let mut new_para = Paragraph::new();

    // Set to Heading3 style
    static RE_HEADING3: &Lazy<Regex> = regex!(r"^(第.*集|人物.*|[0-9].*)$");
    if RE_HEADING3.is_match(&text) {
        // 三号 bold
        return new_para
            .add_run(Run::new().add_text(text).size(32).bold())
            .line_spacing(
                LineSpacing::new()
                    .before(260) // 13磅
                    .after(260) // 13磅
                    .line_rule(LineSpacingType::Auto)
                    .line(413), // 1.72 * 240
            );
    }

    // Set 【】 to bold
    if text.starts_with("【") {
        return new_para.add_run(Run::new().add_text(text).bold());
    }

    // Now it should only be dialog line

    // find colon
    text = text.replace(':', "：");
    let split_pos = text.find('：');
    let Some(pos) = split_pos else {
        return new_para.add_run(Run::new().add_text(text));
    };

    const COLON_BYTE_LEN: usize = 3;
    let before = &text[..(pos + COLON_BYTE_LEN)]; // include colon
    let after = &text[(pos + COLON_BYTE_LEN)..];
    new_para = new_para.add_run(Run::new().add_text(before));

    // Content inside brackets stay normal
    static RE_BRACKET: &Lazy<Regex> = regex!(r"（[^）]*）");
    let mut last_end = 0;
    for mat in RE_BRACKET.find_iter(after) {
        // Text before bracket: red bold
        if mat.start() > last_end {
            let red_text = &after[last_end..mat.start()];
            new_para = new_para.add_run(Run::new().add_text(red_text).color("FF0000").bold());
        }
        // Text inside bracket: default
        let bracket_text = mat.as_str();
        new_para = new_para.add_run(Run::new().add_text(bracket_text));
        last_end = mat.end();
    }

    // Text after last bracket
    if last_end < after.len() {
        let remain = &after[last_end..];
        new_para = new_para.add_run(Run::new().add_text(remain).color("FF0000").bold());
    }

    return new_para;
}
