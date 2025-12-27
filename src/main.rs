use docx_rs::*;
use regex::Regex;
use rfd::FileDialog;
use std::fs::File;
use std::path::Path;

fn main() {
    // 1. 选取文件
    let file_path = FileDialog::new()
        .add_filter("Word 文档", &["docx"])
        .set_title("请选择要标红对话的剧本文件")
        .pick_file();

    if let Some(path) = file_path {
        match colorize_after_colon(&path, "标红") {
            Ok(out) => println!("处理完成！已保存至: {}", out),
            Err(e) => eprintln!("处理失败: {:?}", e),
        }
    }
}

fn colorize_after_colon(
    input_path: &Path,
    prefix: &str,
) -> Result<String, Box<dyn std::error::Error>> {
    // 读取文档
    let file = File::open(input_path)?;
    let mut reader = read_docx(&file)?;

    // 创建新文档用于存放结果
    let mut new_doc = Docx::new();

    // 初始化正则（处理括号）
    let re_bracket = Regex::new(r"（[^）]*）")?;
    // 排除前缀
    let skip_prefixes = ["【", "△", "人物"];

    // 遍历旧文档中的每一个段落
    for child in reader.document.children.drain(..) {
        if let DocumentChild::Paragraph(para) = child {
            let mut new_para = Paragraph::new();

            // --- 处理软回车逻辑 ---
            // 将段落内的所有 Run 合并处理，并识别其中的 Break (Shift+Enter)
            // 在 docx-rs 中，我们需要手动处理这些 children

            let mut current_line_text = String::new();

            // 我们通过一个临时容器来模拟“行”的分隔
            // 如果遇到 Break，我们就把之前的文字当做一行处理
            for run_child in &para.children {
                match run_child {
                    ParagraphChild::Run(run) => {
                        for run_content in &run.children {
                            match run_content {
                                RunChild::Text(t) => current_line_text.push_str(&t.text),
                                RunChild::Break(_) => {
                                    // 遇到软回车，处理当前积累的文字，并添加一个 Break
                                    process_line(
                                        &mut new_para,
                                        &current_line_text,
                                        &re_bracket,
                                        &skip_prefixes,
                                    );
                                    new_para = new_para
                                        .add_run(Run::new().add_break(BreakType::TextWrapping));
                                    current_line_text.clear();
                                }
                                _ => {}
                            }
                        }
                    }
                    _ => {}
                }
            }

            // 处理段落最后剩余的文字（或者没有软回车的普通段落）
            if !current_line_text.is_empty() {
                process_line(
                    &mut new_para,
                    &current_line_text,
                    &re_bracket,
                    &skip_prefixes,
                );
            }

            new_doc = new_doc.add_paragraph(new_para);
        }
    }

    // 保存文件
    let file_name = input_path.file_name().unwrap().to_str().unwrap();
    let parent = input_path.parent().unwrap();
    let output_path = parent.join(format!("{}_{}", prefix, file_name));
    let out_file = File::create(&output_path)?;
    new_doc.build().pack(out_file)?;

    Ok(output_path.to_str().unwrap().to_string())
}

/// 处理单行文本逻辑（冒号分割、括号排除、标红）
fn process_line(para: &mut Paragraph, text: &str, re_bracket: &Regex, skips: &[&str]) {
    if text.is_empty() {
        return;
    }

    // 1. 检查排除前缀
    if skips.iter().any(|&s| text.starts_with(s)) {
        *para = para.clone().add_run(Run::new().add_text(text));
        return;
    }

    // 2. 寻找冒号 (支持中文和英文)
    let split_pos = text.find('：').or_else(|| text.find(':'));

    if let Some(pos) = split_pos {
        let before = &text[..=pos]; // 包含冒号本身
        let after = &text[pos + 1..];

        // 添加前半部分（原始格式）
        *para = para.clone().add_run(Run::new().add_text(before));

        // 处理后半部分（括号逻辑）
        let mut last_end = 0;
        for mat in re_bracket.find_iter(after) {
            // 括号前的文字：标红加粗
            if mat.start() > last_end {
                let red_text = &after[last_end..mat.start()];
                *para = para
                    .clone()
                    .add_run(Run::new().add_text(red_text).color("FF0000").bold());
            }
            // 括号内容：默认格式
            let bracket_text = mat.as_str();
            *para = para.clone().add_run(Run::new().add_text(bracket_text));
            last_end = mat.end();
        }

        // 剩余部分：标红加粗
        if last_end < after.len() {
            let remain = &after[last_end..];
            *para = para
                .clone()
                .add_run(Run::new().add_text(remain).color("FF0000").bold());
        }
    } else {
        // 没有冒号，直接添加
        *para = para.clone().add_run(Run::new().add_text(text));
    }
}
