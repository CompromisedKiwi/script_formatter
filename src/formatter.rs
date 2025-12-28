use std::fs;
use std::path::{Path, PathBuf};

use docx_rs::*;
use anyhow::Result;
use lazy_regex::{Lazy, regex, Regex};

pub struct Formatter {
    output_dir: PathBuf,
}

impl Formatter {
    pub fn new(output_dir: PathBuf) -> Self {
        Self { output_dir }
    }

    pub fn formatting(&self, input_path: &Path) -> Result<String> {
        // Load doc
        let file_bytes = fs::read(input_path)?;
        let mut reader = read_docx(&file_bytes)?;

        // New doc
        // Set default font to 宋体（正文）, size 小四（12磅）
        let mut new_doc = Docx::new()
            .default_fonts(RunFonts::new().east_asia("宋体").ascii("Times New Roman"))
            .default_size(24); // Unit: half point

        new_doc = new_doc.add_style(Self::get_style_h3());

        // Extract lines
        for child in reader.document.children.drain(..) {
            // Process paragraph only
            let DocumentChild::Paragraph(para) = child else {
                continue;
            };
            let raw = Self::get_text_with_breaks(&para);
            let lines = raw.split('\n').collect::<Vec<&str>>();
            for line in lines {
                if line.is_empty() {
                    continue;
                }
                let new_para = Self::process_line(line.to_string());
                new_doc = new_doc.add_paragraph(new_para);
            }
        }

        // Save new doc
        let file_name = input_path
            .file_name()
            .ok_or_else(|| anyhow::anyhow!("无法获取文件名"))?
            .to_str()
            .ok_or_else(|| anyhow::anyhow!("文件名包含无效UTF-8字符"))?;
        let output_path = self.output_dir.join(file_name);
        let out_file = fs::File::create(&output_path)?;
        new_doc.build().pack(out_file)?;

        Ok(output_path
            .to_str()
            .ok_or_else(|| anyhow::anyhow!("文件名包含无效UTF-8字符"))?
            .to_string())
    }

    fn get_style_h3() -> Style {
        Style::new("Heading3", StyleType::Paragraph)
            .name("heading 3")
            .size(32) // 三号字
            .bold()
            .line_spacing(
                LineSpacing::new()
                    .before(260)
                    .after(260)
                    .line(413)
                    .line_rule(LineSpacingType::Auto),
            )
            .fonts(RunFonts::new().east_asia("宋体").ascii("Times New Roman"))
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

    fn process_line(mut text: String) -> Paragraph {
        // Replace @ with △ in the beginning
        text = text.replace('@', "△");

        // Create new paragraph
        let mut new_para = Paragraph::new();

        // Set to Heading3 style
        static RE_HEADING3: &Lazy<Regex> = regex!(r"^(第.*集|人物.*|[0-9].*)$");
        if RE_HEADING3.is_match(&text) {
            return new_para
                .add_run(Run::new().add_text(text))
                .style("Heading3");
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
}
