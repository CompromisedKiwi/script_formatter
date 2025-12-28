#![windows_subsystem = "windows"]
mod formatter;

use std::fs;
use std::path::PathBuf;

use anyhow::Result;
use rfd::FileDialog;
use formatter::Formatter;

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

    let fmtr = Formatter::new(out_dir);
    let mut output_paths = Vec::new();
    for path in paths {
        let p = fmtr.formatting(path)?;
        output_paths.push(p);
    }
    Ok(output_paths.join("\n"))
}
