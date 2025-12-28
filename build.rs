fn main() {
    if std::env::var("CARGO_CFG_WINDOWS").is_ok() {
        let mut res = winres::WindowsResource::new();
        res.set_icon("icon.ico");

        res.set_language(0x0804); // 中文 (简体)
        res.set("FileDescription", "剧本格式化工具");
        res.set("ProductName", "小牛马");
        res.set("OriginalFilename", "小牛马-剧本格式化工具.exe");
        res.set("LegalCopyright", "Copyright © 2024 Mick");

        res.compile().unwrap();
    }
}
