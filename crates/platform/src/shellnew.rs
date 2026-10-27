//! Explorer-style "New" menu entries, read from the registry the same way
//! Explorer does it: `.<ext>\ShellNew` with a `FileName` (template to copy)
//! or `NullFile` (an empty file is valid) value.

use std::path::PathBuf;

use windows_core::PCWSTR;

use crate::bindings::{RegGetValueW, HKEY};
use crate::wide::to_wide;

const RRF_RT_REG_SZ: u32 = 2;

/// `HKEY_CLASSES_ROOT`: `HKCU\Software\Classes` merged with the machine keys,
/// which is where `.ext\ShellNew` lives. `bindings` only generates the two
/// real root constants.
const HKEY_CLASSES_ROOT: HKEY = HKEY(-2_147_483_648i32 as _);

fn reg_sz(root: HKEY, subkey: &str, value: &str) -> Option<String> {
    let subkey = to_wide(subkey);
    let value = to_wide(value);
    let mut buf = [0u16; 1024];
    let mut len = (buf.len() * 2) as u32;
    // SAFETY: strings outlive the call; `buf` is a live out buffer sized above.
    let status = unsafe {
        RegGetValueW(
            root,
            PCWSTR(subkey.as_ptr()),
            PCWSTR(value.as_ptr()),
            RRF_RT_REG_SZ,
            None,
            Some(buf.as_mut_ptr().cast()),
            Some(&mut len),
        )
    };
    (status.0 == 0).then(|| {
        let end = buf.iter().position(|&c| c == 0).unwrap_or(buf.len());
        String::from_utf16_lossy(&buf[..end])
    })
}

/// A "New" menu entry: display name, extension (with dot), and how to create
/// the initial file.
pub struct NewEntry {
    pub name: String,
    pub ext: String,
    /// Copy this template file when creating; `None` = an empty file is valid.
    pub template: Option<PathBuf>,
}

/// Curated extensions worth offering, with display names. Explorer offers one
/// entry per registered ShellNew key; a fixed candidate list keeps the menu
/// short and the registry reads cheap.
const CANDIDATES: [(&str, &str); 6] = [
    ("docx", "Word 文档"),
    ("xlsx", "Excel 工作表"),
    ("pptx", "PowerPoint 演示文稿"),
    ("md", "Markdown 文件"),
    ("rtf", "RTF 文档"),
    ("bmp", "位图图像"),
];

/// Entries the installed programs actually registered, in `CANDIDATES` order.
/// Office registers ShellNew under the extension's ProgId subkey
/// (`.docx\Word.Document.12\ShellNew`) rather than `.docx\ShellNew`, so both
/// are consulted, ProgId first as Explorer does.
pub fn entries() -> Vec<NewEntry> {
    let mut out = Vec::new();
    for (ext, name) in CANDIDATES {
        let ext_key = format!(".{ext}");
        let mut subkeys = vec![format!("{ext_key}\\ShellNew")];
        if let Some(progid) = reg_sz(HKEY_CLASSES_ROOT, &ext_key, "").filter(|s| !s.is_empty()) {
            subkeys.insert(0, format!("{ext_key}\\{progid}\\ShellNew"));
        }
        for subkey in subkeys {
            let file = reg_sz(HKEY_CLASSES_ROOT, &subkey, "FileName").filter(|s| !s.is_empty());
            let null_file = reg_sz(HKEY_CLASSES_ROOT, &subkey, "NullFile").is_some();
            let data = reg_sz(HKEY_CLASSES_ROOT, &subkey, "Data").is_some();
            let template = file.map(|f| PathBuf::from(expand_env(&f)));
            if template.is_some() || null_file || data {
                out.push(NewEntry {
                    name: name.to_string(),
                    ext: format!(".{ext}"),
                    template,
                });
                break;
            }
        }
    }
    out
}

/// Expand the `%VAR%` prefixes template paths actually use; anything else is
/// returned unchanged.
fn expand_env(s: &str) -> String {
    let lower = s.to_ascii_lowercase();
    for (var, val) in [("%userprofile%", "USERPROFILE"), ("%appdata%", "APPDATA")] {
        if let Some(rest) = lower.strip_prefix(var) {
            if let Ok(base) = std::env::var(val) {
                return format!("{base}{}", &s[s.len() - rest.len()..]);
            }
        }
    }
    s.to_string()
}
