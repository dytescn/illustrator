use windows::Win32::System::Com;
use crate::sdk::application::IllApplication;

pub fn open_illust_file(ver: &str,file_path: &str)->Option<()> {
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            return None;
        }
        let app_res = IllApplication::new(ver);
        if app_res.is_none() {
            return None;
        }
        let app = app_res.unwrap();
        let res = app.do_open(file_path);
        if res.is_none() {
            return None;
        }
        let _ = Com::CoUninitialize();
        return Some(());
    }
}

pub fn get_documents_sum(ver: &str)->Option<usize>{
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            return None;
        }
        let app_res = IllApplication::new(ver);
        if app_res.is_none() {
            return None;
        }
        let app = app_res.unwrap();
        let docs = app.get_documents().unwrap();
        let sum = docs.get_count();
        let _ = Com::CoUninitialize();
        return Some(sum as usize);
    }
}

pub fn get_illust_excute_path(ver:&str)->Option<String>{
    unsafe{
        // 初始化系统
        let res =Com::CoInitialize(None);
        if res.is_err() {
            return None;
        }
        let app_res = IllApplication::new(ver);
        if app_res.is_none() {
            return None;
        }
        let app = app_res.unwrap();
        let exe_path = app.get_path();
        let exe_src = format!("{}Support Files\\Contents\\Windows\\Illustrator.exe",exe_path);
        let _ = Com::CoUninitialize();
        return Some(exe_src);
    }
}