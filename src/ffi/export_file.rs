use dyteslogs::logs::LogError;
use windows::Win32::System::Com;

use crate::sdk::application;
// 导出文件
pub fn export_ill_file(ver:&str,file_src:&str)->Option<()>{
    // 初始化系统
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = application::IllApplication::new(ver);
        let app = app_res.unwrap();
        let act_doc = app.get_activedocument().unwrap();
        let doc_file_path = act_doc.get_fullname();
        let parts: Vec<&str> = doc_file_path.split('.').collect();
        if parts.len()<2 {
            let res = act_doc.do_saveas(&file_src);
            if res.is_some() {
               return Some(());
            }
        }
        // 转移文件
        let res = std::fs::copy(doc_file_path, file_src);
        if res.is_err() {
            println!("{:?}",res.err());
            return None;
        }
        // 打开文件
        let res =  app.do_open(&file_src);
        if res.is_none() {
            return None;
        }
        let res = act_doc.do_close();
        if res.is_none() {
            return None;
        }
        let _ = Com::CoUninitialize();
        return Some(());

    }
}

pub fn save_ill_file(ver:&str,file_src:&str)->Option<()>{
    // 初始化系统
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = application::IllApplication::new(ver);
        let app = app_res.unwrap();
        let act_doc = app.get_activedocument().unwrap();
        let res = act_doc.do_saveas(&file_src);
        if res.is_none() {
            return None;
        }
        let _ = Com::CoUninitialize();
        return Some(());

    }
}


