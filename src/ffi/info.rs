use dyteslogs::logs::LogError;
use windows::Win32::System::Com;
use crate::sdk::application::IllApplication;
use windows::core;
pub fn get_version(ver:&str)->Option<String>{
    unsafe{
        // 初始化系统
        Com::CoInitialize(None).log_error("init com error").unwrap();
        let app_res = IllApplication::new(ver);
        match app_res{
            Some(app)=>{
                let ver = app.get_version();
                println!("{:?}",ver);
                return None;
            }
            None=>{
                return None;
            }
        }
        // Com::CoUninitialize();
    }
}

// 获取coreldraw
pub fn get_version_info(ver:&str)->Option<u32>{
    unsafe{
        // 初始化系统
        let com_res = Com::CoInitialize(None);
        if com_res.is_err() {
            println!("com init error");
            return None;
        }
    }
    let id_name = "Illustrator.Application.".to_string() + ver; // Can't use + with two &str
    let lpsz = core::HSTRING::from(id_name);
    let rclsid_res = unsafe { Com::CLSIDFromProgID(&lpsz) };
    if rclsid_res.is_err(){
        return None;
    }
    println!("{:?}",rclsid_res);
    let res_ver_n = ver.parse();
    if res_ver_n.is_err() {
        return None;
    }
    let ver_n = res_ver_n.unwrap();
    unsafe{
        Com::CoUninitialize();
    }
    return Some(ver_n);
}
