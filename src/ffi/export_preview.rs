use windows::Win32::System::Com;
use crate::sdk::{application::IllApplication, struct_pngoption::ExportOptionsPNG24};


// 保存封面图
pub fn export_illust_preview(ver:&str,preview_src:&str)->Option<usize>{
    unsafe {
        let _ = Com::CoInitialize(None);
    }
    let app_res = IllApplication::new(ver);
    if app_res.is_none() {
        return None;
    }
    let app = app_res.unwrap();
    let act_doc = app.get_activedocument().unwrap();
    let pages = act_doc.get_artboards().unwrap();
    // // 保存封面

    let pngopts = ExportOptionsPNG24::new();
    pngopts.set_ArtBoardClipping(true);
    let sum = pages.get_count();
    for i in  0..sum{
        let _ = pages.SetActiveArtboardIndex(i);
        let src = format!("{}{}_preview.png",preview_src,i+1);
        let res =  act_doc.do_export_png(&src,5,&pngopts);
        if res.is_none() {
            println!("{}",src);
        }
    }
    unsafe {
        let _ = Com::CoUninitialize();
    }        // 
    return Some(sum as usize);
}


// 保存封面图
pub fn export_illust_preview_sum(ver:&str)->Option<usize>{
    unsafe {
        let res = Com::CoInitialize(None);
        if res.is_err() {
            return None;
        }
    }
    let app_res = IllApplication::new(ver);
    if app_res.is_none() {
        return None;
    }
    let app = app_res.unwrap();
    let doc = app.get_activedocument().unwrap();
    let pages = doc.get_artboards().unwrap();
    // let cur_page = pages.get
    let sum = pages.get_count();
    unsafe {
        let _ = Com::CoUninitialize();
    }
    return Some(sum as usize);
}