use windows::Win32::System::Com;
use crate::sdk::application::IllApplication; 
use crate::sdk::struct_jpgoption::ExportOptionsJPEG;

pub fn export_illust_cover(ver:&str,file_src:&str)->Option<()>{
    unsafe{
        // 初始化系统
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
    let act_doc = app.get_activedocument().unwrap();
    let boards = act_doc.get_artboards().unwrap();
    let res = boards.SetActiveArtboardIndex(0);
    let jpg_opt = ExportOptionsJPEG::new();
    jpg_opt.set_ArtBoardClipping(true);
    let res =  act_doc.do_export_jpg(file_src,1,jpg_opt);
    if res.is_none() {
        return None;
    }
    unsafe{
        let _ = Com::CoUninitialize();
    }
    return Some(());
}