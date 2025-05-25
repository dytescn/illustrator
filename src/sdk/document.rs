use windows::Win32::System::{Com, Variant::VARIANT};
use wincom::{dispatch::ComObject, utils::VariantExt};
use windows::core::GUID;

use super::artboards::IllArtboards;
use super::struct_jpgoption::ExportOptionsJPEG;
use super::struct_pngoption::ExportOptionsPNG24;

pub struct IllDocument {
    disp:ComObject
}

impl IllDocument{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    // 属性
    pub fn get_fullname(&self)->String{

        let fullname_res = self.disp.get_property("fullname");
        if fullname_res.is_err() {
            return "".to_string();
        }
        let fullname = fullname_res.unwrap();
        let res_str = fullname.to_string();
        if res_str.is_err() {
            return "".to_string();
        }
        return res_str.unwrap();
    }  
    // 方法
    pub fn do_save(&self)->Option<()>{
        let opts = Vec::new();
        let hres = self.disp.invoke_method("save", opts);
        if hres.is_err() {
            return None;
        }
        return Some(());
    }  // readonly	Saves the document.
    pub fn do_saveas(&self,file_src:&str)->Option<()>{
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(file_src);
        opts.push(vart_file);
        let hres = self.disp.invoke_method("saveas", opts);
        if hres.is_err() {
            return None;
        }
        return Some(());
    }  // readonly	Saves the document with the specified save options.
    pub fn do_export_jpg(&self,img_src:&str,file_type:i32,opt:ExportOptionsJPEG)->Option<()>{
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(img_src);
        opts.push(vart_file);
        let vart_file_type = <VARIANT as VariantExt>::from_i32(file_type);
        opts.push(vart_file_type);
        // let save_opts = opt.to_variant();
        // opts.push(save_opts);
        let hres = self.disp.invoke_method("export", opts);
        if hres.is_err() {
            // println!("{:?}",hres.err());
            print!("zazaza");
            return None;
        }
        return Some(());
    }
    pub fn do_export_png(&self,img_src:&str,file_type:i32,opt:&ExportOptionsPNG24)->Option<()>{
        let mut opts = Vec::new();
        let vart_file = <VARIANT as VariantExt>::from_str(img_src);
        opts.push(vart_file);
        let vart_file_type = <VARIANT as VariantExt>::from_i32(file_type);
        opts.push(vart_file_type);
        let save_opts = opt.to_variant();
        opts.push(save_opts);
        let hres = self.disp.invoke_method("export", opts);
        if hres.is_err() {
            println!("zzzzzzzzzzzzz{:?}",hres.err());
            return None;
        }
        return Some(());
    }
    pub fn do_export_pdf(&self){

    }
    pub fn get_artboards(&self)->Option<IllArtboards>{
        let doc_res = self.disp.get_property("artboards");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let ivg_doc = IllArtboards::new(disp.clone());
                        return Some(ivg_doc);
                    }
                    Err(_e)=>{
                        return None
                    }
                }
            }
            Err(_)=>{
                return None
            }
        }
    }
    pub fn do_close(&self)->Option<()>{
        let opts = Vec::new();
        let hres = self.disp.invoke_method("close", opts);
        if hres.is_err() {
            return None;
        }
        return Some(());
    }
}
