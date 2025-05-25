#![allow(dead_code,non_snake_case)]
use dyteslogs::logs::LogError;
use wincom::utils::VariantExt;
use windows::Win32::System::Com;
use windows::Win32::System::Variant::VARIANT;
use wincom::dispatch::ComObject;
use crate::sdk::types::*;

pub struct ExportOptionsPNG24 {
    obj:ComObject
}

impl ExportOptionsPNG24{
    pub fn new()-> Self {
       let data=  ComObject::new_from_name("Illustrator.ExportOptionsPNG24", APP_INTER_IID).expect("got app err");
       return  Self{
            obj:data
       };
    }
    pub fn from_disp(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, APP_INTER_IID).expect("init core error");
        Self{
            obj:obj
        }
    }
    // readonly	If true, the color profile is embedded in the document.
    pub fn embedColorProfile(&self){

    }
    // readonly	The download format to use.
    pub fn formatOptions(&self){

    }
    // readonly	The color to use to fill anti-aliased edges adjacent to transparent areas of the image. Default: white.
    pub fn matte(&self){

    }
    // readonly	The quality of the produced image.
    pub fn quality(&self){

    }
    // readonly	The number of scans. Valid only for progressive type JPEG files.
    pub fn scans(&self){

    }
    pub fn get_ArtBoardClipping(&self)->Option<bool>{
        let info =  self.obj.get_property("ArtBoardClipping").unwrap();
        let res = info.to_bool().unwrap();
        return Some(res);
    }
    pub fn set_ArtBoardClipping(&self,art_board_clipping:bool)->Option<bool>{
        let mut opts = Vec::new();
        let vart_bool = VARIANT::from_bool(art_board_clipping); 
        opts.push(vart_bool);
        let info =  self.obj.set_property("ArtBoardClipping",opts);
        if info.is_err() {
            return None;
        }
        return Some(true);
    }
    pub fn to_variant(&self)->VARIANT{
        // self.obj.
        self.obj.get_variant().log_error("got variant err:").unwrap()
    }
}