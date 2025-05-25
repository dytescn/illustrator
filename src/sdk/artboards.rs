use windows::Win32::System::{Com, Variant::VARIANT};
use wincom::{dispatch::ComObject, utils::VariantExt};
use windows::core::GUID;

pub struct IllArtboards {
    disp:ComObject
}

impl IllArtboards{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    
    pub fn get_count(&self)->i32{
        let res =  self.disp.get_property("Count").expect("got name err:");
        res.to_i32().unwrap()
    }
    pub fn SetActiveArtboardIndex(&self,index:i32)->Option<()>{
        let mut opts = Vec::new();
        let vart_bool = VARIANT::from_i32(index); 
        opts.push(vart_bool);
        let info =  self.disp.invoke_method("SetActiveArtboardIndex",opts);
        if info.is_err() {
            println!("set active artboard index err:{}",info.err().unwrap());
            return None;
        }
        return Some(());
    }
}