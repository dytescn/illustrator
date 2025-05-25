use windows::{core, Win32::System::Com};
use wincom::dispatch::ComObject;

pub struct ExportOptions {
    obj:ComObject
}

impl ExportOptions{
    pub fn new()-> Self {
       let data=  ComObject::new_from_name("CorelDRAW.StructExportOptions", core::GUID::zeroed()).unwrap();
       return  Self{
            obj:data
       };
    }
    pub fn from_disp(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, core::GUID::zeroed()).expect("init core error");
        Self{
            obj:obj
        }
    }
}
