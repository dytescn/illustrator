use windows::Win32::System::{Com, Variant::VARIANT};
use wincom::{dispatch::ComObject, utils::VariantExt};
use windows::core::GUID;

pub struct IllPageItem {
    disp:ComObject
}

impl IllPageItem{
   pub fn new(disp:Com::IDispatch)->Self {
        let obj: ComObject =  ComObject::clone_from(disp, GUID::zeroed()).expect("init core error");
        Self{
            disp:obj
        }
    }
    
}
