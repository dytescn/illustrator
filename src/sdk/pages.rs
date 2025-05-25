use windows::Win32::System::Com;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::core::GUID;
pub struct IllPageItems {
    disp:ComObject
}

impl IllPageItems{
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

}