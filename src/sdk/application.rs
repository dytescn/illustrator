use dyteslogs::logs::LogError;
use wincom::dispatch::ComObject;
use wincom::utils::VariantExt;
use windows::Win32::System::Variant::VARIANT;
use super::document::IllDocument;
use super::documents::IllDocuments;
use windows::core::GUID;
pub struct IllApplication {
    disp:ComObject
}

impl IllApplication {
    pub fn new(ver:&str)-> Option<Self> {
       let app_name = format!("Illustrator.application.{}", ver);
       let res_data=  ComObject::new_from_name(&app_name, GUID::zeroed());
       match res_data {
           Some(data)=>{
                return Some(Self{
                    disp:data
                });
           }
           None=>{
                return None;
           }
       }
    }

    pub fn get_version(&self)->String{
        let app_vers = self.disp.get_property("version").expect("got err");

        app_vers.to_string().log_error("got error").unwrap()
    }
    // 方法
    pub fn get_path(&self)->String{
        let app_path = self.disp.get_property("path").expect("got error");
        app_path.to_string().log_error("got error").unwrap()
    }

    pub fn get_activedocument(&self)->Option<IllDocument>{
        let doc_res = self.disp.get_property("ActiveDocument");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let ivg_doc = IllDocument::new(disp.clone());
                        return Some(ivg_doc);
                    }
                    Err(_)=>{
                        return None
                    }
                }
            }
            Err(_)=>{
                return None
            }
        }
    }
    
    pub fn get_documents(&self)->Option<IllDocuments>{
        let doc_res = self.disp.get_property("documents");
        match doc_res {
            Ok(doc)=>{
                match doc.to_idispatch(){
                    Ok(disp)=>{
                        let ivg_doc = IllDocuments::new(disp.clone());
                        return Some(ivg_doc);
                    }
                    Err(_)=>{
                        return None
                    }
                }
            }
            Err(_)=>{
                return None
            }
        }
    }
    // 方法
    pub fn do_open(&self,file_src:&str)->Option<()>{
        let mut opts = Vec::new();
        let vart_src: VARIANT = VARIANT::from_str(file_src);
        opts.push(vart_src);
        let hres = self.disp.invoke_method("open", opts);
        if hres.is_err() {
            return None;
        }
            return  Some(());
    }
    pub fn visible(){

    }

    // 执行退出
    pub fn do_quit(&self)->bool {
       let res = self.disp.invoke_method("quit", vec![]);
       match res {
           Ok(_)=>{
                return true;
           }
           Err(e)=>{
                println!("{:?}",e);
                return false;
           }
       }
    }
}

