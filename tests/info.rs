// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use illustrator::ffi::{app, info};
    #[test]
    fn test_do_cdr_info() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr".to_string();
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        // println!("{:?}",res);
        let info = info::get_version_info("26");
        let info_str = app::get_illust_excute_path("26");   
        println!("{:?}",info);
        println!("{:?}",info_str);
    }
}