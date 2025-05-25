// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use illustrator::ffi::app;
    #[test]
    fn test_open_ill_info() {
        // C:\Users\dowell\Desktop\2222.cdr
        // let src = "C:\\Users\\dowell\\codes\\dytes\\cdrsdk\\cache\\123123.cdr".to_string();
        // let src ="C:\\Users\\dowell\\Desktop\\2222.cdr";
        
        let src ="C:\\Users\\Lenovo\\Desktop\\imgcache\\333.ai";
        // let res = app::do_cdr_app_open("测试".to_string(),src,"25".to_string());
        // println!("{:?}",res);
        // let info = export_file::export_ill_file("26",src);
        let res = app::open_illust_file("26",src);
        println!("{:?}",res);
    }
}