// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use illustrator::ffi::export_cover::export_illust_cover;
    #[test]
    fn test_do_cdr_info() {        
        let src ="C:\\Users\\Lenovo\\Desktop\\imgcache\\333.jpg";
        let info = export_illust_cover("26",src);
        println!("{:?}",info);
    }
}