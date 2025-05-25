// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use illustrator::ffi::export_preview::export_illust_preview;
    #[test]
    fn test_do_cdr_info() {        
        let src ="C:\\Users\\Lenovo\\Desktop\\imgcache\\aaa_";
        let info = export_illust_preview("26",src);
        println!("{:?}",info);
    }
}