// 获取coreldraw的状态
#[cfg(test)]
mod tests {
    use illustrator::ffi::export_file;
    #[test]
    fn test_do_cdr_info() {        
        let src ="C:\\Users\\Lenovo\\Desktop\\imgcache\\333.ai";
        let info = export_file::export_ill_file("26",src);
        println!("{:?}",info);
    }
}