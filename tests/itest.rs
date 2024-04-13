// 集成测试
use xlsx_batch_reader::{get_num_from_ord, get_ord_from_num, get_tuple_from_ord};

#[test]
pub fn test_ord_to_num(){
    assert!(get_num_from_ord("B3".as_bytes()).unwrap() == 2);
    assert!(get_num_from_ord("Z".as_bytes()).unwrap() == 26);
    assert!(get_num_from_ord("AB".as_bytes()).unwrap() == 28);

    
    assert!(get_ord_from_num(1).unwrap() == "A".to_string());
    assert!(get_ord_from_num(27).unwrap() == "AA".to_string());
    assert!(get_ord_from_num(37).unwrap() == "AK".to_string());


    assert!(get_tuple_from_ord("A1".as_bytes()).unwrap() == (1, 1));
    assert!(get_tuple_from_ord("B3".as_bytes()).unwrap() == (3, 2));
}
