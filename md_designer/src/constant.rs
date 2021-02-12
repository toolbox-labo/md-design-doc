use lazy_static::lazy_static;

lazy_static! {
    pub static ref AUTO_INCREMENT_KEY: String = String::from("AUTOINCREMENT");
    pub static ref CUSTOM_PREFIX_KEY: String = String::from("CUSTOMPREFIX");
}
