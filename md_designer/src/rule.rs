use yaml_rust::{YamlLoader, YamlEmitter};

pub struct Rule {
    doc: Doc,
}

impl Rule {
    pub fn marshal(input: &str) -> Self {
        let docs = YamlLoader::load_from_str(input).unwrap();
    }
}

struct Doc {
    rows: Vec<Row>,
}

struct Row {
    columns: Vec<Column>,
}

struct Column {
    title: String,
    auto_increment: bool,
    group: Option<Box<Group>>,
}

struct Group {
    title: String,
}
