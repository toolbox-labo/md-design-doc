#![warn(rust_2018_idioms)]

use anyhow::Result;
use clap::{crate_authors, crate_description, crate_name, crate_version, App as ClapApp};

use md_designer::app::App;

fn main() -> Result<()> {
    // setup clap
    let _clap = ClapApp::new(crate_name!())
        .author(crate_authors!())
        .version(crate_version!())
        .about(crate_description!())
        .get_matches();

    let app = App::new(
        r#"
# Sheet Name
## Test Variation - 1
### Test Variation - 2-1
#### Test Variation - 3-1 [priority(Optional: # is low, None is High)]
* Test Description
  more lines...
  more lines...
- Test Procedure(1)
- Test Procedure(2)
- Test Procedure(3)
- more procedures...
- [ ] Confirmation item(1)
- [ ] Confirmation item(2)
- [ ] Confirmation item(3)
- [ ] more items...
"#,
    )?;

    Ok(())
}
