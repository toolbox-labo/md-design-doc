# md-design-doc
WIP

## Getting Started
- install Rust toolchain

https://www.rust-lang.org/tools/install

```
# check Rust installation
$ cargo -v
```

- clone this repo and cd

```
$ git clone https://github.com/toolbox-labo/md-design-doc.git
$ cd md-design-doc/md_designer
```

- execute command to convert your `.md` into `.xlsx`

```
# your markdown file
$ cargo run --features excel -- [path(.md)]
# or example file
$ cargo run --features excel -- test.md
```

Fow now, the output file name is `test.xlsx` .

## Markdown Pattern

WIP

```
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
```
