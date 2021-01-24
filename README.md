# md-design-doc
WIP

## Getting Started
- install Rust toolchain

https://www.rust-lang.org/tools/install

```
# check Rust installation
$ cargo -v
```

- install LLVM (for Windows)

For Windows, install [LLVM Pre-built binaries](https://releases.llvm.org/download.html#11.0.0) of Windows(32bit or 64bit).

- clone this repo and cd

```
$ git clone https://github.com/toolbox-labo/md-design-doc.git
$ cd md-design-doc/md_designer
```

- execute command to convert your `.md` into `.xlsx`

```
# your markdown file and rule file
$ cargo run --features excel -- [path(.md)] [rule path(.yml)]
# or example files
$ cargo run --features excel -- test.md test_rule.yml
```

Fow now, the output file name is `test.xlsx` .

## Custom Parsing Rule

WIP

```yml
# TODO: general settings
# general:
#   copyright: hogehoge
#   prefix: IT
doc:
  blocks:
    - block:
      - column: No
        isNum: true
      - group: Variation
        columns:
          - column: Variation 1
            md: Heading2
          - column: Variation 2
            md: Heading3
          - column: Variation 3
            md: Heading4
          - column: Variation 4
            md: Heading5
          - column: Variation 5
            md: Heading6
          - column: Variation 6
            md: Heading7
          - column: Variation 7
            md: Heading8
      - column: Description
        md: List
      # TODO: support variable list patterns
      #   customPrefix: "*"
      # - column: Procedure
      #   md: List
      # - column: Confirmation Items
      #   md: TaskList
    # TODO: multi blocks support
    # - block:
    #   - column: another block's column 1
    #     md: Heading2
    #   - column: another block's column 2
    #     md: Heading3
```

## Markdown Pattern

WIP

```markdown
# Sheet Name
## Test Variation - 1
### Test Variation - 2-1
#### Test Variation - 3-1
* Test Description
  more lines...
  more lines...
```