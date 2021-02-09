# md-design-doc
[![Build Status](https://travis-ci.com/toolbox-labo/md-design-doc.svg?branch=master)](https://travis-ci.com/toolbox-labo/md-design-doc) [![codecov](https://codecov.io/gh/toolbox-labo/md-design-doc/branch/master/graph/badge.svg?token=258Z1OJCOY)](https://codecov.io/gh/toolbox-labo/md-design-doc)

WIP

## Getting Started

### Prerequisites

- Rust toolchain (nightly and stable, default to nightly)

First, install Rust.

https://www.rust-lang.org/tools/install

```sh
# check Rust installation
$ cargo -v
```

Install nightly toolchain and set it to default.

```sh
$ rustup install nightly
$ rustup default nightly
```

- Rust components
  - clippy (for nightly)
  - rustfmt (for stable)

```sh
$ rustup component add clippy --toolchain nightly
$ rustup component add rustfmt --toolchain stable
```

- LLVM (for Windows)

For Windows, install [LLVM Pre-built binaries](https://releases.llvm.org/download.html#11.0.0) of Windows(32bit or 64bit).

### Run locally

- clone this repo and cd

```sh
$ git clone https://github.com/toolbox-labo/md-design-doc.git
$ cd md-design-doc/md_designer
```

- execute command to convert your `.md` into `.xlsx`

```sh
# your markdown file and rule file
$ cargo run --features excel -- [path(.md)] [rule path(.yml)]
# or example files
$ cargo run --features excel -- test.md test_rule.yml
```

Fow now, the output file name is same as input file name .

### Run test and check the code coverage
#### Test

```sh
$ cd md_designer
$ cargo test --features excel
```

#### Coverage report

If this is your first time, install required modules.

```sh
$ cargo install grcov rust-covfix
```

To generate the coverage report,

```sh
$ cd md_designer
$ bash coverage.sh
```

It will be created to `report` and you can see the whole coverage report by accessing `report/index.html` .

## Examples

### Parsing Rule

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

### Markdown Pattern

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

## Contributing
Contributions are what make the open source community such an amazing place to be learn, inspire, and create. Any contributions you make are greatly appreciated.

1. Fork the Project
1. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
1. Change the code
1. Format your code (`cargo +stable fmt`)
1. Run lint tool (clippy) , if errors, fix them (`cargo clippy --features excel -Z unstable-options -- -D warnings`)
1. Run test (`cargo test --features excel`)
1. Commit your Changes (`git commit`). **Please follow [Angular Commit Message Format](https://github.com/angular/angular/blob/master/CONTRIBUTING.md#-commit-message-format) for your commit message.**
1. Push to the Branch (`git push origin feature/AmazingFeature`)
1. Open a Pull Request