# tsv2xlsx

![GitHub License](https://img.shields.io/github/license/azukichi67/tsv2xlsx)
![GitHub top language](https://img.shields.io/github/languages/top/azukichi67/tsv2xlsx)

## Introduction

converting tsv to xlsx.

## Usage

```
Usage:
  tsv2xlsx [flags]

Flags:
  -i, --input string          input tsv file path
  -o, --output string         output xlsx file path
  -f, --filter                set filter to header
  -c, --column-width string   change columns width (e.g. A:50,B100)
  -h, --help                  help for tsv2xlsx
```

`-i` & `-o` is required.

## Sample

```
tsv2xlsx -i ./sample.tsv -p ./output.xlsx -f -c "A:30,B:20"
```

## License

MIT license
