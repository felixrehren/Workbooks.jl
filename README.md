# Workbooks.jl

Spreadsheets with the power of Julia

[![Build status (Github Actions)](https://github.com/sylvaticus/MyAwesomePackage.jl/workflows/CI/badge.svg)](https://github.com/sylvaticus/MyAwesomePackage.jl/actions)
[![codecov.io](http://codecov.io/github/sylvaticus/MyAwesomePackage.jl/coverage.svg?branch=main)](http://codecov.io/github/sylvaticus/MyAwesomePackage.jl?branch=main)

## Installation

```julia
julia> Pkg.add("JLX")
```


## Documentation

This package has the following aims:

1. Capture fundamental spreadsheet functionality
    - Variables are represented by `Cell`s; there are no other/global variables
    - A `Cell` is defined by its formula, calling functions that refer to other `Cell`s
    - Updating a `Cell` causes its dependents to update (if the graph of references contains no cycles!)
    - A workbook is a self-contained collection of sheets

2. Let the full power of Julia be used
    - Cells can have any value of any type, including any `Number`, an `Array` or a `Struct`
    - Any function can be called in a cell's formula (but no side-effects, please?)
    - The workbook comes with a "config" Julia file that imports modules and defines and custom functions
    - Standard tooling allows graphing, benchmarking, and other analysis

3. Other features
    - Simple file format
    - Simple interface in VS Code
    - Interoperability through CSV and a subset of Excel


## Related packages

* [XLSX.jl](https://github.com/felipenoris/XLSX.jl)