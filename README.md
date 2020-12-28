# Workbooks.jl

Spreadsheets with the power of Julia

[![Build status (Github Actions)](https://github.com/sylvaticus/MyAwesomePackage.jl/workflows/CI/badge.svg)](https://github.com/sylvaticus/MyAwesomePackage.jl/actions)
[![codecov.io](http://codecov.io/github/sylvaticus/MyAwesomePackage.jl/coverage.svg?branch=main)](http://codecov.io/github/sylvaticus/MyAwesomePackage.jl?branch=main)

## Installation

```julia
julia> Pkg.add("Workbooks")
```


## Documentation

This package has the following aims:

1. **Capture fundamental spreadsheet functionality**
    - Variables are represented by `Cell`s; there are no other/global variables
    - A `Cell` is defined by its formula, potentially calling functions that refer to other `Cell`s
    - Updating a `Cell` causes its dependents to update (if the graph of references contains no cycles!)
    - A `Workbook` is a self-contained collection of `Sheet`s

2. **Let the full power of Julia be used**
    - `Cell`s can have any value of any type, including any `Number`, an `Array` or a `Struct`
    - Any function can be called in a `Cell`'s formula (but no side-effects, please?)
    - The `Workbook` comes with a "config" file that imports modules and defines and custom functions
    - Standard tooling can add plotting, benchmarking, and other analysis

3. **Other features**
    - Simple file format: zipped collection of sheets + config file
    - Simple interface in VS Code
    - Interoperability through CSV and a subset of Excel


## Related packages

* [XLSX.jl](https://github.com/felipenoris/XLSX.jl)