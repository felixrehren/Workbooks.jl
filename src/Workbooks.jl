module Workbooks

using DataStructures, LightGraphs
using ZipFile, DelimitedFiles
using PrettyTables, GraphPlot

export plusTwo, greet
export Workbook, Sheet, Cell

greet() = print("Hello World!")
plusTwo(x) = x + 2

include("types.jl")
include("reference.jl")
include("cell.jl")
include("sheet.jl")
include("workbook.jl")
include("io.jl")

end # module

# todo: 
# 1) File IO -- improve and make robust, including more testing
# 2) check cross-sheet references
# 3) convert from real XL
# 
# larger questions:
# * work out these `Project.toml` and `include.jl` functionalities. What if you have two workbooks loaded?
# * think about what kind of formulas are allowed. Side effects? global variables?