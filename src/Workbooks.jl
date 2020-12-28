module Workbooks

using DataStructures, LightGraphs
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

end # module

# todo: 
# 0.5) make it work with ranges! e.g. sum(A1:B3)
# 1) develop interface
#   a) convert a sheet to CSV (just the formulas in a table)
#   b) read a CSV to make a sheet
#   c) make a workbook from a collection of sheets
# 2) convert from real XL
# 3) think about what kind of formulas are allowed. Side effects? global variables?

# you have to change DynamicCell.ancestors; it should not just take GlobalPositions, but also GlobalRanges. A GlobalRange should be replaced with an array of values prior to execution!

# maybe swap it so that ConstCells have formulas rather than values, and create a "values" accessor function that works for all sorts of cells and refs (incl. positions and ranges!)