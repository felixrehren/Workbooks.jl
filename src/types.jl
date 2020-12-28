## P O S I T I O N S   A N D   R E F E R E N C E S
"""
Wrapper for cell and range references
"""
abstract type JLXref end
"""
A cell or range reference within the same sheet
"""
abstract type LocalRef <: JLXref end # within a single unnamed sheet
"""
A cell or range reference that specifies the sheet as well as the local position
"""
abstract type GlobalRef <: JLXref end # with reference to the sheet

struct LocalPosition <: LocalRef
    row::UInt
    column::UInt
end
struct CellRange <: LocalRef
    start::LocalPosition
    stop::LocalPosition
end
struct ColumnRange <: LocalRef
    start::UInt
    stop::UInt
end
struct RowRange <: LocalRef
    start::UInt
    stop::UInt
end
const LocalRangeType = Union{CellRange,ColumnRange,RowRange}

struct GlobalRange <: GlobalRef
    sheet::String
    lref::LocalRangeType
end
struct GlobalPosition <: GlobalRef
    sheet::String
    lref::LocalPosition
end
struct SheetRef <: GlobalRef
    sheet::String
end
const Position = Union{LocalPosition,GlobalPosition}

## C E L L S
abstract type JLXstyle end # not yet done

abstract type Cell end
mutable struct ConstCell <: Cell
    position::GlobalPosition
    value::Any
end
mutable struct DynamicCell <: Cell
    position::GlobalPosition
    formula::String
    func::Function
    ancestors::Array{GlobalRef}
    value::Any
end
# mutable struct CloneCell <: Cell
#     position::GlobalPosition
#     cloneOf::DynamicCell
#     ancestors::Array{GlobalPosition}
#     value::Any
#     #style::JLXstyle
# end

abstract type AbstractSheet end
abstract type AbstractWorkbook end

## S H E E T S
mutable struct Sheet <: AbstractSheet
    name::String
    wb::AbstractWorkbook # its "owner"
    map::OrderedDict{LocalPosition, Cell} # mapping positions to cells
end

## W O R K B O O K
mutable struct Workbook <: AbstractWorkbook
    sheets::Dict{String,AbstractSheet}
    map::OrderedDict{GlobalPosition, Cell} # mapping positions to cells
    graph::SimpleDiGraph # directed edge A -> B means A's value is an input to B's formula
    order::Vector{Int} # topological ordering, telling us how to cascade computation (using map)
end
