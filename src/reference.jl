## C E L L   P O S I T I O N S
const colsyntax = "([A-Z]+)"
const rowsyntax = "([1-9]+)"
const cellsyntax = colsyntax*rowsyntax
const cellregex = Regex("^"*cellsyntax*"\$")
const rangeregex = Regex("([1-9A-Z][0-9A-Z!]*:[1-9A-Z][0-9A-Z]*)")
const localsyntax = "([1-9A-Z][0-9A-Z:]*)"

const sheetsyntax = "([a-zA-Z_][0-9A-Za-z]*)"
const sheetregex = Regex("^"*sheetsyntax*"\$")

const globalposregex = Regex("^"*sheetsyntax * "!" * cellsyntax*"\$")
const globalrefregex = Regex("^"*sheetsyntax * "!" * localsyntax*"\$")
const refregex = Regex("^"*sheetsyntax)

function LocalPosition(r::AbstractString) # accept "A1"-syntax, convert to "A",1 syntax
    m = match(cellregex,r)
    @assert !isnothing(m) ("That's not a valid local reference: " * r)
    colStr = string(m.captures[1])
    rowNum = parse(UInt,m.captures[2])
    LocalPosition(colStr,rowNum)
end
function colStrAsNum(c::AbstractString) # accept "A", convert to 1
    colsbase26 = map(x -> Int(x)-64,collect(c))
    col = zero(UInt)
    l = length(colsbase26)
    for i in 1:l
        col += colsbase26[i]*26^(l-i)
    end
    col
end
function LocalPosition(colStr::AbstractString,rowNum::Number) # accept "A",1-syntax, convert to 1,1-syntax
    LocalPosition(rowNum,colStrAsNum(colStr))
end

# arithmetic of local positions
Base.isless(p::LocalPosition,q::LocalPosition) = (p.row < q.row) && (p.column < q.column)
Base.:-(p::LocalPosition,q::LocalPosition) = (p.row-q.row, p.column-q.column)
Base.:+(p::LocalPosition,q::LocalPosition) = (p.row+q.row, p.column+q.column)
Base.maximum(pp::Array{LocalPosition}) = LocalPosition(maximum.(getfield.(pp,[:row,:column])))

# display functions
function colNumAsStr(colNum)
    if colNum > 26
        colNumAsStr(div(colNum,26)) * colNumAsStr(rem(colNum,26))
    else
        string(Char(64+colNum))
    end
end
Base.string(p::LocalPosition) = colNumAsStr(p.column) * string(p.row)
Base.show(io::IO, p::LocalPosition) = Base.print(io, string(p))

# ranges
const cellrangeregex = Regex("^"*cellsyntax*":"*cellsyntax*"\$")
CellRange(p::AbstractString,q::AbstractString) = CellRange(LocalPosition(p),LocalPosition(q))
Base.string(r::CellRange) = string(r.start) * ":" * string(r.stop)
Base.show(io::IO, r::CellRange) = Base.print(io, string(r))

const columnrangeregex = Regex("^"*colsyntax*":"*colsyntax*"\$")
ColumnRange(a::AbstractString,b::AbstractString) = ColumnRange(colStrAsNum(a),colStrAsNum(b))
Base.string(r::ColumnRange) = colNumAsStr(r.start) * ":" * colNumAsStr(r.stop)
Base.show(io::IO, r::ColumnRange) = Base.print(io, string(r))

const rowrangeregex = Regex("^"*rowsyntax*":"*rowsyntax*"\$")
RowRange(a::AbstractString,b::AbstractString) = RowRange(parse(UInt,a),parse(UInt,b))
Base.string(r::RowRange) = string(r.start) * ":" * string(r.stop)
Base.show(io::IO, r::RowRange) = Base.print(io, string(r))

function LocalRange(s::AbstractString)
    m = match(cellrangeregex,s)
    isnothing(m) || return CelllRange(m.captures[1]*m.captures[2],m.captures[3]*m.captures[4])
    m = match(columnrangeregex,s)
    isnothing(m) || return ColumnRange(m.captures[1],m.captures[2])
    m = match(rowrangeregex,s)
    isnothing(m) || return RowRange(m.captures[1],m.captures[2])
    error("That's not a valid local range: " * s)
end
function LocalRef(s::AbstractString) # pure concatenation of LocalPosition and LocalRange ... better way?
    m = match(cellregex,r)
    if !isnothing(m)
        colStr = string(m.captures[1])
        rowNum = parse(UInt,m.captures[2])
        return LocalPosition(colStr,rowNum)
    end
    m = match(cellrangeregex,s)
    isnothing(m) || return CelllRange(m.captures[1]*m.captures[2],m.captures[3]*m.captures[4])
    m = match(columnrangeregex,s)
    isnothing(m) || return ColumnRange(m.captures[1],m.captures[2])
    m = match(rowrangeregex,s)
    isnothing(m) || return RowRange(m.captures[1],m.captures[2])
    error("That's not a valid local ref: " * s)
end

## G L O B A L
function GlobalPosition(s::AbstractString)
    m = match(globalposregex,s)
    isnothing(m) || return GlobalPosition(m.captures[1],LocalPosition(m.captures[2] * m.captures[3]))
    error("That's not a valid global position: " * s)
end
LocalPosition(gp::GlobalPosition) = gp.lref
GlobalRef(s::String, lref::LocalPosition) = GlobalPosition(s,lref)
GlobalRef(s::String, lref::LocalRangeType) = GlobalRange(s,lref)
function GlobalRef(s::AbstractString)
    m = match(globalrefregex,s)
    isnothing(m) || GlobalRef(m.captures[1], LocalRef(m.captures[2]))
    error("That's not a valid global reference: " * s)
end

Base.string(r::GlobalRef) = r.sheet * "!" * string(r.lref)
Base.show(io::IO, r::GlobalRef) = Base.print(io, string(r))


## P A R S I N G
# make sure that ranges are correctly grabbed by the Julia AST parser, by adding brackets around them
bracketedRanges(f::AbstractString) = replace(f, rangeregex => s"(\1)")

"""
    extractMatchFromSymbol(symbol)

Identifies and returns the string representing a reference to a position or range in the `symbol`
"""
function extractMatchFromSymbol(s::Symbol)
    t = string(s)
    m = match(globalrefregex,t)
    isnothing(m) || return m.match
    m = match(cellregex,t)
    isnothing(m) || return m.match
    nothing
end

function exprIsRange(e::Expr)
    (e.head == :call) &&
    (length(e.args) >= 3) &&
    (e.args[1] == :(:)) &&
    (isa(e.args[2],Symbol) == Symbol) &&
    (isa(e.args[3],Symbol) == Symbol)
end
"""
    extractRefsFromExpression(expression)

Identifies and returns the set of strings representing references to positions and ranges in the `expression`
"""
function extractRefsFromExpression(e::Expr)
    rr = String[]
    
    exprIsRange(e) && return string(e)

    m = match(cellregex,string(e.head))
    isnothing(m) || push!(rr,m.match)
    
    for (index,arg) in enumerate(e.args)
        if isa(arg,Symbol)
            m = match(cellregex,string(arg))
            isnothing(m) || push!(rr,m.match)
        elseif isa(arg,Expr)
            push!(rr,extractRefsFromExpression(arg)...)
        end
    end
    
    unique(rr)
end