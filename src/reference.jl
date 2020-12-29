###  R E F E R E N C E S
# Scope is to recognise and manage local and global references to positions or ranges

## C E L L   P O S I T I O N S
const colsyntax = "([A-Z]+)"
const rowsyntax = "([1-9]+)"
const cellsyntax = colsyntax*rowsyntax
const cellregex = Regex("^"*cellsyntax*"\$")

# Recognise a cell in the same sheet
function _LocalPosition(r::AbstractString)  # accept "A1"-syntax, convert to "A",1 syntax
                                            # returns Union{LocalPosition,Nothing}
    m = match(cellregex,r)
    isnothing(m) && return nothing 
    colStr = string(m.captures[1])
    rowNum = parse(UInt,m.captures[2])
    LocalPosition(colStr,rowNum)
end
function LocalPosition(r::AbstractString)   # accept "A1"-syntax, convert to "A",1 syntax
                                            # returns LocalPosition or throws error
    s = _LocalPosition(r)
    @assert !isnothing(s) ("That's not a valid local reference: " * r)
    s
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
Base.:-(p::LocalPosition,q::LocalPosition) = LocalPosition(p.row-q.row, p.column-q.column)
Base.:+(p::LocalPosition,q::LocalPosition) = LocalPosition(p.row+q.row, p.column+q.column)
Base.maximum(pp::Array{LocalPosition}) = LocalPosition(maximum(getfield.(pp,:row)),maximum(getfield.(pp,:column)))
Base.minimum(pp::Array{LocalPosition}) = LocalPosition(minimum(getfield.(pp,:row)),minimum(getfield.(pp,:column)))

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


## C E L L   R A N G E S
const rangeregex = Regex("([1-9A-Z][0-9A-Z!]*:[1-9A-Z][0-9A-Z]*)")

# cell ranges
const cellrangeregex = Regex("^"*cellsyntax*":"*cellsyntax*"\$")
CellRange(p::AbstractString,q::AbstractString) = CellRange(LocalPosition(p),LocalPosition(q))
Base.:(:)(p::LocalPosition,q::LocalPosition) = CellRange(p,q)
Base.string(r::CellRange) = string(r.start) * ":" * string(r.stop)
Base.show(io::IO, r::CellRange) = Base.print(io, string(r))

function Base.collect(R::CellRange)
    m = min(R.start,R.stop)
    r = m.row
    c = m.column
    M = max(R.start,R.stop)
    delta = M - m
    I = delta.row + 1
    J = delta.column + 1
    pp = Array{LocalPosition}(undef,I,J) # empty array
    for i in 1:I, j in 1:J
        pp[i,j] = LocalPosition(r+i-1,c+j-1)
    end
    pp
end
Base.collect(p::LocalPosition) = [p]
expanded_refs(rr) = vcat(collect.(rr)...)

# larger ranges --> not really there yet
const columnrangeregex = Regex("^"*colsyntax*":"*colsyntax*"\$")
ColumnRange(a::AbstractString,b::AbstractString) = ColumnRange(colStrAsNum(a),colStrAsNum(b))
Base.string(r::ColumnRange) = colNumAsStr(r.start) * ":" * colNumAsStr(r.stop)
Base.show(io::IO, r::ColumnRange) = Base.print(io, string(r))

const rowrangeregex = Regex("^"*rowsyntax*":"*rowsyntax*"\$")
RowRange(a::AbstractString,b::AbstractString) = RowRange(parse(UInt,a),parse(UInt,b))
Base.string(r::RowRange) = string(r.start) * ":" * string(r.stop)
Base.show(io::IO, r::RowRange) = Base.print(io, string(r))


## L O C A L   R E F E R E N C E S
const localsyntax = "([1-9A-Z][0-9A-Z:]*)"

function _LocalRange(s::AbstractString) # returns Union{LocalRangeType,Nothing}
    m = match(cellrangeregex,s)
    isnothing(m) || return CellRange(m.captures[1]*m.captures[2],m.captures[3]*m.captures[4])
    m = match(columnrangeregex,s)
    isnothing(m) || return ColumnRange(m.captures[1],m.captures[2])
    m = match(rowrangeregex,s)
    isnothing(m) || return RowRange(m.captures[1],m.captures[2])
    nothing
end
function LocalRange(s::AbstractString) # returns LocalRangeType or throws error
    r = _LocalRange(s)
    @assert !isnothing(r) ("That's not a valid local range: " * s)
    r
end

function _LocalRef(s::AbstractString) # returns Union{LocalRef,Nothing}
    t = _LocalPosition(s)
    isnothing(t) || return t
    u = _LocalRange(s)
    isnothing(u) || return u
    nothing
end
function LocalRef(s::AbstractString) # returns LocalRef or throws error
    r = _LocalRef(s)
    @assert !isnothing(r) ("That's not a valid local ref: " * s)
    r
end


## G L O B A L   P O S I T I O N
const sheetsyntax = "([a-zA-Z_][0-9A-Za-z]*)"
const sheetregex = Regex("^"*sheetsyntax*"\$")

const globalposregex = Regex("^"*sheetsyntax * "!" * cellsyntax*"\$")
const globalrefregex = Regex("^"*sheetsyntax * "!" * localsyntax*"\$")
const refregex = Regex("^"*sheetsyntax)

# "Sheet1", "A1"-syntax
GlobalPosition(s::AbstractString,p::AbstractString) = GlobalPosition(s, LocalPosition(p))
function _GlobalPosition(s::AbstractString)
    m = match(globalposregex,s)
    isnothing(m) || return GlobalPosition(m.captures[1],LocalPosition(m.captures[2] * m.captures[3]))
    nothing
end
function GlobalPosition(s::AbstractString)
    t = _GlobalPosition(s)
    @assert !isnothing(t) ("That's not a valid global position: " * s)
    t
end
LocalPosition(gp::GlobalPosition) = gp.lref


## G L O B A L   R E F E R E N C E
GlobalRef(s::AbstractString, lref::LocalPosition) = GlobalPosition(s,lref)

function Base.:(:)(p::GlobalPosition,q::GlobalPosition)
    @assert (p.sheet == q.sheet) ("Cannot range between cells of two different sheets!: " * string(p) * " and " * string(q))
    GlobalRange(p.sheet, CellRange(p.lref,q.lref))
end
Base.collect(R::GlobalRange) = GlobalPosition.(R.sheet, collect(R.lref))
GlobalRef(s::AbstractString, lref::LocalRangeType) = GlobalRange(s,lref)

GlobalRef(s::AbstractString, lref::AbstractString) = GlobalRef(s, LocalRef(lref))
function _GlobalRef(s::AbstractString)
    m = match(globalrefregex,s)
    isnothing(m) || return GlobalRef(m.captures[1], LocalRef(m.captures[2]))
    nothing
end
function GlobalRef(s::AbstractString)
    t = _GlobalRef(s)
    @assert !isnothing(t) ("That's not a valid global reference: " * s)
    t
end
function GlobalisedRef(s::AbstractString, default_Sheet::String)
    t = _GlobalRef(s)
    isnothing(t) || return t
    u = LocalRef(s)
    isnothing(u) || return GlobalRef(default_Sheet,u)
    error("That's not a reference: " * s)
end

# display functions
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
    isa(e.args[2],Symbol) &&
    isa(e.args[3],Symbol)
end
"""
    extractRefsFromExpression(expression)

Identifies and returns the set of strings representing references to positions and ranges in the `expression`
"""
function extractRefsFromExpression(e::Expr)
    rr = String[]
    
    exprIsRange(e) && return [string(e)]

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