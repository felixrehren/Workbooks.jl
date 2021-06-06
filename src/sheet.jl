## S H E E T
# Empty sheet
Sheet(name::String, wb::Workbook) = Sheet(name, wb,
    OrderedDict{LocalPosition,Cell}() # map
)

# Getters
get(S::Sheet, p::LocalPosition, default = missing) = Base.get(S.map, p, default)
get(S::Sheet, R::LocalRangeType) = get.([S],collect(R))
Base.getindex(S::Sheet, p::LocalRef) = get(S, p)
Base.getindex(S::Sheet, p::String) = get(S, LocalRef(p))
Base.getindex(S::Sheet, i, j) = get(S, LocalPosition(i,j))

# No setters for cells: only set at the workbook level
Base.setindex!(S::Sheet, f, p::Union{LocalPosition,AbstractString}) = setindex!(S.wb, f, GlobalPosition(S.name,p))
Base.setindex!(S::Sheet, f, i,j)   = setindex!(S.wb, f, GlobalPosition(S.name,LocalPosition(i,j)))
Base.setindex!(S::Sheet, f, xs...) = setindex!(S,string(f),xs...)

## O U T P U T
const minPos = LocalPosition(1,1)
const maxPos = LocalPosition(typemax(UInt),typemax(UInt))
function hull(pp) 
    isempty(pp) && return LocalPosition(0,0)
    LocalPosition(maximum(getfield.(pp,:row)),maximum(getfield.(pp,:column)))
end

function array(S::Sheet, fn = value)
    isempty(S.map.keys) && return nothing
    fn.(S[LocalPosition(1,1):hull(S.map.keys)])
end

function Base.string(S::Sheet, field = value)
    R = array(S,field)
    isnothing(R) && return "Empty sheet"
    pretty_table(String, R, colNumAsStr.(1:size(R,2)), row_names=1:size(R,1))
end
Base.show(io::IO, S::Sheet) = Base.print(io, string(S))

function Base.size(S::Sheet)
    h = hull(S.map.keys)
    (h.row,h.column)
end