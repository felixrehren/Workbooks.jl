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

# No setters for cells: only set at the workbook level
set!(S::Sheet, p::LocalPosition, f::AbstractString) = set!(S.wb, GlobalPosition(S.name,p), f)
Base.setindex!(S::Sheet, f::AbstractString, p::AbstractString) = set!(S.wb, GlobalPosition(S.name,p), f) # note the order: setindex!(object, value, position)
# todo: refresh formulas and values tables

## O U T P U T
const minPos = LocalPosition(1,1)
const maxPos = LocalPosition(typemax(UInt),typemax(UInt))
function hull(pp) 
    isempty(pp) && return LocalPosition(0,0)
    LocalPosition(maximum(getfield.(pp,:row)),maximum(getfield.(pp,:column)))
end

function Base.string(S::Sheet, field = :value)
    isempty(S.map.keys) && return "Empty sheet"
    R = map(x -> ismissing(x) ? missing : getfield(x,field), S[LocalPosition(1,1):hull(S.map.keys)])
    pretty_table(String, R, colNumAsStr.(1:size(R,2)), row_names=1:size(R,1))
end
Base.show(io::IO, S::Sheet) = Base.print(io, string(S))
