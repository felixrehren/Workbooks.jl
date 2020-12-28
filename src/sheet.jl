## S H E E T
# Empty sheet
Sheet(name::String, wb::Workbook) = Sheet(name, wb,
    OrderedDict{LocalPosition,Cell}() # map
)

# Getters
function get(S::Sheet, p::LocalPosition, default = missing)
    Base.get(S.map, p, default)
end
function get(S::Sheet, R::LocalRef) # if LocalRef but not LocalPosition --> LocalRange
    vv = Array{Any}(undef,(R.stop - R.start)...)
    for c in 1:size(vv,2), r in 1:size(vv,1)
        vv[r,c] = get(S,LocalPosition(start.row+r,start.column+c))
    end
    vv
end
Base.getindex(S::Sheet, p::LocalRef) = get(S, p)
Base.getindex(S::Sheet, p::String) = get(S, LocalPosition(p))

# No setters for cells: only set at the workbook level
function set!(S::Sheet, p::LocalPosition, f::AbstractString)
    set!(S.wb, GlobalPosition(S.name,p), f)
end
# todo: refresh formulas and values tables

## O U T P U T
const minPos = LocalPosition(1,1)
const maxPos = LocalPosition(typemax(UInt),typemax(UInt))
function hull(pp) 
    isempty(pp) && return LocalPosition(0,0)
    LocalPosition(maximum(getfield.(pp,:row)),maximum(getfield.(pp,:column)))
end

function Range(S::Sheet, start::LocalPosition = minPos, stop::LocalPosition = hull(keys(S.map)), field = :value)
    diff = stop-start
    r = Array{Any}(undef,(1 .+ diff)...)
    for col = 0:diff[2], row = 0:diff[1]
        c = get(S,LocalPosition(start.row+row,start.column+col))
        r[1+row,1+col] = ismissing(c) ? missing : getfield(c,field)
    end
    r
end

function Base.string(S::Sheet, field = :value)
    isempty(S.map.keys) && return "Empty sheet"
    R = Range(S, LocalPosition(1,1), hull(S.map.keys), field)
    pretty_table(String, R, colNumAsStr.(1:size(R,2)), row_names=1:size(R,1))
end
Base.show(io::IO, S::Sheet) = Base.print(io, string(S))
