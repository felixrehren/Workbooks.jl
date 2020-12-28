## W O R K B O O K
# Empty workbook
Workbook() = Workbook(
    Dict{String,Sheet}(), # sheets
    OrderedDict{GlobalRef, Cell}(), # map of positions to cells
    SimpleDiGraph(), # mapping dependencies
    Int[] # topological ordering
)

# getters
# get a cell
get(wb::Workbook, p::GlobalPosition, default::Any = EmptyCell(p)) = Base.get(wb.map, p, default)
get!(wb::Workbook, p::GlobalPosition, default::Cell = EmptyCell(p)) = Base.get!(wb.map, p, default)
get!(wb::Workbook, r::GlobalRange, default::Cell = EmptyCell(p)) = Base.get!(wb[r.sheet], r.lref, default)

# get a sheet
Base.get(wb::Workbook, s::AbstractString) = getindex(wb, s)
function Base.getindex(wb::Workbook, s::AbstractString, default = missing)
    m = match(sheetregex,s)
    if !isnothing(m)
        if !(s in keys(wb.sheets))
            if ismissing(default)
                wb.sheets[s] = Sheet(s, wb)
            else
                @assert default.wb === wb "Sheet must belong to this workbook"
                wb.sheets[s] = default
            end
        end
        return wb.sheets[s]
    else
        return get(wb, GlobalPosition(s))
    end
end

# calculation agents
# calculate the value of the cell based on current state of sheet
refresh!(p::AbstractString, wb::Workbook) = refresh!(GlobalPosition(p), wb)
refresh!(p::GlobalPosition, wb::Workbook) = refresh!(get(wb,p), wb)
function refresh!(_::ConstCell,_::Workbook) end # nothing to do
function refresh!(c::DynamicCell,wb::Workbook)
    c.value = c.func(getfield.(broadcast(p -> wb.map[p],c.ancestors),:value)...)
end
function refresh_latest!(_::ConstCell,_::Workbook) end
function refresh_latest!(c::DynamicCell,wb::Workbook)
    c.value = Base.invokelatest(c.func, getfield.(broadcast(p -> wb.map[p],c.ancestors),:value)...)
end

function orderedFromSuperset(xx,xxx)
    xx[sortperm(findfirst.(isequal.(xx),[xxx]))]
end
linear_location(c::Cell, wb::Workbook) = linear_location(c.position, wb)
linear_location(p::GlobalPosition, wb::Workbook) = findfirst(isequal(p),wb.map.keys)
function cascade(p::GlobalPosition,wb::Workbook)
    loc = linear_location(p,wb)
    refreshNeeded = orderedFromSuperset(neighbors(wb.graph, loc),wb.order)
    for key in refreshNeeded
        refresh!(wb.map[wb.map.keys[key]],wb)
        push!(refreshNeeded, orderedFromSuperset(neighbors(wb.graph, key),wb.order)...) # add cells that are further downstream to the list
    end
end
# get the cell into the overall weave of the thing
function extendgraph(_::ConstCell, _::Workbook) end # nothing to do
function extendgraph(c::DynamicCell, wb::Workbook)
    # basically: make sure cell's ancestors exist & are recorded
    
    # create nonexistent ancestors
    for anc in c.ancestors # iterating over GlobalPositions
        ismissing(get(wb, anc, missing)) && set!(wb, EmptyCell(anc))
    end
    # extend graph
    incrementalNodes = length(wb.map) - nv(wb.graph)
    add_vertices!(wb.graph,incrementalNodes)
    # add edges
    loc = linear_location(c.position,wb)
    for anc in c.ancestors # iterating over GlobalPositions
        ancloc = findfirst(isequal(anc),wb.map.keys)
        add_edge!(wb.graph, ancloc, loc) # ancestor -> descendant
        if is_cyclic(wb.graph) || has_self_loops(wb.graph)
            rem_edge!(wb.graph, ancloc, loc) # let's undo that
            error("There should not be any cycles ... now what?")
        end
    end
    
    # update total ordering
    wb.order = topological_sort_by_dfs(wb.graph)    
end

# write a cell
function set!(wb::Workbook, p::Union{AbstractString,GlobalPosition}, f::AbstractString)
    set!(wb::Workbook, Cell(p,f))
end

function setcellbasics(wb::Workbook, c::Cell)
    p = c.position
    # add the cell to the wb
    push!(wb.map, p => c)
    # add the cell to the graph
    is_new_node = length(wb.map) - nv(wb.graph) # integer
    add_vertices!(wb.graph,is_new_node)

    # take care of the sheet
    S = wb[p.sheet]
    # add the cell to the sheet
    lp = p.lref
    push!(S.map, lp => c)

    return Bool(is_new_node)
end
function cleanup(wb::Workbook, c::Cell)
    loc = linear_location(c, wb)
    for anc in vertices(wb.graph)
        rem_edge!(wb.graph, anc, loc) # many false positives
    end
end
function set!(wb::Workbook, c::Cell)
    is_new_node = setcellbasics(wb,c)
    is_new_node || cleanup(wb,c)
    extendgraph(c,wb)
    refresh_latest!(c,wb)
    # update downstream
    is_new_node || cascade(c.position,wb) # only if an existing cell has been replaced
end

## O U T P U T
const minPos = LocalPosition(1,1)
const maxPos = LocalPosition(typemax(UInt),typemax(UInt))

function Base.string(wb::Workbook) 
    if isempty(wb.sheets)
        s = "Empty workbook"
    else
        s = "Workbook with " 
        s *= string(length(wb.sheets)) * " sheet" * ((length(wb.sheets) > 1) ? "s" : "")
        s *= ": " * join(string.(keys(wb.sheets)),", "," and ")
    end
    s
end
Base.show(io::IO, wb::Workbook) = Base.print(io, string(wb))

### you got this far last time ... ###

# todo: 1) find name; 2) convert from real XL; 3) write/read file; 4) develop interface