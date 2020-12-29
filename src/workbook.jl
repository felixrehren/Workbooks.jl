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
get(wb::Workbook, R::GlobalRange) = get.([wb], collect(R))
get!(wb::Workbook, p::GlobalPosition, default::Cell = EmptyCell(p)) = Base.get!(wb.map, p, default)
get!(wb::Workbook, R::GlobalRange) = get!.([wb], collect(R))

# get a sheet
Base.get(wb::Workbook, s::AbstractString) = getindex(wb, s)
function Base.getindex(wb::Workbook, s::AbstractString, default = missing)
    # synonym of wb[s], e.g. wb["Sheet1"] or wb["Sheet1!A1"]
    m = match(sheetregex,s) # when s = e.g. "Sheet1"
    if !isnothing(m)
        if !(s in keys(wb.sheets))
            if ismissing(default)
                wb.sheets[s] = Sheet(s, wb)
            else
                @assert (default.wb === wb) "Sheet must belong to this workbook"
                wb.sheets[s] = default
            end
        end
        return wb.sheets[s]
    else
        return get(wb, GlobalPosition(s))
    end
end

## calculation agents
# get values
function value(wb::Workbook, p::GlobalPosition) 
    c = wb.map[p]
    c.value
end
function value(wb::Workbook, r::GlobalRange)
    map(p -> value(wb,p), collect(r))
end

# calculate the value of the cell based on current state of sheet
refresh!(p::AbstractString, wb::Workbook) = refresh!(GlobalPosition(p), wb)
refresh!(p::GlobalPosition, wb::Workbook) = refresh!(get(wb,p), wb)
function refresh!(_::ConstCell,_::Workbook) end # nothing to do
function refresh!(c::DynamicCell,wb::Workbook)
    c.value = c.func(value.([wb],c.ancestors)...)
end
function refresh_latest!(_::ConstCell,_::Workbook) end
function refresh_latest!(c::DynamicCell,wb::Workbook)
    c.value = Base.invokelatest(c.func, value.([wb],c.ancestors)...)
end

function orderedFromSuperset(xx,xxx)
    xx[sortperm(findfirst.(isequal.(xx),[xxx]))]
end
linear_location(c::Cell, wb::Workbook) = linear_location(c.position, wb)
linear_location(p::GlobalPosition, wb::Workbook) = findfirst(isequal(p),wb.map.keys)
function cascade(p::GlobalPosition,wb::Workbook,latest::Bool=false)
    loc = linear_location(p,wb)
    refreshNeeded = orderedFromSuperset(neighbors(wb.graph, loc),wb.order)
    for key in refreshNeeded
        latest ? refresh_latest!(wb.map[wb.map.keys[key]],wb) : refresh!(wb.map[wb.map.keys[key]],wb)
        push!(refreshNeeded, orderedFromSuperset(neighbors(wb.graph, key),wb.order)...) # add cells that are further downstream to the list
    end
end
# get the cell into the overall weave of the thing
function extendgraph(_::ConstCell, _::Workbook) end # nothing to do
function extendgraph(c::DynamicCell, wb::Workbook)
    # basically: make sure cell's ancestors exist & are recorded
    
    # create nonexistent ancestors
    for anc in c.ancestors # iterating over GlobalPositions
        get!(wb, anc) # creates empty cells where necessary
    end
    # extend graph
    incrementalNodes = length(wb.map) - nv(wb.graph)
    add_vertices!(wb.graph,incrementalNodes)
    # add edges
    for ancestor in c.ancestors # iterating over GlobalPositions
        add_ancestor(wb,ancestor,c.position)
    end
    
    # update total ordering
    wb.order = topological_sort_by_dfs(wb.graph)    
end
function add_ancestor(wb::Workbook, ancestor::GlobalPosition, descendant::GlobalPosition)
    loc = linear_location(descendant,wb)
    ancloc = findfirst(isequal(ancestor),wb.map.keys)
    add_edge!(wb.graph, ancloc, loc) # ancestor -> descendant
    if is_cyclic(wb.graph) || has_self_loops(wb.graph)
        rem_edge!(wb.graph, ancloc, loc) # let's undo that
        error("There should not be any cycles ... now what?")
    end
end
add_ancestor(wb::Workbook, ancestor::GlobalRange, descendant::GlobalPosition) = add_ancestor.([wb], collect(ancestor), [descendant])

# write a cell
Base.setindex!(wb::Workbook, f::AbstractString, p::GlobalPosition) = set!(wb, Cell(p,f))
Base.setindex!(wb::Workbook, f::AbstractString, p::AbstractString) = set!(wb, Cell(p,f))  # note the order: setindex!(object, value, position)
Base.setindex!(wb::Workbook, f, x) = setindex!(wb,string(f),x)

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
    is_new_node || cascade(c.position,wb,true) # only if an existing cell has been replaced
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

