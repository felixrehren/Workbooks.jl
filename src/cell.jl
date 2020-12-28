## C E L L

EmptyCell(p::AbstractString) = EmptyCell(GlobalPosition(p))
EmptyCell(p::GlobalPosition) = ConstCell(p,missing)

# construction of a cell given user input as string
Cell(p::AbstractString, formula) = Cell(GlobalPosition(p),formula)
function Cell(p::GlobalPosition, formula::AbstractString)
    isempty(formula) && EmptyCell(p)
    if first(formula) == '='
        content = Meta.parse(chop(formula,head=1,tail=0))
        if isa(content, Expr)
            #mm = extractMatchesFromExpr(content) # local expressions only for now
            #refs = getfield.(mm,:match)
            refs = extractRefsFromExpression(content)
            sanitised_content = content
            for r in filter(contains(":"),refs)
                sanitised_r = replace(r, ":" => "to")
                sanitised_content = replace(sanitised_content, r => sanitised_r)
            end
            sanitised_refs = replace.(refs,[":" => "to"])
            if isempty(refs)
                func = @eval () -> $content
            else
                func = @eval $(Meta.parse(join(sanitised_refs,", "))) -> $content
            end
            
            ancestors = GlobalPosition.([p.sheet],LocalPosition.(refs))
            return DynamicCell(p, formula, func, ancestors, missing)
        end
    else
        content = Meta.parse(formula)
    end
    value = isa(content,Number) ? content : formula # number or string
    ConstCell(p,value)
end

Base.show(io::IO, c::ConstCell) = Base.print(io, "Cell = " * string(c.value))
Base.show(io::IO, c::DynamicCell) = Base.print(io, "Cell = " * string(c.value) * " " * c.formula)
# Base.show(io::IO, c::CloneCell) = Base.print(io, "Cell = " * string(c.value) * " " * c.cloneOf.formula)

formula(c::DynamicCell) = c.formula
formula(c::ConstCell) = string(c.value)