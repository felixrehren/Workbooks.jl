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
            refs = extractRefsFromExpression(content)

            # sanitise spreadsheet-style references that are not compatible with Julia
            sanitised_refs = replace.(refs,[":" => "to"])
            sanitised_formula = formula
            for r in filter(contains(":"),refs)
                sanitised_r = replace(r, ":" => "to")
                sanitised_formula = replace(sanitised_formula, r => sanitised_r)
            end
            sanitised_content = Meta.parse(chop(sanitised_formula,head=1,tail=0))

            if isempty(refs)
                func = @eval () -> $sanitised_content
            else
                func = @eval $(Meta.parse(join(sanitised_refs,", "))) -> $sanitised_content
            end
            
            ancestors = GlobalisedRef.(refs,[p.sheet])
            return DynamicCell(p, formula, func, ancestors, missing)
        elseif isa(content,Symbol)
            func = identity
            ancestors = [GlobalisedRef(string(content),p.sheet)]
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
formula(_::Missing) = missing

value(c::Cell) = c.value
value(_::Missing) = missing