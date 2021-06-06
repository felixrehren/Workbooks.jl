## J e W e L s
function JWL(name::AbstractString)
    Pkg.project(Pkg.Types.Context(),name,tempdir())
    folder = joinpath(tempdir(),name)
    open(joinpath(folder,"include.jl"), "w") do f
        Base.write(f, """
        
        [compat]
        julia = "1.6"
        """ * string(VERSION))
    end
    open(joinpath(folder,"include.jl"), "w") do f
        Base.write(f, """
        # Use this file to `import` any relevant packages*, 
        # and to define your own functions for use in the workbook.
        # It will be `include`d when the jwl workbook is opened.
        # * (don't forget to add packages to the `Project.toml` too)

        """)
    end
    return JWL(name, folder, Workbook())
end

Base.show(io::IO, j::JWL) = Base.print(io, "JWL " * string(j.name) * " with " * string(j.wb) * " in " * dirname(j.folder))