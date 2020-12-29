## reading and writing these files
const folder = dirname(pathof(Workbooks)) * "/../io/"

function write(wb::Workbook, location::String=folder, filename::AbstractString="Workbook1", overwrite=true)
    extension = ".jwl"
    location = folder * filename
    location *= (filename[end-3:end]==extension) ? "" : extension
    overwrite || @assert !isfile(location) ("File exists already!")
    
    w = ZipFile.Writer(location)
    for (s, S) in wb.sheets # sheet name, Sheet
        sheetloc = folder * s * ".csv"
    #    t = Tables.table(array(S,formula))
    #    CSV.write(sheetloc, t; writeheader=false)
        f = ZipFile.addfile(w, sheetloc)
    #    write(f,sprint(show, "text/csv", t))
        Base.write(f,sprint(show, "text/csv", array(S,formula)))
    end
    # add Project.toml, if any
    f = ZipFile.addfile(w, folder * "Project.toml")
    # ??
    # add include.jl, if any
    f = ZipFile.addfile(w, folder * "include.jl")
    # ??

    ZipFile.close(w)
    # move to its target location?
end

function read(filelocation::String=folder * "Workbook1.jwl")
    wb = Workbook()

     iobuffer = IOBuffer(Base.read(filelocation,String)) # workaround to no method matching `seekend` on ZipFile.ReadableFile
    r = ZipFile.Reader(iobuffer)

    # import sheets
    isCSV(fname) = fname[end-3:end] == ".csv"
    for f in filter(f -> isCSV(f.name),r.files)
        sheetname = basename(f.name)[1:end-4]
        addSheet(wb, sheetname, DelimitedFiles.readdlm(f,','))
    end
    # read `Project.toml`, if any
    f = findfirst(f -> isequal("Project.toml",f.name),r.files)
    if !isnothing(f)
        # P = Pkg.read_project(f)
        # ?? find a way to `using` all the dependencies! ## ASK SOMEONE
    end
    # include `include.jl`, if any
    f = findfirst(f -> isequal("include.jl",f.name),r.files)
    if !isnothing(f)
        write(f.name, sprint(show,Base.read(f)))
        include(f.name)
        rm(f.name)
    end

    close(r)
   
    wb
end

function addSheet(wb::Workbook, name::String, formulas::Matrix{Any})
    S = wb[name]
    for i in axes(formulas,1), j in axes(formulas,2)
        f = formulas[i,j]
        f == "missing" || (S[i,j] = string(f))
    end
end