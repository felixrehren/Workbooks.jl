## reading and writing these files

# copy the workbook into the relevant folder so it's up-to-date
"""
    updatefiles(jwl)

Write the JWL workbook `jwl.wb` to `jwl.folder`, ensuring that the JWL file is up-to-date.
"""
function updatefiles(jwl::JWL)
    for (sname, S) in jwl.wb.sheets
        sheetloc = joinpath(jwl.folder, sname * ".csv")
        t = Tables.table(array(S,formula))
        CSV.write(sheetloc, t; header=false)
    end
end

# write the workbook as a .jwl file somewhere
"""
    write(jwl[, targetfolder, overwrite])

Copies the JWL workbook to `targetfolder` as a `.jwl` file.
By default, `overwrite = true`.
"""
function write(jwl::JWL, targetfolder::String=tempdir(), overwrite=true)
    location = joinpath(targetfolder, jwl.name * ".jwl")
    overwrite || @assert !isfile(location) ("File exists already, and overwrite is set to `false`!")
    w = ZipFile.Writer(location)
    for (sname, S) in jwl.wb.sheets # sheet name, Sheet
        f = ZipFile.addfile(w, sname * ".csv")
        Base.write(f,sprint(show, "text/csv", array(S,formula)))
    end
    for x in ("include.jl", "Project.toml", "Manifest.toml")
        p = joinpath(jwl.folder,x)
        if isfile(p)
            f = ZipFile.addfile(w, x)
            Base.write(f, Base.read(p,String))
        end
    end
    ZipFile.close(w)
end

# read a .jwl file as a JWL
"""
    read(filename[, folder])

Reads a `.jwl` file and returns a JWL.
"""
read(filename::String,folder::String=tempdir()) = readabs(joinpath(folder,filename))
function readabs(filelocation::String)
    name = replace(basename(filelocation),".jwl" => "")
    targetpath = joinpath(tempdir(),name)
    isdir(targetpath) || mkdir(targetpath)

    # turn zipped jwl file into temporary location
    # r = ZipFile.Reader(filelocation)
    iobuffer = IOBuffer(Base.read(filelocation,String)) # workaround to no method matching `seekend` on ZipFile.ReadableFile
    r = ZipFile.Reader(iobuffer)
    for f in r.files
        Base.write(joinpath(targetpath,f.name), Base.read(f)) # this assumes that there are no folders in the folder
    end
    close(r)

    # start the environment
    Pkg.activate(targetpath) # ?
    # include `include.jl`, if any
    pinc = joinpath(targetpath,"include.jl")
    isfile(pinc) && include(pinc)
    
    # import sheets
    wb = Workbook()
    isCSV(f) = f[end-3:end] == ".csv"
    for f in filter(isCSV,readdir(targetpath))
        filepath = joinpath(targetpath,f)
        sheetname = basename(f)[1:end-4]
        addSheet(wb, sheetname, DelimitedFiles.readdlm(filepath,','))
    end
   
    JWL(name, targetpath, wb)
end

"""
    addSheet(wb, name, formulas)

Given an array `formulas`, it creates a new sheet named `name` in the workbook `wb`.
(It partially overwrites any existing sheet with the same `name`.)
"""
function addSheet(wb::Workbook, name::String, formulas::Matrix{Any})
    S = wb[name]
    for i in axes(formulas,1), j in axes(formulas,2)
        f = formulas[i,j]
        f == "missing" || (S[i,j] = string(f))
    end
end