using Documenter
using Workbooks

makedocs(
    sitename = "Workbooks",
    format = Documenter.HTML(),
    modules = [Workbooks]
)

# Documenter can also automatically deploy documentation to gh-pages.
# See "Hosting Documentation" and deploydocs() in the Documenter manual
# for more information.
#=deploydocs(
    repo = "<repository url>"
)=#
