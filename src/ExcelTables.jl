VERSION >= v"0.6.0" && __precompile__()

module ExcelTables

################################################################################
##
## Dependencies
##
################################################################################

using DataFrames, PyCall, Distributions, GLM, StatsBase, DataStructures, HypothesisTests, NamedArrays, FreqTables, Stella, Labels, AUC

##############################################################################
##
## Exported methods and types (in addition to everything reexported above)
##
##############################################################################

export  univariatexls, # output univariate statistics in an excel worksheet
        bivariatexls,  # output bivariate statistics in an excel worksheet
        glmxls,        # output GLM models to an excel worksheet
        mglmxls,       # output multiple GLM regression models to an excel spreadsheet
        dfxls          # output dataframe in an excel file

##############################################################################
##
## Load files
##
##############################################################################
include("xlsout.jl")
include("mglmxls.jl")
include("formats.jl")

end # module
