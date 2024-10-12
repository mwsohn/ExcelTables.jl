module ExcelTables

################################################################################
##
## Dependencies
##
################################################################################

using DataFrames, PyCall, CategoricalArrays, Distributions, GLM, Survival, StatsBase, DataStructures,
    HypothesisTests, NamedArrays, FreqTables, Stella, TableMetadataTools, ROC, Dates, OrderedCollections

##############################################################################
##
## Exported methods and types (in addition to everything reexported above)
##
##############################################################################

export  univariatexls, # output univariate statistics in an excel worksheet
        bivariatexls,  # output bivariate statistics in an excel worksheet
        glmxls,        # output GLM models to an excel worksheet
        mglmxls,       # output multiple GLM regression models to an excel spreadsheet
        dfxls,         # output dataframe in an excel file
        hltest
       
##############################################################################
##
## Load files
##
##############################################################################
include("xlsout.jl")
include("mglmxls.jl")
include("formats.jl")

end # module
