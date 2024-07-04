###########################################################################
#
# This file contains following functions
#   glmxls - GLM regression models
#   bivariatexls - for two-way tables for discrete or continuous variables
#   univariatexls - for continuous variables
#   dfxls - for exporting a dataframe
#
###########################################################################

"""
"""
function hltest(glmout,q = 10)
    d = DataFrame(yhat = predict(glmout),y = glmout.model.rr.y)
    d[:group] = xtile(d[:yhat],nq = q)
    df3 = by(d,:group) do subdf
        n = size(subdf,1)
        o1 = sum(subdf[:y])
        e1 = sum(subdf[:yhat])
        DataFrame(
            observed1 = o1,
            observed0 = n - o1,
            expected1 = e1,
            expected0 = n - e1
        )
    end

    hlstat = sum((df3[:observed1] .- df3[:expected1]).^2 ./ df3[:expected1]
        .+ (df3[:observed0] .- df3[:expected0]).^2 ./ df3[:expected0])

    dof = q - 2

    pval = Distributions.ccdf(Distributions.Chisq(dof),hlstat)

    return (hlstat, dof, pval)
end

# For these functions, use r2(GLM, :MacFadden) or r2(GLM,:Nagelkerke)
# function nagelkerke(glmout)
#     return StatsBase.r2(glmout,:Nagelkerke)
# end

# function macfadden(glmout)
#     return StatsBase.r2(glmout,:McFadden)
# end


"""
    glmxls(glmout::DataFrames.DataFrameRegressionModel, workbook::PyObject, worksheet::AbstractString; eform=false,ci=true, row = 0, col =0)

Outputs a GLM regression table to an excel spreadsheet.
To use this function, `PyCall` is required with a working version python and
a python package called `xlsxwriter` installed. If a label is found for a variable
or a value of a variable in a `Label`, the label will be output. Options are:

- `glmout`: returned value from a GLM regression model
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name
- `eform`: use `eform = true` to get exponentiated estimates, standard errors, or 95% confidence intervals
- `ci`: use `ci = true` (default) to get 95% confidence intervals. `ci = false` will produce standard errors and Z values instead.
- `row`: specify the row of the workbook to start the output table (default = 0 (for row 1))
- `col`: specify the column of the workbook to start the output table (default = 0 (for column A))

# Example 1
This example is useful when one wants to append a worksheet to an existing workbook.
It is responsibility of the user to open a workbook before the function call and close it
to actually create the physical file by close the workbook.

```
julia> using PyCall

julia> @pyimport xlsxwriter

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> glmxls(ols1,wb,"OLS1",labels = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate")

Julia> wb.close()
```

# Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `@pyimport`
does not need to be called before the function.

```
julia> glmxls(ols1,"test_workbook.xlsx","OLS1")
```

"""
function glmxls(glmout,wbook::PyObject,wsheet::AbstractString; labels::Dict = nothing,
    eform::Bool = false, ci = true, row = 0, col = 0, robust::Symbol = nothing)

    if (typeof(glmout) <: StatsModels.TableRegressionModel) == false
        error("This is not a regression model output.")
    end

    if isa(glmout.model,GeneralizedLinearModel)
        distrib = glmout.model.rr.d
        linkfun = Link(glmout.model.rr)
    else
        distrib = nothing
        linkfun = nothing
    end

    # create a worksheet
    t = wbook.add_worksheet(wsheet)

    # attach formats to the workbook
    formats = ExcelTables.attach_formats(wbook)

    # starting location in the worksheet
    r = row
    c = col

    # set column widths
    t.set_column(c,c,40)
    t.set_column(c+1,c+4,9)

    # headings ---------------------------------------------------------
    # if ci == true, Estimate (95% CI) P-Value
    # if eform == true, Estimate is OR for logit, IRR for poisson
    otype = "Estimate"
    if eform == true
        otype = Stella.coeflab(distrib,linkfun)
    end

    t.write_string(r,c,"Variable",formats[:heading])
    if ci == true
        t.merge_range(r,c+1,r,c+3,string(otype," (95% CI)"),formats[:heading])
        t.write_string(r,c+4,"P-Value",formats[:heading])
    else
        t.write_string(r,c+1,otype,formats[:heading])
        t.write_string(r,c+2,"SE",formats[:heading])
        t.write_string(r,c+3,"Z Value",formats[:heading])
        t.write_string(r,c+4,"P-Value",formats[:heading])
    end

    #------------------------------------------------------------------
    r += 1
    c = col

    # go through each variable and construct variable name and value label arrays
    tdata = coeftable(glmout)
    nrows = length(tdata.rownms)
    varname = Vector{String}(undef,nrows)
    vals = Vector{String}(undef,nrows)
    nlev = zeros(Int,nrows)
    #nord = zeros(Int,nrow)
    for i = 1:nrows
    	# variable name
        # parse varname to separate variable name from value
        if occursin(":",tdata.rownms[i])
            (varname[i],vals[i]) = split(tdata.rownms[i],": ")
        elseif occursin(" - ",tdata.rownms[i])
            (varname[i],vals[i]) = split(tdata.rownms[i]," - ")
        else
            varname[i] = tdata.rownms[i]
            vals[i] = ""
        end

        # column labels
        if labels != nothing
            if haskey(labels,Symbol(varname[i]))
                varname[i] = labels[Symbol(varname[i])]
            else
                varname[i] = string(varname[i])
            end
        end
    end
    for i = 1:nrows
        nlev[i] = countlev(varname[i],varname)
    end

    # write table
    tconfint = confint(glmout)
    lastvarname = ""

    for i = 1:nrows
        if otype == "OR" && varname[i] == "(Intercept)"
            continue
        end
        if varname[i] != lastvarname
            # output cell boundaries only and go to the next line
            if nlev[i] > 1
                t.write_string(r,c,varname[i],formats[:heading_left])

                if ci == true
                    t.write(r,c+1,"",formats[:empty_right])
                    t.write(r,c+2,"",formats[:empty_both])
                    t.write(r,c+3,"",formats[:empty_left])
                    t.write(r,c+4,"",formats[:p_fmt])
                else
                    t.write(r,c+1,"",formats[:empty_border])
                    t.write(r,c+2,"",formats[:empty_border])
                    t.write(r,c+3,"",formats[:empty_border])
                    t.write(r,c+4,"",formats[:p_fmt])
                end
                r += 1
                t.write_string(r,c,vals[i],formats[:varname_1indent])

            else
                if vals[i] != "" && vals[i] != "Yes"
                    t.write_string(r,c,string(varname[i]," - ",vals[i]),formats[:heading_left])
                else
                    t.write_string(r,c,varname[i],formats[:heading_left])
                end
            end
        else
            t.write_string(r,c,vals[i],formats[:varname_1indent])
        end

    	# estimates
        if eform == true
    	    t.write(r,c+1,exp(tdata.cols[1][i]),formats[:or_fmt])
        else
            t.write(r,c+1,tdata.cols[1][i],formats[:or_fmt])
        end

        # standard errors
        if robust != nothing
        end

        if ci == true

            if eform == true
            	# 95% CI Lower
            	t.write(r,c+2,exp(tconfint[i,1]),formats[:cilb_fmt])

            	# 95% CI Upper
            	t.write(r,c+3,exp(tconfint[i,2]),formats[:ciub_fmt])
            else
                # 95% CI Lower
            	t.write(r,c+2,tconfint[i,1],formats[:cilb_fmt])

            	# 95% CI Upper
            	t.write(r,c+3,tconfint[i,2],formats[:ciub_fmt])
            end
        else
            # SE
            if eform == true
        	    t.write(r,c+2,exp(tdata.cols[1][i])*tdata.cols[2][i],formats[:or_fmt])
            else
                t.write(r,c+2,tdata.cols[1][i],formats[:or_fmt])
            end

            # Z value
            t.write(r,c+3,tdata.cols[3][i],formats[:or_fmt])

        end

        # P-Value
        t.write(r,c+4,tdata.cols[4][i] < 0.001 ? "< 0.001" : tdata.cols[4][i] ,formats[:p_fmt])

        lastvarname = varname[i]

        # update row
        r += 1
    end

    # Write model characteristics and goodness of fit statistics
    c = col

    # N
    t.write(r,c,"N",formats[:model_name])
    t.merge_range(r,c+1,r,c+4,nobs(glmout),formats[:n_fmt_center])

    # degress of freedom
    r += 1
    t.write(r,c,"DF",formats[:model_name])
    t.merge_range(r,c+1,r,c+4,dof(glmout),formats[:n_fmt_center])

    # R² or pseudo R²
    r += 1
    if isa(linkfun,LogitLink)
        t.write(r,c,"Pseudo R² (MacFadden)",formats[:model_name])
        t.merge_range(r,c+1,r,c+4,r2(glmout,:MacFadden),formats[:p_fmt_center])
        t.write(r+1,c,"Pseudo R² (Nagelkerke)",formats[:model_name])
        t.merge_range(r+1,c+1,r+1,c+4,r2(GLM,:Nagelkerke),formats[:p_fmt_center])

        # -2 log-likelihood
        t.write(r+2,c,"-2 Log-Likelihood",formats[:model_name])
        t.merge_range(r+2,c+1,r+2,c+4,deviance(glmout),formats[:p_fmt_center])

        # Hosmer-Lemeshow GOF test
        t.write(r+3,c,"Hosmer-Lemeshow Chisq Test (df), p-value",formats[:model_name])
        hl = hltest(glmout)
        t.merge_range(r+3,c+1,r+3,c+4,string(round(hl[1],digits=4)," (",hl[2],"); p = ",round(hl[3],digits=4)),formats[:p_fmt_center])

        # ROC (c-statistic)
        t.write(r+4,c,"Area under the ROC Curve",formats[:model_name])
        roc = auc(roc(glmout.model.rr.y,predict(glmout)))
        n1 = sum(glmout.model.rr.y) # number of positive responses
        n2 = nobs(glmout) - n1
        q1 = roc / (2 - roc)
        q2 = (2*roc^2) / (1 + roc)
        rocse = sqrt((roc*(1-roc) + (n1-1)*(q1 - roc^2) + (n2 - 1)*(q2 - roc^2)) / (n1*n2))
        t.merge_range(r+4,c+1,r+4,c+4,string(round(roc,digits=4)," (95% CI, ", round(roc - 1.96*rocse,digits=4), " - ", round(roc + 1.96*rocse,digits=4),")"),formats[:p_fmt_center])

        r += 5
    else
        t.write(r,c,"R²",formats[:model_name])
        t.merge_range(r,c+1,r,c+4,r2(glmout),formats[:p_fmt_center])
        t.write(r+1,c,"Adjusted R²",formats[:model_name])
        t.merge_range(r+1,c+1,r+1,c+4,adjr2(glmout),formats[:p_fmt_center])

        r += 2
    end

    # AIC & BIC
    t.write(r,c,"AIC",formats[:model_name])
    t.merge_range(r,c+1,r,c+4,aic(glmout),formats[:p_fmt_center])

    r += 1
    t.write(r,c,"BIC",formats[:model_name])
    t.merge_range(r,c+1,r,c+4,bic(glmout),formats[:p_fmt_center])
end
function glmxls(glmout,
    wbook::AbstractString,
    wsheet::AbstractString;
    labels::Dict = nothing,
    eform::Bool = false,
    ci = true,
    row = 0,
    col = 0,
    robust = nothing)

    xlsxwriter = pyimport("xlsxwriter")

    glmxls(glmout,xlsxwriter.Workbook(wbook),wsheet,labels=labels,eform=eform,ci=ci,row=row,col=col,robust=robust)
end


"""
    bivariatexls(df::DataFrame,colvar::Symbol,rowvars::Vector{Symbol},workbook::PyObject,worksheet::AbstractString; row=0,col=0,column_percent = true)

 Creates bivariate statistics and appends it in a nice tabular format to an existing workbook.
 To use this function, `PyCall` is required with a working version python and
 a python package called `xlsxwriter` installed.  If a label is found for a variable
 or a value of a variable in a `Label`, the label will be output. Options are:

- `df`: a DataFrame
- `colvar`: a categorical variable whose values will be displayed on the columns
- `rowvars`: a Vector of Symbols for variables to be displayed on the rows. Both continuous and categorical variables are allowed.
    For continuous variables, mean and standard deviations will be output and a p-value will be based on an ANOVA test. For categorical variables,
    a r x c table with cell counts and row percentages will be output with a p-value based on a chi-square test.
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name
- `row`: specify the row of the workbook to start the output table (default = 0 (for row 1))
- `col`: specify the column of the workbook to start the output table (default = 0 (for column A))
- `column_percent`: set this to `false` if you want row percentages in the output table (default = true)

### Example 1
This example is useful when one wants to append a worksheet to an existing workbook.
It is responsibility of the user to open a workbook before the function call and close it
to actually create the physical file by close the workbook.

```
julia> using PyCall

julia> xlsxwriter = pyimport("xlsxwriter")

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> glmxls(ols1,wb,"OLS1",labels = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate")

Julia> wb.close()
```

### Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `@pyimport`
does not need to be called before the function.

```
julia> bivariatexls(df,:incomecat,[:age,:race,:male,:bmicat],"test_workbook.xlsx","Bivariate")
```
"""
function bivariatexls(df::AbstractDataFrame, colvar::Symbol, rowvars::Vector{Symbol}, wbook::PyObject, wsheet::AbstractString;
    wt::Symbol = nothing, row::Int = 0, col::Int = 0, column_percent::Bool = true, verbose::Bool = false)

    # colvar has to be a CategoricalArray and must have 2 or more categories
    if isa(df[!,colvar], CategoricalArray) == false || length(levels(df[!,colvar])) < 2
        error("`",colvar,"` is not a CategoricalArray or does not have two or more levels")
    end

    # create a worksheet
    t = wbook.add_worksheet(wsheet)

    # attach formats to the workbook
    formats = ExcelTables.attach_formats(wbook)

    # starting row and column
    r = row
    c = col

    # drop NAs in colvar
    df2 = df[completecases(df[!,[colvar]]),:]

    # number of columns
    # column values
    if wt == nothing
        collev = freqtable(df2,colvar,skipmissing=true)
    else
        collev = freqtable(df2,colvar,skipmissing=true,weights=df2[wt])
    end
    nlev = length(collev.array)
    colnms = names(collev,1)
    coltot = sum(collev.array,dims=1)

    # set column widths
    t.set_column(c,c,40)
    t.set_column(c+1,c+(nlev+1)*2+1,9)

    # create header
    # column variable name
    # It uses three rows
    t.merge_range(r,c,r+2,c,"Variable",formats[:heading])

    # header 1st row = variable name
    t.merge_range(r,c+1,r,c+(nlev+1)*2+1,label(df,colvar),formats[:heading])

    # header 2nd and 3rd rows
    r += 1

    t.merge_range(r,1,r,2,"All",formats[:heading])
    t.write_string(r+1,1,"N",formats[:n_fmt_right])
    t.write_string(r+1,2,"(%)",formats[:pct_fmt_parens])

    # 
    c += 3
    for i = 1:nlev
        t.merge_range(r,c+(i-1)*2,r,c+(i-1)*2+1,colnms[i],formats[:heading])
        t.write_string(r+1,c+(i-1)*2,"N",formats[:n_fmt_right])
        t.write_string(r+1,c+(i-1)*2+1,"(%)",formats[:pct_fmt_parens])
    end

    # P-value
    t.merge_range(r,c+nlev*2,r+1,c+nlev*2,"P-Value",formats[:heading])

    # total
    c = col
    r += 2
    t.write_string(r,c,"All, n (Row %)",formats[:model_name])
    if wt == nothing
        x = freqtable(df2,colvar,skipmissing=true)
    else
        x = freqtable(df2,colvar,skipmissing=true,weights=df2[wt])
    end
    tot = sum(x)
    t.write(r,c+1,tot,formats[:n_fmt_right])
    t.write(r,c+2,1.0,formats[:pct_fmt_parens])
    for i = 1:nlev
        t.write(r,c+i*2+1,x.array[i],formats[:n_fmt_right])
        t.write(r,c+i*2+2,x.array[i]/tot,formats[:pct_fmt_parens])
    end
    t.write(r,c+(nlev+1)*2+1,"",formats[:empty_border])

    # covariates
    c = col
    r += 1
    for varname in rowvars

        if verbose == true
            println("Processing ",varname)
        end

        # variable name
        vars = label(df,varname)

        # determine if varname is categorical or continuous
        if isa(df2[!,varname], CategoricalArray) || eltype(df2[!,varname]) == String

            # categorial
            df3=df2[completecases(df2[:,[varname]]),[varname,colvar]]
            if wt == nothing
                x = freqtable(df3,varname,colvar,skipmissing=true)
            else
                x = freqtable(df3,varname,colvar,skipmissing=true,weights=df3[wt])
            end
            rowval = names(x,1)
            rowtot = sum(x.array,dims=2)
            coltot = sum(x.array,dims=1)

            # variable name
            # if there only two levels and one of the values is 1 or true
            # and the other values is 0 or false,
            # just output the frequency and percentage of the 1/true row

            # variable name
            t.write_string(r,c,vars,formats[:model_name])

            # two levels with [0,1] or [false,true]
            if length(rowval) <= 2 && rowval in ([1],[true],["Yes"],[0,1],[false,true],["No","Yes"])

                nrow = length(rowval) 
                # row total
                t.write(r,c+1,rowtot[nrow],formats[:n_fmt_right])
                t.write(r,c+2,rowtot[nrow]/tot,formats[:pct_fmt_parens])

                for j = 1:nlev
                    t.write(r,c+j*2+1,x.array[nrow,j],formats[:n_fmt_right])
                    if column_percent
                        t.write(r,c+j*2+2, coltot[j] > 0 ? x.array[nrow,j]/coltot[j] : "",formats[:pct_fmt_parens])
                    else # elseif rowtot[2] > 0
                        t.write(r,c+j*2+2, rowtot[nrow] > 0 ? x.array[nrow,j]/rowtot[nrow] : "",formats[:pct_fmt_parens])
                    end
                end
                pval = pvalue(ChisqTest(x.array))
                if isnan(pval) || isinf(pval)
                    pval = ""
                elseif pval < 0.001
                    pval = "< 0.001"
                end
                t.write(r,c+(nlev+1)*2+1, pval,formats[:p_fmt])
                r += 1
            else
                for i = 1:nlev+1
                    t.write_string(r,c+(i-1)*2+1,"",formats[:empty_right])
                    t.write_string(r,c+(i-1)*2+2,"",formats[:empty_left])
                end
                t.write_string(r,c+(nlev+1)*2+1,"",formats[:empty_border])

                r += 1
                for i = 1:length(rowval)
                    # row value
                    t.write_string(r,c,rowval[i],formats[:varname_1indent])

                    # row total
                    t.write(r,c+1,rowtot[i],formats[:n_fmt_right])
                    t.write(r,c+2,rowtot[i]/tot,formats[:pct_fmt_parens])

                    for j = 1:nlev
                        t.write(r,c+j*2+1,x.array[i,j],formats[:n_fmt_right])
                        if column_percent
                            t.write(r,c+j*2+2,coltot[j] > 0 ? x.array[i,j]/coltot[j] : "",formats[:pct_fmt_parens])
                        else
                            t.write(r,c+j*2+2,rowtot[i] > 0 ? x.array[i,j]/rowtot[i] : "",formats[:pct_fmt_parens])
                        end
                    end
                    # p-value - output only once
                    pval = pvalue(ChisqTest(x.array))

                    if isnan(pval) || isinf(pval)
                        pval = ""
                    elseif pval < 0.001
                        pval = "< 0.001"
                    end
                    if length(rowval) == 1
                        t.write(r,c+(nlev+1)*2,pval,formats[:p_fmt])
                    elseif i == 1
                        t.merge_range(r,c+(nlev+1)*2+1,r+length(rowval)-1,c+(nlev+1)*2+1,pval,formats[:p_fmt])
                    end
                    r += 1
                end
            end
        else
            # continuous variable
            df3=df2[completecases(df2[!,[varname]]),[varname,colvar]]
            y = tabstat(df3, varname, colvar, table=false, wt=df3[wt])

            # variable name
            t.write_string(r,c,string(vars,", mean (SD)"),formats[:model_name])

            # All
            tmpvec = collect(skipmissing(df3[!,varname]))
            if length(tmpvec) == 0
                amean = ""
                astd = ""
            else
                amean = mean(tmpvec)
                if isnan(amean)
                    amean = ""
                end
                astd = std(tmpvec)
                if isnan(astd)
                    astd = ""
                end
            end
            t.write(r,c+1,amean,formats[:f_fmt_right])
            t.write(r,c+2,astd,formats[:f_fmt_left_parens])

            # colvar levels
            for i = 1:nlev
                if i <= size(y,1) && y[i,:N] > 1
                    t.write(r,c+i*2+1,y[i,:mean],formats[:f_fmt_right])
                    t.write(r,c+i*2+2,y[i,:sd],formats[:f_fmt_left_parens])
                else
                    t.write(r,c+i*2+1,"",formats[:f_fmt_right])
                    t.write(r,c+i*2+2,"",formats[:f_fmt_left_parens])
                end
            end
            if size(y,1) > 1
                pval = Stella.anova(df3,varname,colvar; pval=true)
                if isnan(pval) || isinf(pval)
                    pval = ""
                elseif pval < 0.001
                    pval = "< 0.001"
                end
                t.write(r,c+(nlev+1)*2+1,pval,formats[:p_fmt])
            else
                t.write(r,c+(nlev+1)*2+1,"",formats[:p_fmt])
            end

            r += 1
        end
    end
end
# function bivariatexls(_df::AbstractDataFrame,
#     colvar::Symbol,
#     rowvars::Vector{Symbol},
#     wbook::AbstractString,
#     wsheet::AbstractString;
#     row::Int = 0,
#     col::Int = 0)

#     xlsxwriter = pyimport("xlsxwriter")

#     wb = xlsxwriter.Workbook(wbook)

#     bivariatexls(_df,colvar,rowvars,wb,wsheet,row=row,col=col)

#     wb.close()
# end


"""
    univariatexls(df::DataFrame,contvars::Vector{Symbol},workbook::PyObject,worksheet::AbstractString; labels::Union{Nothing,Label}=nothing,row=0,col=0)

Creates univariate statistics for a vector of continuous variable and
appends it to an existing workbook.
To use this function, `PyCall` is required with a working version python and
a python package called `xlsxwriter` installed.  If a label is found for a variable
in a `Label` object, the label will be output. Options are:

- `df`: a DataFrame
- `contvars`: a vector of continuous variables
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name
- `labels`: an option to specify a `label` dictionary (see an example below)
- `row`: specify the row of the workbook to start the output table (default = 0 (for row 1))
- `col`: specify the column of the workbook to start the output table (default = 0 (for column A))

# Example 1
This example is useful when one wants to append a worksheet to an existing workbook.
It is responsibility of the user to open a workbook before the function call and close it
to actually create the physical file by close the workbook.

```
julia> using PyCall

julia> @pyimport xlsxwriter

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> glmxls(ols1,wb,"OLS1",labels = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate",labels = label)

julia> univariatexls(df,[:age,:income_amt,:bmi],wb,"Univariate",labels = label)

Julia> wb.close()
```

# Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `@pyimport`
does not need to be called before the function.

```jldoctest
julia> univariatexls(df,[:age,:income_amt,:bmi],"test_workbook.xlsx","Bivariate",labels = label)
```

"""
function univariatexls(df::DataFrame,
    contvars::Vector{Symbol},
    wbook::PyObject,
    wsheet::AbstractString;
    wt::Union{Nothing,Symbol} = nothing,
    row = 0,
    col = 0)

    # create a worksheet
    t = wbook.add_worksheet(wsheet)

    # attach formats to the workbook
    formats = ExcelTables.attach_formats(wbook)

    # starting row and column
    r = row
    c = col

    # column width
    t.set_column(0,0,20)
    t.set_column(1,length(contvars),12)

    # output the row names
    rownms = ["N Total", "N Miss", "N Used", "Sum", "Mean",
            "SD", "Variance", "Minimum", "P25", "Median", "P75", "Maximum",
            "Skewness","Kurtosis","Smalles","","","","","Largest","","","",""]

    t.write_string(r,c,"Statistic",formats[:heading])
    for i in 1:24
        t.write_string(r+i,c,rownms[i],formats[:heading_left])
    end

    col = 1
    for vsym in contvars

        # if symbol is not found in the DataFrame
        if !in(vsym,propertynames(df))
            continue
        end

        # vsym is not a real number
        if (nonmissingtype(eltype(df[!,vsym])) <: Real) == false
            continue
        end

        # non-missing values
        len = size(df,1) - count(ismissing, df[!,vsym])  # sum(ismissing.(df[!,vsym]) .== false)

        # pick up the variable label
        varstr = label(df,vsym)
        t.write_string(0,col,varstr,formats[:heading])
        u = Stella.univariate(df[!,vsym]) #,wt=df[wt])
        for j = 1:14
            if j<4
                fmttype = :n_fmt
            else
                fmttype = :p_fmt
            end
            if isnan(u[j,:Value]) || isinf(u[j,:Value])
                t.write(j,col,"",formats[fmttype])
            else
                t.write(j,col,u[j,:Value],formats[fmttype])
            end
        end

        len = len < 5 ? len : 5
        smallest=Stella.smallest(df[!,vsym], n = len)
        if nonmissingtype(eltype(df)) <: Integer
            fmttype = :n_fmt
        else
            fmttype = :p_fmt
        end
        for j = 1:5
            if j <= len
                t.write(j+14,col,smallest[j],formats[fmttype])
            else
                t.write(j+14,col,"",formats[fmttype])
            end
        end
        largest=Stella.largest(df[!,vsym], n = len)
        for j = 1:5
            if j <= len
                t.write(j+19,col,largest[j],formats[fmttype])
            else
                t.write(j+19,col,"",formats[fmttype])
            end
        end
        col += 1
    end
end
function univariatexls(df::DataFrame,contvars::Vector{Symbol},wbook::AbstractString,wsheet::AbstractString;
    wt::Union{Nothing,Symbol} = nothing,row = 0, col = 0)

    xlsxwriter=pyimport("xlsxwriter")

    wb = wlsxwriter.Workbook(wbook)

    univariatexls(df,contvars,wb,wsheet,wt=wt,row=row,col=col)

    wb.close()
end


"""
    dfxls(df::DataFrame, workbook::PyObject, worksheet::AbstractString; nrows = 500, start = 1, row=0, col=0)

 To use this function, `PyCall` is required with a working version python and
 a python package called `xlsxwriter` installed. Options are:

- `df`: a DataFrame
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name (default: "Data1")
- `start`: specify the row number from which data will be output (default: 1)
- `nrows`: specify the number of rows to output (default: 500). If nows = 0, the entire dataframe will be output.
- `row`: specify the row of the workbook to start the output table (default = 0 (for row 1))
- `col`: specify the column of the workbook to start the output table (default = 0 (for column A))

# Example 1
This example is useful when one wants to append a worksheet to an existing workbook.
It is responsibility of the user to open a workbook before the function call and close it
to actually create the physical file by close the workbook.

```
julia> using PyCall

julia> xlsxwriter = pyimport("xlsxwriter")

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> glmxls(ols1,wb,"OLS1",labels = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate",labels = label)

julia> univariatexls(df,[:age,:income_amt,:bmi],wb,"Univariate",labels = label)

julia> dfxls(df,wb,"dataframe",nrows = 0)

Julia> wb.close()
```

# Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `@pyimport`
does not need to be called before the function.

```
julia> dfxls(df,"test_workbook.xlsx","df",nrows = 0)
```

"""
function dfxls(df::DataFrame,
    wbook::PyObject,
    worksheet::AbstractString;
    nrows::Int64 = 0, start::Int64 = 1, col::Int64 = 0, row::Int64 = 0)

    # create a worksheet
    t = wbook.add_worksheet(worksheet)

    # attach formats to the workbook
    formats = ExcelTables.attach_formats(wbook)

    # starting row and column
    c = col

    # Eltype
    typ = Vector{DataType}(undef,size(df,2))
    for i=1:ncol(df)
        if typeof(df[!,i]) <: CategoricalArray
            typ[i] = eltype(df[!,i].pool.index)
        else
            typ[i] = nonmissingtype(eltype(df[!,i]))
        end
    end
    varnames = propertynames(df)

    # if nrows = 0, output the full data set
    if nrows == 0
        start = 1
        nrows = size(df,1)
    else
        nrows = (start+nrows-1) < size(df,1) ? nrows : (size(df,1)-start+1)
    end

    for i = 1:size(df,2)

        r = row
        t.set_column(c,c,10)
        t.write_string(r,c,string(varnames[i]),formats[:heading])
        r += 1

        for j in start:(start+nrows-1)
            # println("i = ",i,"; j = ",j,". value = ",df[j,i])

            if ismissing(df[j,i])
                t.write_string(r,c," ",formats[:n_fmt])
            elseif typ[i] <: AbstractString
                if df[j,i] == ""
                    t.write_string(r,c," ",formats[:text])
                else
                    t.write(r,c,df[j,i],formats[:text])
                end
            elseif typ[i] <: Number
                if isnan(df[j,i]) || isinf(df[j,i])
                    t.write_string(r,c," ",formats[:text])
                elseif typ[i] <: Integer
                    t.write(r,c,df[j,i],formats[:n_fmt])
                elseif typ[i] <: AbstractFloat
                    t.write(r,c,df[j,i],formats[:f_fmt])
                end
            elseif typ[i] <: Date
                t.write(r,c,Dates.value(df[j,i] - Date(1899,12,30)),formats[:f_date])
            elseif typ[i] <: DateTime
                t.write(r,c,(Dates.value(df[j,i] - DateTime(1899,12,30,0,0,0)))/86400000,formats[:f_datetime])
            elseif typ[i] == Symbol || typ[i] == DataType
                t.write(r,c,string(df[j,i]),formats[:text])
            else
                # skip
            end

            r += 1
        end
        c += 1
    end

end
function dfxls(df::DataFrame,
    wbook::AbstractString,
    worksheet::AbstractString;
    nrows::Int64 = 0, start::Int64 = 1, col::Int64 = 0, row::Int64 = 0)

    # import xlsxwriter
    XlsxWriter = pyimport("xlsxwriter")

    # create a workbook
    wb = XlsxWriter.Workbook(wbook)

    dfxls(df,wb,worksheet, nrows = nrows, start = start, col = col, row = row)

    wb.close()
end

function newfilename(filen::AbstractString)
    while (isfile(filen))
        # separate the name into three parts, basename (number).extension
        (basename,ext) = splitext(filen)
        m = match(r"(.*)(\(.*\))$",basename)
        if m == nothing
            filen = string(filen," (1)")
        else
            m2 = parse(Int,replace(m[2],r"\((.*)\)",s"\1"))
            filen = string(m[1]," (",m2+1,")")
        end
    end
    return filen
end

function countlev(str::AbstractString,sarray::Vector{String})
    k = 0
    for i = 1:length(sarray)
        if str == sarray[i]
            k += 1
        end
    end
    return k
end
