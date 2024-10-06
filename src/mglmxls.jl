"""
    mglmxls(glmout::Vector{DataFrames.DataFrameRegressionModel}, workbook::PyObject, worksheet::AbstractString; labels::Union{Nothing,Dict}=nothing,mtitle::Union{Vector,Nothing}=nothing,eform=false,ci=true, row = 0, col =0)

Outputs multiple GLM regression tables side by side to an excel spreadsheet.
To use this function, `PyCall` is required with a working version python and
a python package called `xlsxwriter` installed. If a label is found for a variable
or a value of a variable in a `labels`, the label will be output. Options are:

- `glmout`: a vector of GLM regression models
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name
- `labels`: an option to specify a `label` dictionary (see an example below)
- `mtitle`: header label for GLM models. If not specified, the dependent variable name will be used.
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

julia> xlsxwriter = pyimport("xlsxwriter")

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> mglmxls(olsmodels,wb,"OLS1",labels = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate",labels = label)

Julia> wb.close()
```

# Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `pyimport()`
does not need to be called before the function.

```
julia> mglmxls(olsmodels,"test_workbook.xlsx","OLS1",labels = label)
```

"""
function mglmxls(glmout,
    wbook::PyObject,
    wsheet::AbstractString;
    mtitle::Union{Vector,Nothing} = nothing,
    eform::Bool = false,
    ci = true,
    row = 0,
    col = 0,
    robust = nothing,
    adjust = true,
    labels = nothing)

    # take care of issues with nan/inf
    wbook.nan_inf_to_errors = true

    num_models = length(glmout)
    otype = Vector(undef,num_models)
    if isa(glmout[1].model,GeneralizedLinearModel)
	    linkfun = Vector{Link}(undef,num_models)
    	distrib = Vector{UnivariateDistribution}(undef,num_models)
    end

    for i=1:num_models

        # assuming that all models have the same family and link function
        otype[i] = "Estimate"
        if isa(glmout[i].model,GeneralizedLinearModel)
            distrib[i] = glmout[i].model.rr.d
            linkfun[i] = Link(glmout[i].model.rr)
            if eform == true
                otype[i] = Stella.coeflab(distrib[i], linkfun[i])
            end
        elseif isa(glmout[i].model,CoxModel)
            if eform == true
                otype[i] = "HR"
            end
        end

    end

    if mtitle == nothing
        mtitle = Vector{String}(undef,num_models)

        # assign dependent variables
        for i=1:num_models
            ysim = glmout[i].mf.f.lhs.sym #terms.eterms[1]
            mtitle[i] = labels != nothing && haskey(labels, ysim) ? varlab(labels,ysim) : string(ysim)
        end
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
    t.set_column(c+1,c+4*num_models,7)

    t.merge_range(r,c,r+1,c,"Variable",formats[:heading])
    for i=1:num_models
        if ci == true
            t.merge_range(r,c+1,r,c+4,mtitle[i],formats[:heading])
            t.merge_range(r+1,c+1,r+1,c+3,string(otype[i]," (95% CI)"),formats[:heading])
            t.write_string(r+1,c+4,"P-Value",formats[:heading])
        else
            t.write_string(r,c+1,otype[i],formats[:heading])
            t.write_string(r,c+2, robust != nothing ? "Robsut SE" : "SE",formats[:heading])
            t.write_string(r,c+3,"Z Value",formats[:heading])
            t.write_string(r,c+4,"P-Value",formats[:heading])
        end
        c += 4
    end

    #------------------------------------------------------------------
    r += 2
    c = col

    # collate variables from multiple regressions
    covariates = Vector{String}()
    vvalues = Dict()
    tdata = Vector(undef,num_models)
    tconfint = Vector(undef,num_models)
    loc = Dict()

    for i=1:num_models
        if isa(glmout[i].model, CoxModel)
            tdata[i] = Survival.coeftable(glmout[i])
        else
            if robust == nothing
                tdata[i] = coeftable(glmout[i])
            else
                tdata[i] = rcoeftable(glmout[i], robust=robust, adjust=adjust)
            end
        end
        if robust == nothing
            tconfint[i] = confint(glmout[i]) 
        else
            tconfint[i] = rconfint(glmout[i], robust=robust, adjust=adjust) 
        end

        # build a vector of covariate names
        for (k,nm) in enumerate(tdata[i].rownms)

            # save the location of the variable in loc dictionary
            if haskey(loc,i) == false
                loc[i] = Dict()
            end

            if occursin(":",nm)
                (vn,val) = split(nm,": ")
            else
                vn = nm
                val = ""
            end
            loc[i][(vn,val)] = k # location of the variable in the tdata.cols[1]

            if in(vn, covariates) == false
                push!(covariates,vn)
            end
            
            if val != ""
                if haskey(vvalues,vn)
                    vvalues[vn] = OrderedSet(vcat(collect(vvalues[vn]),val))
                else
                    vvalues[vn] = OrderedSet([val])
                end
            end
        end
    end

    # expand covariates and vvalues
    covars = Vector{String}()
    valvec = Vector{String}()
    for vn in covariates
        ncat = 1
        vval = [""]
        if haskey(vvalues, vn)
            ncat = length(vvalues[vn])
            vval = collect(vvalues[vn])
        end
        for i = 1:ncat
            push!(covars, vn)
            push!(valvec, vval[i])
        end
    end

    # go through each variable and construct variable name and value label arrays
    # vals = Vector{String}(undef,nrows)
    nrows = length(covars)
    nlev = zeros(Int, nrows)
    npred = [dof(m) for m in glmout]
    varname = deepcopy(covars)

    for i = 1:length(covars)

        vn = covars[i]

        # count the number of levels in a categorical variable
        if haskey(vvalues, vn)
            nlev[i] = length(vvalues[vn])
        end

        # use labels if exist
        if labels != nothing && haskey(labels, Symbol(vn))
            varname[i] = labels[Symbol(vn)]
        else
            varname[i] = vn
        end

    end

    # write table
    lastvarname = ""

    for i = 1:length(covars)
        if covars[i] != lastvarname
            # output cell boundaries only and go to the next line

            # variable name - expanded if `labels` is provided
            t.write_string(r,c,varname[i],formats[:model_name])

            if nlev[i] > 1

                for k = 1:num_models
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
                    c += 4
                end
                c = col
                r += 1

                t.write_string(r,c,valvec[i],formats[:varname_1indent])

            else
                if valvec[i] != "" && in(valvec[i], ["Yes","1"]) == false
                    t.write_string(r,c,string(varname[i], ": ", valvec[i]),formats[:model_name])
                else
                    t.write_string(r,c,varname[i],formats[:model_name])
                end
            end
        else
            t.write_string(r,c,valvec[i],formats[:varname_1indent])
        end

        for j=1:num_models

            # find the index number for each coeftable row
            if haskey(loc[j], (covars[i],valvec[i]))
                ri = loc[j][(covars[i],valvec[i])] # findfirst(x->x == covariates[i],tdata[j].rownms)
            else
                ri = 0
            end

            if ri == 0
                # this variable is not in the model
                # print empty cells and then move onto the next model

                t.write_string(r,c+1,"",formats[:or_fmt])
                t.write_string(r,c+2,"",formats[:cilb_fmt])
                t.write_string(r,c+3,"",formats[:ciub_fmt])
                t.write_string(r,c+4,"",formats[:p_fmt])

                c += 4
                continue
            end


    	    # estimates
            ret = ri <= npred[j] ? tdata[j].cols[1][ri] : NaN
            if eform == true
                t.write(r,c+1,exp(ret),formats[:or_fmt])
            else
                t.write(r,c+1,ret,formats[:or_fmt])
            end

            if ci == true

                retlo = ri <= npred[j] ? tconfint[j][ri,1] : NaN
                rethi = ri <= npred[j] ? tconfint[j][ri,2] : NaN
                if eform == true

                	# 95% CI Lower
                    t.write(r, c + 2, exp(retlo), formats[:cilb_fmt])

                	# 95% CI Upper
                    t.write(r, c + 3, exp(rethi), formats[:ciub_fmt])
                else
                    # 95% CI Lower
                    t.write(r, c + 2, retlo, formats[:cilb_fmt])

                	# 95% CI Upper
                    t.write(r, c + 3, rethi, formats[:ciub_fmt])
                end
            else
                # SE
                if eform == true
            	    t.write(r,c+2,ri <= npred[j] ? exp(tdata[j].cols[1][ri])*tdata[j].cols[2][ri] : "",formats[:or_fmt])
                else
                    t.write(r,c+2,ri <= npred[j] ? tdata[j].cols[1][ri] : "",formats[:or_fmt])
                end

                # Z value
                t.write(r,c+3,ri <= npred[j] ? tdata[j].cols[3][ri] : "",formats[:or_fmt])

            end

            # P-Value
	        t.write(r,c+4,ri <= npred[j] ? (tdata[j].cols[4][ri] < 0.001 ? "< 0.001" : tdata[j].cols[4][ri]) : "" ,formats[:p_fmt])

            c += 4

        end

        lastvarname = covars[i]

        # update row
        r += 1
        c = col
    end

    # Write model characteristics and goodness of fit statistics
    c = col
    row2 = r
    for i=1:num_models

        # N
        t.write(r,c,"N",formats[:model_name])
        t.merge_range(r,c+1,r,c+4,nobs(glmout[i]),formats[:n_fmt_center])

        # degress of freedom
        r += 1
        t.write(r,c,"DF",formats[:model_name])
        t.merge_range(r,c+1,r,c+4,dof(glmout[i]),formats[:n_fmt_center])

        # R² or pseudo R²
        r += 1
        if isa(glmout[i].model,LinearModel)
            t.write(r,c,"R²",formats[:model_name])
            t.merge_range(r,c+1,r,c+4,r2(glmout[i]),formats[:p_fmt_center])
            t.write(r+1,c,"Adjusted R²",formats[:model_name])
            t.merge_range(r+1,c+1,r+1,c+4,adjr2(glmout[i]),formats[:p_fmt_center])
        
            # r += 2
	        r = row2
	        c += 4
	
	    elseif !isa(glmout[i].model, CoxModel)

            # Logistic regression
            if isa(linkfun[i],LogitLink)
                t.write(r,c,"Pseudo R² (MacFadden)",formats[:model_name])
                t.merge_range(r,c+1,r,c+4,r2(glmout[i], :McFadden),formats[:p_fmt_center])
                t.write(r+1,c,"Pseudo R² (Nagelkerke)",formats[:model_name])
                t.merge_range(r+1,c+1,r+1,c+4,r2(glmout[i], :Nagelkerke),formats[:p_fmt_center])

                # -2 log-likelihood
                t.write(r+2,c,"-2 Log-Likelihood",formats[:model_name])
                t.merge_range(r+2,c+1,r+2,c+4,deviance(glmout[i]),formats[:p_fmt_center])

                # Hosmer-Lemeshow GOF test
                t.write(r+3,c,"Hosmer-Lemeshow Chisq Test (df), p-value",formats[:model_name])
                hl = hltest(glmout[i])
                t.merge_range(r+3,c+1,r+3,c+4,string(round(hl[1],digits=4)," (",hl[2],"); p = ",round(hl[3],digits=4)),formats[:p_fmt_center])

                # ROC (c-statistic)
                t.write(r+4,c,"Area under the ROC Curve",formats[:model_name])
                _roc = auc(roc(glmout[i].model.rr.y,predict(glmout[i])))
                t.merge_range(r+4,c+1,r+4,c+4,round(_roc,digits=4),formats[:p_fmt_center])

                r += 5
            end

            # AIC & BIC
            t.write(r,c,"AIC",formats[:model_name])
            t.merge_range(r,c+1,r,c+4,aic(glmout[i]),formats[:p_fmt_center])

            r += 1
            t.write(r,c,"BIC",formats[:model_name])
            t.merge_range(r,c+1,r,c+4,bic(glmout[i]),formats[:p_fmt_center])

            r = row2
            c += 4
        else

            # For coxModel
            # AIC & BIC
            t.write(r,c,"AIC",formats[:model_name])
            t.merge_range(r,c+1,r,c+4,aic(glmout[i]),formats[:p_fmt_center])

            r += 1
            t.write(r,c,"BIC",formats[:model_name])
            t.merge_range(r,c+1,r,c+4,bic(glmout[i]),formats[:p_fmt_center])

            r = row2
            c += 4
        end
    end
end
function mglmxls(glmout,
    wbook::AbstractString,
    wsheet::AbstractString;
    mtitle::Union{String,Nothing} = nothing,
    labels::Dict = nothing,
    eform::Bool = false,
    ci = true,
    row = 0,
    col = 0)

    xlsxwriter = pyimport("xlsxwriter")

    mglmxls(glmout,xlsxwriter.Workbook(wbook),wsheet,labels=labels,mtitle=mtitle,eform=eform,ci=ci,row=row,col=col)
end
