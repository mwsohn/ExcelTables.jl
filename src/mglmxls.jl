"""
    mglmxls(glmout::Vector{DataFrames.DataFrameRegressionModel}, workbook::PyObject, worksheet::AbstractString; label_dict::Union{Nothing,Dict}=nothing,mtitle::Union{Vector,Nothing}=nothing,eform=false,ci=true, row = 0, col =0)

Outputs multiple GLM regression tables side by side to an excel spreadsheet.
To use this function, `PyCall` is required with a working version python and
a python package called `xlsxwriter` installed. If a label is found for a variable
or a value of a variable in a `label_dict`, the label will be output. Options are:

- `glmout`: a vector of GLM regression models
- `workbook`: a returned value from xlsxwriter.Workbook() function (see an example below)
- `worksheet`: a string for the worksheet name
- `label_dict`: an option to specify a `label` dictionary (see an example below)
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

julia> @pyimport xlsxwriter

julia> wb = xlsxwriter.Workbook("test_workbook.xlsx")
PyObject <xlsxwriter.workbook.Workbook object at 0x000000002A628E80>

julia> glmxls(olsmodels,wb,"OLS1",label_dict = label)

julia> bivairatexls(df,:incomecat,[:age,:race,:male,:bmicat],wb,"Bivariate",label_dict = label)

Julia> wb[:close]()
```

# Example 2
Alternatively, one can create a spreadsheet file directly. `PyCall` or `@pyimport`
does not need to be called before the function.

```
julia> glmxls(olsmodels,"test_workbook.xlsx","OLS1",label_dict = label)
```

# Example 3
A `label` dictionary is a collection of dictionaries, two of which are `variable` and `value`
dictionaries. The label dictionary can be created as follows:
```jldoctest
julia> label = Dict()
Dict{Any,Any} with 0 entries

julia> label["variable"] = Dict(
    "age" => "Age at baseline",
    "race" => "Race/ethnicity",
    "male" => "Male sex",
    "bmicat" => "Body Mass Index Category")
Dict{String,String} with 4 entries:
  "male"   => "Male sex"
  "race"   => "Race/ethnicity"
  "bmicat" => "Body Mass Index Category"
  "age"    => "Age at baseline"

julia> label["value"] = Dict()
Dict{Any,Any} with 0 entries

julia> label["value"]["race"] = Dict(
    1 => "White",
    2 => "Black",
    3 => "Hispanic",
    4 => "Other")
Dict{Int64,String} with 4 entries:
  4 => "Other"
  2 => "Black"
  3 => "Hispanic"
  1 => "White"

julia> label["value"]["bmicat"] = Dict(
    1 => "< 25 kg/m²",
    2 => "25 - 29.9",
    3 => "20 or over")
Dict{Int64,String} with 3 entries:
  2 => "25 - 29.9"
  3 => "20 or over"
  1 => "< 25 kg/m²"

 julia> label["variable"]["male"]
 "Male sex"

 julia> label["value"]["bmicat"][1]
 "< 25 kg/m²"
 ```
"""
function mglmxls(glmout,
    wbook::PyObject,
    wsheet::AbstractString;
    mtitle::Union{Vector,Nothing} = nothing,
    labels::Union{Nothing,Label} = nothing,
    eform::Bool = false,
    ci = true,
    row = 0,
    col = 0)

    num_models = length(glmout)

    modelstr = Vector(num_models)
    otype = Vector(num_models)
    linkfun = Vector(num_models)

    for i=1:num_models

        # check if the models are GLM outputs
        modelstr[i] = string(typeof(glmout[i]))

        # if (typeof(modelstr[i]) <: RegressionModel) == false
        #     error("This is not a regression model: ",i)
        # end

        if ismatch(r"GLM\.GeneralizedLinearModel",modelstr[i])
            distrib = replace(modelstr[i],r".*Distributions\.(Normal|Bernoulli|Binomial|Bernoulli|Gamma|Normal|Poisson)\{.*",s"\1")
            linkfun = replace(modelstr[i],r".*,GLM\.(CauchitLink|CloglogLink|IdentityLink|InverseLink|LogitLink|LogLink|ProbitLink|SqrtLink)\}.*",s"\1")
        end

        otype[i] = "Estimate"
        if eform == true
            if distrib in ("Bernoulli","Binomial") && linkfun == "LogitLink"
                otype[i] = "OR"
            elseif distrib == "Binomial" && linkfun == "LogLink"
                otype[i] = "RR"
            elseif distrib == "Poisson" && linkfun == "LogLink"
                otype[i] = "IRR"
            else
                otype[i] = "exp(Est)"
            end
        end
    end

    if mtitle == nothing
        mtitle = Vector{String}(num_models)

        # assign dependent variables
        for i=1:num_models
            mtitle[i] = label_dict != nothing ? varlab[glmout[i].mf.terms.eterms[1]] : glmout[i].mf.terms.eterms[1]
        end
    end

    # create a worksheet
    t = wbook[:add_worksheet](wsheet)

    # attach formats to the workbook
    formats = ExcelTables.attach_formats(wbook)

    # starting location in the worksheet
    r = row
    c = col

    # set column widths
    t[:set_column](c,c,40)
    t[:set_column](c+1,c+4*num_models,7)

    t[:merge_range](r,c,r+1,c,"Variable",formats[:heading])
    for i=1:num_models
        if ci == true
            t[:merge_range](r,c+1,r,c+4,mtitle[i],formats[:heading])
            t[:merge_range](r+1,c+1,r+1,c+3,string(otype[i]," (95% CI)"),formats[:heading])
            t[:write_string](r+1,c+4,"P-Value",formats[:heading])
        else
            t[:write_string](r,c+1,otype[i],formats[:heading])
            t[:write_string](r,c+2,"SE",formats[:heading])
            t[:write_string](r,c+3,"Z Value",formats[:heading])
            t[:write_string](r,c+4,"P-Value",formats[:heading])
        end
        c += 4
    end

    #------------------------------------------------------------------
    r += 2
    c = col

    # collate variables
    covariates = Vector{String}()
    tdata = Vector(num_models)
    tconfint = Vector(num_models)

    for i=1:num_models
        tdata[i] = coeftable(glmout[i])
        tconfint[i] = confint(glmout[i])

        for nm in tdata[i].rownms
            if in(nm, covariates) == false
                push!(covariates,nm)
            end
        end
    end

    # go through each variable and construct variable name and value label arrays
    nrows = length(covariates)
    varname = Vector{String}(nrows)
    vals = Vector{String}(nrows)
    nlev = zeros(Int,nrows)

    for i = 1:nrows
    	# variable name
        # parse varname to separate variable name from value
        if contains(covariates[i]," & ")
            varname[i] = covariates[i]
            vals[i] = ""
        elseif contains(covariates[i],":")
            (varname[i],vals[i]) = split(covariates[i],": ")
        else
            varname[i] = covariates[i]
            vals[i] = ""
        end

        # use labels if exist
        if labels != nothing
            sv = Symbol(varname[i])
            if vals[i] != ""
                valn = vals[i] == "true" ? 1 : parse(Int,vals[i])
                vals[i] = vallab(labels,sv,valn)
            end
            varname[i] = varlab(labels,sv)
        end
    end
    for i = 1:nrows
        nlev[i] = ExcelTables.countlev(varname[i],varname)
    end

    # write table
    lastvarname = ""

    for i = 1:nrows
        if varname[i] != lastvarname
            # output cell boundaries only and go to the next line
            if nlev[i] > 1

                # variable name
                t[:write_string](r,c,varname[i],formats[:model_name])

                for k = 1:num_models
                    if ci == true
                        t[:write](r,c+1,"",formats[:empty_right])
                        t[:write](r,c+2,"",formats[:empty_both])
                        t[:write](r,c+3,"",formats[:empty_left])
                        t[:write](r,c+4,"",formats[:p_fmt])
                    else
                        t[:write](r,c+1,"",formats[:empty_border])
                        t[:write](r,c+2,"",formats[:empty_border])
                        t[:write](r,c+3,"",formats[:empty_border])
                        t[:write](r,c+4,"",formats[:p_fmt])
                    end
                    c += 4
                end
                c = col
                r += 1
                t[:write_string](r,c,vals[i],formats[:varname_1indent])

            else
                if vals[i] != "" && vals[i] != "Yes"
                    t[:write_string](r,c,string(varname[i],": ",vals[i]),formats[:model_name])
                else
                    t[:write_string](r,c,varname[i],formats[:model_name])
                end
            end
        else
            t[:write_string](r,c,vals[i],formats[:varname_1indent])
        end

        for j=1:num_models

            # find the index number for each coeftable row
            ri = findfirst(x->x == covariates[i],tdata[j].rownms)

            if ri == 0
                # this variable is not in the model
                # print empty cells and then move onto the next model

                t[:write_string](r,c+1,"",formats[:or_fmt])
                t[:write_string](r,c+2,"",formats[:cilb_fmt])
                t[:write_string](r,c+3,"",formats[:ciub_fmt])
                t[:write_string](r,c+4,"",formats[:p_fmt])

                c += 4
                continue
            end


    	    # estimates
            if eform == true
        	    t[:write](r,c+1,exp(tdata[j].cols[1][i]),formats[:or_fmt])
            else
                t[:write](r,c+1,tdata[j].cols[1][i],formats[:or_fmt])
            end

            if ci == true

                if eform == true
                	# 95% CI Lower
                	t[:write](r,c+2,exp(tconfint[j][i,1]),formats[:cilb_fmt])

                	# 95% CI Upper
                	t[:write](r,c+3,exp(tconfint[j][i,2]),formats[:ciub_fmt])
                else
                    # 95% CI Lower
                	t[:write](r,c+2,tconfint[j][i,1],formats[:cilb_fmt])

                	# 95% CI Upper
                	t[:write](r,c+3,tconfint[j][i,2],formats[:ciub_fmt])
                end
            else
                # SE
                if eform == true
            	    t[:write](r,c+2,exp(tdata[j].cols[1][i])*tdata[j].cols[2][i],formats[:or_fmt])
                else
                    t[:write](r,c+2,tdata[j].cols[1][i],formats[:or_fmt])
                end

                # Z value
                t[:write](r,c+3,tdata[j].cols[3][i],formats[:or_fmt])

            end

            # P-Value
	        t[:write](r,c+4,tdata[j].cols[4][i].v < 0.001 ? "< 0.001" : tdata[j].cols[4][i].v ,formats[:p_fmt])




            c += 4

        end

        lastvarname = varname[i]

        # update row
        r += 1
        c = col
    end

    # Write model characteristics and goodness of fit statistics
    c = col
    row2 = r
    for i=1:num_models

        # N
        t[:write](r,c,"N",formats[:model_name])
        t[:merge_range](r,c+1,r,c+4,nobs(glmout[i]),formats[:n_fmt_center])

        # degress of freedom
        r += 1
        t[:write](r,c,"DF",formats[:model_name])
        t[:merge_range](r,c+1,r,c+4,dof(glmout[i]),formats[:n_fmt_center])

        # R² or pseudo R²
        r += 1
        if linkfun == "LogitLink"
            t[:write](r,c,"Pseudo R² (MacFadden)",formats[:model_name])
            t[:merge_range](r,c+1,r,c+4,macfadden(glmout[i]),formats[:p_fmt_center])
            t[:write](r+1,c,"Pseudo R² (Nagelkerke)",formats[:model_name])
            t[:merge_range](r+1,c+1,r+1,c+4,nagelkerke(glmout[i]),formats[:p_fmt_center])

            # -2 log-likelihood
            t[:write](r+2,c,"-2 Log-Likelihood",formats[:model_name])
            t[:merge_range](r+2,c+1,r+2,c+4,deviance(glmout[i]),formats[:p_fmt_center])

            # Hosmer-Lemeshow GOF test
            t[:write](r+3,c,"Hosmer-Lemeshow Chisq Test (df), p-value",formats[:model_name])
            hl = hltest(glmout[i])
            t[:merge_range](r+3,c+1,r+3,c+4,string(round(hl[1],4)," (",hl[2],"); p = ",round(hl[3],4)),formats[:p_fmt_center])

            # ROC (c-statistic)
            t[:write](r+4,c,"Area under the ROC Curve",formats[:model_name])
            roc = auc(glmout[i].model.rr.y,predict(glmout[i]))
            t[:merge_range](r+4,c+1,r+4,c+4,round(roc,4),formats[:p_fmt_center])

            r += 5
        else
            t[:write](r,c,"R²",formats[:model_name])
            t[:merge_range](r,c+1,r,c+4,r2(glmout[i]),formats[:p_fmt_center])
            t[:write](r+1,c,"Adjusted R²",formats[:model_name])
            t[:merge_range](r+1,c+1,r+1,c+4,adjr2(glmout[i]),formats[:p_fmt_center])

            r += 2
        end

        # AIC & BIC
        t[:write](r,c,"AIC",formats[:model_name])
        t[:merge_range](r,c+1,r,c+4,aic(glmout[i]),formats[:p_fmt_center])

        r += 1
        t[:write](r,c,"BIC",formats[:model_name])
        t[:merge_range](r,c+1,r,c+4,bic(glmout[i]),formats[:p_fmt_center])

        r = row2
        c += 4
    end
end
function mglmxls(glmout,
    wbook::AbstractString,
    wsheet::AbstractString;
    mtitle::Union{String,Nothing} = nothing,
    labels::Union{Nothing,Label} = nothing,
    eform::Bool = false,
    ci = true,
    row = 0,
    col = 0)

    xlsxwriter = pyimport("xlsxwriter")

    mglmxls(glmout,xlsxwriter[:Workbook](wbook),wsheet,labels=labels,mtitle=mtitle,eform=eform,ci=ci,row=row,col=col)
end
