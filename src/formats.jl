
format_defs = Dict()

format_defs[:heading] = Dict(
	"bold" => true,
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "center",
	"border" => true
)

format_defs[:text] = Dict(
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"border" => true
)

format_defs[:heading_right] = Dict(
	"bold" => true,
	"font" => "Arial",
	"size" => 10,
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)

format_defs[:heading_left] = Dict(
	"bold" => true,
	"font" => "Arial",
	"size" => 10,
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)

format_defs[:model_name] = Dict(
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"border" => true
)

format_defs[:varname_1indent] = Dict(
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"border" => true,
	"indent" => 1
)


format_defs[:varname_2indent] = Dict(
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"border" => true,
	"indent" => 2
)


format_defs[:n_fmt_right] = Dict(
	"num_format" => "#,##0",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)


format_defs[:n_fmt_left_parens] = Dict(
	"num_format" => "(#,##0)",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"right" => true, "bottom" => true, "top" => true
)

format_defs[:f_date] = Dict(
	"num_format" => "mm/dd/yyyy",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:f_datetime] = Dict(
	"num_format" => "mm/dd/yyyy hh:mm:ss",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:f_fmt_right] = Dict(
	"num_format" => "#,##0.00",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)

format_defs[:f_fmt_center] = Dict(
	"num_format" => "#,##0.0000",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "center",
	"border" => true
)

format_defs[:f_fmt] = Dict(
	"num_format" => "#,##0.00",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)


format_defs[:n_fmt] = Dict(
	"num_format" => "#,##0",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:n_fmt_center] = Dict(
	"num_format" => "#,##0",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "center",
	"border" => true
)

format_defs[:f_fmt_left_parens] = Dict(
	"num_format" => "(#,##0.00)",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)


format_defs[:pct_fmt_parens] = Dict(
	"num_format" => "(0.00%)",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)


format_defs[:pct_fmt] = Dict(
	"num_format" => "0.00%",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:or_fmt] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)

format_defs[:cilb_fmt] = Dict(
	"num_format" => "(0.000\,;(-0.000\,",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"bottom" => true, "top" => true
)

format_defs[:ciub_fmt] = Dict(
	"num_format" => "0.000)",
	"font" => "Arial",
	"font_size" => 9,
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)


format_defs[:or_fmt_red] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"font_color" => "red",
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)

format_defs[:cilb_fmt_red] = Dict(
	"num_format" => "(0.000 -",
	"font" => "Arial",
	"font_size" => 9,
	"font_color" => "red",
	"valign" => "vcenter",
	"align" => "right",
	"bottom" => true, "top" => true
)

format_defs[:ciub_fmt_red] = Dict(
	"num_format" => "0.000)",
	"font" => "Arial",
	"font_size" => 9,
	"font_color" => "red",
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)


format_defs[:or_fmt_bold] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"bold" => true,
	"valign" => "vcenter",
	"align" => "right",
	"left" => true, "bottom" => true, "top" => true
)

format_defs[:cilb_fmt_bold] = Dict(
	"num_format" => "(0.000 -",
	"font" => "Arial",
	"font_size" => 9,
	"bold" => true,
	"valign" => "vcenter",
	"align" => "right",
	"bottom" => true, "top" => true
)

format_defs[:ciub_fmt_bold] = Dict(
	"num_format" => "0.000)",
	"font" => "Arial",
	"font_size" => 9,
	"bold" => true,
	"valign" => "vcenter",
	"align" => "left",
	"right" => true, "bottom" => true, "top" => true
)

format_defs[:p_fmt] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:p_fmt_center] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "center",
	"border" => true
)

format_defs[:p_fmt2] = Dict(
	#"num_format" => "0.000",
	"font" => "Arial",
	"size" => 9,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)


format_defs[:p_fmt_red] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"font_color" => "red",
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:p_fmt2_red] = Dict(
	#"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"font_color" => "red",
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)


format_defs[:p_fmt_bold] = Dict(
	"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"bold" => true,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:p_fmt2_bold] = Dict(
	#"num_format" => "0.000",
	"font" => "Arial",
	"font_size" => 9,
	"bold" => true,
	"valign" => "vcenter",
	"align" => "right",
	"border" => true
)

format_defs[:empty_border] = Dict(
	#"num_format" => "0.000",
	"valign" => "vcenter",
	"border" => true
)

format_defs[:empty_left] = Dict(
	#"num_format" => "0.000",
	"valign" => "vcenter",
	"right" => true, "bottom" => true, "top" => true
)

format_defs[:empty_right] = Dict(
	#"num_format" => "0.000",
	"valign" => "vcenter",
	"left" => true, "bottom" => true, "top" => true
)


format_defs[:empty_both] = Dict(
	#"num_format" => "0.000",
	"valign" => "vcenter",
	"bottom" => true, "top" => true
)


function attach_formats(workbook;formats::Dict = format_defs)
	newfmts = Dict()
	for key in keys(formats)
		newfmts[key] = workbook.add_format(formats[key])
	end
	return newfmts
end
