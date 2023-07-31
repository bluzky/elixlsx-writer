defmodule ElixlsxWriter.SheetRenderer do
  @moduledoc false
  alias Elixlsx.Compiler.CellStyleDB
  alias Elixlsx.Compiler.StringDB
  alias Elixlsx.Style.CellStyle
  alias Elixlsx.Util, as: U

  def render_sheet_header(sheet) do
    grouping_info = get_grouping_info(sheet.group_rows)

    """
    <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
    <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
    <sheetPr filterMode="false">
      <pageSetUpPr fitToPage="false"/>
    </sheetPr>
    <dimension ref="A1"/>
    <sheetViews>
    <sheetView workbookViewId="0" #{make_sheet_show_grid(sheet)}>
    """ <>
      make_sheetview(sheet) <>
      """
      </sheetView>
      </sheetViews>
      <sheetFormatPr defaultRowHeight="12.8" #{make_max_outline_level_row(grouping_info.outline_lvs)} />
      """ <>
      make_cols(sheet) <>
      "<sheetData>"
  end

  def render_sheet_data(sheet, rows, wci, start_idx \\ 1) do
    grouping_info = get_grouping_info(sheet.group_rows)
    xl_sheet_rows(rows, sheet.row_heights, grouping_info, wci, start_idx)
  end

  def render_sheet_footer(sheet) do
    "</sheetData>" <>
      xl_merge_cells(sheet.merge_cells) <>
      make_data_validations(sheet.data_validations) <>
      """
      <pageMargins left="0.75" right="0.75" top="1" bottom="1.0" header="0.5" footer="0.5"/>
      </worksheet>
      """
  end

  defp make_sheet_show_grid(sheet) do
    if not sheet.show_grid_lines do
      "showGridLines=\"0\""
    end
  end

  defp make_sheetview(sheet) do
    # according to spec:
    # * when only horizontal split is applied we need to use bottomLeft
    # * when only vertical split is applied we need to use topRight
    # * and when both splits is applied, we can use bottomRight
    pane =
      case sheet.pane_freeze do
        {_row_idx, 0} ->
          "bottomLeft"

        {0, _col_idx} ->
          "topRight"

        {col_idx, row_idx} when col_idx > 0 and row_idx > 0 ->
          "bottomRight"

        _any ->
          nil
      end

    {selection_pane_attr, panel_xml} =
      case sheet.pane_freeze do
        {row_idx, col_idx} when col_idx > 0 or row_idx > 0 ->
          top_left_cell = U.to_excel_coords(row_idx + 1, col_idx + 1)

          {"pane=\"#{pane}\"",
           """
           <pane xSplit="#{col_idx}" ySplit="#{row_idx}" topLeftCell="#{top_left_cell}" activePane="#{pane}" state="frozen" />
           """}

        _any ->
          {"", ""}
      end

    panel_xml <> "<selection " <> selection_pane_attr <> " activeCell=\"A1\" sqref=\"A1\" />"
  end

  defp make_max_outline_level_row(row_outline_levels) do
    if not Enum.empty?(row_outline_levels) do
      max_outline_level_row =
        row_outline_levels
        |> Map.values()
        |> Enum.max()

      " outlineLevelRow=\"#{max_outline_level_row}\""
    end
  end

  defp xl_merge_cells([]) do
    ""
  end

  defp xl_merge_cells(merge_cells) do
    """
    <mergeCells count="#{Enum.count(merge_cells)}">
    #{Enum.map(merge_cells, fn {fromCell, toCell} -> "<mergeCell ref=\"#{fromCell}:#{toCell}\"/>" end)}
    </mergeCells>
    """
  end

  defp make_data_validations([]) do
    ""
  end

  defp make_data_validations(data_validations) do
    """
    <dataValidations count="#{Enum.count(data_validations)}">
    #{Enum.map(data_validations, &make_data_validation/1)}
    </dataValidations>
    """
  end

  defp make_data_validation({start_cell, end_cell, values}) when is_bitstring(values) do
    """
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="#{start_cell}:#{end_cell}">
      <formula1>#{values}</formula1>
    </dataValidation>
    """
  end

  defp make_data_validation({start_cell, end_cell, values}) do
    joined_values =
      values
      |> Enum.join(",")
      |> String.codepoints()
      |> Enum.chunk_every(255)
      |> Enum.join("&quot;&amp;&quot;")

    """
    <dataValidation type="list" allowBlank="1" showErrorMessage="1" sqref="#{start_cell}:#{end_cell}">
      <formula1>&quot;#{joined_values}&quot;</formula1>
    </dataValidation>
    """
  end

  defp xl_sheet_rows(data, row_heights, grouping_info, wci, start_idx) do
    rows =
      data
      |> Enum.with_index(start_idx)
      |> Enum.map(fn {row, rowidx} ->
        """
        <row r="#{rowidx}" #{get_row_height_attr(row_heights, rowidx)} #{get_row_grouping_attr(grouping_info, rowidx)}>
        #{xl_sheet_cols(row, rowidx, wci)}
        </row>
        """
      end)
    |> Enum.join()

    if (length(data) + 1) in grouping_info.collapsed_idxs do
      rows <>
        """
        <row r="#{length(data) + 1}" collapsed="1"></row>
        """
    else
      rows
    end
  end

  defp make_col({k, width, outline_level, hidden, collapsed}) do
    width_attr = if width, do: " width=\"#{width}\" customWidth=\"1\"", else: ""
    hidden_attr = if hidden, do: " hidden=\"1\"", else: ""
    outline_level_attr = if outline_level, do: " outlineLevel=\"#{outline_level}\"", else: ""
    collapsed_attr = if collapsed, do: " collapsed=\"1\"", else: ""

    ~c"<col min=\"#{k}\" max=\"#{k}\"#{width_attr}#{hidden_attr}#{outline_level_attr}#{collapsed_attr} />"
  end

  defp xl_sheet_cols(row, rowidx, wci) do
    {updated_row, _id} =
      List.foldl(row, {"", 1}, fn cell, {acc, colidx} ->
        {content, styleID, cellstyle} = split_into_content_style(cell, wci)

        if is_nil(content) do
          {acc, colidx + 1}
        else
          content =
            if CellStyle.is_date?(cellstyle) do
              U.to_excel_datetime(content)
            else
              content
            end

          cv = get_content_type_value(content, wci)

          {content_type, content_value, content_opts} =
            case cv do
              {t, v} ->
                {t, v, []}

              {t, v, opts} ->
                {t, v, opts}

              :error ->
                raise %ArgumentError{
                  message: "Invalid column content at " <> U.to_excel_coords(rowidx, colidx) <> ": " <> inspect(content)
                }
            end

          cell_xml =
            case content_type do
              :formula ->
                value = if not is_nil(content_opts[:value]), do: "<v>#{content_opts[:value]}</v>"

                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}" s="#{styleID}">
                <f>#{content_value}</f>
                #{value}
                </c>
                """

              :empty ->
                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}" s="#{styleID}"></c>
                """

              type ->
                """
                <c r="#{U.to_excel_coords(rowidx, colidx)}" s="#{styleID}" t="#{type}">
                <v>#{content_value}</v>
                </c>
                """
            end

          {acc <> cell_xml, colidx + 1}
        end
      end)

    updated_row
  end

  defp make_cols(sheet) do
    grouping_info = get_grouping_info(sheet.group_cols)

    col_indices =
      [
        Map.keys(sheet.col_widths),
        Map.keys(grouping_info.outline_lvs),
        grouping_info.hidden_idxs,
        grouping_info.collapsed_idxs
      ]
      |> Stream.concat()
      |> Enum.sort()
      |> Enum.dedup()

    if not Enum.empty?(col_indices) do
      cols =
        col_indices
        |> Enum.map(
          &make_col({
            &1,
            Map.get(sheet.col_widths, &1),
            Map.get(grouping_info.outline_lvs, &1),
            &1 in grouping_info.hidden_idxs,
            &1 in grouping_info.collapsed_idxs
          })
        )
      |> Enum.join()

      "<cols>#{cols}</cols>"
    else
      ""
    end
  end

  defp get_row_grouping_attr(gr_info, rowidx) do
    outline_level = Map.get(gr_info.outline_lvs, rowidx)

    if(outline_level, do: " outlineLevel=\"#{outline_level}\"", else: "") <>
      if(rowidx in gr_info.hidden_idxs, do: " hidden=\"1\"", else: "") <>
      if rowidx in gr_info.collapsed_idxs, do: " collapsed=\"1\"", else: ""
  end

  @typep grouping_info :: %{
           outline_lvs: %{optional(idx :: pos_integer) => lv :: pos_integer},
           hidden_idxs: MapSet.t(pos_integer),
           collapsed_idxs: MapSet.t(pos_integer)
         }
  @spec get_grouping_info([Sheet.rowcol_group()]) :: grouping_info
  defp get_grouping_info(groups) do
    ranges =
      Enum.map(groups, fn
        {%Range{} = range, _opts} -> range
        %Range{} = range -> range
      end)

    collapsed_ranges =
      groups
      |> Enum.filter(fn
        {%Range{} = _range, opts} -> opts[:collapsed]
        %Range{} = _range -> false
      end)
      |> Enum.map(fn {range, _opts} -> range end)

    # see ECMA Office Open XML Part1, 18.3.1.73 Row -> attributes -> collapsed for examples
    %{
      outline_lvs:
        ranges
        |> Stream.concat()
        |> Enum.group_by(& &1)
        |> Map.new(fn {k, v} -> {k, length(v)} end),
      hidden_idxs: collapsed_ranges |> Stream.concat() |> MapSet.new(),
      collapsed_idxs: collapsed_ranges |> Enum.map(&(&1.last + 1)) |> MapSet.new()
    }
  end

  defp get_row_height_attr(row_heights, rowidx) do
    row_height = Map.get(row_heights, rowidx)

    if row_height do
      "customHeight=\"1\" ht=\"#{row_height}\""
    end
  end

  defp split_into_content_style([h | t], wci) do
    cellstyle = CellStyle.from_props(t)

    {
      h,
      CellStyleDB.get_id(wci.cellstyledb, cellstyle),
      cellstyle
    }
  end

  defp split_into_content_style(cell, _wci), do: {cell, 0, nil}

  defp get_content_type_value(content, wci) do
    case content do
      {:excelts, num} ->
        {"n", to_string(num)}

      {:formula, x} ->
        {:formula, x}

      {:formula, x, opts} when is_list(opts) ->
        {:formula, x, opts}

      x when is_number(x) ->
        {"n", to_string(x)}

      x when is_binary(x) ->
        id = StringDB.get_id(wci.stringdb, x)

        if id == -1 do
          {:empty, :empty}
        else
          {"s", to_string(id)}
        end

      x when is_boolean(x) ->
        {"b",
         if x do
           "1"
         else
           "0"
         end}

      :empty ->
        {:empty, :empty}

      true ->
        :error
    end
  end
end
