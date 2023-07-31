defmodule ElixlsxWriter.SheetWriter do
  @moduledoc false
  alias Elixlsx.Compiler.SheetCompInfo
  alias ElixlsxWriter.SheetRenderer
  alias ElixlsxWriter.SheetWriter

  defstruct file_path: nil, sheet: nil, last_idx: 0, file_handler: nil

  def new(sheet, %SheetCompInfo{} = sci, directory) do
    %SheetWriter{sheet: sheet, file_path: Path.join(directory, sheet_full_path(sci))}
  end

  def initialize(%SheetWriter{} = writer, wci) do
    data = SheetRenderer.render_sheet_header(writer.sheet)
    ElixlsxWriter.Helpers.create_directory_if_not_exists(writer.file_path)
    file = File.open!(writer.file_path, [:write])
    writer = %{writer | file_handler: file}
    IO.write(file, data)
    write(writer, writer.sheet.rows, wci)
  end

  def write(%SheetWriter{} = writer, data, wci) do
    row_count = length(data)
    data = SheetRenderer.render_sheet_data(writer.sheet, data, wci, writer.last_idx + 1)
    IO.write(writer.file_handler, data)
    %{writer | last_idx: writer.last_idx + row_count}
  end

  def finalize(%SheetWriter{} = writer) do
    data = SheetRenderer.render_sheet_footer(writer.sheet)
    IO.write(writer.file_handler, data)
    File.close(writer.file_handler)
  end

  def discard(%SheetWriter{} = writer) do
    File.close(writer.file_handler)
    writer
  end

  defp sheet_full_path(sci) do
    "xl/worksheets/#{sci.filename}"
  end
end
