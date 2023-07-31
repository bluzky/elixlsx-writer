defmodule ElixlsxWriter do
  @moduledoc false
  alias Elixlsx.Compiler
  alias ElixlsxWriter.SheetWriter

  defstruct workbook: nil, wci: nil, sheet_writers: %{}, temp_dir: nil, output_file: nil

  def init(workbook, output_file) do
    wci = Compiler.make_workbook_comp_info(workbook)
    tmp_dir = Path.join(System.tmp_dir!(), "xlsx-#{:rand.uniform(10_000_000)}")

    sheet_writers =
      workbook.sheets
      |> Enum.zip(wci.sheet_info)
      |> Map.new(fn {sheet, sci} ->
        {sheet.name, SheetWriter.init(sheet, sci, wci, tmp_dir)}
      end)

    %ElixlsxWriter{
      workbook: workbook,
      wci: wci,
      temp_dir: tmp_dir,
      sheet_writers: sheet_writers,
      output_file: output_file
    }
  end

  def increment_write(%ElixlsxWriter{} = writer, sheet_name, rows) do
    wci = Compiler.compinfo_from_rows(writer.wci, rows)
    sheet_writer = SheetWriter.write(writer.sheet_writers[sheet_name], rows, wci)
    %{writer | wci: wci, sheet_writers: Map.put(writer.sheet_writers, sheet_name, sheet_writer)}
  end

  def finalize(%ElixlsxWriter{} = writer) do
    # complete writing all sheets
    Enum.each(writer.sheet_writers, fn {_, sheet_writer} ->
      SheetWriter.finalize(sheet_writer)
    end)

    sheet_files =
      Enum.map(writer.sheet_writers, fn {_, sheet_writer} ->
        sheet_writer.file_path
      end)

    files =
      writer.workbook
      |> Elixlsx.Writer.create_files(writer.wci)
      |> Enum.map(fn {file, content} ->
        path = Path.join(writer.temp_dir, file)

        if path not in sheet_files do
          ElixlsxWriter.Helpers.create_directory_if_not_exists(path)
          File.write!(path, content)
          path
        end
      end)
      # remove sheet files from the list of files to be zipped
      |> Enum.reject(&is_nil(&1))
      |> Enum.concat(sheet_files)
      |> Enum.map(fn path -> path |> Path.relative_to(writer.temp_dir) |> String.to_charlist() end)

    rs = :zip.create(to_charlist(writer.output_file), files, cwd: writer.temp_dir)

    cleanup(writer)
    rs
  end

  def discard(%ElixlsxWriter{} = writer) do
    Enum.each(writer.sheet_writers, fn {_, sheet_writer} ->
      SheetWriter.discard(sheet_writer)
    end)

    cleanup(writer)
  end

  defp cleanup(%ElixlsxWriter{} = writer) do
    File.rm_rf(writer.temp_dir)
  end
end
