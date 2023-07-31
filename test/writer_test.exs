defmodule WriterTest do
  use ExUnit.Case

  alias Elixlsx.Sheet
  alias Elixlsx.Workbook

  test "write_file" do
    wb = Workbook.append_sheet(%Workbook{}, %Sheet{name: "Sheet 1", rows: [["a", "b"]]})

    writer = ElixlsxWriter.init(wb, "output/test.xlsx")

    rs =
      1..100
      |> Enum.reduce(writer, fn i, writer ->
        ElixlsxWriter.increment_write(writer, "Sheet 1", [[i, i + 1]])
      end)
      |> ElixlsxWriter.finalize()

    assert {:ok, ~c"output/test.xlsx"} = rs
  end
end
