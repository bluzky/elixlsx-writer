NimbleCSV.define(MyParser, separator: ",", escape: "\"")



alias Elixlsx.Workbook
alias Elixlsx.Sheet


Benchee.run(
  %{
    "ElixlsxWriter" => fn ->
      sheet = %Sheet{name: "Sheet 1", rows: []}
      sheet2 = %Sheet{name: "Sheet 2", rows: []}
      sheet3 = %Sheet{name: "Sheet 3", rows: []}
      wb = Workbook.append_sheet(%Workbook{}, sheet)
      |> Workbook.append_sheet(sheet2)
      |> Workbook.append_sheet(sheet3)

      writer = ElixlsxWriter.init(wb, "output/hello2.xlsx")

      writer = "priv/test.csv"
      |> File.stream!(read_ahead: 100_000)
      |> MyParser.parse_stream()
      |> Stream.chunk_every(1000)
      |> Enum.reduce(writer, fn chunk, writer ->
        ElixlsxWriter.increment_write(writer, "Sheet 1", chunk)
        |> ElixlsxWriter.increment_write( "Sheet 2", chunk)
        |> ElixlsxWriter.increment_write( "Sheet 3", chunk)
      end)
      |> ElixlsxWriter.finalize()

      IO.puts "Done writing"
    end,
    "Elixlsx" => fn ->
      data = "priv/test.csv"
      |> File.stream!(read_ahead: 100_000)
      |> MyParser.parse_stream()
      |> Enum.to_list()
      sheet = %Sheet{name: "Sheet 1", rows: data}
      sheet2 = %Sheet{name: "Sheet 2", rows: data}
      sheet3 = %Sheet{name: "Sheet 3", rows: data}
      wb = Workbook.append_sheet(%Workbook{}, sheet)
      |> Workbook.append_sheet(sheet2)
      |> Workbook.append_sheet(sheet3)

      Elixlsx.write_to(wb, "output/hello.xlsx")
      IO.puts "Done writing"
    end,
  },
  time: 1,
  memory_time: 2
)
