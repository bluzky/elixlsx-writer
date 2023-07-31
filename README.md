# A writer for Elixlsx library that support writing large excel file


## Example

```elixir
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
```

## Benchmark

Code is in "benchmark.exs"
Result for large file:

```
Operating System: macOS
CPU Information: Apple M1
Number of Available Cores: 8
Available memory: 16 GB
Elixir 1.14.3
Erlang 25.2

Benchmark suite executing with the following configuration:
warmup: 2 s
time: 1 s
memory time: 2 s
reduction time: 0 ns
parallel: 1
inputs: none specified
Estimated total run time: 10 s

Name                    ips        average  deviation         median         99th %
ElixlsxWriter          0.65         1.55 s     ±0.00%         1.55 s         1.55 s
Elixlsx                0.24         4.20 s     ±0.00%         4.20 s         4.20 s

Comparison:
ElixlsxWriter          0.65
Elixlsx                0.24 - 2.71x slower +2.65 s

Memory usage statistics:

Name             Memory usage
ElixlsxWriter       769.77 MB
Elixlsx             829.07 MB - 1.08x memory usage +59.30 MB

**All measurements for memory usage were the same**
```
