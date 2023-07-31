defmodule Elixlsx2.Mixfile do
  use Mix.Project

  @source_url "https://github.com/bluzky/elixlsx-writer"
  @version "0.1.0"

  def project do
    [
      app: :elixlsx_writer,
      version: @version,
      elixir: "~> 1.12",
      package: package(),
      description:
        "Elixlsx-writer is a writer for Elixlsx library, supporting writing large data to xlsx file by chunks. So you don't have to load all data into memory before writing to file.",
      build_embedded: Mix.env() == :prod,
      start_permanent: Mix.env() == :prod,
      deps: deps(),
      docs: docs()
    ]
  end

  def application do
    []
  end

  defp deps do
    [
      {:elixlsx, "~> 0.5.1"},
      {:credo, "~> 1.7", only: [:dev, :test]},
      {:ex_doc, ">= 0.0.0", only: [:dev], runtime: false},
      {:dialyxir, "~> 1.3", only: [:dev], runtime: false},
      {:benchee, "~> 1.1", only: [:dev], runtime: false},
      {:nimble_csv, "~> 1.2", only: [:dev], runtime: false},
      {:styler, "~> 0.8", only: [:dev, :test], runtime: false}
    ]
  end

  defp docs do
    [
      extras: ["CHANGELOG.md", "README.md"],
      main: "readme",
      source_url: @source_url,
      source_ref: "v#{@version}"
    ]
  end

  defp package do
    [
      maintainers: ["Dzung Nguyen <bluesky.1289@gmail.com>"],
      licenses: ["MIT"],
      links: %{
        "GitHub" => @source_url
      }
    ]
  end
end
