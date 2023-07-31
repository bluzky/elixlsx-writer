defmodule ElixlsxWriter.Helpers do
  @moduledoc false
  @doc """
  Give a file path, create the directory if it doesn't exist.
  """
  def create_directory_if_not_exists(file) do
    dir = Path.dirname(file)

    if !File.dir?(dir) do
      File.mkdir_p!(dir)
    end
  end
end
