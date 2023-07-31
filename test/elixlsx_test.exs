defmodule ElixlsxTest do
  use ExUnit.Case

  alias Elixlsx2.Compiler.StringDB
  alias Elixlsx2.Sheet
  alias Elixlsx2.Style.Font
  alias Elixlsx2.Workbook
  alias Elixlsx2.XMLTemplates

  require Record

  Record.defrecord(
    :xmlAttribute,
    Record.extract(:xmlAttribute, from_lib: "xmerl/include/xmerl.hrl")
  )

  Record.defrecord(:xmlElement, Record.extract(:xmlElement, from_lib: "xmerl/include/xmerl.hrl"))
  Record.defrecord(:xmlText, Record.extract(:xmlText, from_lib: "xmerl/include/xmerl.hrl"))

  doctest Elixlsx2
  doctest Elixlsx2.Sheet
  doctest Elixlsx2.Util, import: true
  doctest Elixlsx2.XMLTemplates
  doctest Elixlsx2.Color
  doctest Elixlsx2.Style.Border
  doctest Elixlsx2.Style.BorderStyle

  def xpath(el, path) do
    :xmerl_xpath.string(to_charlist(path), el)
  end

  defp xml_inner_strings(xml, path) do
    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    Enum.map(
      xpath(xmerl, path),
      fn element ->
        Enum.reduce(xmlElement(element, :content), "", fn text, acc ->
          acc <> to_text(text)
        end)
      end
    )
  end

  defp to_text(xml_text) do
    xmlText(value: value) = xml_text
    to_string(value)
  end

  test "basic StringDB functionality" do
    sdb =
      %StringDB{}
      |> StringDB.register_string("Hello")
      |> StringDB.register_string("World")
      |> StringDB.register_string("Hello")

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(sdb))

    assert xml_inner_strings(xml, ~c"/sst/si/t") == ["Hello", "World"]
  end

  test "xml escaping StringDB functionality" do
    sdb = StringDB.register_string(%StringDB{}, "Hello World & Goodbye Cruel World")

    xml = XMLTemplates.make_xl_shared_strings(StringDB.sorted_id_string_tuples(sdb))

    assert xml_inner_strings(xml, ~c"/sst/si/t") == ["Hello World & Goodbye Cruel World"]
  end

  test "font color" do
    xml =
      [color: "#012345"]
      |> Font.from_props()
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [color] = :xmerl_xpath.string(~c"/font/color/@rgb", xmerl)

    assert xmlAttribute(color, :value) == ~c"FF012345"
  end

  test "font name" do
    xml =
      [font: "Arial"]
      |> Font.from_props()
      |> Font.get_stylexml_entry()

    {xmerl, []} = :xmerl_scan.string(String.to_charlist(xml))

    [name] = :xmerl_xpath.string(~c"/font/name/@val", xmerl)

    assert xmlAttribute(name, :value) == ~c"Arial"
  end

  test "blank sheet name" do
    sheet1 = Sheet.with_name("")

    assert_raise ArgumentError, "The sheet name cannot be blank.", fn ->
      Elixlsx2.write_to(%Workbook{sheets: [sheet1]}, "test.xlsx")
    end
  end

  test "too long sheet name" do
    sheet1 = Sheet.with_name("This is a very looong sheet name")

    assert_raise ArgumentError,
                 ~r/The sheet name .* is too long. Maximum 31 chars allowed for name./,
                 fn ->
                   Elixlsx2.write_to(%Workbook{sheets: [sheet1]}, "test.xlsx")
                 end
  end

  test "invalid chars in sheet name" do
    sheet_names = ~W(: \ / ? * [ ])

    Enum.each(sheet_names, fn name ->
      sheet1 = Sheet.with_name(name)

      assert_raise ArgumentError,
                   ~r/The sheet name .* contains following invalid characters:/,
                   fn ->
                     Elixlsx2.write_to(%Workbook{sheets: [sheet1]}, "test.xlsx")
                   end
    end)
  end
end
