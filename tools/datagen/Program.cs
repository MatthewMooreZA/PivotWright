using PivotWright;
string dir = args[0];
// Modes that hide the columns of Excel's samples: "hidden" = H1_hidden_columns.xlsx (ten columns), "h8" = H8_hide_one_column.xlsx
// (Sales.Qty), "h9" = H9_hide_one_hierarchy_level.xlsx (Products.Product).
var hiddenSet = (args.Length > 2 ? args[2] : "") switch
{
    "hidden" => new HashSet<string> { "Dates.Date", "Dates.Year", "Dates.Month", "Dates.MonthNumber", "Products.ProductKey", "Products.Category", "Products.CategoryOrder", "Products.Product", "Sales.ProductKey", "Sales.Date" },
    "h8" => new HashSet<string> { "Sales.Qty" },
    "h9" => new HashSet<string> { "Products.Product" },
    _ => new HashSet<string>(),
};
ColumnDefinition Col(string table, string name, ColumnDataType type, string? sortBy = null) => new(name, type) { SortByColumn = sortBy, IsHidden = hiddenSet.Contains(table + "." + name) };
using var m = new DataModelWriter(new DataModelOptions { Lcid = 7177 });
var p = m.CreateTable("Products", new[] { Col("Products", "ProductKey", ColumnDataType.Int64), Col("Products", "Category", ColumnDataType.String, "CategoryOrder"), Col("Products", "CategoryOrder", ColumnDataType.Int64), Col("Products", "Product", ColumnDataType.String) });
string[] cats = { "Bikes", "Clothing", "Accessories" };
int k = 1;
for (int c = 0; c < 3; c++) for (int i = 1; i <= 4; i++) p.AddRow(k++, cats[c], c + 1, $"{cats[c]} item {i}");
bool generated = args.Length > 2 && args[2] == "gendate";
if (generated)
    m.AddDateTable("Dates",
        new DateOnly(args.Length > 4 ? int.Parse(args[4]) : 2025, 1, 1), new DateOnly(args.Length > 5 ? int.Parse(args[5]) : 2025, 12, 31),
        new DateTableOptions
        {
            YearEndMonth = args.Length > 3 ? int.Parse(args[3]) : 12,
            YearLabel = args.Length > 6 ? args[6] : "Fiscal",
            CompleteYears = !(args.Length > 7 && args[7].Contains("exact")),
            IncludeYearMonth = !(args.Length > 7 && args[7].Contains("noym")),
            HideSortColumns = !(args.Length > 7 && args[7].Contains("nohide")),
        });
else
{
    var d = m.CreateTable("Dates", new[] { Col("Dates", "Date", ColumnDataType.DateTime), Col("Dates", "Year", ColumnDataType.Int64), Col("Dates", "Month", ColumnDataType.String, "MonthNumber"), Col("Dates", "MonthNumber", ColumnDataType.Int64) });
    for (int i = 0; i < 365; i++) { var dt = new DateTime(2025, 1, 1).AddDays(i); d.AddRow(dt, dt.Year, dt.ToString("MMMM", System.Globalization.CultureInfo.InvariantCulture), dt.Month); }
}
var s = m.CreateTable("Sales", new[] { Col("Sales", "ProductKey", ColumnDataType.Int64), Col("Sales", "Date", ColumnDataType.DateTime), Col("Sales", "Qty", ColumnDataType.Int64), Col("Sales", "Amount", ColumnDataType.Currency) });
for (int i = 0; i < 1000; i++) s.AddRow(1 + i % 12, new DateTime(2025, 1, 1).AddDays(i % 365), 1 + i % 5, (decimal)(10 + i % 40));
m.AddRelationship("Sales", "ProductKey", "Products", "ProductKey");
m.AddRelationship("Sales", "Date", "Dates", "Date");
m.AddHierarchy("Products", "Product Hierarchy", "Category", "Product");
if (!generated) m.AddHierarchy("Dates", "Calendar", "Year", "Month");
m.AddMeasure("Sales", "Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
if (args.Length > 2 && args[2] == "date") m.MarkAsDateTable("Dates", "Date");
m.Save(Path.Combine(dir, args[1] + ".xlsx"));
