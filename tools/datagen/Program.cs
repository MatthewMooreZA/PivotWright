using System.Globalization;
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
using var m = new DataModelWriter(new DataModelOptions { Culture = CultureInfo.GetCultureInfo(7177) });
if (args.Length > 2 && args[2] == "twotable")   // two tables, one relationship, no measure, no hierarchy
{
    var pr = m.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Product", ColumnDataType.String) });
    for (int i = 1; i <= 12; i++) pr.AddRow(i, $"Item {i}");
    var sl = m.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Qty", ColumnDataType.Int64) });
    for (int i = 0; i < 500; i++) sl.AddRow(1 + i % 12, 1 + i % 5);
    m.AddRelationship(sl["ProductKey"], pr["ProductKey"]);
    m.Save(Path.Combine(dir, args[1] + ".xlsx"));
    return;
}
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
var dates = m.Tables["Dates"];
m.AddRelationship(s["ProductKey"], p["ProductKey"]);
m.AddRelationship(s["Date"], dates["Date"]);
bool plain = args.Skip(2).Contains("plain");   // no hierarchies at all ("plain", or a trailing "plain" after a calc mode)
if (!plain) p.AddHierarchy("Product Hierarchy", "Category", "Product");
if (!generated && !plain) dates.AddHierarchy("Calendar", "Year", "Month");
s.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
if (args.Length > 2 && args[2] == "date") dates.MarkAsDateTable("Date");
if (args.Length > 2 && args[2] == "calc")
{
    // calc [currency|whole|decimal|date|boolean|text]: one calculated column, as in Excel's H10 to H15 samples
    var (calcName, calcFormula, calcType) = (args.Skip(3).FirstOrDefault(a => a != "plain") ?? "currency") switch
    {
        "whole" => ("Double Qty", "[Qty]*2", ColumnDataType.Int64),
        "decimal" => ("Half Qty", "[Qty]/2", ColumnDataType.Double),
        "date" => ("Next Day", "[Date]+1", ColumnDataType.DateTime),
        "boolean" => ("Big Order", "[Qty]>2", ColumnDataType.Boolean),
        "text" => ("Label", "\"Qty \"&[Qty]", ColumnDataType.String),
        _ => ("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency),
    };
    s.AddCalculatedColumn(calcName, calcFormula, calcType);
}
// calc2, calc3, calc4: that many columns on Sales (H18, H25, H26); calcp: one on Products (H19); calcsp: one on Sales and one on
// Products (H20); calc2p: two on Sales and one on Products (H27). Add "plain" for a model without hierarchies.
var salesColumns = new[]
{
    ("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency), ("Double Qty", "[Qty]*2", ColumnDataType.Int64),
    ("Triple Qty", "[Qty]*3", ColumnDataType.Int64), ("Quad Qty", "[Qty]*4", ColumnDataType.Int64),
};
string calcMode = args.Length > 2 ? args[2] : "";
int onSales = calcMode switch { "calc2" or "calc2p" => 2, "calc3" => 3, "calc4" => 4, "calcsp" => 1, _ => 0 };
for (int n = 0; n < onSales; n++) s.AddCalculatedColumn(salesColumns[n].Item1, salesColumns[n].Item2, salesColumns[n].Item3);
if (calcMode is "calcp" or "calcsp" or "calc2p") p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
m.Save(Path.Combine(dir, args[1] + ".xlsx"));
