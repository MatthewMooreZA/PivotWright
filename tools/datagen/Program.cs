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
int salesRows = args.Select(a => a.StartsWith("rows=") ? int.Parse(a[5..]) : 0).FirstOrDefault(n => n > 0);   // "rows=1100000" for a table of several segments
if (salesRows == 0) salesRows = 1000;
for (int i = 0; i < salesRows; i++) s.AddRow(1 + i % 12, new DateTime(2025, 1, 1).AddDays(i % 365), 1 + i % 5, (decimal)(10 + i % 40));
var dates = m.Tables["Dates"];
m.AddRelationship(s["ProductKey"], p["ProductKey"]);
m.AddRelationship(s["Date"], dates["Date"]);
// calcmf: a second fact table, Stock, related to Products (a second one to Dates would make a loop), with a calculated column on Stock and one on Sales (four tables, two many-side tables)
if (args.Length > 2 && args[2] == "calcmf")
{
    var stock = m.CreateTable("Stock", new[] { Col("Stock", "ProductKey", ColumnDataType.Int64), Col("Stock", "Date", ColumnDataType.DateTime), Col("Stock", "OnHand", ColumnDataType.Int64) });
    for (int i = 0; i < 500; i++) stock.AddRow(1 + i % 12, new DateTime(2025, 1, 1).AddDays(i % 365), 100 + i % 50);
    m.AddRelationship(stock["ProductKey"], p["ProductKey"]);
    s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
    stock.AddCalculatedColumn("Stock Doubled", "[OnHand]*2", ColumnDataType.Int64);
}
bool plain = args.Skip(2).Contains("plain");   // no hierarchies at all ("plain", or a trailing "plain" after a calc mode)
if (!plain) p.AddHierarchy("Product Hierarchy", "Category", "Product");
if (!generated && !plain) dates.AddHierarchy("Calendar", "Year", "Month");
s.AddMeasure("Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
if ((args.Length > 2 && args[2] == "date") || args.Skip(2).Contains("datetable")) dates.MarkAsDateTable("Date");   // "datetable" after a calc mode
if (args.Length > 2 && args[2] == "calc")
{
    // calc [currency|whole|decimal|date|boolean|text|related]: one calculated column, as in Excel's H10 to H15 samples
    var (calcName, calcFormula, calcType) = (args.Skip(3).FirstOrDefault(a => a != "plain") ?? "currency") switch
    {
        "whole" => ("Double Qty", "[Qty]*2", ColumnDataType.Int64),
        "decimal" => ("Half Qty", "[Qty]/2", ColumnDataType.Double),
        "date" => ("Next Day", "[Date]+1", ColumnDataType.DateTime),
        "boolean" => ("Big Order", "[Qty]>2", ColumnDataType.Boolean),
        "text" => ("Label", "\"Qty \"&[Qty]", ColumnDataType.String),
        "related" => ("Category", "RELATED(Products[Category])", ColumnDataType.String),
        _ => ("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency),
    };
    s.AddCalculatedColumn(calcName, calcFormula, calcType);
}
// calcN: N columns on Sales (calc2 is H18, calc3 H25, calc4 H26); calcp: one on Products (H19); calcsp: one on Sales and one on
// Products (H20); calcNp: N on Sales and one on Products (calc2p is H27). Add "plain" for a model without hierarchies.
var salesColumns = new[]
{
    ("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency), ("Double Qty", "[Qty]*2", ColumnDataType.Int64),
    ("Triple Qty", "[Qty]*3", ColumnDataType.Int64), ("Quad Qty", "[Qty]*4", ColumnDataType.Int64),
};
string calcMode = args.Length > 2 ? args[2] : "";
// calcN (N columns on Sales), calcNp (N on Sales and one on Products), calcp, calcsp
var countMatch = System.Text.RegularExpressions.Regex.Match(calcMode, @"^calc(\d+)(p?)$");
int onSales = countMatch.Success ? int.Parse(countMatch.Groups[1].Value) : calcMode == "calcsp" ? 1 : 0;
for (int n = 0; n < onSales; n++)
{
    var (name, formula, type) = n < salesColumns.Length ? salesColumns[n] : ($"Qty x{n + 1}", $"[Qty]*{n + 1}", ColumnDataType.Int64);
    s.AddCalculatedColumn(name, formula, type);
}
// calch: hide Sales[Qty] and add Line Total (H33); calchl: hide Products[Product], a hierarchy level, and add Line Total (H36);
// calchc: add Line Total and hide it (H35 shows the calc-then-hide order, which also drops two information files)
if (calcMode is "calch" or "calchl" or "calchc") s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
if (calcMode == "calch") s["Qty"].IsHidden = true;
if (calcMode == "calchl") p["Product"].IsHidden = true;
if (calcMode == "calchc") s.CalculatedColumns[0].IsHidden = true;
// calcfmt: the eight formatted columns of Excel's H37 sample, on Sales
if (calcMode == "calcfmt")
{
    var us = new System.Globalization.CultureInfo("en-US");
    var za = new System.Globalization.CultureInfo("en-ZA");
    s.AddCalculatedColumn("Fmt Whole", "[Qty]*2", ColumnDataType.Int64, MeasureFormat.WholeNumber);
    s.AddCalculatedColumn("Fmt Decimal3", "[Qty]/3*10000", ColumnDataType.Double, MeasureFormat.Decimal(3));
    s.AddCalculatedColumn("Fmt DecimalNoSep", "[Qty]/7*10000", ColumnDataType.Double, MeasureFormat.Decimal(2, thousandSeparator: false));
    s.AddCalculatedColumn("Fmt Percent", "[Qty]/10", ColumnDataType.Double, MeasureFormat.Percentage(1));
    s.AddCalculatedColumn("Fmt Currency R", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.Currency(2, za));
    s.AddCalculatedColumn("Fmt Currency USD", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.Currency(2, us));
    s.AddCalculatedColumn("Fmt Current No Unit", "[Qty]*[Amount]", ColumnDataType.Currency, MeasureFormat.CurrencyWithoutSymbol());
    s.AddCalculatedColumn("Fmt Date Short", "[Date]+1", ColumnDataType.DateTime, MeasureFormat.ShortDate(us));
    s.AddCalculatedColumn("Fmt Date Long", "[Date]+2", ColumnDataType.DateTime, MeasureFormat.LongDate(us));
}
// calcdatefmt: calculated date columns with a range of .NET date format strings 
if (calcMode == "calcdatefmt")
{
    var us = new System.Globalization.CultureInfo("en-US");
    var gb = new System.Globalization.CultureInfo("en-GB");
    var za = new System.Globalization.CultureInfo("en-ZA");
    var formats = new (string Name, string Format, System.Globalization.CultureInfo Culture)[]
    {
        ("ShortDate us", "d", us), ("ShortDate gb", "d", gb), ("ShortDate za", "d", za), ("LongDate us", "D", us), ("LongDate gb", "D", gb),
        ("ShortTime us", "t", us), ("LongTime us", "T", us), ("General short us", "g", us), ("General long gb", "G", gb), ("Full short us", "f", us),
        ("Full long us", "F", us), ("Month day us", "M", us), ("Year month us", "Y", us), ("Sortable", "s", us), ("Universal sortable", "u", us),
        ("Custom 1", "dd MMM yyyy", us), ("Custom 2", "yyyy-MM-dd HH:mm", us), ("Custom 3", "dddd d MMMM yyyy", gb),
        ("Custom 4", "MMM-yy", us), ("Custom 5", "h:mm tt", us),
    };
    for (int n = 0; n < formats.Length; n++)
        s.AddCalculatedColumn(formats[n].Name, $"[Date]+{n}", ColumnDataType.DateTime, MeasureFormat.Date(formats[n].Format, formats[n].Culture));
}
// calcd: one column on Dates (H31); calcsd: one on Sales and one on Dates (H32; with "datetable"); calcpd: one on Products and one on Dates
if (calcMode is "calcd" or "calcsd" or "calcpd" or "calcdh") dates.AddCalculatedColumn("Year Month", "[Year]*100+[MonthNumber]", ColumnDataType.Int64);
// calcdh: calcd plus a hierarchy Year > Year Month over the calculated column (H31B; with "datetable")
if (calcMode == "calcdh") dates.AddHierarchy("Hierarchy1", "Year", "Year Month");
if (calcMode == "calcsd") s.AddCalculatedColumn("Line Total", "[Qty]*[Amount]", ColumnDataType.Currency);
// calcpN: N columns on Products (calcp is one, H19)
var productsMatch = System.Text.RegularExpressions.Regex.Match(calcMode, @"^calcp(\d+)$");
if (productsMatch.Success)
    for (int n = 1; n <= int.Parse(productsMatch.Groups[1].Value); n++)
        p.AddCalculatedColumn(n == 1 ? "Name Length" : $"Name Length x{n}", n == 1 ? "LEN([Product])" : $"LEN([Product])*{n}", ColumnDataType.Int64);
if (calcMode is "calcp" or "calcsp" or "calcpd" || (countMatch.Success && countMatch.Groups[2].Value == "p")) p.AddCalculatedColumn("Name Length", "LEN([Product])", ColumnDataType.Int64);
m.Save(Path.Combine(dir, args[1] + ".xlsx"));
