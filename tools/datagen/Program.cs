using PivotWright;
string dir = args[0];
using var m = new DataModelWriter(new DataModelOptions { Lcid = 7177 });
var p = m.CreateTable("Products", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Category", ColumnDataType.String) { SortByColumn = "CategoryOrder" }, new ColumnDefinition("CategoryOrder", ColumnDataType.Int64), new ColumnDefinition("Product", ColumnDataType.String) });
string[] cats = { "Bikes", "Clothing", "Accessories" };
int k = 1;
for (int c = 0; c < 3; c++) for (int i = 1; i <= 4; i++) p.AddRow(k++, cats[c], c + 1, $"{cats[c]} item {i}");
bool generated = args.Length > 2 && args[2] == "gendate";
if (generated) m.AddDateTable("Dates", new DateOnly(args.Length > 4 ? int.Parse(args[4]) : 2025, 1, 1), new DateOnly(args.Length > 5 ? int.Parse(args[5]) : 2025, 12, 31), args.Length > 3 ? int.Parse(args[3]) : 12, args.Length > 6 ? args[6] : "Fiscal",
    completeYears: !(args.Length > 7 && args[7].Contains("exact")), includeYearMonth: !(args.Length > 7 && args[7].Contains("noym")), hideSortColumns: !(args.Length > 7 && args[7].Contains("nohide")));
else
{
    var d = m.CreateTable("Dates", new[] { new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Year", ColumnDataType.Int64), new ColumnDefinition("Month", ColumnDataType.String) { SortByColumn = "MonthNumber" }, new ColumnDefinition("MonthNumber", ColumnDataType.Int64) });
    for (int i = 0; i < 365; i++) { var dt = new DateTime(2025, 1, 1).AddDays(i); d.AddRow(dt, dt.Year, dt.ToString("MMMM", System.Globalization.CultureInfo.InvariantCulture), dt.Month); }
}
var s = m.CreateTable("Sales", new[] { new ColumnDefinition("ProductKey", ColumnDataType.Int64), new ColumnDefinition("Date", ColumnDataType.DateTime), new ColumnDefinition("Qty", ColumnDataType.Int64), new ColumnDefinition("Amount", ColumnDataType.Currency) });
for (int i = 0; i < 1000; i++) s.AddRow(1 + i % 12, new DateTime(2025, 1, 1).AddDays(i % 365), 1 + i % 5, (decimal)(10 + i % 40));
m.AddRelationship("Sales", "ProductKey", "Products", "ProductKey");
m.AddRelationship("Sales", "Date", "Dates", "Date");
m.AddHierarchy("Products", "Product Hierarchy", "Category", "Product");
if (!generated) m.AddHierarchy("Dates", "Calendar", "Year", "Month");
m.AddMeasure("Sales", "Total Qty", "SUM(Sales[Qty])", MeasureFormat.WholeNumber);
if (args.Length > 2 && args[2] == "date") m.MarkAsDateTable("Dates", "Date");
m.Save(Path.Combine(dir, args[1] + ".xlsx"));
