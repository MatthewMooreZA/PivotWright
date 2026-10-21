using System.Data;
using System.Diagnostics;
using System.Globalization;
using PivotWright;

// Usage:
//   demo csv <input.csv> <output.xlsx> <type,type,...> [culture] load a CSV with explicit column types (culture e.g. en-ZA)
//   demo synthetic <rows> <output.xlsx> [segmentRows]           generate a multi-type table of N rows
//   demo sales <rows> <output.xlsx>                             two tables built from DataTables
string mode = args.Length > 0 ? args[0] : "synthetic";
var sw = Stopwatch.StartNew();

switch (mode)
{
    case "csv":
    {
        var types = args[3].Split(',').Select(Enum.Parse<ColumnDataType>).ToArray();
        using var model = new DataModelWriter(new DataModelOptions { Culture = CultureInfo.GetCultureInfo(args.Length > 4 ? args[4] : "en-US") });
        var lines = File.ReadAllLines(args[1]);
        var header = lines[0].Split(',');
        var table = model.CreateTable(Path.GetFileNameWithoutExtension(args[1]),
            header.Select((h, i) => new ColumnDefinition(h, types[i])));
        foreach (var line in lines.Skip(1))
        {
            var cells = line.Split(',');
            var row = new object?[cells.Length];
            for (int i = 0; i < cells.Length; i++)
            {
                string c = cells[i];
                if (c.Length == 0) { row[i] = null; continue; }
                row[i] = types[i] switch
                {
                    ColumnDataType.Int64 => (object)(long)double.Parse(c, CultureInfo.InvariantCulture),
                    ColumnDataType.Double => double.Parse(c, CultureInfo.InvariantCulture),
                    ColumnDataType.Currency => decimal.Parse(c, NumberStyles.Float, CultureInfo.InvariantCulture),
                    ColumnDataType.DateTime => DateTime.Parse(c, CultureInfo.InvariantCulture),
                    ColumnDataType.Boolean => bool.Parse(c),
                    _ => c,
                };
            }
            table.AddRow(row);
        }
        model.Save(args[2]);
        Console.WriteLine($"{table.RowCount} rows -> {args[2]} in {sw.Elapsed}");
        break;
    }
    case "synthetic":
    {
        long rows = args.Length > 1 ? long.Parse(args[1]) : 100_000;
        string output = args.Length > 2 ? args[2] : "synthetic.xlsx";
        using var model = new DataModelWriter(new DataModelOptions
        {
            SegmentRowCount = args.Length > 3 ? int.Parse(args[3]) : 1 << 20,
        });
        var table = model.CreateTable("Facts", new[]
        {
            new ColumnDefinition("Id", ColumnDataType.Int64),
            new ColumnDefinition("Category", ColumnDataType.String),
            new ColumnDefinition("Amount", ColumnDataType.Double),
            new ColumnDefinition("Price", ColumnDataType.Currency),
            new ColumnDefinition("OrderDate", ColumnDataType.DateTime),
            new ColumnDefinition("Shipped", ColumnDataType.Boolean),
            new ColumnDefinition("Sparse", ColumnDataType.Int64),
            new ColumnDefinition("Note", ColumnDataType.String),
        });
        var rnd = new Random(42);
        string[] cats = { "Alpha", "Beta", "Gamma", "Delta", "Épsilon", "beta", "ZETA" };
        var start = new DateTime(2020, 1, 1);
        using var expected = args.Length > 4 ? new StreamWriter(args[4]) : null;
        expected?.WriteLine("Id,Amount,Price,OrderDate,Shipped,Sparse");
        var row = new object?[8];
        for (long i = 0; i < rows; i++)
        {
            row[0] = i + 1;
            row[1] = cats[(int)(i / 1000 % cats.Length)];
            row[2] = i % 17 == 0 ? null : Math.Round(rnd.NextDouble() * 1000, 3);
            row[3] = (decimal)(rnd.Next(0, 100000) / 100.0);
            row[4] = start.AddDays(rnd.Next(0, 1500)).AddHours(i % 3 == 0 ? 0 : rnd.Next(0, 24));
            row[5] = i % 5 == 0 ? null : rnd.Next(2) == 0;
            row[6] = i % 3 == 0 ? (long)rnd.NextInt64(-4_000_000_000_000, 4_000_000_000_000) : null;
            row[7] = "Note " + (i % 5003);
            table.AddRow(row);
            expected?.WriteLine(string.Join(",", row[0], ((double?)row[2])?.ToString("R", CultureInfo.InvariantCulture),
                ((decimal)row[3]!).ToString(CultureInfo.InvariantCulture), ((DateTime)row[4]!).ToString("s"),
                row[5], row[6]));
        }
        model.Save(output);
        Console.WriteLine($"{rows} rows -> {output} ({new FileInfo(output).Length:N0} bytes) in {sw.Elapsed}");
        break;
    }
    case "sales":
    {
        long rows = args.Length > 1 ? long.Parse(args[1]) : 10_000;
        string output = args.Length > 2 ? args[2] : "sales.xlsx";
        using var model = new DataModelWriter(new DataModelOptions { DecimalAsCurrency = true });

        var products = new DataTable("Products");
        products.Columns.Add("ProductKey", typeof(int));
        products.Columns.Add("Product", typeof(string));
        products.Columns.Add("Colour", typeof(string));
        products.Columns.Add("ListPrice", typeof(decimal));
        for (int p = 1; p <= 50; p++)
            products.Rows.Add(p, "Product " + p, p % 3 == 0 ? "Red" : p % 3 == 1 ? "Blue" : null, 10m + p * 1.25m);
        var productTable = model.AddTable(products);

        var sales = new DataTable("Sales");
        sales.Columns.Add("SaleId", typeof(long));
        sales.Columns.Add("ProductKey", typeof(int));
        sales.Columns.Add("Date", typeof(DateTime));
        sales.Columns.Add("Quantity", typeof(int));
        sales.Columns.Add("Discount", typeof(double));
        var rnd = new Random(7);
        for (long i = 0; i < rows; i++)
            sales.Rows.Add(i + 1, rnd.Next(1, 51), new DateTime(2024, 1, 1).AddDays(rnd.Next(0, 365)), rnd.Next(1, 10), rnd.Next(0, 4) * 0.05);
        var salesTable = model.AddTable(sales);
        model.AddRelationship(salesTable["ProductKey"], productTable["ProductKey"]);

        model.Save(output);
        Console.WriteLine($"products + {rows} sales -> {output} in {sw.Elapsed}");
        break;
    }
    default:
        Console.Error.WriteLine("Unknown mode " + mode);
        return 1;
}
return 0;
