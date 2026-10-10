"""Read a generated workbook back with pbixray (an independent VertiPaq reader).

    pip install pbixray
    python tools/verify_with_pbixray.py model.xlsx                 # list tables, schema and a preview
    python tools/verify_with_pbixray.py model.xlsx Sales sales.csv # also compare table Sales with a CSV

The CSV comparison matches columns by name; numbers are compared with a small tolerance, dates
to the millisecond, text exactly. Note that pbixray turns BLANK booleans into True.
"""
import sys

import numpy as np
import pandas as pd
from pbixray import PBIXRay


def main() -> int:
    model = PBIXRay(sys.argv[1])
    print("Tables:", list(model.tables))
    print(model.schema[["TableName", "ColumnName", "PandasDataType"]].to_string())
    for name in model.tables:
        df = model.get_table(name)
        print(f"\n{name}: {len(df):,} rows")
        print(df.head(5).to_string())

    if len(sys.argv) < 4:
        return 0

    table, csv = sys.argv[2], sys.argv[3]
    actual = model.get_table(table)
    expected = pd.read_csv(csv)
    if len(actual) != len(expected):
        print(f"Row count differs: model {len(actual)}, CSV {len(expected)}")
        return 1
    failures = 0
    for col in expected.columns:
        a, e = actual[col], expected[col]
        if pd.api.types.is_bool_dtype(a):
            # pbixray cannot represent BLANK booleans, so only non-blank rows are compared.
            mask = e.notna().values
            ok = (a.values[mask] == (e[mask].astype(str).str.lower() == "true").values).all()
        elif pd.api.types.is_datetime64_any_dtype(a):
            e = pd.to_datetime(e)
            ok = ((a.isna() & e.isna()) | ((a - e).abs() < pd.Timedelta(milliseconds=5))).all()
        elif pd.api.types.is_numeric_dtype(e) and not pd.api.types.is_bool_dtype(e):
            ok = np.allclose(a.astype(float).fillna(-1e300), e.astype(float).fillna(-1e300), rtol=1e-12, atol=1e-9)
        else:
            ok = (a.astype(str).fillna("<NA>").values == e.astype(str).fillna("<NA>").values).all()
        print(f"{col}: {'OK' if ok else 'MISMATCH'}")
        failures += 0 if ok else 1
    return 1 if failures else 0


if __name__ == "__main__":
    sys.exit(main())
