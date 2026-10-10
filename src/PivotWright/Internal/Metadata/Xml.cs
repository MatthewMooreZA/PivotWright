using System.Globalization;
using System.Text;

namespace PivotWright.Internal.Metadata;

internal static class Xml
{
    public static string Esc(string s)
    {
        if (s.AsSpan().IndexOfAny("&<>\"'") < 0) return s;
        var sb = new StringBuilder(s.Length + 16);
        foreach (char c in s)
        {
            switch (c)
            {
                case '&': sb.Append("&amp;"); break;
                case '<': sb.Append("&lt;"); break;
                case '>': sb.Append("&gt;"); break;
                case '"': sb.Append("&quot;"); break;
                case '\'': sb.Append("&apos;"); break;
                default: sb.Append(c); break;
            }
        }
        return sb.ToString();
    }

    public static string Num(long v) => v.ToString(CultureInfo.InvariantCulture);
    public static string Bool(bool v) => v ? "true" : "false";

    /// <summary>Timestamp format used by Analysis Services, e.g. 2026-07-30T20:47:15.723333.</summary>
    public static string Time(DateTime t) => t.ToString("yyyy-MM-dd'T'HH:mm:ss.FFFFFF", CultureInfo.InvariantCulture);

    /// <summary>Double formatting used for Magnitude, e.g. "1." or "1.E-2".</summary>
    public static string Magnitude(double d)
    {
        if (d == 1.0) return "1.";
        string r = d.ToString("0.################E0", CultureInfo.InvariantCulture);
        int e = r.IndexOf('E');
        string mant = r[..e];
        if (!mant.Contains('.')) mant += ".";
        return mant + "E" + r[(e + 1)..];
    }

    public static string Guid() => System.Guid.NewGuid().ToString("D").ToUpperInvariant();
}

/// <summary>Helpers for the XMObject (imbi) documents used by .tbl.xml files ([MS-XLDM] 2.5).</summary>
internal sealed class XmBuilder
{
    public const string RootNamespaces =
        "xmlns=\"http://schemas.microsoft.com/analysisservices/imbi\" xmlns:imbi200=\"http://schemas.microsoft.com/analysisservices/2010/imbi/200\" xmlns:imbi200_200=\"http://schemas.microsoft.com/analysisservices/2010/imbi/200/200\" xmlns:xsd=\"http://www.w3.org/2001/XMLSchema\" xmlns:xsi=\"http://www.w3.org/2001/XMLSchema-instance\"";

    private readonly StringBuilder _sb = new();

    public XmBuilder Raw(string s) { _sb.Append(s); return this; }

    public XmBuilder Open(string cls, string? name = null, int providerVersion = 0, bool root = false)
    {
        _sb.Append("<XMObject ");
        if (root) _sb.Append(RootNamespaces).Append(' ');
        _sb.Append("class=\"").Append(cls).Append('"');
        if (name is not null) _sb.Append(" name=\"").Append(Xml.Esc(name)).Append('"');
        _sb.Append(" ProviderVersion=\"").Append(providerVersion).Append("\">");
        return this;
    }

    public XmBuilder Empty(string cls, string? name = null)
    {
        _sb.Append("<XMObject class=\"").Append(cls).Append('"');
        if (name is not null) _sb.Append(" name=\"").Append(Xml.Esc(name)).Append('"');
        _sb.Append(" ProviderVersion=\"0\"/>");
        return this;
    }

    public XmBuilder Close() { _sb.Append("</XMObject>"); return this; }

    public XmBuilder Prop(string name, string type, string value)
    {
        _sb.Append('<').Append(name).Append(" xsi:type=\"xsd:").Append(type).Append("\">")
           .Append(value).Append("</").Append(name).Append('>');
        return this;
    }

    public XmBuilder Long(string name, long v) => Prop(name, "long", Xml.Num(v));
    public XmBuilder Int(string name, long v) => Prop(name, "int", Xml.Num(v));
    public XmBuilder Short(string name, long v) => Prop(name, "short", Xml.Num(v));
    public XmBuilder Bool(string name, bool v) => Prop(name, "boolean", Xml.Bool(v));
    public XmBuilder UInt(string name, long v) => Prop(name, "unsignedInt", Xml.Num(v));
    public XmBuilder Double(string name, string v) => Prop(name, "double", v);

    public XmBuilder Properties(Action<XmBuilder> body)
    {
        _sb.Append("<Properties>");
        body(this);
        _sb.Append("</Properties>");
        return this;
    }

    public XmBuilder Member(string name, Action<XmBuilder> body)
    {
        _sb.Append("<Member><Name>").Append(name).Append("</Name>");
        body(this);
        _sb.Append("</Member>");
        return this;
    }

    public XmBuilder Members(Action<XmBuilder> body)
    {
        _sb.Append("<Members>");
        body(this);
        _sb.Append("</Members>");
        return this;
    }

    public XmBuilder Collection(string name, Action<XmBuilder>? body)
    {
        _sb.Append("<Collection><Name>").Append(name).Append("</Name>");
        body?.Invoke(this);
        _sb.Append("</Collection>");
        return this;
    }

    public XmBuilder Collections(Action<XmBuilder> body)
    {
        _sb.Append("<Collections>");
        body(this);
        _sb.Append("</Collections>");
        return this;
    }

    public XmBuilder DataObjects(Action<XmBuilder> body)
    {
        _sb.Append("<DataObjects>");
        body(this);
        _sb.Append("</DataObjects>");
        return this;
    }

    public XmBuilder DataObject(Action<XmBuilder> body)
    {
        _sb.Append("<DataObject>");
        body(this);
        _sb.Append("</DataObject>");
        return this;
    }

    public override string ToString() => _sb.ToString();
}
