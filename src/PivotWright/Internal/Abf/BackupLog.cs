using System.Globalization;
using System.Text;
using PivotWright.Internal.Metadata;

namespace PivotWright.Internal.Abf;

/// <summary>A file group of the backup log ([MS-XLDM] 2.1.2.3.1.3.1): one model object and its files.</summary>
internal sealed class FileGroup
{
    public required int Class { get; init; }
    public required string Id { get; init; }
    public required string Name { get; init; }
    public required int ObjectVersion { get; init; }
    public required string PersistLocationPath { get; init; }
    public required string ObjectId { get; init; }
    public List<LoggedFile> Files { get; } = new();
}

internal sealed record LoggedFile(string Path, string StorageName, long LastWriteTime, long Size);

/// <summary>Builds the backup log XML that maps logical model paths to stored files.</summary>
internal sealed class BackupLog
{
    public const int ClassDatabase = 100002;
    public const int ClassDataSource = 100003;
    public const int ClassDimension = 100006;
    public const int ClassCube = 100010;
    public const int ClassMeasureGroup = 100016;
    public const int ClassPartition = 100021;
    public const int ClassDataSourceView = 100053;
    public const int ClassMdxScript = 100060;

    public required string ServerRoot { get; init; }
    public required string DatabaseId { get; init; }
    public required string DatabaseName { get; init; }
    public required int Lcid { get; init; }
    public List<FileGroup> Groups { get; } = new();

    public string ToXml()
    {
        var sb = new StringBuilder();
        sb.Append("<BackupLog><BackupRestoreSyncVersion>11.53</BackupRestoreSyncVersion>");
        sb.Append("<ServerRoot>").Append(Xml.Esc(ServerRoot)).Append("</ServerRoot>");
        sb.Append("<SvrEncryptPwdFlag>true</SvrEncryptPwdFlag><ServerEnableBinaryXML>false</ServerEnableBinaryXML>");
        sb.Append("<ServerEnableCompression>false</ServerEnableCompression><CompressionFlag>true</CompressionFlag><EncryptionFlag>false</EncryptionFlag>");
        sb.Append("<ObjectName>").Append(Xml.Esc(DatabaseName)).Append("</ObjectName><ObjectId>").Append(DatabaseId).Append("</ObjectId>");
        sb.Append("<Write>ReadWrite</Write><OlapInfo>false</OlapInfo><Collations/><Languages><Language>")
          .Append(Lcid.ToString(CultureInfo.InvariantCulture)).Append("</Language></Languages><FileGroups>");
        foreach (var g in Groups)
        {
            sb.Append("<FileGroup><Class>").Append(g.Class).Append("</Class><ID>").Append(Xml.Esc(g.Id)).Append("</ID><Name>")
              .Append(Xml.Esc(g.Name)).Append("</Name><ObjectVersion>").Append(g.ObjectVersion).Append("</ObjectVersion>");
            sb.Append("<PersistLocation>0</PersistLocation><PersistLocationPath>").Append(Xml.Esc(g.PersistLocationPath))
              .Append("</PersistLocationPath><StorageLocationPath/><ObjectID>").Append(g.ObjectId).Append("</ObjectID><FileList>");
            foreach (var f in g.Files)
            {
                sb.Append("<BackupFile><Path>").Append(Xml.Esc(f.Path)).Append("</Path><StoragePath>").Append(f.StorageName)
                  .Append("</StoragePath><LastWriteTime>").Append(f.LastWriteTime.ToString(CultureInfo.InvariantCulture))
                  .Append("</LastWriteTime><Size>").Append(f.Size.ToString(CultureInfo.InvariantCulture)).Append("</Size></BackupFile>");
            }
            sb.Append("</FileList></FileGroup>");
        }
        sb.Append("</FileGroups></BackupLog>");
        return sb.ToString();
    }

    /// <summary>A storage name: 20 upper-case hex characters, like Excel's.</summary>
    public static string NewStorageName() => Guid.NewGuid().ToString("N")[..20].ToUpperInvariant();
}
