using System.Buffers.Binary;
using System.IO.Compression;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;

namespace PivotWright.Tests;

/// <summary>
/// Minimal, independent reader for the Spreadsheet Data Model stream, used to check what the
/// writer produced: parses the header, virtual directory and backup log, validates CRCs and
/// un-chunks (and Xpress-decompresses) every file.
/// </summary>
internal sealed class ModelReader
{
    public Dictionary<string, byte[]> Files { get; } = new();
    public string HeaderXml { get; private set; } = "";
    public XElement Log { get; private set; } = null!;
    public string RawLogXml { get; private set; } = "";
    public byte[] PartitionsRaw { get; private set; } = Array.Empty<byte>();
    public List<(string Storage, byte[] Content)> InLogOrder { get; } = new();

    public static byte[] ItemData(string xlsxPath)
    {
        using var zip = ZipFile.OpenRead(xlsxPath);
        var e = zip.GetEntry("xl/model/item.data") ?? throw new InvalidDataException("No data model part.");
        using var s = e.Open();
        var ms = new MemoryStream();
        s.CopyTo(ms);
        return ms.ToArray();
    }

    public static ModelReader Read(byte[] data)
    {
        var r = new ModelReader();
        Check.True(data[0] == 0xFF && data[1] == 0xFE, "BOM");
        string sig = Encoding.Unicode.GetString(data, 2, 70);
        Check.Equal("STREAM_STORAGE_SIGNATURE_)!@#$%^&*(", sig, "signature");
        r.HeaderXml = Encoding.Unicode.GetString(data, 72, 4096 - 72).TrimEnd('\0');
        var header = XElement.Parse(r.HeaderXml);
        long vdOffset = long.Parse(header.Element("m_cbOffsetHeader")!.Value);
        int vdSize = int.Parse(header.Element("DataSize")!.Value);
        Check.True(vdOffset % 4096 == 0, "virtual directory page aligned");
        Check.Equal(data.Length % 4096, 0, "stream page aligned");
        var vd = XElement.Parse(Encoding.Unicode.GetString(data, (int)vdOffset, vdSize));
        var entries = vd.Elements("BackupFile").Select(e => (
            Path: e.Element("Path")!.Value,
            Offset: long.Parse(e.Element("m_cbOffsetHeader")!.Value),
            Size: long.Parse(e.Element("Size")!.Value))).ToList();
        Check.Equal(entries.Count, int.Parse(header.Element("Files")!.Value), "file count");
        Check.Equal("PARTITIONS", entries[0].Path, "first file");
        Check.Equal(4096L, entries[0].Offset, "files section offset");

        var stored = new Dictionary<string, byte[]>();
        foreach (var (path, offset, size) in entries)
        {
            var body = data.AsSpan((int)offset, (int)size - 4);
            uint crc = BinaryPrimitives.ReadUInt32LittleEndian(data.AsSpan((int)(offset + size - 4)));
            Check.Equal(Crc(body), crc, "CRC of " + path);
            stored[path] = body.ToArray();
        }
        var logBytes = stored["LOG"];
        r.RawLogXml = Encoding.Unicode.GetString(logBytes, 2, logBytes.Length - 2);
        r.PartitionsRaw = stored["PARTITIONS"];
        r.Log = XElement.Parse(r.RawLogXml);
        string root = r.Log.Element("ServerRoot")!.Value + "\\";
        foreach (var bf in r.Log.Descendants("BackupFile"))
        {
            string path = bf.Element("Path")!.Value;
            string sp = bf.Element("StoragePath")!.Value;
            long size = long.Parse(bf.Element("Size")!.Value);
            var content = Unchunk(stored[sp]);
            Check.Equal(size, (long)content.Length, "size of " + path);
            string rel = path.StartsWith(root, StringComparison.Ordinal) ? path[root.Length..] : path;
            r.Files[rel] = content;
            r.InLogOrder.Add((sp, content));
        }
        return r;
    }

    /// <summary>Finds a file by a regular expression over its logical path (GUIDs replaced by "G").</summary>
    public byte[] Find(string pattern)
    {
        var rx = new Regex(pattern);
        var hits = Files.Where(kv => rx.IsMatch(Normalize(kv.Key))).ToList();
        Check.Equal(1, hits.Count, "files matching " + pattern);
        return hits[0].Value;
    }

    public static string Normalize(string path) =>
        Regex.Replace(path, "[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "G");

    public static byte[] Unchunk(byte[] stored)
    {
        var output = new MemoryStream();
        int pos = 0;
        while (pos < stored.Length)
        {
            int uncompressed = BinaryPrimitives.ReadUInt16LittleEndian(stored.AsSpan(pos));
            int size = BinaryPrimitives.ReadUInt16LittleEndian(stored.AsSpan(pos + 2));
            pos += 4;
            var chunk = stored.AsSpan(pos, size);
            if (size == uncompressed) output.Write(chunk);
            else output.Write(Xpress.Decompress(chunk, uncompressed));
            pos += size;
        }
        return output.ToArray();
    }

    public static uint Crc(ReadOnlySpan<byte> data)
    {
        uint c = 0xFFFFFFFF;
        foreach (byte b in data)
        {
            c ^= (uint)b << 24;
            for (int k = 0; k < 8; k++) c = (c & 0x80000000) != 0 ? (c << 1) ^ 0x04C11DB7 : c << 1;
        }
        return c ^ 0xFFFFFFFF;
    }
}

/// <summary>Reference Xpress (plain LZ77) decoder, written from [MS-XCA] 2.4, for tests.</summary>
internal static class Xpress
{
    public static byte[] Decompress(ReadOnlySpan<byte> input, int outputSize)
    {
        var output = new byte[outputSize];
        int ip = 0, op = 0, nibble = -1;
        uint flags = 0;
        int flagCount = 0;
        while (op < outputSize)
        {
            if (flagCount == 0)
            {
                flags = BinaryPrimitives.ReadUInt32LittleEndian(input[ip..]);
                ip += 4;
                flagCount = 32;
            }
            flagCount--;
            if ((flags & (1u << flagCount)) == 0)
            {
                output[op++] = input[ip++];
                continue;
            }
            int word = BinaryPrimitives.ReadUInt16LittleEndian(input[ip..]);
            ip += 2;
            int offset = (word >> 3) + 1;
            int length = word & 7;
            if (length == 7)
            {
                if (nibble < 0) { nibble = ip; length = input[ip++] & 15; }
                else { length = input[nibble] >> 4; nibble = -1; }
                if (length == 15)
                {
                    length = input[ip++];
                    if (length == 255)
                    {
                        length = BinaryPrimitives.ReadUInt16LittleEndian(input[ip..]);
                        ip += 2;
                        length -= 15 + 7;
                    }
                    length += 15;
                }
                length += 7;
            }
            length += 3;
            if (offset > op) throw new InvalidDataException("Xpress offset before start of buffer.");
            for (int k = 0; k < length && op < outputSize; k++, op++) output[op] = output[op - offset];
        }
        return output;
    }
}
