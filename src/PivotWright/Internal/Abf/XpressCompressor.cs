using System.Buffers.Binary;

namespace PivotWright.Internal.Abf;

/// <summary>
/// Xpress (LZ77, "direct2" coding) compressor for ABF chunks ([MS-XLDM] 2.7.5). The bit stream is
/// the one decoded by the Extensible Storage Engine's Xpress8 decoder: 32-bit flag words consumed
/// most-significant bit first (0 = literal byte, 1 = match), matches as a 16-bit word holding a
/// 13-bit offset-1 and a 3-bit length-3, with nibble/byte/word extensions for long matches.
/// </summary>
internal sealed class XpressCompressor : IChunkCompressor
{
    private const int MinMatch = 3;
    private const int MaxOffset = 1 << 13;
    private const int HashBits = 12;
    private const int MaxChain = 24;

    private readonly int[] _head = new int[1 << HashBits];
    private readonly int[] _prev = new int[AbfWriter.ChunkSize];

    private static int Hash(ReadOnlySpan<byte> s, int i) => ((s[i] << 8) ^ (s[i + 1] << 4) ^ s[i + 2]) & ((1 << HashBits) - 1);

    public int TryCompress(ReadOnlySpan<byte> input, Span<byte> output)
    {
        int n = input.Length;
        if (n < 16 || n > _prev.Length) return 0;
        Array.Fill(_head, -1);

        int limit = Math.Min(output.Length, n) - 8; // give up once we are no smaller than the input
        int op = 4;                 // first flag word reserved at 0
        int flagPos = 0;
        uint flags = 0;
        int flagCount = 0;
        int nibblePos = -1;         // byte holding a pending high nibble for the next long match

        void PushFlag(bool match, Span<byte> outp, ref int o)
        {
            flags = (flags << 1) | (match ? 1u : 0u);
            if (++flagCount == 32)
            {
                BinaryPrimitives.WriteUInt32LittleEndian(outp[flagPos..], flags);
                flagPos = o;
                o += 4;
                flags = 0;
                flagCount = 0;
            }
        }

        int i = 0;
        while (i < n)
        {
            if (op >= limit) return 0;
            int bestLen = 0, bestOff = 0;
            if (i + MinMatch <= n)
            {
                int h = Hash(input, i);
                int cand = _head[h];
                int chain = 0;
                int maxLen = n - i;
                while (cand >= 0 && i - cand <= MaxOffset && chain++ < MaxChain)
                {
                    // Quick reject: a better match must also agree at position bestLen.
                    if (bestLen == 0 || (bestLen < maxLen && input[cand + bestLen] == input[i + bestLen]))
                    {
                        int len = 0;
                        while (len < maxLen && input[cand + len] == input[i + len]) len++;
                        if (len > bestLen)
                        {
                            bestLen = len;
                            bestOff = i - cand;
                            if (len == maxLen) break;
                        }
                    }
                    cand = _prev[cand];
                }
            }

            if (bestLen >= MinMatch)
            {
                int l = bestLen - MinMatch;
                ushort word = (ushort)(((bestOff - 1) << 3) | Math.Min(l, 7));
                BinaryPrimitives.WriteUInt16LittleEndian(output[op..], word);
                op += 2;
                if (l >= 7)
                {
                    l -= 7;
                    int nib = Math.Min(l, 15);
                    if (nibblePos < 0)
                    {
                        nibblePos = op;
                        output[op++] = (byte)nib;
                    }
                    else
                    {
                        output[nibblePos] |= (byte)(nib << 4);
                        nibblePos = -1;
                    }
                    if (l >= 15)
                    {
                        l -= 15;
                        if (l < 255)
                        {
                            output[op++] = (byte)l;
                        }
                        else
                        {
                            output[op++] = 255;
                            BinaryPrimitives.WriteUInt16LittleEndian(output[op..], (ushort)(bestLen - MinMatch));
                            op += 2;
                        }
                    }
                }
                PushFlag(true, output, ref op);
                for (int k = 0; k < bestLen; k++, i++)
                {
                    if (i + MinMatch <= n)
                    {
                        int h = Hash(input, i);
                        _prev[i] = _head[h];
                        _head[h] = i;
                    }
                }
            }
            else
            {
                output[op++] = input[i];
                PushFlag(false, output, ref op);
                if (i + MinMatch <= n)
                {
                    int h = Hash(input, i);
                    _prev[i] = _head[h];
                    _head[h] = i;
                }
                i++;
            }
        }

        // Flush the last flag word; unused positions are set (they are never read).
        if (flagCount > 0)
        {
            uint final = (flags << (32 - flagCount)) | ((1u << (32 - flagCount)) - 1);
            BinaryPrimitives.WriteUInt32LittleEndian(output[flagPos..], final);
        }
        else
        {
            BinaryPrimitives.WriteUInt32LittleEndian(output[flagPos..], 0xFFFFFFFFu);
        }
        return op < n ? op : 0;
    }
}
