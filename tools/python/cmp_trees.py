"""Compare two extracted model trees (tools/python/extract_tree.py output): file lists, then a raw diff of every XML file.

    python3 tools/python/cmp_trees.py <treeA> <treeB> [max lines per file]

Unlike xd.py this keeps the version counters (ObjectVersion, StructureVersion, CacheVersion, DataVersion, Ordinal ...), which is what
matters when a profile is compared with an Excel-written file. Ignored: GUIDs, timestamps, the 20-digit database id, namespace
declarations, whitespace between tags, &apos; and &quot;, the order of a DataFileList. Expect noise from the encrypted connection string,
sizes that depend on the data, and `EstimatedSize`. Used in 2026-11 to match perspectives against H64, H65 and H66 (SPEC section 13).
"""
import difflib
import glob
import os
import re
import sys


def root(tree):
    return (glob.glob(os.path.join(tree, "*.db")) or [tree])[0]


def normalise_name(s):
    return re.sub(r"[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}", "G", s)


def files(r):
    out = {}
    for d, _, fs in os.walk(r):
        for f in fs:
            if re.search(r"\.(idf|hidx|dictionary)$", f):
                continue
            out[normalise_name(os.path.relpath(os.path.join(d, f), r))] = os.path.join(d, f)
    return out


def sort_data_file_list(lines):
    return [re.sub(r"<DataFileList>(.*)</DataFileList>", lambda m: "<DataFileList>" + ";".join(sorted(m.group(1).split(";"))) + "</DataFileList>", x) for x in lines]


def text(path):
    t = normalise_name(open(path, "rb").read().decode("utf-8", "replace"))
    t = re.sub(r"\d{4}-\d\d-\d\dT[\d:.]+", "T", t)
    t = re.sub(r"[0-9A-F]{20}", "DB", t)
    t = t.replace("&apos;", "'").replace("&quot;", '"')
    t = re.sub(r' xmlns(:\w+)?="[^"]*"', "", t)
    t = re.sub(r">\s+<", "><", t)
    t = re.sub(r"<(\w+)></\1>", r"<\1/>", t)
    return sort_data_file_list(re.sub(r"><", ">\n<", t).split("\n"))


def main():
    a, b = sys.argv[1], sys.argv[2]
    limit = int(sys.argv[3]) if len(sys.argv) > 3 else 12
    fa, fb = files(root(a)), files(root(b))
    print("only in", a, sorted(set(fa) - set(fb)))
    print("only in", b, sorted(set(fb) - set(fa)))
    for key in sorted(set(fa) & set(fb)):
        if not key.endswith(".xml"):
            continue
        diff = [l for l in difflib.unified_diff(text(fa[key]), text(fb[key]), lineterm="", n=0) if not l.startswith(("---", "+++", "@@"))]
        if diff:
            print("==", key, len(diff))
            for l in diff[:limit]:
                print("  ", l[:200])


if __name__ == "__main__":
    main()
