import sys, zipfile, re, os
sys.path.insert(0, os.path.expanduser('~/research/pbixray'))
src, out = sys.argv[1], sys.argv[2]
if zipfile.is_zipfile(src):
    d = zipfile.ZipFile(src).read('xl/model/item.data')
else:
    d = open(src,'rb').read()
hdr = d[72:4096].decode('utf-16le').rstrip('\x00')
print("HEADER:", hdr)
off = int(re.search(r'<m_cbOffsetHeader>(\d+)', hdr).group(1))
size = int(re.search(r'<DataSize>(\d+)', hdr).group(1))
vd = d[off:off+size]
os.makedirs(out, exist_ok=True)
open(os.path.join(out,'_VirtualDirectory.xml'),'wb').write(vd)
import xml.etree.ElementTree as ET
root = ET.fromstring(vd)
files = root.findall('BackupFile')
log = files[-1]
lo, ls = int(log.findtext('m_cbOffsetHeader')), int(log.findtext('Size'))
logbytes = d[lo:lo+ls]
open(os.path.join(out,'_BackupLog.raw'),'wb').write(logbytes)
print("LOG first bytes", logbytes[:8], "last", logbytes[-8:])
for f in files:
    p, o, s = f.findtext('Path'), int(f.findtext('m_cbOffsetHeader')), int(f.findtext('Size'))
    open(os.path.join(out, p.replace('\\','_')), 'wb').write(d[o:o+s])
print(len(files), "files; total len", len(d))
