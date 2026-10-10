import sys,os,re,zipfile
import xml.etree.ElementTree as ET
from xpress8 import Xpress8
x=Xpress8()
src,out=sys.argv[1],sys.argv[2]
d=zipfile.ZipFile(src).read('xl/model/item.data') if zipfile.is_zipfile(src) else open(src,'rb').read()
hdr=d[72:4096].decode('utf-16le').rstrip('\x00')
off=int(re.search(r'<m_cbOffsetHeader>(\d+)',hdr).group(1)); size=int(re.search(r'<DataSize>(\d+)',hdr).group(1))
vd=ET.fromstring(d[off:off+size].decode('utf-16le') if d[off+1]==0 else d[off:off+size])
vdf={f.findtext('Path'):(int(f.findtext('m_cbOffsetHeader')),int(f.findtext('Size'))) for f in vd.findall('BackupFile')}
lo,ls=vdf['BACKUP.LOG'] if 'BACKUP.LOG' in vdf else list(vdf.values())[-1]
log=ET.fromstring(d[lo:lo+ls-4].decode('utf-16'))
root=None
for fg in log.findall('FileGroups/FileGroup'):
    for bf in fg.findall('FileList/BackupFile'):
        p=bf.findtext('Path'); sp=bf.findtext('StoragePath')
        o,s=vdf[sp]; body=d[o:o+s-4]
        u=x.decompress_chunked(body)
        assert len(u)==int(bf.findtext('Size')),(p,len(u))
        m=re.search(r'VertiPaq_[0-9A-F]+\\(.*)$',p)
        rel=m.group(1) if m else p.split('\\')[-1]
        fp=os.path.join(out,*rel.split('\\')); os.makedirs(os.path.dirname(fp),exist_ok=True)
        open(fp,'wb').write(u)
print('ok')
