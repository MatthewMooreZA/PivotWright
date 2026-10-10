import os,re,sys,difflib
from xml.dom import minidom
def norm(t):
    t=re.sub(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}','GUID',t)
    t=re.sub(r'\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?','TIME',t)
    t=t.replace('TheTable_GUID','TBL').replace('null_table_GUID','TBL').replace('TheTable','TBLNAME').replace('null_table','TBLNAME')
    t=re.sub(r'(\[EncryptedDataStart\]).*?(\[EncryptedDataEnd\])',r'\1ENC\2',t,flags=re.S)
    t=re.sub(r'<Password>.*?</Password>','<Password>ENC</Password>',t,flags=re.S)
    try:
        t=minidom.parseString(t.encode()).toprettyxml(indent=' ')
    except Exception as e: t='PARSEERR '+str(e)+'\n'+t
    return [l for l in t.splitlines() if l.strip()]
def files(root):
    out={}
    for dp,dn,fn in os.walk(root):
        for f in fn:
            if not f.endswith('.xml'): continue
            p=os.path.join(dp,f)
            key=re.sub(r'[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}','G',os.path.relpath(p,root))
            key=key.replace('TheTable_G','T').replace('null_table_G','T')
            out[key]=open(p,encoding='utf-8-sig').read()
    return out
a=files(sys.argv[1]); b=files(sys.argv[2])
only=sys.argv[3] if len(sys.argv)>3 else None
for k in sorted(set(a)&set(b)):
    if only and only not in k: continue
    d=list(difflib.unified_diff(norm(a[k]),norm(b[k]),lineterm='',n=1))
    if d:
        print('=====',k); print('\n'.join(d[2:120]))
