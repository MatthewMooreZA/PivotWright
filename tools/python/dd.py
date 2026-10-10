import sys,re,glob,difflib
from xml.dom import minidom
def norm(t):
    t=re.sub(r'[0-9a-fA-F]{8}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{4}-[0-9a-fA-F]{12}','GUID',t)
    t=re.sub(r'\d{4}-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d+)?','TIME',t)
    t=re.sub(r'(\[EncryptedDataStart\]).*?(\[EncryptedDataEnd\])',r'\1ENC\2',t,flags=re.S)
    t=re.sub(r'<Password>.*?</Password>','<Password>ENC</Password>',t,flags=re.S)
    t=re.sub(r' xmlns(:\w+)?="[^"]*"','',t)  # drop ns decls
    t=re.sub(r'<(/?)(ddl\d+(_\d+)?):','<\\1\\2_',t); t=re.sub(r'(\s)(xsi:type|ddl\d+(_\d+)?:\w+)=',lambda m:m.group(1)+m.group(2).replace(':','_')+'=',t)
    t=t.replace('xsi_type','type')
    t=re.sub(r'<(\/?)xsi:','<\\1',t)
    try: t=minidom.parseString(('<r>'+re.sub(r'^<\?xml[^>]*\?>','',t)+'</r>').encode()).toprettyxml(indent=' ')
    except Exception as e: return ['PARSEERR '+str(e)]
    return [l for l in t.splitlines() if l.strip()]
def load(root):
    out={}
    for p in glob.glob(root+'/*.db/**/*.xml',recursive=True):
        k=p[len(root):]
        k=re.sub(r'^/[^/]*\.db/','',k)
        k=re.sub(r'(Products|Sales|Dates)_[0-9a-f-]{36}',r'\1',k)
        k=re.sub(r'\.\d+\.(dim|det|cub|scr|dsv|ds|prt)\.xml',r'.\1.xml',k)
        k=re.sub(r'\.dim\.xml|\.0\.dim/','.dim/',k)
        k=re.sub(r'R\$Sales_[0-9a-f-]{36}\$[0-9a-f-]{36}','R$Sales$REL',k)
        k=re.sub(r'R\$Sales\$[0-9a-f-]{36}','R$Sales$REL',k)
        k=re.sub(r'(\$)(Products|Sales|Dates)_[0-9a-f-]{36}',r'\1\2',k)
        out[k]=open(p,encoding='utf-8-sig').read()
    return out
a=load(sys.argv[1]); b=load(sys.argv[2]); only=sys.argv[3] if len(sys.argv)>3 else ''
print('only A',sorted(set(a)-set(b))); print('only B',sorted(set(b)-set(a)))
for k in sorted(set(a)&set(b)):
    if only not in k: continue
    d=list(difflib.unified_diff(norm(a[k]),norm(b[k]),lineterm='',n=0))
    if d: print('=====',k,len(d)); print('\n'.join(x[:230] for x in d[2:int(sys.argv[4]) if len(sys.argv)>4 else 60]))
