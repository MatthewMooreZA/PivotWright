import os,re,difflib,sys
def load(r):
    d={}
    for p,_,fs in os.walk(r):
        for f in fs:
            if f.endswith('.xml') and not f.endswith('ds.xml') and not f.endswith('db.xml'):
                k=re.sub(r'\.\d+\.(dim|det|cub|scr|prt|tbl)\.xml',r'.\1.xml',f)
                k=re.sub(r'^info\.\d+','info',k)
                k=re.sub(r'[0-9a-f]{8}-[0-9a-f-]{27}','G',k)
                if k=='info.xml': k=os.path.basename(p)[:12]+'/info'
                t=open(os.path.join(p,f),encoding='utf-8-sig').read()
                t=re.sub(r'[0-9a-fA-F]{8}-[0-9a-fA-F-]{27}','G',t)
                t=re.sub(r'\d{4}-\d\d-\d\dT[\d:.]+','T',t)
                t=re.sub(r'<(EstimatedSize|Ordinal|CacheVersion|DatabaseID)>[^<]*<','<X<',t)
                t=re.sub(r'<DataFileList>[^<]*','<DataFileList>',t)
                t=re.sub(r' xmlns(:\w+)?="[^"]*"','',t); t=re.sub(r'<(ObjectVersion|StructureVersion|DataVersion|LastSchemaUpdate|LastProcessed)>[^<]*<','<V<',t); t=re.sub(r'\s+<','<',t)
                d[k]=t.replace('><','>\n<').split('\n')
    return d
a=load(sys.argv[1]);b=load(sys.argv[2])
for k in sorted(set(a)|set(b)):
    if k not in a or k not in b: print('ONLY',k,'in', 'A' if k in a else 'B'); continue
    df=[l for l in difflib.unified_diff(a[k],b[k],lineterm='',n=1) if not l.startswith(('---','+++'))]
    if df: print('==',k);print('\n'.join(df[:int(sys.argv[3])]))
