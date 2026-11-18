# Step 1: extract inputs from the original manuscript (run first)
# Requires: pandoc, python-docx. Edit DOCX to your file path.
import subprocess, re, json, docx
DOCX='MetApps_Main_Document_Anonymised_1611.docx'
subprocess.run(['pandoc','-t','markdown','--wrap=none','--extract-media=media',DOCX,'-o','ms2.md'],check=True)
t=open('ms2.md',encoding='utf8').read()
# references -> refs_raw.json
refs=[p.strip() for p in t[t.index('## References'):].split('\n\n')[1:] if p.strip()]
out=[]
for r in refs:
    m=re.match(r'^(.*?)\s\((\d{4}[a-z]?)\)\.\s(.*)$',r,re.S)
    out.append({'authors':m.group(1),'year':m.group(2),'rest':m.group(3)})
json.dump(out,open('refs_raw.json','w'),ensure_ascii=False)
# sections/captions -> secs.json
t=t[:t.index('## References')]
secs={};cur=None;order=[];caps=[]
for b in re.split(r'\n\n+',t):
    if b.startswith('#'):
        cur=b.lstrip('#').strip(); secs[cur]=[]; order.append(cur); continue
    if b.startswith('  ') or b.startswith('!['): continue
    if re.match(r'\*\*(Table|Fig)',b): caps.append(b); continue
    if b.startswith('**(a)') or b.startswith('**(b)'): continue
    secs[cur].append(b)
json.dump({'secs':secs,'order':order,'caps':caps},open('secs.json','w'),ensure_ascii=False)
# tables -> tables.json
d=docx.Document(DOCX); tabs=[]
for tb in d.tables:
    rows=[]
    for r in tb.rows:
        row=[];seen=set()
        for c in r.cells:
            if id(c._tc) in seen: continue
            seen.add(id(c._tc))
            row.append([[(ru.text,bool(ru.bold),ru.font.superscript) for ru in p.runs if ru.text] for p in c.paragraphs if p.runs])
        rows.append(row)
    tabs.append(rows)
json.dump(tabs,open('tables.json','w'),ensure_ascii=False)
