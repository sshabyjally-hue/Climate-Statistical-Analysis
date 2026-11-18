import subprocess, tempfile, os, zipfile, copy, hashlib
from lxml import etree
NS={'m':'http://schemas.openxmlformats.org/officeDocument/2006/math','w':'http://schemas.openxmlformats.org/wordprocessingml/2006/main'}
_cache={}
M='{%s}'%NS['m']
def fix(el):
    for d in el.iter(M+'dPr'):
        order=['begChr','sepChr','endChr','grow','shp','ctrlPr']
        kids=sorted(list(d),key=lambda c:order.index(c.tag.replace(M,'')) if c.tag.replace(M,'') in order else 99)
        for k in list(d): d.remove(k)
        for k in kids: d.append(k)
    for d in el.iter(M+'mcPr'):
        kids=sorted(list(d),key=lambda c:0 if c.tag==M+'count' else 1)
        for k in list(d): d.remove(k)
        for k in kids: d.append(k)
    for r in el.iter(M+'rPr'):
        tags=[c.tag.replace(M,'') for c in r]
        if 'nor' in tags:
            for c in list(r):
                if c.tag.replace(M,'') in ('sty','scr'): r.remove(c)
        else:
            order=['lit','scr','sty','brk','aln']
            kids=sorted(list(r),key=lambda c:order.index(c.tag.replace(M,'')) if c.tag.replace(M,'') in order else 99)
            for k in list(r): r.remove(k)
            for k in kids: r.append(k)

def omml(latex):
    if latex in _cache: return copy.deepcopy(_cache[latex])
    with tempfile.TemporaryDirectory() as d:
        md=os.path.join(d,'a.md'); out=os.path.join(d,'a.docx')
        open(md,'w',encoding='utf8').write('$'+latex+'$\n')
        r=subprocess.run(['pandoc',md,'-f','markdown+tex_math_dollars','-o',out],capture_output=True,text=True)
        if r.returncode: raise RuntimeError(r.stderr+latex)
        x=zipfile.ZipFile(out).read('word/document.xml')
    root=etree.fromstring(x)
    el=root.find('.//m:oMath',NS)
    if el is None: raise RuntimeError('no omml for '+latex)
    fix(el)
    _cache[latex]=el
    return copy.deepcopy(el)
if __name__=='__main__':
    tests=[r'\begin{cases}1/d & a\\ 1 & b\end{cases}',r'\mathrm{arg\,min}_\theta \frac15\sum_{k=1}^5 x',r'\hat{y}=\mathbf{w}^{\top}h',r'\text{distance weights}',r'\mathcal{N}_k\in\mathbb{R}^{18}',r'\left\|\mathbf{w}\right\|^{2}']
    for t in tests:
        e=omml(t); print(len(etree.tostring(e)),t[:30])
