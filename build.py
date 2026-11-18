import re, json, copy, unicodedata
from docx import Document
from docx.shared import Pt, Inches, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH, WD_COLOR_INDEX, WD_BREAK
from docx.enum.table import WD_TABLE_ALIGNMENT, WD_CELL_VERTICAL_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement
from lxml import etree
from omml import omml, NS
from ieee import format_ref

secs=json.load(open('secs.json',encoding='utf8'))['secs']
TABLES=json.load(open('tables.json',encoding='utf8'))
REFS=json.load(open('refs_raw.json',encoding='utf8'))
FONT='Times New Roman'

# ---------------- reference machinery ----------------
def norm(s): return re.sub(r'[^a-z]','',unicodedata.normalize('NFKD',s).encode('ascii','ignore').decode().lower())
REFKEY={}
for i,r in enumerate(REFS):
    first=r['authors'].split(',')[0] if 'Intergovernmental' not in r['authors'] else 'intergovernmental'
    REFKEY[(norm(first),r['year'])]=i
cited=[]   # order of first citation -> ref index
def cnum(i):
    if i not in cited: cited.append(i)
    return cited.index(i)+1
def fmt_group(nums):
    nums=sorted(set(nums)); out=[];k=0
    while k<len(nums):
        j=k
        while j+1<len(nums) and nums[j+1]==nums[j]+1: j+=1
        if j-k>=2: out.append(f'[{nums[k]}]–[{nums[j]}]')
        else: out+= [f'[{n}]' for n in nums[k:j+1]]
        k=j+1
    return ', '.join(out)
def lookup(name,year):
    name=name.strip()
    if name.startswith('Intergovernmental'): key='intergovernmental'
    else:
        key=re.split(r' et al\.| and | & ',name)[0]
        key=norm(key)
    return REFKEY[(key,year)]
PAR_ITEM=re.compile(r'^(.+?),\s*(\d{4}[a-z]?(?:,\s*\d{4}[a-z]?)*)$')
def cite(text):
    text=text.replace('Zhou et al. (2020, Table 3)','Zhou et al. (2020)§T3§')
    def par(m):
        items=m.group(1).split(';'); nums=[]
        for it in items:
            mm=PAR_ITEM.match(it.strip())
            if not mm or not re.match(r"^[A-ZÀ-Ž]",mm.group(1)): return m.group(0)
            for y in re.split(r',\s*',mm.group(2)): nums.append(cnum(lookup(mm.group(1),y)))
        return fmt_group(nums)
    text=re.sub(r'\(([^()]*\d{4}[a-z]?[^()]*)\)',par,text)
    NAME=r"[A-ZÀ-Ž][\w’'\-ăţ]+"
    def nar(m):
        n=cnum(lookup(m.group(1),m.group(2)))
        return f'{m.group(1)} [{n}]'
    text=re.sub(rf"((?:Intergovernmental Panel on Climate Change)|(?:{NAME}(?: et al\.|(?: and| &) {NAME})?)) \((\d{{4}}[a-z]?)\)",nar,text)
    text=text.replace('§T3§','§T3§')
    text=re.sub(r'(\[\d+\])§T3§',lambda m:m.group(1)[:-1]+', Table 3]',text)
    return text

# ---------------- document + styles ----------------
doc=Document()
sec=doc.sections[0]
sec.page_width=Cm(21);sec.page_height=Cm(29.7)
for a in ('left_margin','right_margin','top_margin','bottom_margin'): setattr(sec,a,Cm(2.5))
TW=16.0 # text width cm
def setfont(style,size,bold=None,italic=None,color=RGBColor(0,0,0)):
    f=style.font; f.name=FONT; f.size=Pt(size)
    if bold is not None: f.bold=bold
    if italic is not None: f.italic=italic
    f.color.rgb=color
    rpr=style.element.get_or_add_rPr(); rf=rpr.find(qn('w:rFonts'))
    if rf is None: rf=OxmlElement('w:rFonts'); rpr.append(rf)
    for k in ('w:ascii','w:hAnsi','w:eastAsia','w:cs'): rf.set(qn(k),FONT)
    for k in ('w:asciiTheme','w:hAnsiTheme','w:eastAsiaTheme','w:cstheme'):
        if rf.get(qn(k)) is not None: del rf.attrib[qn(k)]
st=doc.styles
setfont(st['Normal'],11); st['Normal'].paragraph_format.space_after=Pt(6); st['Normal'].paragraph_format.line_spacing=1.15
for name,size,b,i,sb,sa in (('Heading 1',13,True,False,14,6),('Heading 2',11.5,True,False,10,4),('Heading 3',11,True,True,8,3)):
    setfont(st[name],size,b,i); pf=st[name].paragraph_format; pf.space_before=Pt(sb); pf.space_after=Pt(sa); pf.keep_with_next=True
setfont(st['Title'],17,True,False)
st['Title'].paragraph_format.space_after=Pt(8); st['Title'].paragraph_format.alignment=WD_ALIGN_PARAGRAPH.LEFT
# remove title bottom border
ppr=st['Title'].element.get_or_add_pPr()
for b in ppr.findall(qn('w:pBdr')): ppr.remove(b)

TOK=re.compile(r'(\$[^$]+\$|\*\*.+?\*\*|\*[^*]+?\*|_\{.+?\}|\^\{.+?\}|\[\[.+?\]\])')
def clean(s):
    s=s.replace('---','—').replace('--','–')
    s=re.sub(r'\\([<>\[\]*_#$&%])',r'\1',s)
    return s
def run(p,text,bold=None,italic=None,sup=False,sub=False,size=None,hl=False):
    r=p.add_run(text)
    if bold is not None: r.bold=bold
    if italic is not None: r.italic=italic
    if sup: r.font.superscript=True
    if sub: r.font.subscript=True
    if size: r.font.size=Pt(size)
    if hl: r.font.highlight_color=WD_COLOR_INDEX.YELLOW
    return r
def rich(p,text,size=None,bold=None,italic=None):
    text=clean(text)
    for part in TOK.split(text):
        if not part: continue
        if part.startswith('$'):
            el=omml(part[1:-1]); p._p.append(el)
        elif part.startswith('**'): run(p,part[2:-2],bold=True,italic=italic,size=size)
        elif part.startswith('*'): run(p,part[1:-1],italic=True,bold=bold,size=size)
        elif part.startswith('_{'): run(p,part[2:-1],sub=True,size=size,bold=bold,italic=italic)
        elif part.startswith('^{'): run(p,part[2:-1],sup=True,size=size,bold=bold,italic=italic)
        elif part.startswith('[['): run(p,'['+part[2:-2]+']',hl=True,size=size,bold=bold,italic=italic)
        else: run(p,part,size=size,bold=bold,italic=italic)
    return p

# ---------------- label registry ----------------
LAB={}
def reg(kind,names,prefix=''):
    for i,n in enumerate(names): LAB[(kind,n)]=f'{prefix}{i+1}'
EQS=['clim', 'r2', 'rmse', 'mae', 'bias', 'nse', 'knn', 'svrf', 'kern', 'rf', 'gbu', 'hgbleaf', 'dnnl', 'opt', 'dm', 'mk', 'sen', 'pet']
reg('eq',EQS)
reg('fig',['hist','flow','imp','cv','scatx','scatn','heat','ts','err','opt','trend'])
reg('tab',['desc','pred','hyper','perf','dm','abl','loro','trend'])  # S3,S4 assigned below
def resolve(t):
    def f(m):
        k,n=m.group(1),m.group(2)
        v=LAB[(k,n)]
        return {'eq':v,'fig':f'Figure {v}','tab':f'Table {v}'}[k]
    return re.sub(r'\{(eq|fig|tab):(\w+)\}',f,t)
def T(text):  # full text pipeline
    return rich_ready(text)
def rich_ready(t): return cite(resolve(t))

def para(text,style=None,align=None,size=None,after=None,keep=False,indent=True,first_line=None):
    p=doc.add_paragraph(style=style)
    if align: p.alignment=align
    if after is not None: p.paragraph_format.space_after=Pt(after)
    if keep: p.paragraph_format.keep_with_next=True
    rich(p,rich_ready(text),size=size)
    return p
def body(text,**k):
    p=para(text,align=WD_ALIGN_PARAGRAPH.JUSTIFY,**k); return p
def H(level,text): 
    p=doc.add_heading(level=level); rich(p,text); return p

# ---------------- equations ----------------
TBLPR_ORDER=['tblStyle','tblpPr','tblOverlap','bidiVisual','tblStyleRowBandSize','tblStyleColBandSize','tblW','jc','tblCellSpacing','tblInd','tblBorders','shd','tblLayout','tblCellMar','tblLook','tblCaption','tblDescription']
def tblpr_insert(tblPr,el):
    name=el.tag.split('}')[1]; idx=TBLPR_ORDER.index(name)
    for ch in list(tblPr):
        n=ch.tag.split('}')[1]
        if n in TBLPR_ORDER and TBLPR_ORDER.index(n)>idx:
            ch.addprevious(el); return
    tblPr.append(el)
def set_cell_borders_none(tbl):
    tblPr=tbl._tbl.tblPr
    b=OxmlElement('w:tblBorders')
    for e in ('top','left','bottom','right','insideH','insideV'):
        x=OxmlElement(f'w:{e}'); x.set(qn('w:val'),'nil'); b.append(x)
    tblpr_insert(tblPr,b)
def set_col_widths(tbl,widths_cm):
    tbl.autofit=False
    tblPr=tbl._tbl.tblPr
    grid=tbl._tbl.tblGrid
    for gc,w in zip(grid.findall(qn('w:gridCol')),widths_cm): gc.set(qn('w:w'),str(int(w*567)))
    for row in tbl.rows:
        for c,w in zip(row.cells,widths_cm): c.width=Cm(w)
def equation(label,latex,where=None):
    n=LAB[('eq',label)]
    t=doc.add_table(rows=1,cols=3); t.alignment=WD_TABLE_ALIGNMENT.CENTER
    set_cell_borders_none(t); set_col_widths(t,[1.2,13.6,1.2])
    c=t.rows[0].cells
    for cc in c: cc.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
    p=c[1].paragraphs[0]; p.paragraph_format.space_after=Pt(2); p.paragraph_format.space_before=Pt(2)
    mp=OxmlElement('m:oMathPara'); jc=OxmlElement('m:oMathParaPr'); j=OxmlElement('m:jc'); j.set(qn('m:val'),'center'); jc.append(j); mp.append(jc)
    mp.append(omml(latex)); p._p.append(mp)
    p3=c[2].paragraphs[0]; p3.alignment=WD_ALIGN_PARAGRAPH.RIGHT; run(p3,f'({n})')
    # cantSplit
    trPr=t.rows[0]._tr.get_or_add_trPr(); cs=OxmlElement('w:cantSplit'); trPr.append(cs)
    if where:
        w=para(where,align=WD_ALIGN_PARAGRAPH.JUSTIFY,after=6)
        w.paragraph_format.space_before=Pt(2)

# ---------------- tables ----------------
def border(cell,**kw):
    tcPr=cell._tc.get_or_add_tcPr(); b=tcPr.find(qn('w:tcBorders'))
    if b is None: b=OxmlElement('w:tcBorders'); tcPr.append(b)
    for edge,val in kw.items():
        e=OxmlElement(f'w:{edge}'); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),str(val)); e.set(qn('w:color'),'000000'); b.append(e)
    va=tcPr.find(qn('w:vAlign'))
    if va is not None: va.addprevious(b)
def caption(kind,label,text,keep=False):
    n=LAB[(kind,label)]
    p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(4 if kind=='tab' else 10); p.paragraph_format.space_before=Pt(8 if kind=='tab' else 2)
    p.paragraph_format.keep_with_next=keep; p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p,('Table ' if kind=='tab' else 'Figure ')+n+'. ',bold=True,size=10)
    rich(p,rich_ready(text),size=10)
    return p
def data_table(idx,fs=9,first_left=True,widths=None):
    rows=TABLES[idx]; nc=len(rows[0])
    t=doc.add_table(rows=len(rows),cols=nc); t.alignment=WD_TABLE_ALIGNMENT.CENTER
    set_cell_borders_none(t)
    # widths from content
    if widths is None:
        mx=[max(max((len(''.join(r[0] for r in pp)) for pp in row[j]) or [1]) if row[j] else 1 for row in rows) for j in range(nc)]
        mx=[min(max(m,4),26) for m in mx]; tot=sum(mx); widths=[TW*m/tot for m in mx]
    set_col_widths(t,widths)
    for i,row in enumerate(rows):
        for j,cell in enumerate(row):
            c=t.rows[i].cells[j]; c.vertical_alignment=WD_CELL_VERTICAL_ALIGNMENT.CENTER
            first=True
            for pp in cell:
                p=c.paragraphs[0] if first else c.add_paragraph(); first=False
                p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(1); p.paragraph_format.line_spacing=1.0
                p.alignment=WD_ALIGN_PARAGRAPH.LEFT if (j==0 and first_left) else WD_ALIGN_PARAGRAPH.CENTER
                for txt,b,sup in pp:
                    run(p,clean(txt).replace('\\','') if False else txt,bold=b or (i==0),sup=bool(sup),size=fs)
            if i==0: border(c,top=12,bottom=6)
            if i==len(rows)-1: border(c,bottom=12)
        t.rows[i]._tr.get_or_add_trPr().append(OxmlElement('w:cantSplit'))
    for pp_ in t.rows[-1].cells:
        for q in pp_.paragraphs: q.paragraph_format.keep_with_next=True
    return t
def note(text,size=9):
    p=doc.add_paragraph(); p.paragraph_format.space_before=Pt(3); p.paragraph_format.space_after=Pt(10); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    rich(p,rich_ready(text),size=size); return p
def figure(label,path,width_in,cap):
    p=doc.add_paragraph(); p.alignment=WD_ALIGN_PARAGRAPH.CENTER; p.paragraph_format.keep_with_next=True; p.paragraph_format.space_before=Pt(6); p.paragraph_format.space_after=Pt(2)
    p.add_run().add_picture(path,width=Inches(width_in))
    caption('fig',label,cap)
ABB='KNN, k-nearest neighbours; SVR, support vector regression; RF, random forest; GBDT, gradient boosting decision tree; HGB, histogram-based gradient boosting; DNN, deep neural network; RMSE, root-mean-square error; MAE, mean absolute error; NSE, Nash–Sutcliffe efficiency; CI, confidence interval.'
M='media/media/'

# ======================= CONTENT =======================
# ---- title block
TITLE='Leakage-aware machine learning for one-month-ahead air temperature prediction in semi-arid Tanzania'
p=doc.add_paragraph(style='Title'); run(p,TITLE)
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); rich(p,'[Author names and ORCID iDs]')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(2); rich(p,'[Institutional affiliations]')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(8); rich(p,'[Corresponding author e-mail]')
p=doc.add_paragraph(); p.paragraph_format.space_after=Pt(8)
rich(p,'**Keywords:** air temperature prediction; machine learning; gradient boosting; temporal cross-validation; Tanzania')
H(1,'Abstract')
ABSTRACT=("Reliable one-month-ahead predictions of air temperature are needed for agriculture, water and health planning in semi-arid Tanzania, where weather-station networks are sparse. However, no study has compared machine-learning algorithms for predicting monthly maximum (Tmax) and minimum (Tmin) temperature in central Tanzania, and most comparisons rely on random validation that ignores the time ordering of climate data. This study compares six machine-learning algorithms—k-nearest neighbours, support vector regression (SVR), random forest, gradient boosting decision tree, histogram-based gradient boosting (HGB) and a deep neural network—for one-month-ahead prediction of monthly Tmax and Tmin in the Dodoma, Singida and Tabora regions (1981–2024). Models were trained on 1982–2018 data, evaluated out of time on a 2019–2024 hold-out against climatology and persistence benchmarks, and tested for transferability to an unseen region by leave-one-region-out validation. HGB was most accurate for Tmax (coefficient of determination R² = 0.777, Nash–Sutcliffe efficiency NSE = 0.774, root-mean-square error RMSE = 1.351 °C, bias = 0.029 °C) and SVR for Tmin (R² = 0.885, NSE = 0.876, RMSE = 0.822 °C). SVR, random forest, gradient boosting decision tree and HGB did not differ significantly, whereas k-nearest neighbours and the deep neural network were significantly less accurate. Random 10-fold cross-validation underestimated out-of-time Tmax error by 5.6–15.1 %, concealed warm biases of up to 0.79 °C and changed the model ranking. Tmin warmed by 0.15–0.19 °C per decade at Dodoma and Singida, whereas the apparent Tmax decline reflects a step change after 2017. Tree-ensemble and kernel methods evaluated out of time are recommended for monthly temperature outlooks in data-sparse East Africa.")
nw=len(ABSTRACT.split()); print('ABSTRACT WORDS',nw)
assert nw<=300
para(ABSTRACT,align=WD_ALIGN_PARAGRAPH.JUSTIFY)

# ---- 1 Introduction
H(1,'1. Introduction')
intro=list(secs['1 Introduction'])
rep=[("Zhou et al. (2020) compared MLR, GBDT, KNN, RF, extreme gradient boosting (XGB) and a DNN using","Zhou et al. (2020) compared multiple linear regression (MLR), gradient boosting decision tree (GBDT), k-nearest neighbours (KNN), random forest (RF), extreme gradient boosting (XGB) and a deep neural network (DNN) using"),
("For temporal forecasting, SVR has been favoured","For temporal forecasting, support vector regression (SVR) has been favoured"),
("report RMSE and R² without trivial","report the root-mean-square error (RMSE) and the coefficient of determination (R²) without trivial"),
("six ML algorithms (KNN, SVR, RF, GBDT, HGB and DNN)","six ML algorithms (KNN, SVR, RF, GBDT, histogram-based gradient boosting [HGB] and DNN)"),
("Section 2 introduces the study area, data and methods. Section 3 presents variable importance, model performance, validation, transferability and trend results with discussion. Section 4 gives the conclusions.","Section 2 describes the study area, data, performance metrics, algorithms and validation methods. Section 3 presents the results on variable importance, model performance, validation, transferability and trends. Section 4 discusses the significance of the results and compares them with previous work, Section 5 notes the limitations and Section 6 gives the conclusions.")]
for ptxt in intro:
    for a,b in rep:
        if a in ptxt: ptxt=ptxt.replace(a,b); rep=[x for x in rep if x[0]!=a]
    body(ptxt)
assert not rep, rep

# ---- 2 Method
H(1,'2. Method')
H(2,'2.1. Study area and data')
s21=secs['2.1 Study area and data'][0]
a,b=s21.split('As Table 1 shows')
a=a.replace("from which DTR = Tmax − Tmin was derived.","from which the diurnal temperature range (DTR = Tmax − Tmin) was derived.")
body(a.strip()+' As {tab:desc} shows'+b)
caption('tab','desc','Descriptive statistics of monthly temperature data, 1981–2024 (n = 528 per region).',keep=True)
data_table(0,fs=8,widths=[1.7,1.7,2.0,2.2,1.3,2.0,2.2,1.3,1.6])
note('Abbreviations: Tmax, maximum temperature; Tmin, minimum temperature; DTR, diurnal temperature range; SD, standard deviation.')

H(2,'2.2. Preparation of the training dataset')
p1,p2=secs['2.2.1 Preparation of the training dataset']
first,rest=p1.split('Eighteen predictors')
body("Each target (Tmax or Tmin) at month $t$ in region $r$ was predicted from information available at the end of month $t-1$, so that every model is a genuine one-month-ahead predictor.")
rest=rest.replace('(Table 2)','({tab:pred})')
body('Eighteen predictors'+rest.replace('Following Zhou et al.','The seasonal terms are $\\sin(2\\pi m/12)$ and $\\cos(2\\pi m/12)$ of the calendar month $m$, the anomaly is the previous month’s value minus the climatology of Eq. ({eq:clim}), and the 3-month mean is the average of the three preceding months. Following Zhou et al.'))
caption('tab','pred','Predictor groups used for the Tmax model (the Tmin model mirrors it with Tmax and Tmin interchanged).',keep=True)
data_table(1,fs=9,widths=[3.6,5.4,1.6,5.4])
note('Abbreviations: SST, sea-surface temperature.')
equation('clim',r'\bar{T}_{r,m}=\frac{1}{\left|\mathcal{Y}_{\mathrm{tr}}\right|}\sum_{y\in\mathcal{Y}_{\mathrm{tr}}}T_{r,y,m}','where $\\mathcal{Y}_{\\mathrm{tr}}$ is the set of training years (1982–2018) and $T_{r,y,m}$ is the value for region $r$, year $y$ and calendar month $m$. This region-month climatology is both a predictor and the climatology reference forecast.')
p2=p2.replace('Figure 1 compares the distributions','{fig:hist} compares the distributions').replace('The workflow is shown in Fig. 2.','The workflow is shown in {fig:flow}.')
assert '{fig:hist}' in p2 and '{fig:flow}' in p2
body(p2)
figure('hist',M+'image1.png',5.6,'Histograms of monthly (a–c) maximum temperature and (d–f) minimum temperature for all data, the development set (1982–2018) and the test set (2019–2024).')
figure('flow',M+'image2.png',5.3,'Flowchart of the air temperature prediction and evaluation framework used in this study.')

H(2,'2.3. Performance metrics')
body("Five statistics were used to evaluate predictions: the coefficient of determination (R²), the root-mean-square error (RMSE), the mean absolute error (MAE), the mean bias and the Nash–Sutcliffe efficiency (NSE) [Nash & Sutcliffe, 1970]. They are given by Eqs. ({eq:r2})–({eq:nse}).".replace('[Nash & Sutcliffe, 1970]','(Nash & Sutcliffe, 1970)'))
equation('r2',r'R^{2}=\frac{\left[\sum_{i=1}^{N}\left(T_{e,i}-\bar{T}_{e}\right)\left(T_{o,i}-\bar{T}_{o}\right)\right]^{2}}{\sum_{i=1}^{N}\left(T_{e,i}-\bar{T}_{e}\right)^{2}\sum_{i=1}^{N}\left(T_{o,i}-\bar{T}_{o}\right)^{2}}')
equation('rmse',r'\mathrm{RMSE}=\sqrt{\frac{1}{N}\sum_{i=1}^{N}\left(T_{e,i}-T_{o,i}\right)^{2}}')
equation('mae',r'\mathrm{MAE}=\frac{1}{N}\sum_{i=1}^{N}\left|T_{e,i}-T_{o,i}\right|')
equation('bias',r'\mathrm{Bias}=\frac{1}{N}\sum_{i=1}^{N}\left(T_{e,i}-T_{o,i}\right)')
equation('nse',r'\mathrm{NSE}=1-\frac{\sum_{i=1}^{N}\left(T_{e,i}-T_{o,i}\right)^{2}}{\sum_{i=1}^{N}\left(T_{o,i}-\bar{T}_{o}\right)^{2}}','where $T_e$ is the predicted and $T_o$ the observed temperature, overbars denote means and $N$ is the sample size.')
body("R² and NSE have an upper limit of 1, which indicates a perfect prediction. An NSE of 0 means that a model is no better than the observed mean, and a negative NSE means that it is worse. RMSE and MAE are reported together because they emphasize different aspects of error (Chai & Draxler, 2014; Willmott & Matsuura, 2005), and the bias shows systematic over- or under-prediction. The relative RMSE reduction of a model with respect to a reference forecast is $100\\times\\left(\\mathrm{RMSE}_{\\mathrm{ref}}-\\mathrm{RMSE}_{\\mathrm{model}}\\right)/\\mathrm{RMSE}_{\\mathrm{ref}}$.")

H(2,'2.4. Machine-learning algorithms')
body("Six algorithms were compared, implemented in scikit-learn (Pedregosa et al., 2011). Each is formulated below for a training set of $N$ samples $\\left\\{\\left(\\mathbf{x}_{i},y_{i}\\right)\\right\\}$, where $\\mathbf{x}_i$ is the 18-element predictor vector and $y_i$ the target (Tmax or Tmin). The distance-based and kernel-based learners and the neural network are sensitive to predictor scale, so predictors were standardized to zero mean and unit variance using training-period statistics.")
H(3,'2.4.1. k-nearest neighbours')
body("*k-nearest neighbours (KNN)* predicts the (distance-weighted) mean of the $k$ training samples closest in standardized predictor space (Altman, 1992), where closeness is the Euclidean distance and the prediction is given by Eq. ({eq:knn}).")
equation('knn',r'\hat{y}\left(\mathbf{x}\right)=\frac{\sum_{i\in\mathcal{N}_{k}}w_{i}y_{i}}{\sum_{i\in\mathcal{N}_{k}}w_{i}},\qquad w_{i}=\begin{cases}1/d\left(\mathbf{z},\mathbf{z}_{i}\right) & \text{distance weights}\\ 1 & \text{uniform weights}\end{cases}','where $\\mathcal{N}_k$ is the set of the $k$ training samples nearest to the query point.')
H(3,'2.4.2. Support vector regression')
body("*Support vector regression (SVR)* fits a function that keeps most training errors inside an ε-insensitive tube while penalising flatness; with a radial-basis kernel (Eq. ({eq:kern})) it models smooth nonlinear responses (Drucker et al., 1997; Smola & Schölkopf, 2004). The fitted function is given in its dual form by Eq. ({eq:svrf}).")
equation('svrf',r'f\left(\mathbf{x}\right)=\sum_{i=1}^{N}\left(\alpha_{i}-\alpha_{i}^{*}\right)K\left(\mathbf{x}_{i},\mathbf{x}\right)+b')
equation('kern',r'K\left(\mathbf{x},\mathbf{x}^{\prime}\right)=\exp\left(-\gamma\left\|\mathbf{x}-\mathbf{x}^{\prime}\right\|^{2}\right)','where $C$ is the penalty parameter, $\\varepsilon$ the tube width, $\\gamma$ the kernel width, and $\\alpha_i$, $\\alpha_i^{*}$ are the Lagrange multipliers, which are non-zero only for the support vectors.')
H(3,'2.4.3. Random forest')
body("*Random forest (RF)* averages many decision trees grown on bootstrap samples with random feature subsets, reducing variance (Breiman, 2001). The ensemble prediction is the mean of the $B$ tree outputs (Eq. ({eq:rf})); $B$ = 300 trees were used.")
equation('rf',r'\hat{y}_{\mathrm{RF}}\left(\mathbf{x}\right)=\frac{1}{B}\sum_{b=1}^{B}h_{b}\left(\mathbf{x};\Theta_{b}\right)','where $h_b$ is the $b$-th regression tree and $\\Theta_b$ describes its bootstrap sample and random feature subsets.')
H(3,'2.4.4. Gradient boosting decision tree')
body("*Gradient boosting decision tree (GBDT)* adds shallow trees sequentially, each fitted to the negative gradient of the loss of the current ensemble (Friedman, 2001). The ensemble starts from the training mean, each new tree $h_m$ is fitted to the pseudo-residuals of the squared-error loss, and it is added with a shrinkage factor (learning rate) $\\nu$ (Eq. ({eq:gbu})). Each tree was fitted to a random 80 % subsample of the training data.")
equation('gbu',r'F_{m}\left(\mathbf{x}\right)=F_{m-1}\left(\mathbf{x}\right)+\nu\,h_{m}\left(\mathbf{x}\right),\qquad m=1,\dots,M','where $F_m$ is the ensemble after $m$ rounds, $F_0$ the training mean, $h_m$ the tree fitted to the residuals $y_i-F_{m-1}\\left(\\mathbf{x}_i\\right)$, and $M$ the number of boosting rounds.')
H(3,'2.4.5. Histogram-based gradient boosting')
body("*Histogram-based gradient boosting (HGB)* is a regularized boosting implementation that bins inputs into at most 255 histogram bins (the scikit-learn default) and grows trees with L2 regularization, in the same family as XGBoost and LightGBM (Chen & Guestrin, 2016; Ke et al., 2017). It is used here as the counterpart of the extreme gradient boosting model of Zhou et al. (2020). For a leaf $l$, the leaf weight is the regularized Newton step in Eq. ({eq:hgbleaf}).")
equation('hgbleaf',r'w_{l}=-\frac{G_{l}}{H_{l}+\lambda},\qquad G_{l}=\sum_{i\in l}g_{i},\quad H_{l}=\sum_{i\in l}h_{i}','where $g_i$ and $h_i$ are the first and second derivatives of the loss with respect to the current prediction and $\\lambda$ is the L2 regularization strength; splits are chosen to maximize the resulting reduction in loss.')
H(3,'2.4.6. Deep neural network')
body("*Deep neural network (DNN)* is a fully connected multilayer perceptron trained by back-propagation with the Adam optimizer, L2 weight decay and early stopping on a 15 % internal validation split (LeCun et al., 2015). The hidden layers transform the standardized inputs (Eq. ({eq:dnnl})), a linear output unit gives the prediction, and the weights minimize the L2-penalized squared-error loss.")
equation('dnnl',r'\mathbf{h}^{(l)}=\phi\left(\mathbf{W}^{(l)}\mathbf{h}^{(l-1)}+\mathbf{b}^{(l)}\right),\quad l=1,\dots,L,\qquad \mathbf{h}^{(0)}=\mathbf{z}','where $\\phi$ is the rectified linear unit (ReLU) activation, $\\mathbf{W}^{(l)}$ and $\\mathbf{b}^{(l)}$ are the weights and biases of layer $l$, $L$ is the number of hidden layers and $\\mathbf{z}$ is the standardized predictor vector.')
H(3,'2.4.7. Hyperparameter tuning')
body("Hyperparameters $\\theta$ were selected with GridSearchCV using a five-fold forward-chaining split (TimeSeriesSplit) of the development period and RMSE as the criterion. The search spaces and selected values are listed in {tab:hyper}. A fixed random seed (42) was used for all stochastic learners.")
caption('tab','hyper','Hyperparameter search space and selected values.',keep=True)
data_table(2,fs=8.5,widths=[2.2,5.8,4.0,4.0])
note('Abbreviations: '+ABB.split(' RMSE')[0]+' RBF, radial-basis function; lr, learning rate; α, L2 penalty.')
H(3,'2.4.8. Reference forecasts')
body("Two reference forecasts were added. *Climatology* predicts the training-period mean for the same region and calendar month (Eq. ({eq:clim})). *Persistence* predicts the previous month's value.")

H(2,'2.5. Validation design')
body("Four complementary validation designs were used, followed by an analysis of predictor contributions.")
H(3,'2.5.1. Rolling-origin cross-validation')
body("For each year from 2009 to 2018, models were trained on all earlier development years and tested on that year, giving ten out-of-time folds per model, analogous to the 10-fold evaluation of Zhou et al. (2020) but respecting time order (Tashman, 2000). The rolling-origin error is the mean of the ten fold errors.")
H(3,'2.5.2. Random 10-fold cross-validation')
body("The same development data were shuffled into ten folds. The difference between this estimate and the hold-out error measures the optimism of random validation (Bergmeir & Benítez, 2012; Roberts et al., 2017), expressed in Eq. ({eq:opt}) as a percentage of the hold-out error.")
equation('opt',r'O\;(\%)=100\times\frac{\mathrm{RMSE}_{\mathrm{HO}}-\mathrm{RMSE}_{\mathrm{R10}}}{\mathrm{RMSE}_{\mathrm{HO}}}','where $\\mathrm{RMSE}_{\\mathrm{HO}}$ is the independent hold-out error and $\\mathrm{RMSE}_{\\mathrm{R10}}$ the mean error of the random 10-fold cross-validation.')
H(3,'2.5.3. Independent hold-out and significance testing')
body("Final models, trained on 1982–2018, were evaluated on 2019–2024. Uncertainty in RMSE was estimated with 1,000 bootstrap resamples. Differences against the best model were tested with the Diebold–Mariano (DM) statistic (Eq. ({eq:dm})) with the small-sample correction of Harvey et al. (1997), applied to region-averaged monthly loss series (Diebold & Mariano, 1995).")
equation('dm',r'd_{t}=e_{A,t}^{2}-e_{B,t}^{2},\qquad \mathrm{DM}=\frac{\bar{d}}{\sqrt{\hat{\sigma}_{d}^{2}/n}}','where $e_{A,t}$ and $e_{B,t}$ are the prediction errors of the compared model and of the best model in month $t$, $\\bar{d}$ is the mean loss differential, $\\hat{\\sigma}_d^2$ its variance estimate and $n$ = 72 months the length of the test series. The correction of Harvey et al. multiplies DM by $\\sqrt{(n-1)/n}$ for one-step-ahead forecasts. Positive statistics indicate a larger loss than the best model.')
H(3,'2.5.4. Leave-one-region-out validation')
body("In leave-one-region-out (LORO) validation, each region was withheld in turn; models were trained on the other two regions (1982–2018) without region dummies or climatology and tested on the withheld region in 2019–2024, following the target-oriented validation logic of Meyer et al. (2018) and Ploton et al. (2020).")
H(3,'2.5.5. Predictor importance and ablation')
body("Predictor contributions were assessed in three ways: Pearson correlations, impurity-based RF importance and permutation importance of the best model on the test period (Fisher et al., 2019; Strobl et al., 2007), computed as the mean increase in test RMSE over 30 random permutations of each predictor. A predictor-group ablation (G1 to G4) mirrors the stepwise input experiment of Zhou et al. (2020, Table 3).")

H(2,'2.6. Trend and change-point analysis')
body("Trends in annual mean Tmax, Tmin and DTR were tested with the Mann–Kendall test (Kendall, 1975; Mann, 1945) after trend-free pre-whitening to remove lag-1 autocorrelation (Yue et al., 2002), and with the seasonal Mann–Kendall test on monthly data (Hirsch et al., 1982), in which the statistics $S_m$ of the twelve calendar months are summed. For a series $x_1,\\dots,x_n$, the Mann–Kendall statistic and its standardized form are given by Eq. ({eq:mk}). Trend magnitudes were estimated with Sen's slope (Sen, 1968), Eq. ({eq:sen}). Abrupt shifts were located with the Pettitt test (Pettitt, 1979), Eq. ({eq:pet}). Because of an apparent step in recent Tmax, trends were also recomputed for 1981–2017.")
equation('mk',r'S=\sum_{i=1}^{n-1}\sum_{j=i+1}^{n}\mathrm{sgn}\left(x_{j}-x_{i}\right),\qquad Z=\begin{cases}\dfrac{S-1}{\sqrt{\mathrm{Var}(S)}} & S>0\\ 0 & S=0\\ \dfrac{S+1}{\sqrt{\mathrm{Var}(S)}} & S<0\end{cases}','where $\\mathrm{Var}(S)=n(n-1)(2n+5)/18$ for a series without ties.')
equation('sen',r'\beta=\mathrm{median}\left\{\frac{x_{j}-x_{i}}{j-i}:\;1\le i<j\le n\right\}')
equation('pet',r'U_{t}=\sum_{i=1}^{t}\sum_{j=t+1}^{n}\mathrm{sgn}\left(x_{i}-x_{j}\right),\qquad K=\max_{1\le t<n}\left|U_{t}\right|','where the change point is located at the year of $K$.')

# ---- 3 Results
H(1,'3. Results')
body("In this section, the results of variable importance are presented first, the performance of the six ML models is then verified with internal and independent data, and the transferability, validation optimism and underlying trends are analysed.")
def P(key): return secs[key]
H(2,'3.1. Variable importance')
a,b=P('3.1 Variable importance')
a=a.replace('(Fig. 3a, c;','({fig:imp}a, c;'); b=b.replace('(Fig. 3b, d)','({fig:imp}b, d)')
body(a); body(b)
figure('imp',M+'image3.png',5.6,'(a, c) Pearson correlation between each predictor and the target; (b, d) permutation importance (mean ± standard deviation over 30 repeats), shown as the increase in test root-mean-square error when a predictor is permuted, of the best model on the 2019–2024 test period for (a, b) maximum temperature (Tmax) with histogram-based gradient boosting (HGB) and (c, d) minimum temperature (Tmin) with support vector regression (SVR). DTR, diurnal temperature range.')
H(2,'3.2. Cross-validated model performance')
a,b=P('3.2 Cross-validated model performance')
a=a.replace('Figure 4 shows','{fig:cv} shows'); assert '{fig:cv}' in a
body(a); body(b)
figure('cv',M+'image4.png',5.6,'Boxplots of rolling-origin cross-validation (ten annual folds, 2009–2018) for the six models: (a–c) maximum temperature and (d–f) minimum temperature, in terms of root-mean-square error, bias and coefficient of determination. The best hold-out model is shaded orange.')
H(2,'3.3. Independent validation')
body(P('3.3 Independent validation')[0].replace('Table 4 reports the independent 2019–2024 hold-out results, and Figs. 5 and 6 compare','{tab:perf} reports the independent 2019–2024 hold-out results, and {fig:scatx} and {fig:scatn} compare').replace('Table 4 reports the independent 2019--2024 hold-out results, and Figs. 5 and 6 compare','{tab:perf} reports the independent 2019–2024 hold-out results, and {fig:scatx} and {fig:scatn} compare'))
H(3,'3.3.1. Overall accuracy')
caption('tab','perf','Performance on the independent test period (2019–2024, n = 216 per target) for (a) Tmax and (b) Tmin. Best machine-learning values are in bold; the 95 % confidence interval (CI) of RMSE was obtained by bootstrap resampling.',keep=True)
pp=doc.add_paragraph(); pp.paragraph_format.space_after=Pt(2); pp.paragraph_format.keep_with_next=True; run(pp,'(a) Tmax',bold=True,size=9)
data_table(3,fs=9,widths=[3.0,2.0,3.8,2.4,2.4,2.4])
pp=doc.add_paragraph(); pp.paragraph_format.space_after=Pt(2); pp.paragraph_format.space_before=Pt(6); pp.paragraph_format.keep_with_next=True; run(pp,'(b) Tmin',bold=True,size=9)
data_table(4,fs=9,widths=[3.0,2.0,3.8,2.4,2.4,2.4])
note('Abbreviations: '+ABB)
for ptxt in P('3.3.1 Overall accuracy'): body(ptxt)
H(3,'3.3.2. Statistical significance of differences')
body(P('3.3.2 Statistical significance of differences')[0].replace('Table 5 tests','{tab:dm} tests'))
caption('tab','dm','Modified Diebold–Mariano (DM) tests (Harvey et al., 1997) of each model against the best model on the test period. Positive statistics indicate a larger loss than the best model.',keep=True)
data_table(5,fs=9,widths=[4.0,3.0,3.0,3.0,3.0])
note('Abbreviations: '+ABB)
figure('scatx',M+'image5.png',5.6,'Predicted versus observed maximum temperature (Tmax) on the test period (2019–2024) for (a) k-nearest neighbours, (b) support vector regression, (c) random forest, (d) gradient boosting decision tree, (e) histogram-based gradient boosting, (f) deep neural network, (g) climatology and (h) persistence. Colours denote regions; the dashed line is the 1:1 line.')
figure('scatn',M+'image6.png',5.6,'As {fig:scatx}, for minimum temperature.')
H(3,'3.3.3. Regional performance')
body(P('3.3.3 Regional performance')[0].replace('(Fig. 7)','({fig:heat})'))
figure('heat',M+'image7.png',5.6,'Region-by-model heatmaps of test-period (a) maximum temperature root-mean-square error, (b) maximum temperature bias, (c) minimum temperature root-mean-square error and (d) minimum temperature bias.')
H(3,'3.3.4. Temporal behaviour and error structure')
a,b=P('3.3.4 Temporal behaviour and error structure')
body(a.replace('Figure 8 shows','{fig:ts} shows')); body(b.replace('(Fig. 9)','({fig:err})'))
figure('ts',M+'image8.png',5.6,'Observed (black), best-model predicted (orange) and climatological (dotted) monthly maximum temperature (left) and minimum temperature (right) for (a, b) Dodoma, (c, d) Singida and (e, f) Tabora during the test period 2019–2024.')
figure('err',M+'image9.png',5.6,'Error diagnostics of the best models on the test period: (a, c) error versus observed temperature with a least-squares line; (b, d) error by calendar month (blue: cool dry season; orange: hot dry season). (a, b) Maximum temperature with histogram-based gradient boosting; (c, d) minimum temperature with support vector regression.')
H(3,'3.3.5. Contribution of predictor groups')
body(P('3.3.5 Contribution of predictor groups')[0].replace('(Table 6)','({tab:abl})'))
caption('tab','abl','Contribution of predictor groups to test-period RMSE (°C). G1 = calendar, region and climatology; G2 = G1 + own lags; G3 = G2 + cross-variable lags and DTR; G4 = G3 + neighbouring regions (full model). Best value per row in bold.',keep=True)
data_table(6,fs=9,widths=[2.0,2.6,2.3,2.3,2.3,2.3])
note('Abbreviations: '+ABB+' DTR, diurnal temperature range.')
H(2,'3.4. Optimism of random cross-validation')
a,b=P('3.4 Optimism of random cross-validation')
body(a.replace('Figure 10 compares','{fig:opt} compares')); body(b)
figure('opt',M+'image10.png',5.6,'Validation optimism. Mean root-mean-square error (a, b) and bias (c, d) of each model estimated by random 10-fold cross-validation (grey), rolling-origin cross-validation (light blue) and the independent 2019–2024 hold-out (dark blue), for maximum temperature (a, c) and minimum temperature (b, d).')
H(2,'3.5. Spatial transferability')
a,b=P('3.5 Spatial transferability')
body(a.replace('Table 7 reports','{tab:loro} reports')); body(b)
caption('tab','loro','Leave-one-region-out test-period RMSE (°C) and, in parentheses, bias (°C). Best machine-learning value per column in bold.',keep=True)
data_table(7,fs=8,widths=[2.4,2.3,2.4,2.3,2.3,2.4,2.3])
note('Abbreviations: '+ABB)
H(2,'3.6. Trends and change points')
a,b,c=P('3.6 Trends and change points')
body(a.replace('Table 8 and Fig. 11 summarize','{tab:trend} and {fig:trend} summarize'))
caption('tab','trend','Trend and change-point statistics for annual mean temperature, 1981–2024. Z = Mann–Kendall statistic after trend-free pre-whitening; SMK = seasonal Mann–Kendall; slopes are Sen’s slope in °C per decade.',keep=True)
data_table(8,fs=7.5,widths=[1.5,1.4,1.5,1.2,1.2,1.5,1.4,1.5,1.9,2.0])
note('Abbreviations: DTR, diurnal temperature range; Tmax, maximum temperature; Tmin, minimum temperature.')
body(b); body(c)
figure('trend',M+'image11.png',5.8,'Annual mean (a) maximum temperature, (b) minimum temperature and (c) diurnal temperature range for the three regions, 1981–2024, with Sen’s slope lines (dashed) and Pettitt change points (dotted).')

# ---- 4 Discussion, 5 Limitations, 6 Conclusion
H(1,'4. Discussion')
H(2,'4.1. Significance and comparison with previous work')
body("The present results agree with, and extend, earlier work in four respects. First, the superiority of regularized boosting for Tmax mirrors the ranking of Zhou et al. (2020), whereas the advantage of SVR for Tmin is consistent with Salcedo-Sanz et al. (2016) and Cifuentes et al. (2020). Second, the weak performance of the DNN supports the finding that tree ensembles match or exceed deep networks on small tabular datasets (Grinsztajn et al., 2022; Shwartz-Ziv & Armon, 2022). Third, the 5.6–15.1 % optimism of random cross-validation quantifies, for a climate variable, the concerns raised by Bergmeir and Benítez (2012) and Roberts et al. (2017), and the larger errors in unseen regions echo the spatial over-optimism documented by Ploton et al. (2020). Fourth, the late-1990s Tmin shift and the narrowing of the DTR agree with Kebacho et al. (2024), Chang'a et al. (2017) and Vose et al. (2005), whereas the post-2017 Tmax decline contrasts with the increases reported by Gebrechorkos et al. (2019b) and Msilu et al. (2024) and requires confirmation with homogenized station records. Unlike previous Tanzanian applications, which used a single algorithm at a single farm (Tende & Mkiramweni, 2025), this study benchmarks six algorithms across three regions against climatology and persistence.")
H(2,'4.2. Practical implications')
body(secs['3.7.1 Practical implications'][0])
H(1,'5. Limitations')
body("This study has some limitations that point to useful directions for further work. The analysis relies on region-representative gridded series, so cross-checking against homogenized station records would help confirm whether the post-2017 Tmax step reflects the climate itself or the data product. The predictor set was deliberately restricted to temperature-derived variables so that the models can run with minimal inputs; adding rainfall, soil moisture, cloud cover, vegetation indices and large-scale climate drivers such as the El Niño–Southern Oscillation (ENSO) and the Indian Ocean Dipole (IOD) (Nicholson, 2017) may further improve skill, particularly in the rainy-season months. Finally, as with most data-driven approaches, the models learn from the range of conditions seen during training, so periodic retraining as the climate continues to warm is advisable (Zhou et al., 2020).")
H(1,'6. Conclusion')
c1,c2,c3=secs['4 Conclusions']
body(c1); body(c2)
c3=c3.replace('(ENSO, Indian Ocean Dipole)','(ENSO, IOD)').replace('sequence models such as LSTM networks','sequence models such as long short-term memory (LSTM) networks (Hochreiter & Schmidhuber, 1997)')
assert 'Hochreiter' in c3
body(c3)

# ---- back matter
H(1,'Acknowledgments')
body('[Funding agency, grant number and other acknowledgements]')
H(1,'Conflict of Interest')
body('The authors declare no competing interests.')
H(1,"Author's contributions")
body('[Author contributions, following the CRediT taxonomy]')
H(1,'Research Data')
body("The monthly Tmax and Tmin series analysed in this study and the Python code used for feature construction, model training, validation and trend analysis are available from the corresponding author on reasonable request.")
H(1,'Supporting Information')
body("Table S1: Full Pearson correlation matrices of predictors for the Tmax and Tmin models. Table S2: Per-fold rolling-origin and random cross-validation scores for all models. Tables S1 and S2 are supplied as a separate file.")

# ---- references (IEEE)
H(1,'References')
refs_par=[]
for k,idx in enumerate(cited):
    p=doc.add_paragraph(); pf=p.paragraph_format; pf.left_indent=Cm(1.0); pf.first_line_indent=Cm(-1.0); pf.space_after=Pt(3); p.alignment=WD_ALIGN_PARAGRAPH.JUSTIFY
    run(p,f'[{k+1}]\t',size=10)
    pf.tab_stops.add_tab_stop(Cm(1.0))
    for kind,txt in format_ref(REFS[idx]):
        run(p,txt.replace("'","’"),italic=(kind=='i') or None,size=10)
uncited=[i for i in range(len(REFS)) if i not in cited]
print('CITED',len(cited),'UNCITED',[REFS[i]['authors'][:25]+REFS[i]['year'] for i in uncited])

# footer page numbers
fp=sec.footer.paragraphs[0]; fp.alignment=WD_ALIGN_PARAGRAPH.CENTER
def fld(p,code):
    r=p.add_run(); r.font.size=Pt(10)
    for t,txt in (('begin',None),('instr',code),('end',None)):
        if t=='instr':
            e=OxmlElement('w:instrText'); e.set(qn('xml:space'),'preserve'); e.text=txt
        else:
            e=OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),t)
        r._r.append(e)
fld(fp,'PAGE')
doc.core_properties.title=TITLE; doc.core_properties.author=''
z=doc.settings.element.find(qn('w:zoom'))
if z is not None: z.set(qn('w:percent'),'100')
doc.save('out/Manuscript_AI_for_Science_restructured.docx')
print('saved')
