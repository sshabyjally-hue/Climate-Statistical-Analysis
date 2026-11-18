import re, json
ABBR={
'Advances in Space Research':'Adv. Space Res.','PLOS ONE':'PLOS ONE','The American Statistician':'Amer. Statistician',
'Spatial Statistics':'Spatial Statist.','ISPRS Journal of Photogrammetry and Remote Sensing':'ISPRS J. Photogramm. Remote Sens.',
'Information Sciences':'Inf. Sci.','Computational Statistics & Data Analysis':'Comput. Statist. Data Anal.','Machine Learning':'Mach. Learn.',
'Geoscientific Model Development':'Geosci. Model Dev.','Atmospheric and Climate Sciences':'Atmos. Climate Sci.','Energies':'Energies',
'Journal of Business & Economic Statistics':'J. Bus. Econ. Statist.','Earth Perspectives':'Earth Perspect.','Science':'Science',
'Journal of Machine Learning Research':'J. Mach. Learn. Res.','The Annals of Statistics':'Ann. Statist.',
'International Journal of Climatology':'Int. J. Climatol.','Scientific Reports':'Sci. Rep.','Journal of Climate':'J. Climate',
'Atmospheric Research':'Atmos. Res.','Advances in Neural Information Processing Systems':'Adv. Neural Inf. Process. Syst.',
'International Journal of Forecasting':'Int. J. Forecasting','Quarterly Journal of the Royal Meteorological Society':'Quart. J. Roy. Meteorol. Soc.',
'Water Resources Research':'Water Resour. Res.','Neural Computation':'Neural Comput.','Remote Sensing of Environment':'Remote Sens. Environ.',
'Nature':'Nature','Tanzania Journal of Science':'Tanzania J. Sci.','Journal of the Geographical Association of Tanzania':'J. Geogr. Assoc. Tanzania',
'Advances in Meteorology':'Adv. Meteorol.','Frontiers in Environmental Science':'Front. Environ. Sci.','Econometrica':'Econometrica',
'Environmental Modelling & Software':'Environ. Model. Softw.','Ghana Journal of Geography':'Ghana J. Geogr.','Journal of Hydrology':'J. Hydrol.',
'Reviews of Geophysics':'Rev. Geophys.','Proceedings of the National Academy of Sciences':'Proc. Nat. Acad. Sci. USA',
'Theoretical and Applied Climatology':'Theor. Appl. Climatol.','Journal of the American Statistical Association':'J. Amer. Statist. Assoc.',
'Information Fusion':'Inf. Fusion','Statistics and Computing':'Statist. Comput.','Journal of Open Source Software':'J. Open Source Softw.',
'BMC Bioinformatics':'BMC Bioinformatics','East African Journal of Engineering':'East Afr. J. Eng.','Water':'Water',
'Geophysical Research Letters':'Geophys. Res. Lett.','Climate Research':'Climate Res.','Hydrological Processes':'Hydrol. Process.',
'Journal of the Royal Statistical Society: Series C (Applied Statistics)':'J. Roy. Statist. Soc. C, Appl. Statist.',
'Nature Communications':'Nature Commun.','Ecography':'Ecography','Agricultural and Forest Meteorology':'Agric. Forest Meteorol.',
'Science of the Total Environment':'Sci. Total Environ.','Proceedings of the National Academy of Sciences':'Proc. Nat. Acad. Sci. USA',
}
def clean(s): return s.replace('\\','').replace('--','–').strip()
def fmt_authors(a):
    a=a.strip()
    if 'Intergovernmental' in a: return 'Intergovernmental Panel on Climate Change'
    a=a.replace('. . .',',')
    names=re.findall(r"\s*(?:&\s*)?([^,]+?),\s*((?:[A-ZÀ-Ž]\.(?:-[A-ZÀ-Ž]\.)?\s*)+)",a+' ')
    out=[]
    for sn,ini in names:
        ini=re.sub(r'\.(?=[A-Z])','. ',ini.strip())
        out.append(f"{ini} {sn.strip()}")
    # APA with >=21 authors uses ellipsis -> treat as many
    if len(out)>6: return out[0]+' et al.'
    if len(out)==1: return out[0]
    if len(out)==2: return ' and '.join(out)
    return ', '.join(out[:-1])+', and '+out[-1]
def pages(pg):
    pg=clean(pg)
    return pg
def format_ref(r):
    au=fmt_authors(r['authors']); yr=re.sub('[a-z]$','',r['year']); rest=r['rest'].replace('\\<','<').replace('\\>','>')
    doi=None
    m=re.search(r'\s(https?://\S+?)\.?$',rest)
    link=None
    if m:
        link=m.group(1); rest=rest[:m.start()].strip()
        if 'doi.org/' in link: doi=link.split('doi.org/',1)[1]
    # books / chapters
    if rest.startswith('*Rank correlation'):
        return [('t',f"{au}, "),('i','Rank Correlation Methods'),('t',f", 4th ed. London, U.K.: Charles Griffin, {yr}.")]
    if rest.startswith('*Climate change 2021'):
        return [('t',f"{au}, "),('i','Climate Change 2021: The Physical Science Basis. Contribution of Working Group I to the Sixth Assessment Report of the Intergovernmental Panel on Climate Change'),('t',f". Cambridge, U.K.: Cambridge Univ. Press, {yr}, doi: {doi}.")]
    if rest.startswith('Support vector regression machines'):
        return [('t',f"{au}, “Support vector regression machines,” in "),('i','Advances in Neural Information Processing Systems 9'),('t',f", M. C. Mozer, M. I. Jordan, and T. Petsche, Eds. Cambridge, MA, USA: MIT Press, {yr}, pp. 155–161.")]
    if rest.startswith('XGBoost'):
        return [('t',f"{au}, “XGBoost: A scalable tree boosting system,” in "),('i','Proc. 22nd ACM SIGKDD Int. Conf. Knowl. Discovery Data Mining'),('t',f", {yr}, pp. 785–794, doi: {doi}.")]
    m=re.match(r'^(.*?[.?])\s\*([^*]+)\*,\s*(.*)$',rest,re.S)
    title,jr,tail=m.group(1),m.group(2),m.group(3)
    title=title.rstrip('.') if not title.endswith('?') else title
    jabbr=ABBR.get(jr.replace('\\',''),jr)
    vol=iss=None
    mv=re.match(r'^\*([^*]+)\*(?:\(([^)]+)\))?,?\s*(.*)$',tail)
    parts=[]
    if mv:
        vol,iss,pg=mv.group(1),mv.group(2),mv.group(3).strip().rstrip('.')
    else: pg=tail.strip().rstrip('.')
    segs=[]
    if jr=='Advances in Neural Information Processing Systems':
        segs.append(f", vol. {vol}, {yr}, pp. {clean(pg)}"+".")
        return [('t',f"{au}, “{clean(title)}{',' if not title.endswith('?') else ''}” in "),('i','Proc. '+jabbr),('t',''.join(segs))]
    s=''
    if vol: s+=f", vol. {vol}"
    if iss: s+=f", no. {iss}"
    pg=clean(pg)
    if pg.startswith('Article'): s+=f", Art. no. {pg.replace('Article ','')}"
    elif pg: s+=f", pp. {pg}"
    s+=f", {yr}"
    if doi: s+=f", doi: {doi}"
    elif link: s+=f". [Online]. Available: {link}"
    s+='.'
    return [('t',f"{au}, “{clean(title)}{',' if not title.endswith('?') else ''}” "),('i',jabbr),('t',s)]
if __name__=='__main__':
    refs=json.load(open('refs_raw.json'))
    for i,r in enumerate(refs): print(i+1,''.join(x[1] for x in format_ref(r)))
