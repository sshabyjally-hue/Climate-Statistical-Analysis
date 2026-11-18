import numpy as np, matplotlib
matplotlib.use('Agg')
import matplotlib.pyplot as plt
from matplotlib.patches import Circle, FancyBboxPatch, Rectangle, FancyArrowPatch
plt.rcParams['font.family']=['Arial','Liberation Sans','DejaVu Sans']
plt.rcParams['font.size']=9
plt.rcParams['mathtext.fontset']='custom'
plt.rcParams['mathtext.rm']='Liberation Sans'
plt.rcParams['mathtext.it']='Liberation Sans:italic'
plt.rcParams['pdf.fonttype']=42; plt.rcParams['ps.fonttype']=42
C1='#1f77b4';C2='#ff7f0e';G='#8c8c8c';R='#d62728'
fig,axs=plt.subplots(3,2,figsize=(6.3,8.4))
def lab(ax,l,t):
    ax.text(-0.02,1.06,l,transform=ax.transAxes,fontweight='bold',fontsize=11,va='bottom',ha='left')
    ax.text(0.09,1.06,t,transform=ax.transAxes,fontweight='bold',fontsize=9,va='bottom',ha='left')
for a in axs.ravel(): a.axis('off')
rng=np.random.default_rng(3)
# a KNN
ax=axs[0,0]; lab(ax,'a','k-nearest neighbours')
pts=rng.uniform(0,10,(45,2)); q=np.array([5.2,4.8]); d=np.linalg.norm(pts-q,axis=1); idx=np.argsort(d)[:5]; r=d[idx].max()+0.05
ax.scatter(pts[:,0],pts[:,1],c=pts[:,0]*0.3+pts[:,1]*0.2,cmap='Blues',s=22,edgecolor='k',linewidth=0.3,vmin=-1,vmax=5)
for i in idx: ax.plot([q[0],pts[i,0]],[q[1],pts[i,1]],color=C2,lw=1)
ax.scatter(pts[idx,0],pts[idx,1],s=34,facecolor='none',edgecolor=C2,linewidth=1.4)
ax.add_patch(Circle(q,r,fill=False,ls='--',ec=G)); ax.scatter(*q,marker='*',s=130,c=R,zorder=5)
ax.text(q[0]+0.3,q[1]+0.35,'query',color=R,fontsize=9,bbox=dict(fc='white',ec='none',pad=0.5,alpha=0.8))
ax.text(5,-0.6,'k = 5 nearest neighbours in standardized\npredictor space; prediction = distance-\nweighted mean of their targets',fontsize=9,va='top',ha='center')
ax.set_xlim(0,10);ax.set_ylim(-3.6,10.3)
# b SVR
ax=axs[0,1]; lab(ax,'b','Support vector regression')
x=np.linspace(0,10,40); yt=np.sin(x*0.65)*1.6+0.2*x; y=yt+rng.normal(0,0.45,x.size)
xs=np.linspace(0,10,200); f=np.sin(xs*0.65)*1.6+0.2*xs; eps=0.7
ax.fill_between(xs,f-eps,f+eps,color=C1,alpha=0.18,lw=0); ax.plot(xs,f,color=C1,lw=1.6)
ax.plot(xs,f+eps,color=C1,lw=0.6,ls='--'); ax.plot(xs,f-eps,color=C1,lw=0.6,ls='--')
fi=np.interp(x,xs,f); out=np.abs(y-fi)>eps
ax.scatter(x[~out],y[~out],s=16,c='#555555',zorder=3); ax.scatter(x[out],y[out],s=30,facecolor='none',edgecolor=R,linewidth=1.3,zorder=4)
ax.annotate('ε-insensitive tube',xy=(8.6,f[np.argmin(abs(xs-8.6))]+eps),xytext=(5.6,5.6),fontsize=9,arrowprops=dict(arrowstyle='-',color=G))
ax.text(0.1,6.0,'○ support vectors\n   (outside tube)',color=R,fontsize=9,va='top')
ax.text(5,-2.2,'blue line: fitted function f(x) with radial-basis\nkernel; errors inside the tube are not penalized',fontsize=9,va='top',ha='center')
ax.set_xlim(0,10.2);ax.set_ylim(-4.6,6.2)
# c RF
ax=axs[1,0]; lab(ax,'c','Random forest')
ax.set_xlim(0,10);ax.set_ylim(0,11.2)
def tree(cx,cy,s=0.8):
    nodes=[(cx,cy),(cx-0.8*s,cy-1.1*s),(cx+0.8*s,cy-1.1*s),(cx-1.2*s,cy-2.2*s),(cx-0.4*s,cy-2.2*s),(cx+0.4*s,cy-2.2*s),(cx+1.2*s,cy-2.2*s)]
    for a,b in [(0,1),(0,2),(1,3),(1,4),(2,5),(2,6)]: ax.plot([nodes[a][0],nodes[b][0]],[nodes[a][1],nodes[b][1]],color=G,lw=0.8)
    for i,(nx,ny) in enumerate(nodes): ax.add_patch(Circle((nx,ny),0.15,fc=C1 if i<3 else C2,ec='k',lw=0.4,zorder=3))
ax.add_patch(FancyBboxPatch((0.2,3.9),1.8,2.4,boxstyle='round,pad=0.05',fc='#eeeeee',ec=G)); ax.text(1.1,5.1,'Training\ndata',ha='center',va='center',fontsize=9)
tops=[9.0,5.9,2.8]
for k,cy in enumerate(tops):
    ax.annotate('',xy=(3.0,cy-0.9),xytext=(2.1,5.1),arrowprops=dict(arrowstyle='->',color=G,lw=0.8))
    tree(4.3,cy); ax.text(5.5,cy-0.9,f'tree {k+1}' if k<2 else 'tree B',ha='left',va='center',fontsize=9)
    ax.annotate('',xy=(7.7,5.1),xytext=(6.9,cy-0.9),arrowprops=dict(arrowstyle='->',color=G,lw=0.8))
ax.text(5.0,10.9,'bootstrap sample + random features per tree',fontsize=9,va='top',ha='center')
ax.add_patch(FancyBboxPatch((7.8,4.1),2.0,2.0,boxstyle='round,pad=0.05',fc='#fde9d4',ec=C2)); ax.text(8.8,5.1,'average of\nB trees',ha='center',va='center',fontsize=9)
# d GBDT
ax=axs[1,1]; lab(ax,'d','Gradient boosting decision tree')
ax.set_xlim(0,10);ax.set_ylim(0,10)
bx=[(0.1,'F$_0$\n(mean)'),(2.5,'ν·h$_1$'),(4.9,'ν·h$_2$'),(8.3,'ν·h$_M$')]
for i,(x0,t) in enumerate(bx):
    w=1.6
    ax.add_patch(FancyBboxPatch((x0,6.0),w,1.6,boxstyle='round,pad=0.05',fc='#dbe9f6' if i==0 else '#fde9d4',ec=C1 if i==0 else C2)); ax.text(x0+w/2,6.8,t,ha='center',va='center',fontsize=9)
    if i<2: ax.text(x0+w+0.4,6.8,'+',ha='center',va='center',fontsize=11)
ax.text(7.0,6.8,'+ … +',ha='center',va='center',fontsize=9)
ax.text(5,9.4,'trees are added sequentially (M rounds)',ha='center',fontsize=9)
for x0 in (2.5,4.9,8.3):
    ax.annotate('',xy=(x0+0.8,5.85),xytext=(x0+0.8,4.6),arrowprops=dict(arrowstyle='->',color=G,lw=0.8))
ax.text(5.1,4.0,'each tree h$_m$ is fitted to the residuals (negative\ngradient) of the current ensemble F$_{m-1}$',ha='center',va='top',fontsize=9)
ax.annotate('',xy=(9.7,5.0),xytext=(0.1,5.0),arrowprops=dict(arrowstyle='-',color='none'))
ax.add_patch(FancyBboxPatch((2.7,0.5),4.6,1.3,boxstyle='round,pad=0.05',fc='#eeeeee',ec=G)); ax.text(5,1.15,'F$_M$(x) = F$_0$ + ν Σ h$_m$(x)',ha='center',va='center',fontsize=9)
# e HGB
ax=axs[2,0]; lab(ax,'e','Histogram-based gradient boosting')
ax.set_xlim(0,10);ax.set_ylim(0,12)
vals=np.concatenate([rng.normal(3,1.2,300),rng.normal(7,1.0,300)]); h,e=np.histogram(vals,bins=12,range=(0,10))
h=h/h.max()*4.3
for i in range(12): ax.add_patch(Rectangle((0.6+e[i]*0.85,5.0),(e[1]-e[0])*0.85,h[i],fc=C1 if i<6 else C2,ec='k',lw=0.4,alpha=0.85))
ax.plot([0.6,9.2],[5,5],color='k',lw=0.8); ax.plot([0.6+e[6]*0.85]*2,[5,11.0],color=R,ls='--',lw=1.2)
ax.text(0.6+e[6]*0.85,11.4,'best split (maximum gain)',color=R,fontsize=9,va='bottom',ha='center')
ax.text(5.0,4.6,'predictor value, grouped into 12 bins (example)',ha='center',va='top',fontsize=9)
ax.add_patch(FancyBboxPatch((0.3,0.4),4.2,2.0,boxstyle='round,pad=0.05',fc='#dbe9f6',ec=C1)); ax.text(2.4,1.4,'left leaf\nweight w$_L$',ha='center',va='center',fontsize=9)
ax.add_patch(FancyBboxPatch((5.4,0.4),4.2,2.0,boxstyle='round,pad=0.05',fc='#fde9d4',ec=C2)); ax.text(7.5,1.4,'right leaf\nweight w$_R$',ha='center',va='center',fontsize=9)
ax.text(5.0,3.1,'leaf weights shrunk by L2 penalty λ',ha='center',va='center',fontsize=9)
# f DNN
ax=axs[2,1]; lab(ax,'f','Deep neural network')
ax.set_xlim(0,10);ax.set_ylim(0,10)
layers=[(1.2,[8.0,6.8,5.6,3.4,2.2],'Input\n(18 predictors)'),(4.0,[8.2,6.6,5.0,3.4,1.8],'Hidden\nlayer 1'),(6.4,[7.6,6.0,4.4,2.8],'Hidden\nlayer 2'),(8.9,[5.0],'Output\n$\\hat{y}$')]
for (x0,ys,_),(x1,ys1,_) in zip(layers[:-1],layers[1:]):
    for a in ys:
        for b in ys1: ax.plot([x0,x1],[a,b],color='#cccccc',lw=0.5,zorder=1)
for x0,ys,t in layers:
    for y0 in ys: ax.add_patch(Circle((x0,y0),0.32,fc='#dbe9f6' if x0<2 else ('#fde9d4' if x0>8 else 'white'),ec='k',lw=0.6,zorder=3))
    ax.text(x0,0.15 if x0<8 else 3.7,t,ha='center',va='bottom' if x0<8 else 'top',fontsize=9)
ax.text(1.2,4.5,'⋮',ha='center',va='center',fontsize=11)
ax.text(5,9.9,'weights W, biases b, ReLU activation',ha='center',va='top',fontsize=9)
fig.subplots_adjust(left=0.03,right=0.99,top=0.96,bottom=0.02,hspace=0.28,wspace=0.12)
fig.savefig('out/Figure_2_model_schematics.png',dpi=300)
fig.savefig('out/Figure_2_model_schematics.pdf')
fig.savefig('out/Figure_2_model_schematics.eps')
