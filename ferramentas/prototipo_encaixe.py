import math, random, time
from PIL import Image, ImageDraw

def ring(cx,cy,r1,r2,n=40):
    o=[(cx+r1*math.cos(2*math.pi*k/n),cy+r1*math.sin(2*math.pi*k/n)) for k in range(n)]
    i=[(cx+r2*math.cos(2*math.pi*k/n),cy+r2*math.sin(2*math.pi*k/n)) for k in range(n)]
    return [o,i]
def lshape(w,h,t):
    return [[(0,0),(w,0),(w,t),(t,t),(t,h),(0,h)]]
def rect(w,h): return [[(0,0),(w,0),(w,h),(0,h)]]

def depths(polys):
    def inside(x,y,p):
        r=False;j=len(p)-1
        for i in range(len(p)):
            if (p[i][1]>y)!=(p[j][1]>y) and x<(p[j][0]-p[i][0])*(y-p[i][1])/(p[j][1]-p[i][1])+p[i][0]: r=not r
            j=i
        return r
    return [sum(1 for j,q in enumerate(polys) if j!=i and inside(p[0][0],p[0][1],q)) for i,p in enumerate(polys)]

def raster(polys, ang, res, inside_ok, cx, cy):
    """returns (w,h,xmin,ymin,grid set of (c,r))"""
    if not inside_ok:
        d=depths(polys); polys=[p for p,dd in zip(polys,d) if dd==0]
    ca,sa=math.cos(ang),math.sin(ang)
    R=[[(cx+(x-cx)*ca-(y-cy)*sa, cy+(x-cx)*sa+(y-cy)*ca) for x,y in p] for p in polys]
    xs=[x for p in R for x,_ in p]; ys=[y for p in R for _,y in p]
    xmin,ymin=min(xs),min(ys)
    W=int(math.floor((max(xs)-xmin)/res))+1; H=int(math.floor((max(ys)-ymin)/res))+1
    g=bytearray(W*H)
    # scanline fill (even-odd) - cell centers
    rows=[[] for _ in range(H)]
    for p in R:
        n=len(p)
        for i in range(n):
            x1,y1=p[i]; x2,y2=p[(i+1)%n]
            if y1==y2: continue
            if y1>y2: x1,y1,x2,y2=x2,y2,x1,y1
            r0=max(0,int(math.ceil((y1-ymin)/res-0.5))); r1=min(H-1,int(math.ceil((y2-ymin)/res-0.5))-1)
            for r in range(r0,r1+1):
                yc=ymin+(r+0.5)*res
                rows[r].append(x1+(yc-y1)*(x2-x1)/(y2-y1))
    for r in range(H):
        L=sorted(rows[r])
        for k in range(0,len(L)-1,2):
            c0=int(math.floor((L[k]-xmin)/res+0.5)); c1=int(math.floor((L[k+1]-xmin)/res-0.5))
            for c in range(max(0,c0),min(W-1,c1)+1): g[r*W+c]=1
    # edges (conservative)
    for p in R:
        n=len(p)
        for i in range(n):
            x1,y1=p[i]; x2,y2=p[(i+1)%n]
            L=math.hypot(x2-x1,y2-y1); st=max(1,int(L/(res*0.5)))
            for k in range(st+1):
                x=x1+(x2-x1)*k/st; y=y1+(y2-y1)*k/st
                c=min(W-1,int((x-xmin)/res)); r=min(H-1,int((y-ymin)/res)); g[r*W+c]=1
    return W,H,xmin,ymin,g

def runs(W,H,g):
    out=[]
    for r in range(H):
        c=0
        while c<W:
            if g[r*W+c]:
                a=c
                while c<W and g[r*W+c]: c+=1
                out.append((r,a,c-1))
            else: c+=1
    return out
def dilate(W,H,g,k):
    if k<=0: return W,H,g,0
    W2,H2=W+2*k,H+2*k; d=bytearray(W2*H2)
    for r in range(H):
        for c in range(W):
            if g[r*W+c]:
                for rr in range(r,r+2*k+1):
                    base=rr*W2
                    for cc in range(c,c+2*k+1): d[base+cc]=1
    return W2,H2,d,k

def nest(pieces, SW, SH, res, gap, step, inside_ok, timer=0):
    GW=int(SW/res); GH=int(SH/res); k=int(math.ceil(gap/res))
    angs=[math.radians(a) for a in range(0,360,step)] if step else [0]
    R=[]
    for polys in pieces:
        xs=[x for p in polys for x,_ in p]; ys=[y for p in polys for _,y in p]
        cx=(min(xs)+max(xs))/2; cy=(min(ys)+max(ys))/2
        lst=[]
        for a in angs:
            W,H,xm,ym,g=raster(polys,a,res,inside_ok,cx,cy)
            if W>GW or H>GH: lst.append(None); continue
            Wd,Hd,gd,kk=dilate(W,H,g,k)
            lst.append((W,H,xm,ym,runs(W,H,g),runs(Wd,Hd,gd),kk,sum(g)))
        R.append(lst)
    order=sorted(range(len(pieces)),key=lambda i:-max((x[7] for x in R[i] if x),default=0))
    sheets=[]  # each: occ bytearray, pre list of arrays
    res_place=[]
    def newsheet():
        occ=bytearray(GW*GH); pre=[[0]*(GW+1) for _ in range(GH)]; sheets.append((occ,pre))
    def fits(sh,rs,x,y):
        occ,pre=sh
        for (r,a,b) in rs:
            P=pre[y+r]
            if P[x+b+1]-P[x+a]>0: return False
        return True
    tests=0
    for i in order:
        best=None
        if all(x is None for x in R[i]): res_place.append((i,None)); continue
        s=0
        while best is None:
            if s>=len(sheets): newsheet()
            sh=sheets[s]
            for ai,ra in enumerate(R[i]):
                if ra is None: continue
                W,H,xm,ym,rs,rd,kk,_=ra
                ylim=GH-H if best is None else min(GH-H, best[0]-H)
                found=False
                for y in range(0,ylim+1):
                    for x in range(0,GW-W+1):
                        tests+=1
                        if fits(sh,rs,x,y):
                            sc=(y+H, x)
                            if best is None or sc<best[0:2] if False else (best is None or (y+H)*10**6+x < best[0]*10**6+best[1]):
                                best=(y+H,x,ai,x,y)
                            found=True;break
                    if found: break
            if best is None: s+=1
        _,_,ai,x,y=best
        W,H,xm,ym,rs,rd,kk,_=R[i][ai]
        occ,pre=sheets[s]
        for (r,a,b) in rd:
            yy=y+r-kk
            if 0<=yy<GH:
                for c in range(max(0,x+a-kk),min(GW-1,x+b-kk)+1): occ[yy*GW+c]=1
                P=pre[yy]; acc=0
                for c in range(GW): P[c]=acc; acc+=occ[yy*GW+c]
                P[GW]=acc
        res_place.append((i,(s,ai,x,y)))
    return res_place,len(sheets),R,tests,GW,GH

random.seed(3)
pieces=[]
for _ in range(25):
    t=random.choice(['ring','L','rect'])
    if t=='ring': r=random.uniform(30,70); pieces.append(ring(0,0,r,r*0.7))
    elif t=='L': pieces.append(lshape(random.uniform(40,120),random.uniform(40,120),random.uniform(10,20)))
    else: pieces.append(rect(random.uniform(10,40),random.uniform(10,30)))
t0=time.time()
pl,ns,R,tests,GW,GH=nest(pieces,1000,370,4,2,45,True)
print('sheets',ns,'tests',tests,'time %.1f'%(time.time()-t0))
# verify no overlap at raster level (undilated) and spacing
grid={}
bad=0
for i,p in pl:
    if not p: continue
    s,ai,x,y=p; W,H,xm,ym,rs,rd,kk,_=R[i][ai]
    for (r,a,b) in rs:
        for c in range(a,b+1):
            key=(s,x+c,y+r)
            if key in grid: bad+=1
            grid[key]=i
print('overlap cells',bad)
# render
img=Image.new('RGB',(GW*ns,GH),'white'); px=img.load()
for i,p in pl:
    if not p: continue
    s,ai,x,y=p; W,H,xm,ym,rs,rd,kk,_=R[i][ai]
    col=tuple(random.randint(40,200) for _ in range(3))
    for (r,a,b) in rs:
        for c in range(a,b+1): px[s*GW+x+c,GH-1-(y+r)]=col
img=img.resize((img.width*3,img.height*3),Image.NEAREST); img.save('/home/claude/nest/proto.png')
