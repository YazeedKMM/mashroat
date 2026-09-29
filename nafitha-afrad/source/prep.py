from PIL import Image
import numpy as np
IMG='images/'; OUT='scratchpad/assets/'
NAVY=(0x05,0x29,0x77)

def keyout(im, lo=200, hi=245):
    a=np.asarray(im.convert('RGB')).astype(float)
    m=a.min(axis=2)
    alpha=np.clip((hi-m)/(hi-lo),0,1)
    out=np.dstack([a,alpha*255]).astype(np.uint8)
    return Image.fromarray(out,'RGBA')

def trim(im):
    bb=im.getchannel('A').point(lambda v:255 if v>20 else 0).getbbox()
    return im.crop(bb)

# logo mark and wordmark from concept image
src=Image.open(IMG+'1.jpg')
mark=keyout(src.crop((348,560,678,902)))
# drop stray specks (arrow tips / jpeg noise) so trim() hugs the mark itself
from collections import deque
a=np.array(mark); m=a[:,:,3]>20; h,w=m.shape; seen=np.zeros((h,w),bool)
for y0 in range(h):
    for x0 in range(w):
        if m[y0,x0] and not seen[y0,x0]:
            comp=[]; q=deque([(y0,x0)]); seen[y0,x0]=True
            while q:
                cy,cx=q.popleft(); comp.append((cy,cx))
                for dy,dx in((1,0),(-1,0),(0,1),(0,-1)):
                    ny,nx=cy+dy,cx+dx
                    if 0<=ny<h and 0<=nx<w and m[ny,nx] and not seen[ny,nx]:
                        seen[ny,nx]=True; q.append((ny,nx))
            if len(comp)<40:
                for cy,cx in comp: a[cy,cx,3]=0
mark=trim(Image.fromarray(a))
mark.save(OUT+'logo_mark.png'); print('mark',mark.size)
for name,col in (('logo_mark_white',(255,255,255)),('logo_mark_navy',NAVY)):
    b=np.array(mark); b[:,:,:3]=col; Image.fromarray(b).save(OUT+name+'.png')
word=trim(keyout(src.crop((370,905,650,1000))))
word.save(OUT+'logo_word.png'); print('word',word.size)

# city motifs: gray -> navy, white -> transparent
def recolor(im, color, lo=150, hi=235):
    a=np.asarray(im.convert('RGB')).astype(float)
    m=a.min(axis=2)
    alpha=np.clip((hi-m)/(hi-lo),0,1)
    h,w=m.shape
    rgb=np.zeros((h,w,3)); rgb[:]=color
    return Image.fromarray(np.dstack([rgb,alpha*255]).astype(np.uint8),'RGBA')

# split = middle of the blank gap between the pattern tile and the single motif
cities={'jubail':('3.png',354),'raskhair':('4.png',341),'yanbu':('5.png',347),'jazan':('2.png',355)}
for k,(f,split) in cities.items():
    im=Image.open(IMG+f); w,h=im.size
    motif=trim(recolor(im.crop((split,0,w,h)),NAVY)); motif.save(OUT+f'city_{k}.png')
    tile=trim(recolor(im.crop((0,0,split,h)),NAVY)); tile.save(OUT+f'tile_{k}.png')
    print(k,motif.size,tile.size)

# entity logos (full colour, light-background versions)
def to_rgba_keyed(im):
    im=im.convert('RGBA'); a=np.asarray(im).astype(float)
    m=a[:,:,:3].min(axis=2)
    alpha=np.minimum(a[:,:,3], np.clip((252-m)/(252-200),0,1)*255)
    return Image.fromarray(np.dstack([a[:,:,:3],alpha]).astype(np.uint8),'RGBA')
for name,f in (('logo_rcjy','6.png'),('logo_jabeen','7.png'),('logo_rqeem','9.png')):
    im=trim(to_rgba_keyed(Image.open(IMG+f))); im.save(OUT+name+'.png'); print(name, im.size)
