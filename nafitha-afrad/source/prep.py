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
mark=trim(keyout(src.crop((348,560,678,902))))
mark.save(OUT+'logo_mark.png'); print('mark',mark.size)
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

cities={'jubail':'3.png','raskhair':'4.png','yanbu':'5.png','jazan':'2.png'}
for k,f in cities.items():
    im=Image.open(IMG+f); w,h=im.size
    split=int(w*0.62)
    motif=trim(recolor(im.crop((split,0,w,h)),NAVY)); motif.save(OUT+f'city_{k}.png')
    tile=trim(recolor(im.crop((0,0,split,h)),NAVY)); tile.save(OUT+f'tile_{k}.png')
    print(k,motif.size,tile.size)
