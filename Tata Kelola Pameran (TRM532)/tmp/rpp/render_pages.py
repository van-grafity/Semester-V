import sys
from pathlib import Path
sys.path.insert(0,str(Path(__file__).parent/'deps'))
import pymupdf as fitz
from PIL import Image,ImageDraw
folder=Path(__file__).parent/'render'
d=fitz.open(folder/'rpp.pdf')
print('Pages',len(d))
for i,p in enumerate(d):
    p.get_pixmap(matrix=fitz.Matrix(1.5,1.5)).save(folder/f'page-{i+1}.png')
    text=p.get_text()
    print(i+1,len(text),text[:105].replace('\n',' | '), 'END:',text[-100:].replace('\n',' | '))
for start in range(0,len(d),6):
    canvas=Image.new('RGB',(1200,850),'#aaa');draw=ImageDraw.Draw(canvas)
    for i in range(start,min(start+6,len(d))):
        j=i-start
        canvas.paste(Image.open(folder/f'page-{i+1}.png').resize((397,560)),((j%3)*400,(j//3)*425)) if False else None
    canvas=Image.new('RGB',(1200,1140),'#aaa');draw=ImageDraw.Draw(canvas)
    for i in range(start,min(start+6,len(d))):
        j=i-start; canvas.paste(Image.open(folder/f'page-{i+1}.png').resize((397,562)),((j%3)*400,(j//3)*570));draw.text(((j%3)*400+5,(j//3)*570+5),str(i+1),fill='red')
    canvas.save(folder/f'contact-{start+1}.png')
