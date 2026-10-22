"""Validate a drawing plan and its native PPTX; optionally inspect/export PDF pages."""
import argparse,json,math,re,sys,zipfile
from pathlib import Path
from xml.etree import ElementTree as E

NS={'p':'http://schemas.openxmlformats.org/presentationml/2006/main',
    'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
def require(condition,message):
    if not condition:raise ValueError(message)
def numeric(o,keys):
    for key in keys:
        require(key in o and isinstance(o[key],(int,float)) and not isinstance(o[key],bool)
                and math.isfinite(o[key]),f'Invalid {key}: {o.get("id", "canvas")}')
def plain_text(t):return re.sub(r'_\{([^{}]+)\}',r'\1',t).replace('\r','').replace('\n','')
def normalized_font(s):return re.sub(r'[^a-z]','',s.lower()).replace('bold','').replace('italic','').replace('regular','').replace('psmt','')

def validate(plan,pptx=None,pdf=None,allow_images=False,split_dir=None,overwrite_splits=False):
    slides=json.loads(Path(plan).read_text(encoding='utf-8-sig'))
    require(isinstance(slides,list) and slides,'Plan must be a nonempty slide array')
    figures=set();canvas=None;result={'plan':'passed','slides':[]}
    for slide in slides:
        numeric(slide,['width','height'])
        require(slide['width']>0 and slide['height']>0,'Canvas must be positive')
        size=(slide['width'],slide['height'])
        require(canvas in (None,size),'Mixed slide canvas sizes');canvas=size
        n=slide.get('figure');require(type(n) is int and n>0 and n not in figures,'Invalid or duplicate figure number');figures.add(n)
        require(isinstance(slide.get('objects'),list) and slide['objects'],'Missing objects')
        ids=set()
        for o in slide['objects']:
            require(isinstance(o.get('id'),str) and o['id'] and o['id'] not in ids,'Missing or duplicate object ID');ids.add(o['id'])
            k=o.get('kind');require(k in ['shape','text','line','path'],f'Unsupported native object: {k}')
            if k in ['shape','text']:
                numeric(o,['x','y','w','h']);require(o['w']>0 and o['h']>0,'Nonpositive object size')
            if k=='shape':require(o.get('shape') in ['rect','roundRect','ellipse'],'Unsupported shape')
            if k=='text':
                numeric(o,['size']);require(o['size']>0 and isinstance(o.get('text'),str),'Invalid text')
                require(not re.search(r'\\[A-Za-z]+|\^',o['text']),'Complex TeX needs a separate editable math implementation')
                require(not re.search(r'_\{[^}]*\{',o['text']),'Nested subscripts unsupported')
            if k=='line':numeric(o,['x1','y1','x2','y2','sw'])
            if k=='path':
                require(isinstance(o.get('points'),list) and len(o['points'])>=2,'Path needs at least two points')
                for xy in o['points']:require(len(xy)==2 and all(isinstance(v,(int,float)) and math.isfinite(v) for v in xy),'Invalid path point')
                numeric(o,['sw'])
            for key in ['fill','stroke','color','grad']:
                if o.get(key):require(bool(re.fullmatch('[0-9A-Fa-f]{6}',o[key])),f'Invalid color {key}')
            if k in ['text','line','path']:require(o.get('color'),'Missing color')
        result['slides'].append({'figure':n,'planned_objects':len(ids)})
    if pptx:
        with zipfile.ZipFile(pptx) as z:
            require(z.testzip() is None,'Corrupt PPTX archive')
            slide_files=[x for x in z.namelist() if re.fullmatch(r'ppt/slides/slide\d+\.xml',x)]
            require(len(slide_files)==len(slides),'Slide count differs from plan')
            media=[x for x in z.namelist() if x.startswith('ppt/media/')]
            require(allow_images or not media,'PPTX contains media; disclose exceptions explicitly')
            for i,slide in enumerate(slides):
                root=E.fromstring(z.read(f'ppt/slides/slide{i+1}.xml'))
                tree=root.find('p:cSld/p:spTree',NS)
                objects=list(tree.findall('p:sp',NS))+list(tree.findall('p:cxnSp',NS))
                pictures=root.findall('.//p:pic',NS)
                require(allow_images or not pictures,'Slide contains picture objects')
                require(len(objects)==len(slide['objects']),'Native object count differs from plan')
                by_id={obj.find('.//p:cNvPr',NS).get('name'):obj for obj in objects}
                for o in slide['objects']:
                    require(o['id'] in by_id,f'Missing native object: {o["id"]}')
                    if o['kind']=='text':
                        node=by_id[o['id']]
                        actual=''.join(t.text or '' for t in node.findall('.//a:t',NS)).replace('\n','').replace('\r','')
                        require(plain_text(o['text'])==actual,f'Text mismatch: {o["id"]}')
                        fonts={f.get('typeface') for f in node.findall('.//a:rPr/a:latin',NS)}
                        require(fonts=={o.get('font','Times New Roman')},f'Font mismatch: {o["id"]}: {fonts}')
                        count=len(re.findall(r'_\{([^{}]+)\}',o['text']))
                        if count:
                            runs=[e for e in node.findall('.//a:rPr',NS) if int(e.get('baseline','0'))<0]
                            require(len(runs)>=count,f'Missing real subscript formatting: {o["id"]}')
                result['slides'][i].update(native_objects=len(objects),pictures=len(pictures))
            result['pptx']={'structural_checks':'passed','media_entries':len(media)}
    if pdf:
        try:import pymupdf
        except ImportError:result['pdf']={'status':'not_checked','reason':'PyMuPDF unavailable; no dependency installed'}
        else:
            with pymupdf.open(pdf) as doc:
                require(len(doc)==len(slides),'PDF page count differs from plan')
                fonts=[];substitutions=[]
                for i,page in enumerate(doc):
                    require(abs(page.rect.width/page.rect.height-canvas[0]/canvas[1])<0.005,'PDF aspect ratio differs')
                    images=len(page.get_images());require(allow_images or images==0,'PDF contains raster images')
                    expected={normalized_font(o.get('font','Times New Roman')) for o in slides[i]['objects'] if o['kind']=='text'}
                    used={s['font'] for b in page.get_text('dict')['blocks'] for line in b.get('lines',[]) for s in line['spans']}
                    fonts.append(sorted(used))
                    unexpected=[f for f in used if not any(normalized_font(f).startswith(x) for x in expected)]
                    if unexpected:substitutions.append({'figure':slides[i]['figure'],'fonts':unexpected})
                    result['slides'][i].update(pdf_images=images,pdf_drawings=len(page.get_drawings()))
                    if split_dir:
                        Path(split_dir).mkdir(parents=True,exist_ok=True)
                        out=Path(split_dir)/f'Fig{slides[i]["figure"]}_editable.pdf'
                        require(out.resolve()!=Path(pdf).resolve(),'Split output would overwrite input PDF')
                        require(not out.exists() or overwrite_splits,f'Split output exists: {out}; choose a new directory or --overwrite-splits')
                        one=pymupdf.open();one.insert_pdf(doc,from_page=i,to_page=i);one.save(out);one.close()
                result['pdf']={'status':'passed' if not substitutions else 'font_review_required','fonts':fonts,'font_substitutions':substitutions}
    result['visual_review']='required; structural checks do not inspect image semantics or overlap'
    return result

def main():
    p=argparse.ArgumentParser(description=__doc__)
    p.add_argument('--plan',required=True);p.add_argument('--pptx');p.add_argument('--pdf');p.add_argument('--report');p.add_argument('--allow-images',action='store_true')
    p.add_argument('--split-dir');p.add_argument('--overwrite-splits',action='store_true')
    args=p.parse_args()
    try:
        require(not args.split_dir or args.pdf,'--split-dir requires --pdf')
        result=validate(args.plan,args.pptx,args.pdf,args.allow_images,args.split_dir,args.overwrite_splits)
        if args.report:Path(args.report).write_text(json.dumps(result,indent=2,ensure_ascii=False),encoding='utf-8')
        print(json.dumps(result,ensure_ascii=False))
        return 2 if result.get('pdf',{}).get('status')=='font_review_required' else 0
    except (ValueError,KeyError,TypeError,OSError,zipfile.BadZipFile) as e:
        print(f'Validation failed: {e}',file=sys.stderr);return 1

if __name__=='__main__':raise SystemExit(main())
