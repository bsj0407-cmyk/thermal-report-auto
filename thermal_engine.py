"""SmartView report automation for the established two-column Excel template."""
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from collections import Counter
import re, posixpath, math
from lxml import etree as E
from openpyxl import load_workbook
from PIL import Image
W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
A='http://schemas.openxmlformats.org/drawingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
S='http://schemas.openxmlformats.org/spreadsheetml/2006/main'
P='http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing'
PK='http://schemas.openxmlformats.org/package/2006/relationships'
CT='http://schemas.openxmlformats.org/package/2006/content-types'
CX,CY,OFFSET=2184400,1651000,127000
q=lambda n,t:f'{{{n}}}{t}'
def xml(b): return E.fromstring(b,parser=E.XMLParser(resolve_entities=False,no_network=True))
def dump(e): return E.tostring(e,encoding='UTF-8',xml_declaration=True)
def archive(data):
    z=ZipFile(BytesIO(data))
    if sum(i.file_size for i in z.infolist())>750_000_000 or len(z.infolist())>15000:
        z.close(); raise ValueError('파일 내부 용량 제한 초과')
    return z
def extract_word(data):
    entries={}; dates=set()
    with archive(data) as z:
        if 'word/document.xml' not in z.namelist(): raise ValueError('XML 기반 Word 파일만 지원합니다. 구형 DOC는 DOCX로 저장하세요.')
        root=xml(z.read('word/document.xml'))
        paths={x.get('Id'):posixpath.normpath(posixpath.join('word',x.get('Target'))) if not x.get('Target').startswith('/') else x.get('Target').lstrip('/') for x in xml(z.read('word/_rels/document.xml.rels')) if x.get('TargetMode')!='External'}
        tables=root.xpath('.//w:tbl',namespaces={'w':W})
        for i,t in enumerate(tables):
            text=''.join(t.xpath('.//w:t/text()',namespaces={'w':W}))
            m=re.search(r'IR_(\d{8})_(\d+)\.jpe?g',text,re.I)
            if not m: continue
            n=int(m[2]); dates.add(m[1])
            if n in entries: raise ValueError(f'Word 사진 번호 중복: {n}')
            ids=t.xpath('.//a:blip/@r:embed',namespaces={'a':A,'r':R})
            if len(ids)!=2 or i+2>=len(tables): raise ValueError(f'사진 구조 오류: {n}')
            vals={}; temp=tables[i+2]
            if 'Max. temperature' not in ' '.join(temp.xpath('.//w:t/text()',namespaces={'w':W})): raise ValueError(f'온도 표 구조 오류: {n}')
            for tr in temp.xpath('./w:tr',namespaces={'w':W}):
                cs=[''.join(c.xpath('.//w:t/text()',namespaces={'w':W})).strip() for c in tr.xpath('./w:tc',namespaces={'w':W})]
                if len(cs)>1 and cs[0].lower() in ('sp1','sp2','sp3'):
                    if cs[1] not in ('','-'):
                        v=float(cs[1]);
                        if not math.isfinite(v): raise ValueError('유효하지 않은 온도')
                        vals[cs[0].lower()]=v
            if not {'sp1','sp2'}<=vals.keys(): raise ValueError(f'SP1/SP2 누락: {n}')
            entries[n]=(z.read(paths[ids[0]]),z.read(paths[ids[1]]),vals)
    if not entries or len(dates)!=1: raise ValueError('사진이 없거나 촬영 날짜가 여러 개입니다.')
    return entries,dates.pop()
def inspect_excel(data,entries):
    with archive(data) as z:
        if any(re.match(r'xl/drawings/drawing\d+\.xml$',n) for n in z.namelist()): raise ValueError('기존 사진 또는 도형이 있습니다. 제거 후 다시 등록하세요.')
    wb=load_workbook(BytesIO(data),read_only=False,data_only=False)
    try:
        if len(wb.worksheets)!=1: raise ValueError('단일 시트 양식만 지원합니다.')
        ws=wb.active; slots=[]; errors=[]
        for row in ws.iter_rows():
            for c in row:
                if not isinstance(c.value,str) or re.sub(r'\s','',c.value)!='실물사진': continue
                if c.column not in (1,23): raise ValueError(f'지원하지 않는 양식: {c.coordinate}')
                r=c.row-8; col=c.column+3; v=ws.cell(r,col).value
                try:
                    n=int(v); assert float(v)==n and n>0 and not isinstance(v,bool)
                except (ValueError,TypeError,AssertionError,OverflowError): errors.append(f'{c.coordinate}: 사진 번호 없음/오류'); continue
                for j in range(3):
                    if str(ws.cell(r+19+j,col+1).value).strip().upper()!=f'SP{j+1}': raise ValueError(f'온도 입력 위치 불일치: {c.coordinate}')
                slots.append((r,col,n))
        if errors: raise ValueError(' / '.join(errors))
        if not slots: raise ValueError('실물사진 항목이 없습니다.')
        missing=sorted({n for _,_,n in slots if n not in entries})
        dup=sorted(n for n,k in Counter(n for _,_,n in slots).items() if k>1)
        return slots,missing,dup
    finally: wb.close()
def make_excel(data,entries,allow_duplicates=False):
    slots,missing,dup=inspect_excel(data,entries)
    if missing: raise ValueError(f'매칭 누락: {missing}')
    if dup and not allow_duplicates: raise ValueError(f'중복 번호 확인 필요: {dup}')
    with archive(data) as z: files={n:z.read(n) for n in z.namelist()}
    path='xl/worksheets/sheet1.xml'; sheet=xml(files[path]); cells={c.get('r'):c for c in sheet.iter(q(S,'c'))}
    from openpyxl.utils import get_column_letter
    def setcell(r,c,v):
        addr=f'{get_column_letter(c)}{r}'; target=cells.get(addr)
        if target is None: raise ValueError(f'양식 셀 없음: {addr}')
        for child in list(target):
            if child.tag in (q(S,'v'),q(S,'is'),q(S,'f')): target.remove(child)
        target.attrib.pop('t',None)
        if isinstance(v,str):
            target.set('t','inlineStr'); E.SubElement(E.SubElement(target,q(S,'is')),q(S,'t')).text=v
        elif v is not None: E.SubElement(target,q(S,'v')).text=str(v)
    drawing=E.Element(q(P,'wsDr'),nsmap={'xdr':P,'a':A,'r':R}); rels=E.Element(q(PK,'Relationships'),nsmap={None:PK})
    for r,c,n in slots:
        setcell(r,c,None)
        for j,k in enumerate(('sp1','sp2','sp3')): setcell(r+19+j,c+7,entries[n][2].get(k,'-'))
        for top,idx,kind in ((r,1,'visible'),(r+9,0,'infrared')):
            i=len(rels)+1; name=f'thermal_auto_{i}.jpg'; b=BytesIO()
            with Image.open(BytesIO(entries[n][idx])) as im: im.convert('RGB').save(b,format='JPEG',quality=94)
            files['xl/media/'+name]=b.getvalue(); E.SubElement(rels,q(PK,'Relationship'),Id=f'rId{i}',Type=R+'/image',Target='../media/'+name)
            anchor=E.SubElement(drawing,q(P,'oneCellAnchor')); frm=E.SubElement(anchor,q(P,'from'))
            for k,v in [('col',c-1),('colOff',OFFSET),('row',top-1),('rowOff',OFFSET)]: E.SubElement(frm,q(P,k)).text=str(v)
            E.SubElement(anchor,q(P,'ext'),cx=str(CX),cy=str(CY)); pic=E.SubElement(anchor,q(P,'pic')); nv=E.SubElement(pic,q(P,'nvPicPr'))
            E.SubElement(nv,q(P,'cNvPr'),id=str(i+1),name=f'{n}_{kind}.jpg'); E.SubElement(nv,q(P,'cNvPicPr'))
            bf=E.SubElement(pic,q(P,'blipFill')); E.SubElement(bf,q(A,'blip'),{q(R,'embed'):f'rId{i}'}); E.SubElement(E.SubElement(bf,q(A,'stretch')),q(A,'fillRect'))
            sp=E.SubElement(pic,q(P,'spPr')); xf=E.SubElement(sp,q(A,'xfrm')); E.SubElement(xf,q(A,'off'),x='0',y='0'); E.SubElement(xf,q(A,'ext'),cx=str(CX),cy=str(CY)); E.SubElement(E.SubElement(sp,q(A,'prstGeom'),prst='rect'),q(A,'avLst')); E.SubElement(anchor,q(P,'clientData'))
    files['xl/drawings/drawing1.xml']=dump(drawing); files['xl/drawings/_rels/drawing1.xml.rels']=dump(rels)
    rp='xl/worksheets/_rels/sheet1.xml.rels'; sr=xml(files[rp]) if rp in files else E.Element(q(PK,'Relationships'),nsmap={None:PK}); rid='rIdThermalAuto'
    E.SubElement(sr,q(PK,'Relationship'),Id=rid,Type=R+'/drawing',Target='../drawings/drawing1.xml'); files[rp]=dump(sr)
    node=E.Element(q(S,'drawing'),{q(R,'id'):rid}); after={'legacyDrawing','legacyDrawingHF','picture','oleObjects','controls','webPublishItems','tableParts','extLst'}
    for i,child in enumerate(sheet):
        if E.QName(child).localname in after: sheet.insert(i,node); break
    else: sheet.append(node)
    files[path]=dump(sheet); types=xml(files['[Content_Types].xml'])
    if not any(x.get('Extension')=='jpg' for x in types): E.SubElement(types,q(CT,'Default'),Extension='jpg',ContentType='image/jpeg')
    E.SubElement(types,q(CT,'Override'),PartName='/xl/drawings/drawing1.xml',ContentType='application/vnd.openxmlformats-officedocument.drawing+xml'); files['[Content_Types].xml']=dump(types)
    # Force recalculation of existing temperature difference formulas on opening.
    wp='xl/workbook.xml'; wr=xml(files[wp]); calc=wr.find(q(S,'calcPr'))
    if calc is not None: calc.set('fullCalcOnLoad','1'); files[wp]=dump(wr)
    out=BytesIO()
    with ZipFile(out,'w',ZIP_DEFLATED,compresslevel=3) as z:
        for name,b in files.items(): z.writestr(name,b)
    result=out.getvalue(); check=load_workbook(BytesIO(result)); ws=check.active
    try:
        if len(ws._images)!=2*len(slots): raise ValueError('사진 수 검증 실패')
        for r,c,n in slots:
            for j,k in enumerate(('sp1','sp2','sp3')):
                if ws.cell(r+19+j,c+7).value!=entries[n][2].get(k,'-'): raise ValueError('온도 검증 실패')
    finally: check.close()
    return result,slots
