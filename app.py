import streamlit as st
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
from collections import Counter, defaultdict
import re
from thermal_engine import extract_word, inspect_excel, make_excel, verify_photos
st.set_page_config(page_title='열화상 보고서 자동화',page_icon='🌡️',layout='centered')
st.title('🌡️ 열화상 보고서 자동화 V1.2')
st.caption('Word 원본 → 엑셀 자동 입력 → 검증 → 다운로드')
st.info('본 시스템은 열화상 보고서 작성 시간을 단축하기 위한 업무 자동화 프로그램입니다. 파일은 외부 클라우드 서버에서 처리되며 별도의 영구 저장 기능은 없습니다. 회사의 정보보안 규정을 준수하고, 승인 전에는 기밀이 없는 샘플만 사용해 주세요.')
word=st.file_uploader('① Word 원본',type=['doc','docx'])
books=st.file_uploader('② 엑셀 보고서',type=['xlsx'],accept_multiple_files=True)
allow=st.checkbox('중복 사진 번호가 있더라도 지정된 번호 그대로 적용')
st.caption('사진 229.3 × 173.3px · Max. temperature · SP3 누락 시 - · 원본 양식 보존')
if st.button('③ 자동화 작업 시작',type='primary',disabled=not(word and books)):
    st.session_state.pop('audit',None)
    try:
        entries,date=extract_word(word.getvalue()); results=[]; progress=st.progress(0)
        for i,f in enumerate(books):
            report=dict(name=f.name,output=None,slots=[],rows=[],extra=0,error='')
            try:
                m=re.search(r'(\d{4})(?:\(\d+\))?\.xlsx$',f.name)
                if m and m[1]!=date[-4:]: raise ValueError('엑셀 파일명과 Word 촬영일 불일치')
                data=f.getvalue(); slots,missing,dup=inspect_excel(data,entries); report['slots']=slots
                output,slots=make_excel(data,entries,allow)
                rows,extra=verify_photos(output,entries,slots); report.update(rows=rows,extra=extra)
                if extra or not all(r['ok'] for r in rows): raise ValueError('실제 사진 검증 실패: 다운로드 차단')
                report['output']=output
            except Exception as e: report['error']=str(e)
            results.append(report); progress.progress((i+1)/len(books))
        st.session_state['audit']=(sorted(entries),date,results)
    except Exception as e: st.error(f'Word 처리 실패: {e}')

def details(report, universe, counts, requested, locations, image_dups):
    applied={r['number'] for r in report['rows'] if r['ok']}
    wanted={n for _,_,n in report['slots']}; failed=wanted-applied
    failed.update(r['number'] for r in report['rows'] if not r['ok'])
    st.caption('🟢 적용 · ⚪ 미사용 · 🟡 중복 · 🔴 요청했으나 미삽입 (미사용은 정상)')
    numbers=sorted(set(universe)|wanted)
    cells=[]
    for n in numbers:
        color='red' if n in failed else 'orange' if n in applied and (counts[n]>1 or n in image_dups) else 'green' if n in applied else 'gray'
        cells.append(f':{color}-background[{n:04d}]')
    for i in range(0,len(cells),10): st.markdown(' '.join(cells[i:i+10]))
    st.write('적용 번호: '+(', '.join(f'{n:04d}' for n in sorted(applied)) or '-'))
    st.write('미사용 번호: '+(', '.join(f'{n:04d}' for n in sorted(set(universe)-wanted)) or '-'))
    if failed: st.error('미삽입 요청 번호: '+', '.join(map(str,sorted(failed))))
    for n in sorted(wanted):
        if requested[n]>1: st.warning(f'{n:04d} 중복 요청: '+', '.join(locations[n]))
        if n in image_dups: st.warning(f'{n:04d} 동일 이미지 발견 (다른 촬영 번호 포함)')
    if report['error']: st.error(report['error'])
    if report['rows']: st.table([{'번호':r['number'],'셀':r['cell'],'사진 수':r['count'],'검증':r['issue'] or '정상'} for r in report['rows']])

if 'audit' in st.session_state:
    universe,date,reports=st.session_state['audit']; counts=Counter(); requested=Counter(); locations=defaultdict(list); hashes=defaultdict(set)
    for i,rep in enumerate(reports):
        for r,c,n in rep['slots']:
            requested[n]+=1; locations[n].append(f"{rep['name']} ({'D' if c==4 else 'Z'}{r})")
        for row in rep['rows']:
            if row['ok'] and rep['output'] is not None:
                counts[row['number']]+=1
                for kind,h in row['hashes']: hashes[kind,h].add(row['number'])
    image_dups=set().union(*(ns for ns in hashes.values() if len(ns)>1))
    applied=set(counts); unused=set(universe)-set(requested); failed=set(requested)-applied
    failed.update(n for rep in reports if rep['output'] is None for _,_,n in rep['slots'])
    duplicates={n for n,k in counts.items() if k>1}; request_dup={n for n,k in requested.items() if k>1}
    st.subheader('📊 전체 사진 적용 현황'); st.caption(f'촬영일 {date} · 이번 작업에 등록된 엑셀 기준')
    a,b=st.columns(2); a.metric('Word 촬영 번호',len(universe)); b.metric('고유 적용 번호',len(applied)); a.metric('미사용 번호 (정상)',len(unused)); b.metric('중복 적용 번호',len(duplicates)); a.metric('미삽입 요청 번호',len(failed)); b.metric('동일 이미지 의심 번호',len(image_dups))
    st.progress(len(applied)/len(universe)); st.caption('원본 대비 사용률 — 100%일 필요 없음')
    if request_dup: st.warning('중복 요청 번호: '+', '.join(map(str,sorted(request_dup))))
    if duplicates or image_dups or failed or any(r['error'] for r in reports): st.warning('검토가 필요한 작업이 있습니다.')
    else: st.success('요청 사진 삽입 검증 통과')
    if unused: st.caption('전체 미사용: '+', '.join(f'{n:04d}' for n in sorted(unused)))
    for i,rep in enumerate(reports):
        st.divider(); st.write(rep['name']); rows=rep['rows']; output=rep['output']
        if rep['error']: st.error(rep['error'])
        a,b,c=st.columns(3); a.metric('검사 위치',len(rep['slots'])); b.metric('삽입 사진',sum(r['count'] for r in rows) if output else 0); c.metric('미완료 위치',len(rep['slots'])-sum(r['ok'] for r in rows) if output else len(rep['slots']))
        table={}
        for r,c,n in rep['slots']: table.setdefault(r,{'왼쪽':'','오른쪽':''})['왼쪽' if c==4 else '오른쪽']=n
        if table: st.table([{'순서':j+1,**v} for j,v in enumerate(table.values())])
        if st.button('자세히',key=f'detail{i}'):
            if hasattr(st,'dialog'):
                @st.dialog('사진 적용 상세 · '+rep['name'],width='large')
                def show(): details(rep,universe,counts,requested,locations,image_dups)
                show()
            else: details(rep,universe,counts,requested,locations,image_dups)
        if output: st.download_button('완성 엑셀 다운로드',output,file_name=Path(rep['name']).stem+'_자동입력완료.xlsx',key=f'dl{i}',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    if any(r['output'] for r in reports):
        buf=BytesIO()
        with ZipFile(buf,'w',ZIP_DEFLATED) as z:
            for i,rep in enumerate(reports):
                if rep['output']: z.writestr(f"{i+1:02d}_{Path(rep['name']).stem}_자동입력완료.xlsx",rep['output'])
        st.download_button('완료 파일 ZIP 다운로드',buf.getvalue(),'thermal_reports.zip',mime='application/zip')
if st.button('결과 지우기'):
    st.session_state.pop('audit',None); st.rerun()
st.divider(); st.markdown('**📋 V1.2 주요 기능**')
for feature in ['실물사진 바로 위 숫자로 사진 자동 매칭','실물·열화상 사진 자동 삽입 (229.3 × 173.3px)','Max. temperature SP1·SP2·SP3 입력 (SP3 누락 시 -)','기존 엑셀 양식 및 인쇄 설정 보존','실제 삽입 사진 위치 및 이미지 검증','전체 적용·미사용·중복·미삽입 집계','엑셀별 자세히 팝업 및 좌우 번호 비교표','다중 엑셀 작업 및 ZIP 다운로드']: st.caption('✓ '+feature)
st.caption('Thermal Report Automation · Version 1.2')
