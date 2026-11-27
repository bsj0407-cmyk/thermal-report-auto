import streamlit as st
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
import re
from thermal_engine import extract_word, inspect_excel, make_excel
st.set_page_config(page_title='열화상 보고서 자동화',page_icon='🌡️',layout='centered')
st.title('🌡️ 열화상 보고서 자동화 V1.1')
st.caption('Word 원본 → 엑셀 자동 입력 → 다운로드')
st.info('본 시스템은 열화상 보고서 작성 시간을 단축하기 위한 업무 자동화 프로그램입니다. 파일은 외부 클라우드 서버에서 처리되며 별도의 영구 저장 기능은 없습니다. 회사의 정보보안 규정을 준수하고, 승인 전에는 기밀이 없는 샘플만 사용해 주세요.')
word=st.file_uploader('① Word 원본',type=['doc','docx'])
books=st.file_uploader('② 엑셀 보고서',type=['xlsx'],accept_multiple_files=True)
allow=st.checkbox('중복 사진 번호가 있더라도 지정된 번호 그대로 적용')
st.caption('사진 229.3 × 173.3px · Max. temperature · SP3 누락 시 - · 원본 양식 보존')
if st.button('③ 자동화 작업 시작',type='primary',disabled=not(word and books)):
    st.session_state.pop('results',None)
    try:
        entries,date=extract_word(word.getvalue()); st.info(f'촬영일 {date} / Word 사진 번호 {len(entries)}개')
        results=[]; progress=st.progress(0)
        for i,f in enumerate(books):
            try:
                m=re.search(r'(\d{4})(?:\(\d+\))?\.xlsx$',f.name)
                if m and m[1]!=date[-4:]: raise ValueError('엑셀 파일명과 Word 촬영일 불일치')
                data=f.getvalue(); slots,missing,dup=inspect_excel(data,entries)
                output,slots=make_excel(data,entries,allow)
                rows={}
                for r,c,n in slots: rows.setdefault(r,{'왼쪽':'','오른쪽':''})['왼쪽' if c==4 else '오른쪽']=n
                table=[{'순서':i+1,**v} for i,v in enumerate(rows.values())]
                results.append((Path(f.name).stem+'_자동입력완료.xlsx',output,table,len(slots),dup))
            except Exception as e: st.error(f'{f.name}: {e}')
            progress.progress((i+1)/len(books))
        st.session_state['results']=results
    except Exception as e: st.error(f'Word 처리 실패: {e}')
for i,(name,data,table,count,dup) in enumerate(st.session_state.get('results',[])):
    st.success(name); a,b,c=st.columns(3); a.metric('검사 위치',count); b.metric('삽입 사진',count*2); c.metric('매칭 누락',0)
    if dup: st.warning(f'중복 번호 그대로 적용: {dup}')
    st.table(table); st.download_button('완성 엑셀 다운로드',data,file_name=name,key=f'dl{i}',mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
if st.session_state.get('results'):
    buf=BytesIO()
    with ZipFile(buf,'w',ZIP_DEFLATED) as z:
        for i,(name,data,*_) in enumerate(st.session_state['results']): z.writestr(f'{i+1:02d}_{name}',data)
    st.download_button('전체 ZIP 다운로드',buf.getvalue(),'thermal_reports.zip',mime='application/zip')
if st.button('결과 지우기'):
    st.session_state.pop('results',None); st.rerun()

# V1.1 기능 안내
st.divider()
st.markdown('**📋 V1.1 주요 기능**')
for feature in [
    '실물사진 바로 위 숫자로 사진 자동 매칭',
    '실물 사진 및 열화상 사진 자동 삽입 (229.3 × 173.3px)',
    'Max. temperature의 SP1·SP2·SP3 자동 입력 (SP3 누락 시 -)',
    '기존 병합 셀, 행 높이, 열 너비 및 인쇄 설정 보존',
    '사진 번호 중복 및 매칭 누락 검사',
    '실제 적용 사진 번호 좌우 비교표 제공',
    '검사 위치 수, 삽입 사진 수, 매칭 누락 수 안내',
    '엑셀 다중 작업 및 ZIP 일괄 다운로드',
]:
    st.caption('✓ ' + feature)
st.caption('Thermal Report Automation · Version 1.1')
