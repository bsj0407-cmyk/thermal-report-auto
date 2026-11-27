import streamlit as st
from io import BytesIO
from zipfile import ZipFile, ZIP_DEFLATED
from pathlib import Path
import re
from thermal_engine import extract_word, inspect_excel, make_excel
st.set_page_config(page_title='열화상 보고서 자동화',page_icon='🌡️',layout='centered')
st.title('🌡️ 열화상 보고서 자동화 V1.0')
st.caption('Word 원본 → 엑셀 자동 입력 → 다운로드')
st.warning('공개 테스트용입니다. 회사 승인 전에는 기밀이 없는 샘플만 사용하세요. 파일은 서버 메모리에서 처리되며 영구 저장 기능은 없습니다.')
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
