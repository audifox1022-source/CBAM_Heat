import streamlit as st
import pandas as pd
import re
import io

# 앱 제목 설정
st.set_page_config(page_title="CBAM 데이터 통합기", page_icon="🏭")
st.title("🏭 열처리 작업지시서 통합 도구")
st.markdown("여러 개의 **CSV 파일**을 업로드하면 하나로 합쳐줍니다. (월말 정산용)")

# 파일 업로더
uploaded_files = st.file_uploader("CSV 파일들을 여기에 드래그하세요", accept_multiple_files=True, type=['csv'])

if uploaded_files:
    if st.button("데이터 통합 시작"):
        with st.spinner('데이터를 분석하고 합치는 중...'):
            master_df = pd.DataFrame()
            
            # 진행률 표시줄
            progress_bar = st.progress(0)
            
            for i, uploaded_file in enumerate(uploaded_files):
                try:
                    # 파일명 읽기
                    filename = uploaded_file.name
                    date_match = re.search(r"(\d+-\d+)", filename)
                    furnace_match = re.search(r"\((.+?)\)", filename)
                    
                    work_date = date_match.group(1) if date_match else "날짜미상"
                    furnace_no = furnace_match.group(1) if furnace_match else "호기미상"

                    # 헤더 찾기
                    temp_df = pd.read_csv(uploaded_file, header=None, nrows=10)
                    uploaded_file.seek(0) # 파일 포인터 초기화
                    
                    header_row = 0
                    for idx, row in temp_df.iterrows():
                        if row.astype(str).str.contains('수주NO').any():
                            header_row = idx
                            break
                    
                    df = pd.read_csv(uploaded_file, header=header_row)
                    
                    if '수주NO.' in df.columns:
                        df = df[df['수주NO.'].notna()]
                        df.insert(0, '작업지시일', work_date)
                        df.insert(1, '지시서번호', furnace_no)
                        master_df = pd.concat([master_df, df], ignore_index=True)
                
                except Exception as e:
                    st.error(f"{uploaded_file.name} 처리 중 오류: {e}")
                
                # 진행률 업데이트
                progress_bar.progress((i + 1) / len(uploaded_files))

            st.success(f"총 {len(uploaded_files)}개 파일 통합 완료!")
            
            # 엑셀 다운로드 버튼
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                master_df.to_excel(writer, index=False)
                
            st.download_button(
                label="📥 통합 엑셀 파일 다운로드",
                data=buffer,
                file_name="통합_RAW_DATA.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
            
            # 미리보기
            st.write("▼ 데이터 미리보기")
            st.dataframe(master_df.head())