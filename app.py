import streamlit as st
import pandas as pd
import re
import io
import os

# -----------------------------------------------------------
# Streamlit 페이지 설정
# -----------------------------------------------------------
st.set_page_config(page_title="CBAM 데이터 통합기", page_icon="🏭", layout="wide")

st.title("🏭 열처리 작업지시서 통합 도구 (Web)")
st.markdown("""
**CSV 및 Excel 파일**을 업로드하면 하나의 파일로 합쳐줍니다.
1. 아래 영역에 파일을 드래그하거나 선택하세요.
2. [통합 시작] 버튼을 누르세요.
3. 결과 파일을 다운로드하세요.
""")

def read_csv_with_encoding(file_obj, **kwargs):
    """
    CSV 파일을 읽을 때 한글 인코딩(utf-8, cp949 등)을 자동으로 찾아서 읽습니다.
    """
    encodings = ['utf-8', 'cp949', 'euc-kr']
    
    for enc in encodings:
        try:
            file_obj.seek(0)
            return pd.read_csv(file_obj, encoding=enc, **kwargs)
        except UnicodeDecodeError:
            continue
        except Exception:
            continue
            
    # 모든 인코딩 실패 시 다시 utf-8로 시도하여 에러 발생시킴
    file_obj.seek(0)
    return pd.read_csv(file_obj, encoding='utf-8', **kwargs)

def find_header_row(file_obj, file_ext):
    """
    업로드된 파일 객체에서 실제 데이터 헤더(수주NO. 등)가 있는 행 번호를 찾습니다.
    """
    try:
        file_obj.seek(0) # 파일 포인터 초기화
        # 상위 15행만 읽어서 키워드 탐색
        if file_ext == '.csv':
            # 인코딩 자동 감지 함수 사용
            df_temp = read_csv_with_encoding(file_obj, header=None, nrows=15)
        else:
            df_temp = pd.read_excel(file_obj, header=None, nrows=15)

        for i, row in df_temp.iterrows():
            row_str = row.astype(str).values
            # '수주' 또는 'NO.' 라는 단어가 포함된 행을 헤더로 간주
            if any("수주" in s for s in row_str):
                file_obj.seek(0) # 파일 포인터 다시 초기화 (실제 읽기를 위해)
                return i
    except Exception as e:
        # print(f"Header search failed: {e}")
        pass
    
    file_obj.seek(0)
    return 0 # 못 찾으면 첫 번째 줄을 헤더로

# -----------------------------------------------------------
# 파일 업로더 (CSV, Excel 모두 지원)
# -----------------------------------------------------------
uploaded_files = st.file_uploader(
    "여기에 파일을 드래그하세요 (CSV, XLSX, XLS)", 
    accept_multiple_files=True, 
    type=['csv', 'xlsx', 'xls']
)

if uploaded_files:
    if st.button("데이터 통합 시작"):
        master_df = pd.DataFrame()
        success_count = 0
        error_log = []
        
        # 진행 상황바
        progress_bar = st.progress(0)
        status_text = st.empty()

        for idx, uploaded_file in enumerate(uploaded_files):
            try:
                filename = uploaded_file.name
                file_ext = os.path.splitext(filename)[1].lower()
                status_text.text(f"처리 중: {filename}")

                # (1) 파일명에서 날짜와 호기 추출
                # 예: "11월 작업... - 11-03(1).csv"
                date_match = re.search(r"(\d{1,2}-\d{1,2})", filename)
                furnace_match = re.search(r"\((.+?)\)", filename)

                work_date = date_match.group(1) if date_match else "날짜미상"
                furnace_no = furnace_match.group(1) if furnace_match else "호기미상"

                # (2) 헤더 위치 자동 탐색
                header_idx = find_header_row(uploaded_file, file_ext)

                # (3) 데이터 읽기
                if file_ext == '.csv':
                    df = read_csv_with_encoding(uploaded_file, header=header_idx)
                else:
                    df = pd.read_excel(uploaded_file, header=header_idx)

                # (4) 유효한 데이터만 남기기 (수주NO가 있는 행만)
                # 컬럼명에 '수주'가 포함된 컬럼 찾기
                order_col = [c for c in df.columns if "수주" in str(c)]
                
                if order_col:
                    target_col = order_col[0]
                    df = df[df[target_col].notna()] # 수주번호 없는 행 삭제
                    
                    # (5) 메타데이터 열 추가 (맨 앞에 삽입)
                    df.insert(0, '지시서번호(호기)', furnace_no)
                    df.insert(0, '작업지시일', work_date)
                    
                    # (6) 통합
                    master_df = pd.concat([master_df, df], ignore_index=True)
                    success_count += 1
                else:
                    error_log.append(f"⚠️ {filename}: '수주NO' 컬럼을 찾을 수 없음 (헤더 인식 실패)")
                
            except Exception as e:
                error_log.append(f"❌ {filename}: {str(e)}")
            
            # 진행률 업데이트
            progress_bar.progress((idx + 1) / len(uploaded_files))

        status_text.text("처리 완료!")

        # -----------------------------------------------------------
        # 결과 출력 및 다운로드
        # -----------------------------------------------------------
        if not master_df.empty:
            st.success(f"✅ 총 {success_count}개 파일 통합 완료!")
            
            if error_log:
                st.warning(f"⚠️ {len(error_log)}개 파일 처리 실패")
                with st.expander("실패 로그 확인"):
                    for err in error_log:
                        st.write(err)

            # 데이터 미리보기
            st.subheader("📊 통합 데이터 미리보기")
            st.dataframe(master_df.head())

            # 엑셀 다운로드 (메모리 버퍼 사용)
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                master_df.to_excel(writer, index=False)
            
            st.download_button(
                label="📥 통합 엑셀 파일 다운로드",
                data=buffer,
                file_name="통합_RAW_DATA_결과.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        else:
            st.error("통합할 데이터가 없습니다. 아래 로그를 확인해주세요.")
            if error_log:
                with st.expander("에러 상세 내용"):
                    for err in error_log:
                        st.write(err)
