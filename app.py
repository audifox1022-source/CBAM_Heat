import pandas as pd
import glob
import os
import re
import sys
import tkinter as tk
from tkinter import filedialog, messagebox

# -----------------------------------------------------------
# 설정: 윈도우 UI 창 숨기기 (깔끔하게 팝업만 띄우기 위함)
# -----------------------------------------------------------
root = tk.Tk()
root.withdraw()

def find_header_row(file_path, file_ext):
    """
    파일에서 실제 데이터 헤더(수주NO. 등)가 있는 행 번호를 찾습니다.
    CSV와 Excel 모두 지원합니다.
    """
    try:
        # 상위 15행만 읽어서 키워드 탐색
        if file_ext == '.csv':
            df_temp = pd.read_csv(file_path, header=None, nrows=15)
        else:
            # 엑셀 파일인 경우
            df_temp = pd.read_excel(file_path, header=None, nrows=15)

        for i, row in df_temp.iterrows():
            row_str = row.astype(str).values
            # '수주' 또는 'NO.' 라는 단어가 포함된 행을 헤더로 간주
            if any("수주" in s for s in row_str):
                return i
    except:
        pass
    return 0 # 못 찾으면 첫 번째 줄을 헤더로

def main():
    # 1. 시작 안내 메시지
    messagebox.showinfo("CBAM 데이터 통합기", "통합할 파일(CSV, Excel)들이 들어있는 [폴더]를 선택해주세요.")

    # 2. 폴더 선택 창 띄우기
    folder_path = filedialog.askdirectory(title="작업지시서 파일이 있는 폴더 선택")
    
    if not folder_path: # 취소 버튼 누른 경우
        return

    # 3. 파일 목록 가져오기 (CSV, XLSX, XLS 모두 포함)
    extensions = ['*.csv', '*.xlsx', '*.xls']
    all_files = []
    
    for ext in extensions:
        # 대소문자 구분 없이 찾기 위해 패턴 매칭 사용 권장되나, 
        # 간편함을 위해 glob 사용 후 확장자 필터링 방식 사용
        found = glob.glob(os.path.join(folder_path, ext))
        all_files.extend(found)

    # 엑셀 임시 파일(~$로 시작하는 파일) 및 결과 파일 제외
    valid_files = []
    for f in all_files:
        base = os.path.basename(f)
        if not base.startswith('~$') and "통합_RAW_DATA_결과" not in base:
            valid_files.append(f)

    if not valid_files:
        messagebox.showwarning("파일 없음", "선택한 폴더에 처리할 파일(.csv, .xlsx)이 없습니다!")
        return

    master_df = pd.DataFrame()
    success_count = 0
    error_log = []

    # 4. 데이터 통합 루프
    print(f"총 {len(valid_files)}개 파일 처리 시작...")
    
    for filename in valid_files:
        try:
            file_basename = os.path.basename(filename)
            file_ext = os.path.splitext(filename)[1].lower() # 확장자 추출 (.csv, .xlsx 등)
            
            # (1) 파일명에서 날짜와 호기 추출
            # 예: "11월 작업... - 11-03(1).csv" -> 날짜: 11-03, 호기: 1
            date_match = re.search(r"(\d{1,2}-\d{1,2})", file_basename)
            furnace_match = re.search(r"\((.+?)\)", file_basename) # 괄호 안 추출 (1, 단조 등)

            work_date = date_match.group(1) if date_match else "날짜미상"
            furnace_no = furnace_match.group(1) if furnace_match else "호기미상"

            # (2) 헤더 위치 자동 탐색
            header_idx = find_header_row(filename, file_ext)

            # (3) 데이터 읽기 (확장자에 따라 분기)
            if file_ext == '.csv':
                df = pd.read_csv(filename, header=header_idx)
            else:
                df = pd.read_excel(filename, header=header_idx)

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
            
        except Exception as e:
            error_log.append(f"{file_basename}: {str(e)}")
            print(f"오류 발생: {file_basename} - {e}")

    # 5. 결과 저장
    if not master_df.empty:
        output_path = os.path.join(folder_path, "통합_RAW_DATA_결과.xlsx")
        
        # 엑셀로 저장
        try:
            master_df.to_excel(output_path, index=False)
            
            # 완료 메시지
            msg = f"작업 완료!\n\n- 처리 파일: {success_count}/{len(valid_files)}개\n- 저장 위치:\n{output_path}"
            if error_log:
                msg += f"\n\n[주의] {len(error_log)}개 파일 처리 실패 (로그 확인)"
            
            messagebox.showinfo("성공", msg)
        except Exception as e:
            messagebox.showerror("저장 실패", f"파일 저장 중 오류가 발생했습니다.\n파일이 열려있다면 닫아주세요.\n\n{e}")
    else:
        messagebox.showwarning("실패", "통합할 데이터가 없거나 유효한 파일이 없습니다.")

if __name__ == "__main__":
    main()
