import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import json
import os
import math

# =========================================================
# 1. 추후 자동화를 위한 Mock Class
# =========================================================
class MockOnecellUploader:
    def __init__(self):
        # 실제 자동화 시에는 웹드라이버(Selenium)나 API 키 초기화 로직이 들어갑니다.
        pass
        
    def upload(self, file_path):
        # 실제 원셀에 파일을 업로드하는 로직을 대체하는 Mock 함수입니다.
        print(f"[Mock Uploader] '{file_path}' 파일을 원셀 시스템에 업로드 중입니다...")
        print("[Mock Uploader] 업로드 완료!")
        return True

# =========================================================
# 2. 메인 GUI 및 데이터 처리 Class
# =========================================================
class OnecellAutomationApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Onecell 업로드 양식 자동 변환기")
        self.root.geometry("450x250")
        
        self.settings_file = "settings.json"
        self.source_file_path = ""
        self.uploader = MockOnecellUploader()
        
        # 기본 설정값 로드
        self.load_settings()
        
        # UI 구성
        self.create_widgets()

    def load_settings(self):
        # 기본값 설정
        self.tag_string = tk.StringVar(value="")
        self.margin_rate = tk.IntVar(value=15)
        
        # 파일 파싱하여 옵션 적용 (사전준비 사항)
        if os.path.exists(self.settings_file):
            try:
                with open(self.settings_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    self.tag_string.set(data.get("tag_string", ""))
                    self.margin_rate.set(data.get("margin_rate", 15))
            except Exception as e:
                print(f"설정 파일 로드 실패: {e}")

    def save_settings(self):
        # UI에 입력된 값을 settings.json에 저장
        data = {
            "tag_string": self.tag_string.get(),
            "margin_rate": self.margin_rate.get()
        }
        with open(self.settings_file, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=4)

    def create_widgets(self):
        # 파일 선택
        btn_select_file = tk.Button(self.root, text="상품 정보 CSV 파일 선택", command=self.select_file, width=25)
        btn_select_file.pack(pady=10)
        
        self.lbl_file = tk.Label(self.root, text="선택된 파일 없음", fg="blue")
        self.lbl_file.pack()

        # 태그 입력
        frame_tag = tk.Frame(self.root)
        frame_tag.pack(pady=5)
        tk.Label(frame_tag, text="태그 문자열 (예: 26신상):").pack(side=tk.LEFT)
        tk.Entry(frame_tag, textvariable=self.tag_string, width=15).pack(side=tk.LEFT, padx=5)

        # 마진율 입력
        frame_margin = tk.Frame(self.root)
        frame_margin.pack(pady=5)
        tk.Label(frame_margin, text="마진율 (%):").pack(side=tk.LEFT)
        tk.Entry(frame_margin, textvariable=self.margin_rate, width=5).pack(side=tk.LEFT, padx=5)

        # 자동 입력(실행) 버튼
        btn_run = tk.Button(self.root, text="자동 입력 및 저장", command=self.process_and_save, width=25, bg="green", fg="white")
        btn_run.pack(pady=20)

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if file_path:
            self.source_file_path = file_path
            self.lbl_file.config(text=os.path.basename(file_path))

    def process_and_save(self):
        if not self.source_file_path:
            messagebox.showwarning("경고", "먼저 상품 정보 CSV 파일을 선택해주세요.")
            return
            
        # 설정값 저장
        self.save_settings()
        
        # 파일 저장 다이얼로그
        save_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[("Excel Files", "*.xlsx")],
            title="결과 파일을 저장할 위치를 지정하세요",
            initialfile="onecell_업로드_결과.xlsx"
        )
        
        if not save_path:
            return
            
        try:
            self.convert_data(self.source_file_path, save_path)
            # 추후 업로드 자동화 과정 호출
            self.uploader.upload(save_path)
            messagebox.showinfo("완료", "원셀 업로드 양식 변환 및 저장이 완료되었습니다!")
        except Exception as e:
            messagebox.showerror("오류", f"데이터 처리 중 오류가 발생했습니다.\n{str(e)}")

    def convert_data(self, src_file, save_file):
        # 1. 인코딩 에러 방지를 위해 utf-8-sig 또는 cp949 적용
        try:
            df_src = pd.read_csv(src_file, encoding='utf-8-sig')
        except:
            df_src = pd.read_csv(src_file, encoding='cp949')

        tag = self.tag_string.get().strip()
        margin = self.margin_rate.get()

        # 2. 원셀 업로드 양식 헤더 정의 (최소 A ~ BC열=55번째 열까지)
        onecell_cols = [
            '기초상품명', '속성명1', '속성값1', '속성명2', '속성값2', '카테고리템플릿 번호', '소비자가(정가)', '판매가', '공급가', '재고수량', 
            '판매자 관리 코드', '상품바코드', '브랜드', '제조사', '모델명', '가격비교사이트 등록여부', '부가세', '미성년자구매', '원산지1', '원산지2', 
            '원산지3', '수입사', '인증항목', '인증번호', '인증기관', '인증상호', '제조일자', '유효일자', '상세설명', '배송템플릿 번호', 
            'A/S 전화번호', 'A/S 안내', '상품정보제공고시 분류', '상품정보제공고시 상세설명참조', '고시항목1', '고시항목2', '고시항목3', '고시항목4', '고시항목5', '고시항목6', 
            '고시항목7', '고시항목8', '고시항목9', '고시항목10', '고시항목11', '고시항목12', '고시항목13', '고시항목14', '고시항목15', '고시항목16', 
            '고시항목17', '고시항목18', '고시항목19', '고시항목20', '대표이미지'
        ]
        
        df_out = pd.DataFrame(columns=onecell_cols)

        # 3. 데이터 맵핑 및 행 생성
        rows_to_add = []
        for _, row in df_src.iterrows():
            new_row = {col: "" for col in onecell_cols}
            
            # A (기초상품명): [태그] 상품명
            orig_name = str(row.get('상품명', ''))
            new_row['기초상품명'] = f"[{tag}] {orig_name}" if tag else orig_name
            
            # H (판매가): 마진율 적용 
            try:
                base_price = float(row.get('판매가', 0))
                # 판매가 * (1 + 마진율/100) 적용 후 정수 처리
                new_row['판매가'] = math.ceil(base_price * (1 + (margin / 100))) 
            except ValueError:
                new_row['판매가'] = 0
            
            # J (재고수량)
            new_row['재고수량'] = row.get('재고수량', 0)
            
            # AC (상세설명)
            desc_key = '상품 상세정보 (html)' if '상품 상세정보 (html)' in df_src.columns else '상품 상세정보'
            new_row['상세설명'] = row.get(desc_key, '')
            
            # AG (상품정보제공고시 분류)
            new_row['상품정보제공고시 분류'] = row.get('상품정보제공고시 품명', '')
            
            # BC (대표이미지)
            new_row['대표이미지'] = row.get('대표 이미지 파일명', '')
            
            # ==================================
            # AZ, BA열 파싱 (속성명1/2, 속성값1/2) 
            # ==================================
            opt_names_raw = str(row.get('옵션명', ''))
            opt_values_raw = str(row.get('옵션값', ''))
            
            # 개행문자(\n) 기준으로 분리
            opt_names = [n.strip() for n in opt_names_raw.split('\n')] if opt_names_raw and opt_names_raw != 'nan' else []
            opt_values = [v.strip() for v in opt_values_raw.split('\n')] if opt_values_raw and opt_values_raw != 'nan' else []
            
            color_val = "ONE COLOR"
            size_val = "ONE SIZE"
            
            for i, name in enumerate(opt_names):
                val = opt_values[i] if i < len(opt_values) else ""
                if "색상" in name:
                    color_val = val if val else "ONE COLOR"
                elif "사이즈" in name or "타입" in name:
                    size_val = val if val else "ONE SIZE"
                    
            # B, C, D, E 속성 할당
            new_row['속성명1'] = "색상"
            new_row['속성값1'] = color_val
            new_row['속성명2'] = "사이즈"
            new_row['속성값2'] = size_val
            
            rows_to_add.append(new_row)
            
        # 결과 DataFrame 생성 및 엑셀 저장
        df_out = pd.DataFrame(rows_to_add)
        df_out.to_excel(save_file, index=False)

# =========================================================
# 실행 진입점
# =========================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = OnecellAutomationApp(root)
    root.mainloop()