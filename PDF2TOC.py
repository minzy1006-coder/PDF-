import fitz  # PyMuPDF 라이브러리
import re
import os
import tkinter as tk
from tkinter import filedialog, messagebox

class PDFBookmarkerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 책갈피 자동 생성기 (스마트 스캔 버전)")
        self.root.geometry("550x540") # 3개 옵션에 맞게 창 크기 조절
        self.root.resizable(False, False)
        
        # 파일 선택 영역
        frame_file = tk.Frame(root)
        frame_file.pack(pady=15, padx=15, fill="x")
        
        tk.Label(frame_file, text="PDF 파일:").pack(side="left")
        self.file_path_var = tk.StringVar()
        tk.Entry(frame_file, textvariable=self.file_path_var, width=45, state="readonly").pack(side="left", padx=5)
        tk.Button(frame_file, text="찾아보기", command=self.browse_file).pack(side="left")
        
        # 페이지 설정 영역
        frame_page = tk.Frame(root)
        frame_page.pack(pady=5, padx=15, fill="x")
        
        tk.Label(frame_page, text="목차 시작:").pack(side="left")
        self.start_page_var = tk.StringVar()
        tk.Entry(frame_page, textvariable=self.start_page_var, width=5).pack(side="left", padx=5)
        
        tk.Label(frame_page, text="목차 끝:").pack(side="left", padx=(10, 0))
        self.end_page_var = tk.StringVar()
        tk.Entry(frame_page, textvariable=self.end_page_var, width=5).pack(side="left", padx=5)
        
        # 페이지 계산 옵션 설정 영역
        frame_option = tk.Frame(root)
        frame_option.pack(pady=10, padx=15, fill="x")
        
        tk.Label(frame_option, text="페이지 계산 방식:").pack(anchor="w")
        self.page_mode_var = tk.IntVar(value=2) # 기본값 2번으로 설정
        
        tk.Radiobutton(frame_option, text="1. 뷰어 기준 (적힌 숫자 그대로 이동)", variable=self.page_mode_var, value=1).pack(anchor="w", padx=15, pady=2)
        tk.Radiobutton(frame_option, text="2. 본문 기준 (본문 번호 자동 스캔 / 실패 시 직후=1p)", variable=self.page_mode_var, value=2, fg="#1976D2").pack(anchor="w", padx=15, pady=2)
        tk.Radiobutton(frame_option, text="3. 본문 기준 (본문 번호 자동 스캔 / 실패 시 직후=2p)", variable=self.page_mode_var, value=3, fg="#1976D2").pack(anchor="w", padx=15, pady=2)
        
        # 실행 버튼
        tk.Button(root, text="책갈피 자동 생성 시작", command=self.run_process, 
                  bg="#4CAF50", fg="white", font=("맑은 고딕", 10, "bold"), height=2).pack(pady=10, fill="x", padx=15)
        
        # 작업 로그 출력 영역
        tk.Label(root, text="작업 진행 상황:").pack(anchor="w", padx=15)
        
        # 스크롤바와 텍스트 박스
        frame_log = tk.Frame(root)
        frame_log.pack(padx=15, pady=(0, 15), fill="both", expand=True)
        
        scrollbar = tk.Scrollbar(frame_log)
        scrollbar.pack(side="right", fill="y")
        
        self.log_text = tk.Text(frame_log, height=12, state="disabled", bg="#f4f4f4", yscrollcommand=scrollbar.set)
        self.log_text.pack(side="left", fill="both", expand=True)
        scrollbar.config(command=self.log_text.yview)

    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="목차를 만들 PDF 파일 선택",
            filetypes=[("PDF Files", "*.pdf"), ("All Files", "*.*")]
        )
        if file_path:
            self.file_path_var.set(file_path)
            self.log_message(f"파일이 선택되었습니다: {os.path.basename(file_path)}")

    def log_message(self, message):
        """일반 단일 메시지 로그 출력용"""
        self.log_text.config(state="normal")
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end") 
        self.log_text.config(state="disabled")
        self.root.update()

    def extract_candidates(self, input_pdf_path, toc_start_page, toc_end_page, page_mode):
        """PDF에서 텍스트를 읽고 정규표현식으로 후보군(목차 리스트)을 뽑아내는 내부 함수"""
        doc = fitz.open(input_pdf_path)
        page_count = doc.page_count
        
        # --- [옵션 2, 3: 스마트 스캔 처리] 본문의 페이지 번호 수집 ---
        page_mapping = {}
        if page_mode in [2, 3]:
            for i in range(toc_end_page, page_count):
                page = doc[i]
                rect = page.rect
                
                # 상단 15%, 하단 15%의 여백 영역만 스캔 (본문에 등장하는 일반 숫자와 헷갈리지 않게 방지)
                top_clip = fitz.Rect(rect.x0, rect.y0, rect.x1, rect.y0 + rect.height * 0.15)
                bottom_clip = fitz.Rect(rect.x0, rect.y1 - rect.height * 0.15, rect.x1, rect.y1)
                
                # 상/하단 영역에서 단어 추출
                words = page.get_text("words", clip=top_clip) + page.get_text("words", clip=bottom_clip)
                for w in words:
                    text = w[4].strip()
                    # '-', '_', '[', '(' 등으로 감싸진 페이지 번호, 또는 'p', '쪽'이 붙은 숫자 강력 추출
                    m = re.match(r'^[-_\[\<\(]?\s*(\d+)\s*[-_\]\>\)]?\s*(?:p|P|쪽|페이지)?[^a-zA-Z0-9가-힣]*$', text)
                    if m:
                        num = int(m.group(1))
                        # 보통 책 페이지 번호는 3000을 넘지 않으며, 맨 처음 발견된 위치를 우선으로 함
                        if num < 3000 and num not in page_mapping:
                            page_mapping[num] = i + 1  # 뷰어 기준(1번부터 시작) 페이지 번호 기록
        
        # --- 목차 텍스트 분석 시작 ---
        extracted_lines = []
        
        # 사람의 눈처럼 시각적으로 텍스트를 읽어오기 (떨어져 있는 숫자 매칭용)
        for i in range(toc_start_page - 1, toc_end_page):
            page = doc[i]
            words = page.get_text("words")  # 단어 단위로 위치(좌표) 정보와 함께 추출
            
            # 1. 단어들을 세로(Y축) 위치 기준으로 먼저 정렬 (위에서 아래로)
            words.sort(key=lambda w: (w[1] + w[3]) / 2)
            
            current_line_words = []
            current_y = -1
            
            for w in words:
                word_y = (w[1] + w[3]) / 2
                # Y축 중앙값이 6픽셀 이내로 비슷하면 '같은 줄'에 있는 글자로 취급
                if current_y == -1 or abs(word_y - current_y) < 6:
                    current_line_words.append(w)
                    if current_y == -1: current_y = word_y
                else:
                    # 2. 같은 줄로 판별된 단어들을 가로(X축) 위치 기준으로 정렬 (왼쪽에서 오른쪽으로)
                    current_line_words.sort(key=lambda x: x[0])
                    
                    # 3. 단어 이어 붙이기 (제목과 숫자 사이가 멀면 띄어쓰기를 넉넉하게 삽입)
                    text_line = ""
                    last_x1 = -1
                    for cw in current_line_words:
                        if last_x1 != -1:
                            if (cw[0] - last_x1) > 15: # 거리가 멀면 공간 확보
                                text_line += "    "
                            else:
                                text_line += " "
                        text_line += cw[4]
                        last_x1 = cw[2]
                        
                    extracted_lines.append(text_line)
                    current_line_words = [w]
                    current_y = word_y
                    
            if current_line_words:
                current_line_words.sort(key=lambda x: x[0])
                text_line = ""
                last_x1 = -1
                for cw in current_line_words:
                    if last_x1 != -1:
                        if (cw[0] - last_x1) > 15:
                            text_line += "    "
                        else:
                            text_line += " "
                    text_line += cw[4]
                    last_x1 = cw[2]
                extracted_lines.append(text_line)
                
        doc.close()

        candidates_list = []
        pending_titles = []  # 텍스트를 하나로 합치지 않고, 리스트에 따로따로 보관
        ignore_words = ["목차", "차례", "목 차", "차 례", "table of contents", "contents", "index"]

        for line in extracted_lines:
            line = line.strip()
            if not line or line.lower() in ignore_words:
                continue
                
            # 1. 줄 전체가 페이지 번호 하나로만 이루어진 경우 (예: " 15 ", "- 15 -", "15p")
            m_digit = re.match(r'^[-_\[\<\(·]?\s*(\d+)\s*[-_\]\>\)·]?\s*(?:p|P|쪽|페이지)?[^a-zA-Z0-9가-힣]*$', line)
            if m_digit:
                printed_page = int(m_digit.group(1))
                
                if page_mode == 1:
                    calculated_page = printed_page
                elif page_mode in [2, 3]:
                    # 1순위: 스캔된 본문 페이지에서 번호 찾기 (스마트 스캔)
                    if printed_page in page_mapping:
                        calculated_page = page_mapping[printed_page]
                    else:
                        # 2순위: 못 찾은 경우 각 옵션의 기본 수식으로 계산 (백업)
                        if page_mode == 2:
                            calculated_page = printed_page + toc_end_page
                        else:
                            calculated_page = printed_page + toc_end_page - 1
                    
                actual_pdf_page = max(1, min(calculated_page, page_count))
                    
                for pt in pending_titles:
                    candidates_list.append([1, pt, actual_pdf_page])
                pending_titles = []
                continue
                
            # 2. 제목과 페이지 번호가 한 줄에 같이 있는 경우
            # [핵심 수정] 숫자 뒤에 특수기호나 빈칸 등 어떤 찌꺼기가 와도 무조건 숫자만 낚아챔
            match = re.search(r'(?:[\.·…\-\_]{2,}|\t|\s+)([-_\[\<\(]?\s*(\d+)\s*[-_\]\>\)]?\s*(?:p|P|쪽|페이지)?)[^a-zA-Z0-9가-힣]*$', line)
            if match:
                # 완벽하게 분리해낸 순수 페이지 숫자
                printed_page = int(match.group(2))
                # 숫자가 시작되기 직전까지의 문자열만 '제목'으로 가져옴 -> 제목에서 숫자가 완전히 빠짐!
                title_raw = line[:match.start()]
                
                # 마침표(.), 중간점(·), 줄임표(…), 빼기(-) 등 목차 선을 구성하는 모든 형태의 특수문자를 제거
                title_part = re.sub(r'([\.·…\-]\s*){2,}', '', title_raw).strip(' .-_|·…\t')
                
                if page_mode == 1:
                    calculated_page = printed_page
                elif page_mode in [2, 3]:
                    # 1순위: 스캔된 본문 페이지에서 번호 찾기 (스마트 스캔)
                    if printed_page in page_mapping:
                        calculated_page = page_mapping[printed_page]
                    else:
                        # 2순위: 못 찾은 경우 각 옵션의 기본 수식으로 계산 (백업)
                        if page_mode == 2:
                            calculated_page = printed_page + toc_end_page
                        else:
                            calculated_page = printed_page + toc_end_page - 1
                    
                actual_pdf_page = max(1, min(calculated_page, page_count))
                
                # 상위 목차 처리 (번호 없이 대기 중이던 '각각의' 독립된 제목들에게 일괄 페이지 부여)
                for pt in pending_titles:
                    candidates_list.append([1, pt, actual_pdf_page])
                pending_titles = []
                
                # 현재 줄 제목 처리
                if title_part:
                    candidates_list.append([1, title_part, actual_pdf_page])
            else:
                # 3. 번호가 없는 줄 (다음 줄에서 번호가 나오길 기다리며 제목만 보관)
                title_part = re.sub(r'([\.·…\-]\s*){2,}', '', line).strip(' .-_|·…\t')
                if title_part:
                    pending_titles.append(title_part)

        # 끝까지 페이지 번호를 찾지 못한 목차들은 기본 페이지로 설정
        for pt in pending_titles:
            candidates_list.append([1, pt, toc_start_page])

        return candidates_list

    def run_process(self):
        input_file = self.file_path_var.get()
        start_page_str = self.start_page_var.get()
        end_page_str = self.end_page_var.get()
        page_mode = self.page_mode_var.get() # 선택된 계산 방식 (1, 2, 3) 가져오기

        if not input_file:
            messagebox.showwarning("입력 오류", "PDF 파일을 선택해 주세요.")
            return
        
        try:
            start_page = int(start_page_str)
            end_page = int(end_page_str)
        except ValueError:
            messagebox.showwarning("입력 오류", "페이지 번호는 반드시 숫자로 입력해 주세요.")
            return

        if start_page < 1 or end_page < 1 or start_page > end_page:
            messagebox.showwarning("입력 오류", "올바른 페이지 범위를 입력해 주세요.")
            return

        # 로그창 초기화
        self.log_text.config(state="normal")
        self.log_text.delete(1.0, "end")
        self.log_text.config(state="disabled")
        
        if page_mode in [2, 3]:
            self.log_message("▶ 작업 시작: 본문에 인쇄된 페이지 번호를 스캔하여 분석 중입니다...")
        else:
            self.log_message("▶ 작업 시작: 목차 텍스트를 시각적으로 스캔 중입니다...")
        
        # 1. 목차 추출
        candidates = self.extract_candidates(input_file, start_page, end_page, page_mode)
        
        if not candidates:
            self.log_message("\n경고: 목차 패턴을 찾지 못했습니다.")
            messagebox.showwarning("경고", "목차를 찾지 못했습니다. 페이지 범위를 확인해 주세요.")
            return
            
        # 추출된 목록을 한 번에 묶어서 화면에 출력
        self.log_text.config(state="normal")
        log_lines = [f"발견됨: {title} (이동: PDF {page}페이지)" for lvl, title, page in candidates]
        self.log_text.insert("end", "\n".join(log_lines) + "\n")
        self.log_text.see("end")
        self.log_text.config(state="disabled")
        self.root.update()

        # 2. 팝업창 없이 즉시 PDF 파일에 책갈피 저장
        self.log_message(f"\n▶ 총 {len(candidates)}개의 책갈피를 파일에 적용 중...")
        
        file_dir, file_name = os.path.split(input_file)
        name, ext = os.path.splitext(file_name)
        output_file = os.path.join(file_dir, f"{name}_bookmarked{ext}")

        doc = fitz.open(input_file)
        doc.set_toc(candidates)
        
        # [수정] 모든 PDF 뷰어에서 이동 오류가 발생하지 않도록 가장 안정적이고 표준적인 방식으로 저장
        doc.save(output_file)
        doc.close()
        
        self.log_message(f"▶ 완료되었습니다! 저장 위치: {output_file}")
        messagebox.showinfo("작업 완료", f"총 {len(candidates)}개의 책갈피가 성공적으로 저장되었습니다!")

if __name__ == "__main__":
    root = tk.Tk()
    app = PDFBookmarkerApp(root)
    root.mainloop()