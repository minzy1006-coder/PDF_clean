import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import os
import sys
import glob
import fitz  # PyMuPDF: PDF 처리를 위한 강력한 라이브러리

# [Frontend/UX] 고품질 이미지 리사이징을 위해 Pillow 라이브러리 사용
# 설치 필요: pip install Pillow
try:
    from PIL import Image, ImageTk
except ImportError:
    Image = None
    ImageTk = None

def resource_path(relative_path):
    """
    [Backend/Ops] PyInstaller 등으로 exe 빌드 시 임시 폴더(MEIPASS)의 절대 경로를 찾기 위한 헬퍼 함수.
    IDE(VSCode, PyCharm)에서 실행 위치(CWD)가 다를 때를 대비해 __file__ 기준으로 절대 경로를 잡습니다.
    """
    try:
        base_path = sys._MEIPASS
    except Exception:
        # os.path.abspath(".") 대신 실행 중인 스크립트의 위치를 기준으로 잡음 (강력한 경로 탐색)
        base_path = os.path.dirname(os.path.abspath(__file__))
    return os.path.join(base_path, relative_path)

class PDFCleanerApp:
    def __init__(self, root):
        self.root = root
        # 👇 이 타이틀이 뜨는지 꼭 확인하세요!
        self.root.title("데이터클립 PDF 목차 클리너 - 심플 & 고속 (다중 파일 지원)")
        self.root.geometry("650x640") 
        self.root.configure(padx=20, pady=20)
        
        # 상태 관리 (다중 파일 처리를 위한 리스트 구조)
        self.pdf_data = []
        self.total_stats = {}
        
        # Tkinter 이미지 객체는 가비지 컬렉션(GC)에 의해 날아갈 수 있으므로 클래스 변수에 저장해야 합니다.
        self.logo_img = None 
        
        self.setup_ui()

    def setup_ui(self):
        # 헤더 영역
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        # --- [Frontend/UX] 레이아웃 분리: 좌측(텍스트), 우측(로고) ---
        text_frame = ttk.Frame(header_frame)
        text_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        
        logo_frame = ttk.Frame(header_frame)
        logo_frame.pack(side=tk.RIGHT, padx=(10, 0))

        # 타이틀 및 설명 (좌측 프레임에 정렬)
        title_label = ttk.Label(text_frame, text="✨ PDF 목차 클리너", font=("Helvetica", 18, "bold"))
        title_label.pack(anchor=tk.W) # 왼쪽(West) 정렬
        desc_label = ttk.Label(text_frame, text="파일 또는 폴더를 선택하여 여러 PDF의 특수문자를 한 번에 정제하세요.", foreground="gray")
        desc_label.pack(anchor=tk.W, pady=(5, 0))
        
        # ---------------------------------------------------------
        # ✨ 로고 이미지 추가 영역 (우측 프레임에 배치)
        # ---------------------------------------------------------
        logo_filename = "데이터클립_회색로고.png"
        logo_path = resource_path(logo_filename)
        
        if Image and ImageTk and os.path.exists(logo_path):
            try:
                # 이미지 로드
                img = Image.open(logo_path)
                
                # 비율 유지하며 높이 45px로 리사이징
                base_height = 45
                h_percent = (base_height / float(img.size[1]))
                w_size = int((float(img.size[0]) * float(h_percent)))
                
                # 하위 호환성을 고려한 리샘플링 필터
                resample_filter = getattr(Image, 'Resampling', Image).LANCZOS 
                img = img.resize((w_size, base_height), resample_filter)
                
                self.logo_img = ImageTk.PhotoImage(img) # GC 방지
                
                # 로고를 오른쪽 프레임(logo_frame)에 넣고 오른쪽(East) 정렬
                logo_label = ttk.Label(logo_frame, image=self.logo_img)
                logo_label.pack(anchor=tk.E)
            except Exception as e:
                print(f"[Warn] 로고 이미지를 불러오는 중 오류 발생: {e}")
        # ---------------------------------------------------------

        # 파일 업로드 영역
        upload_frame = ttk.LabelFrame(self.root, text=" 1. 파일 및 폴더 선택 ", padding=15)
        upload_frame.pack(fill=tk.X, pady=(0, 15))
        
        # 파일/폴더 버튼 배치
        btn_container = ttk.Frame(upload_frame)
        btn_container.pack(side=tk.LEFT)
        
        self.btn_select_file = ttk.Button(btn_container, text="📄 파일 열기", command=self.load_file)
        self.btn_select_file.pack(side=tk.LEFT, padx=(0, 5))
        
        self.btn_select_folder = ttk.Button(btn_container, text="📁 폴더 열기", command=self.load_folder)
        self.btn_select_folder.pack(side=tk.LEFT, padx=(0, 10))
        
        self.lbl_file_status = ttk.Label(upload_frame, text="현재 선택된 파일 없음", foreground="gray")
        self.lbl_file_status.pack(side=tk.LEFT, fill=tk.X, expand=True)

        # 결과 리스트 영역 (Treeview)
        result_frame = ttk.LabelFrame(self.root, text=" 2. 목차 확인 및 정제 ", padding=15)
        result_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 15))
        
        # 트리뷰 스크롤바 세팅
        tree_scroll = ttk.Scrollbar(result_frame)
        tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.tree = ttk.Treeview(result_frame, columns=("Page", "Status"), yscrollcommand=tree_scroll.set, selectmode="none")
        self.tree.heading("#0", text="파일명 / 목차 제목 (Title)", anchor=tk.W)
        self.tree.heading("Page", text="페이지", anchor=tk.CENTER)
        self.tree.heading("Status", text="상태", anchor=tk.CENTER)
        
        self.tree.column("#0", width=400, stretch=True)
        self.tree.column("Page", width=60, anchor=tk.CENTER)
        self.tree.column("Status", width=80, anchor=tk.CENTER)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        tree_scroll.config(command=self.tree.yview)
        
        # 버튼 영역
        btn_frame = ttk.Frame(result_frame)
        btn_frame.pack(fill=tk.X, pady=(10, 0))
        
        self.btn_clean = ttk.Button(btn_frame, text="🧹 특수문자 일괄 치환", command=self.clean_toc, state=tk.DISABLED)
        self.btn_clean.pack(side=tk.LEFT, padx=(0, 10))
        self.lbl_bookmark_count = ttk.Label(btn_frame, text="추출된 총 목차: 0개", font=("Helvetica", 9, "bold"), foreground="blue")
        self.lbl_bookmark_count.pack(side=tk.LEFT)

        # 저장 영역
        save_frame = ttk.Frame(self.root)
        save_frame.pack(fill=tk.X)
        
        # 🔥 새로 추가된 TOC 내보내기 버튼
        self.btn_save_toc = ttk.Button(save_frame, text="📑 TOC 내보내기", command=self.save_toc, state=tk.DISABLED)
        self.btn_save_toc.pack(side=tk.RIGHT)
        
        self.btn_save = ttk.Button(save_frame, text="💾 정제된 PDF 저장", command=self.save_pdf, state=tk.DISABLED)
        self.btn_save.pack(side=tk.RIGHT, padx=(0, 10))
        
        self.btn_reset = ttk.Button(save_frame, text="초기화", command=self.reset_all)
        self.btn_reset.pack(side=tk.RIGHT, padx=(0, 10))

    def apply_char_replacements(self, text):
        if not text:
            return "", {}
            
        replaced = re.sub(r'^[\s\n\r\t■\-•▶▷*#]+', '', text.strip())
        replaced = re.sub(r'\s{2,}', ' ', replaced)
        
        char_mapping = [
            (r'>', '〉', '> → 〉'), (r'＞', '〉', '＞ → 〉'), (r'<', '〈', '< → 〈'), (r'＜', '〈', '＜ → 〈'),
            (r'·', 'ㆍ', '· → ㆍ'), (r'‘', "'", "‘ → '"), (r'’', "'", "’ → '"), (r'・', 'ㆍ', '・ → ㆍ'),
            (r'‧', 'ㆍ', '‧ → ㆍ'), (r'･', 'ㆍ', '･ → ㆍ'), (r'｢', '「', '｢ → 「'), (r'｣', '」', '｣ → 」'),
            (r'〔', '[', '〔 → ['), (r'〕', ']', '〕 → ]'), (r'\s?～\s?', '~', '～ → ~'), (r'〜', '~', '〜 → ~'),
            (r'․', 'ㆍ', '․ → ㆍ'), (r'…', '...', '… → ...'), (r'∼', '~', '∼ → ~'), (r'⋅', 'ㆍ', '⋅ → ㆍ'),
            (r'–', '-', '– → -'), (r'`', "'", "` → '"), (r'—', '-', '— → -'), (r'´', "'", "´ → '"),
            (r'­', '-', '­ → -'), (r'(?<!^)•', 'ㆍ', '• → ㆍ'), (r'‐', '-', '‐ → -'), (r'∙', 'ㆍ', '∙ → ㆍ'),
            (r'ᆞ', 'ㆍ', 'ᆞ → ㆍ'), (r'“', '"', '“ → "'), (r'”', '"', '” → "'), (r'ž', 'ㆍ', 'ž → ㆍ'),
            (r'【', '[', '【 → ['), (r'】', ']', '】 → ]'), (r'（', '(', '（ → ('), (r'）', ')', '） → )')
        ]
        
        changes = {}
        for pattern, replacement, label in char_mapping:
            matches = len(re.findall(pattern, replaced))
            if matches > 0:
                changes[label] = changes.get(label, 0) + matches
                replaced = re.sub(pattern, replacement, replaced)
                
        return replaced, changes

    def _process_files(self, file_paths):
        self.reset_all()
        valid_files = 0
        
        for path in file_paths:
            try:
                doc = fitz.open(path)
                toc = doc.get_toc()
                doc.close()
                
                # 목차가 있는 파일만 리스트에 추가
                if toc:
                    self.pdf_data.append({
                        'path': path,
                        'filename': os.path.basename(path),
                        'original_toc': toc,
                        'cleaned_toc': [[item[0], item[1], item[2]] for item in toc],
                        'stats': {}
                    })
                    valid_files += 1
            except Exception as e:
                print(f"Error processing {path}: {e}")
                pass
                
        if valid_files == 0:
            messagebox.showinfo("알림", "선택한 항목에서 목차(북마크)가 포함된 PDF를 찾을 수 없습니다.")
            return

        status_msg = f"{os.path.basename(file_paths[0])}" if valid_files == 1 else f"총 {valid_files}개의 파일 로드됨"
        self.lbl_file_status.config(text=status_msg, foreground="black")
        
        self.render_treeview(is_cleaned=False)
        self.btn_clean.config(state=tk.NORMAL)
        self.btn_save.config(state=tk.NORMAL)
        self.btn_save_toc.config(state=tk.NORMAL) # TOC 버튼 활성화

    def load_file(self):
        file_path = filedialog.askopenfilename(
            title="PDF 파일 선택",
            filetypes=(("PDF Files", "*.pdf"), ("All Files", "*.*"))
        )
        if file_path:
            self._process_files([file_path])

    def load_folder(self):
        folder_path = filedialog.askdirectory(title="PDF 파일이 있는 폴더 선택")
        if folder_path:
            # 폴더 내의 모든 pdf 검색
            pdf_files = glob.glob(os.path.join(folder_path, "*.pdf"))
            if not pdf_files:
                messagebox.showinfo("알림", "해당 폴더에 PDF 파일이 없습니다.")
                return
            self._process_files(pdf_files)

    def render_treeview(self, is_cleaned):
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        total_bookmarks = 0
        
        for data in self.pdf_data:
            # 다중 파일일 경우 최상단에 파일명 노드를 삽입 (폴더 뷰 효과)
            file_node = self.tree.insert("", "end", text=f"📄 {data['filename']}", values=("", ""))
            
            for i, (lvl, title, page) in enumerate(data['cleaned_toc']):
                total_bookmarks += 1
                orig_title = data['original_toc'][i][1]
                
                indent = "  " * (lvl - 1)
                display_title = f"{indent}├ {title}" if lvl > 1 else title
                
                status = "정상"
                if is_cleaned and orig_title != title:
                    status = "✨ 치환됨"
                
                # 파일 노드 하위에 자식 노드로 목차 삽입
                self.tree.insert(file_node, "end", text=f"  {display_title}", values=(page, status))
            
            # 기본적으로 파일 노드는 펼쳐진 상태로 유지
            self.tree.item(file_node, open=True)
            
        self.lbl_bookmark_count.config(text=f"추출된 총 목차: {total_bookmarks}개")

    def clean_toc(self):
        if not self.pdf_data:
            return
            
        self.total_stats = {}
        for data in self.pdf_data:
            data['stats'] = {}
            for i, item in enumerate(data['cleaned_toc']):
                orig_title = data['original_toc'][i][1]
                cleaned_text, changes = self.apply_char_replacements(orig_title)
                
                data['cleaned_toc'][i][1] = cleaned_text
                
                for label, count in changes.items():
                    data['stats'][label] = data['stats'].get(label, 0) + count
                    self.total_stats[label] = self.total_stats.get(label, 0) + count
                    
        self.render_treeview(is_cleaned=True)
        self.show_summary_modal()
        
    def show_summary_modal(self):
        if not self.total_stats:
            messagebox.showinfo("치환 완료", "치환할 특수문자가 발견되지 않았습니다. 이미 깨끗합니다!")
            return
            
        msg = "✅ 전체 파일 기준 다음 특수문자들이 정제되었습니다:\n\n"
        for label, count in self.total_stats.items():
            msg += f"- {label} : {count}건\n"
            
        messagebox.showinfo("일괄 치환 완료 보고서", msg)

    # --------------------------------------------------------------------------------
    # 🚀 TOC 내보내기 관련 백엔드 로직 (기존 HTML/JS 버전의 태그 구분 및 들여쓰기 완벽 이식)
    # --------------------------------------------------------------------------------
    def _write_toc_file(self, data, save_path):
        """단일 PDF 데이터의 목차를 추출해 HTML 유사 형식(.toc) 텍스트로 저장하는 코어 로직"""
        toc_list = data['cleaned_toc'] # [level, title, page] 구조
        if not toc_list:
            return
            
        # JS 코드의 `minDepth` 계산 방식과 동일하게 최소 깊이를 구함
        min_depth = min(item[0] for item in toc_list)
        toc_content = ""
        current_tag = None
        
        for lvl, title, page in toc_list:
            next_tag = current_tag or 'body'
            
            # 최상위 레벨 항목일 때 태그 종류 판단
            if lvl == min_depth:
                lower_title = title.replace(" ", "").lower()
                if "표목차" in lower_title:
                    next_tag = "table"
                elif "그림목차" in lower_title or "도목차" in lower_title:
                    next_tag = "figure"
                elif "박스목차" in lower_title or "글상자목차" in lower_title:
                    next_tag = "box"
                else:
                    next_tag = "body"
                    
            # 태그가 바뀌거나 맨 처음 시작할 때 (태그 열고 바로 같은 줄에 텍스트 작성)
            if current_tag != next_tag or not current_tag:
                if current_tag:
                    toc_content += f"</{current_tag}>\n"
                current_tag = next_tag
                toc_content += f"<{current_tag}>{title} {page}\n"
            else:
                # 같은 태그 내의 일반/하위 항목 (들여쓰기 적용)
                rel_depth = lvl - min_depth
                indent = " " * (rel_depth if rel_depth > 0 else 0)
                toc_content += f"{indent}{title} {page}\n"
                
        # 마지막으로 열려있는 태그 닫기
        if current_tag:
            toc_content += f"</{current_tag}>\n"
            
        # [Backend Tip] 한글 인코딩 문제(cp949 에러)를 방지하기 위해 반드시 utf-8을 명시합니다.
        with open(save_path, 'w', encoding='utf-8') as f:
            f.write(toc_content)

    def save_toc(self):
        if not self.pdf_data:
            return
            
        # 1. 단일 파일 TOC 저장
        if len(self.pdf_data) == 1:
            data = self.pdf_data[0]
            base_filename = os.path.splitext(data['filename'])[0]
            save_path = filedialog.asksaveasfilename(
                title="TOC 파일 저장",
                defaultextension=".toc",
                initialfile=f"{base_filename}",
                filetypes=(("TOC Files", "*.toc"), ("Text Files", "*.txt"), ("All Files", "*.*"))
            )
            if not save_path: return
            
            try:
                self._write_toc_file(data, save_path)
                messagebox.showinfo("저장 완료", f"TOC 파일이 성공적으로 저장되었습니다!\n\n경로: {save_path}")
            except Exception as e:
                messagebox.showerror("오류", f"TOC 저장 중 오류가 발생했습니다.\n{str(e)}")
                
        # 2. 다중 파일(폴더) 일괄 TOC 저장
        else:
            save_dir = filedialog.askdirectory(title="TOC 파일들을 저장할 폴더 선택")
            if not save_dir: return
            
            success_count = 0
            try:
                for data in self.pdf_data:
                    base_filename = os.path.splitext(data['filename'])[0]
                    save_path = os.path.join(save_dir, f"{base_filename}.toc")
                    self._write_toc_file(data, save_path)
                    success_count += 1
                    
                messagebox.showinfo("일괄 저장 완료", f"총 {success_count}개의 TOC 파일이 성공적으로 저장되었습니다!\n\n경로: {save_dir}")
            except Exception as e:
                messagebox.showerror("오류", f"일괄 저장 중 오류가 발생했습니다.\n{str(e)}")
    # --------------------------------------------------------------------------------

    def save_pdf(self):
        if not self.pdf_data:
            return
            
        # 1. 단일 파일 저장 로직
        if len(self.pdf_data) == 1:
            data = self.pdf_data[0]
            save_path = filedialog.asksaveasfilename(
                title="정제된 PDF 저장",
                defaultextension=".pdf",
                initialfile=data['filename'],
                filetypes=(("PDF Files", "*.pdf"),)
            )
            if not save_path: return
            
            try:
                doc = fitz.open(data['path'])
                doc.set_toc(data['cleaned_toc'])
                doc.save(save_path)
                doc.close()
                messagebox.showinfo("저장 완료", f"성공적으로 저장되었습니다!\n\n경로: {save_path}")
            except Exception as e:
                messagebox.showerror("오류", f"저장 중 오류가 발생했습니다.\n{str(e)}")
                
        # 2. 다중 파일(폴더) 일괄 저장 로직
        else:
            save_dir = filedialog.askdirectory(title="정제된 파일들을 저장할 폴더 선택")
            if not save_dir: return
            
            success_count = 0
            try:
                for data in self.pdf_data:
                    save_path = os.path.join(save_dir, data['filename'])
                    
                    # 혹시나 원본 폴더를 그대로 선택한 경우, 덮어쓰기 에러를 막기 위해 파일명 변경
                    if save_path == data['path']:
                        save_path = os.path.join(save_dir, f"cleaned_{data['filename']}")
                        
                    doc = fitz.open(data['path'])
                    doc.set_toc(data['cleaned_toc'])
                    doc.save(save_path)
                    doc.close()
                    success_count += 1
                    
                messagebox.showinfo("일괄 저장 완료", f"총 {success_count}개의 파일이 성공적으로 저장되었습니다!\n\n경로: {save_dir}")
            except Exception as e:
                messagebox.showerror("오류", f"일괄 저장 중 오류가 발생했습니다.\n{str(e)}")

    def reset_all(self):
        self.pdf_data = []
        self.total_stats = {}
        
        self.lbl_file_status.config(text="현재 선택된 파일 없음", foreground="gray")
        self.lbl_bookmark_count.config(text="추출된 총 목차: 0개")
        
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        self.btn_clean.config(state=tk.DISABLED)
        self.btn_save.config(state=tk.DISABLED)
        self.btn_save_toc.config(state=tk.DISABLED)


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCleanerApp(root)
    root.mainloop()