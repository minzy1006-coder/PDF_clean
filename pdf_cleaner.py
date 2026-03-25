import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import re
import os
import glob
import fitz  # PyMuPDF: PDF 처리를 위한 강력한 라이브러리

class PDFCleanerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("PDF 목차 클리너 - 심플 & 고속 (다중 파일 지원)")
        self.root.geometry("650x600")
        self.root.configure(padx=20, pady=20)
        
        # 상태 관리 (다중 파일 처리를 위한 리스트 구조)
        # pdf_data 형식: [{'path': str, 'filename': str, 'original_toc': list, 'cleaned_toc': list, 'stats': dict}]
        self.pdf_data = []
        self.total_stats = {}
        
        self.setup_ui()

    def setup_ui(self):
        # 헤더 영역
        header_frame = ttk.Frame(self.root)
        header_frame.pack(fill=tk.X, pady=(0, 20))
        
        title_label = ttk.Label(header_frame, text="✨ PDF 목차 클리너", font=("Helvetica", 18, "bold"))
        title_label.pack()
        desc_label = ttk.Label(header_frame, text="파일 또는 폴더를 선택하여 여러 PDF의 특수문자를 한 번에 정제하세요.", foreground="gray")
        desc_label.pack()

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
        
        self.btn_save = ttk.Button(save_frame, text="💾 정제된 PDF 저장", command=self.save_pdf, state=tk.DISABLED)
        self.btn_save.pack(side=tk.RIGHT)
        
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


if __name__ == "__main__":
    root = tk.Tk()
    app = PDFCleanerApp(root)
    root.mainloop()