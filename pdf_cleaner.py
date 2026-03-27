import sys
import traceback

# 프로그램 전체를 감싸서 에러 발생 시 창이 바로 꺼지지 않도록 방지합니다.
try:
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox
    import re
    import os
    import glob
    import ctypes
    import base64
    import io
    import tempfile
    import shutil
    import threading
    import time  # 스레드 작업 중 UI에 숨통을 틔워주기 위한 모듈
    from pathlib import Path

    # 필수 라이브러리 체크 (더블클릭 실행 시 아무 반응 없이 꺼지는 현상 방지)
    try:
        import fitz
    except ImportError:
        print("\n==================================================")
        print("🚨 필수 라이브러리 누락 오류!")
        print("==================================================")
        print("PDF 처리를 위한 'PyMuPDF' 모듈이 설치되지 않았습니다.")
        print("명령 프롬프트(CMD) 또는 PowerShell을 열고 아래 명령어를 입력해 설치해주세요:\n")
        print("pip install pymupdf\n")
        print("==================================================")
        input("확인하셨으면 엔터 키를 눌러 창을 닫아주세요...")
        sys.exit(1)

    try:
        from PIL import Image, ImageTk
    except ImportError:
        Image = None
        ImageTk = None

    # ── 라이선스 검증 모듈 ─────────────────────────────────────
    import json as _json
    from datetime import date as _date

    def check_startup_license():
        """license.json 파일 존재 여부 및 유효기간 확인"""
        import sys as _sys
        # exe 빌드 시: sys.executable = exe 파일 경로 → 그 폴더 기준
        # .py 실행 시: __file__ = 스크립트 경로 → 그 폴더 기준
        if getattr(_sys, 'frozen', False):
            _base = os.path.dirname(_sys.executable)
        else:
            _base = os.path.dirname(os.path.abspath(__file__))
        _license_path = os.path.join(_base, 'license.json')

        if not os.path.exists(_license_path):
            return False, "license.json 파일을 찾을 수 없습니다.\npdf_2.py와 같은 폴더에 license.json을 넣어주세요."

        try:
            with open(_license_path, 'r', encoding='utf-8') as _f:
                _lic = _json.load(_f)
        except Exception as _e:
            return False, f"라이선스 파일 읽기 오류: {_e}"

        if _lic.get('product_name') != 'pdf_cleaner':
            return False, "이 프로그램용 라이선스가 아닙니다."

        try:
            _today = _date.today()
            _start = _date.fromisoformat(_lic['starts_on'])
            _end   = _date.fromisoformat(_lic['expires_on'])
            if _today < _start:
                return False, f"라이선스 시작일({_lic['starts_on']})이 되지 않았습니다."
            if _today > _end:
                return False, f"라이선스가 만료되었습니다.\n만료일: {_lic['expires_on']}"
        except Exception as _e:
            return False, f"유효기간 확인 오류: {_e}"

        _client = _lic.get('client_name', '')
        _expires = _lic.get('expires_on', '')
        return True, f"라이선스 유효 | {_client} | 만료: {_expires}"

    def resource_path(relative_path):
        """PyInstaller 등으로 exe 빌드 시 임시 폴더(MEIPASS)의 절대 경로를 찾기 위한 헬퍼 함수."""
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.dirname(os.path.abspath(__file__))
        return os.path.join(base_path, relative_path)

    class PDFCleanerApp:
        LOGO_B64 = "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"

        def __init__(self, root):
            self.root = root
            self.root.title('PDF cleaner')
            self.root.geometry('720x850')
            self.root.configure(bg='#3D6AF2')
            try:
                myappid = 'dataclip.pdfcleaner.v2'
                ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
            except:
                pass
            
            # 내장 아이콘을 임시 파일로 저장 후 적용
            try:
                import tempfile
                _ico_data = base64.b64decode(PDFCleanerApp.LOGO_B64)
                with tempfile.NamedTemporaryFile(suffix='.ico', delete=False) as _ico_tmp:
                    _ico_tmp.write(_ico_data)
                    _ico_tmp_path = _ico_tmp.name
                self.root.iconbitmap(_ico_tmp_path)
            except Exception as _ie:
                pass

            self.pdf_data = []
            self.total_stats = {}
            self.logo_img = None
            self.style = ttk.Style()
            self.style.theme_use('clam')
            
            self.default_font = ('Malgun Gothic', 10)
            self.bold_font = ('Malgun Gothic', 10, 'bold')
            self.title_font = ('Malgun Gothic', 20, 'bold')
            self.header_font = ('Malgun Gothic', 11, 'bold')
            
            self.style.configure('Treeview.Heading', background='#3D6AF2', foreground='white', font=self.header_font, relief='flat')
            self.style.map('Treeview.Heading', background=[('active', '#2A52C2')], foreground=[('active', 'white')])
            
            self.setup_ui()

        # [추가됨] 무거운 작업 시 버튼들을 잠그고 로딩 커서를 보여주는 함수
        def set_processing_state(self, is_processing):
            if is_processing:
                self.btn_select_file.config(state=tk.DISABLED)
                self.btn_select_folder.config(state=tk.DISABLED)
                self.btn_clean.config(state=tk.DISABLED)
                self.btn_save_toc.config(state=tk.DISABLED)
                self.btn_save.config(state=tk.DISABLED)
                self.btn_reset.config(state=tk.DISABLED)
                self.root.config(cursor="wait")
            else:
                self.btn_select_file.config(state=tk.NORMAL)
                self.btn_select_folder.config(state=tk.NORMAL)
                self.btn_reset.config(state=tk.NORMAL)
                self.root.config(cursor="")
                if self.pdf_data:
                    self.btn_clean.config(state=tk.NORMAL)
                    self.btn_save.config(state=tk.NORMAL)
                    self.btn_save_toc.config(state=tk.NORMAL)

        def setup_ui(self):
            main_container = tk.Frame(self.root, bg='#3D6AF2', padx=30, pady=15)
            main_container.pack(fill=tk.BOTH, expand=True)
            
            header_frame = tk.Frame(main_container, bg='#3D6AF2')
            header_frame.pack(fill=tk.X, pady=(0, 15))
            
            title_label = tk.Label(header_frame, text='PDF cleaner', font=self.title_font, fg='white', bg='#3D6AF2')
            title_label.pack(side=tk.LEFT, anchor=tk.NW)
            
            try:
                from PIL import Image as _PI, ImageTk as _PT
                
                # [수정] 업로드해주신 '회사로고.png' 파일을 우선적으로 불러오도록 변경
                logo_path = resource_path('회사로고.png')
                if os.path.exists(logo_path):
                    _img = _PI.open(logo_path)
                else:
                    _raw = base64.b64decode(PDFCleanerApp.LOGO_B64)
                    _img = _PI.open(io.BytesIO(_raw))
                    
                _h = 45
                _w = int(_img.size[0] * (_h / _img.size[1]))
                _img = _img.resize((_w, _h), getattr(_PI, 'Resampling', _PI).LANCZOS)
                self.logo_img = _PT.PhotoImage(_img)
                tk.Label(header_frame, image=self.logo_img, bg='#3D6AF2').pack(side=tk.RIGHT, anchor=tk.CENTER, pady=5)
            except Exception as _le:
                print(f'[Warn] 로고 로드 실패: {_le}')
                    
            self.card_frame = tk.Frame(main_container, bg='white', highlightthickness=0, bd=0)
            self.card_frame.pack(fill=tk.BOTH, expand=True)
            
            content_frame = tk.Frame(self.card_frame, bg='white', padx=25, pady=25)
            content_frame.pack(fill=tk.BOTH, expand=True)
            
            section1_label = tk.Label(content_frame, text='1. 파일 및 폴더 선택', font=self.header_font, bg='white', fg='#333333')
            section1_label.pack(anchor=tk.W, pady=(0, 10))
            
            upload_container = tk.Frame(content_frame, bg='#F8F9FA', padx=15, pady=15)
            upload_container.pack(fill=tk.X, pady=(0, 20))
            
            btn_frame = tk.Frame(upload_container, bg='#F8F9FA')
            btn_frame.pack(fill=tk.X)
            
            self.btn_select_file = tk.Button(btn_frame, text='📄 파일 열기', command=self.load_file, bg='#3D6AF2', fg='white', font=self.bold_font, relief=tk.FLAT, padx=20, pady=7, cursor='hand2')
            self.btn_select_file.pack(side=tk.LEFT, padx=(0, 10))
            
            self.btn_select_folder = tk.Button(btn_frame, text='📁 폴더 열기', command=self.load_folder, bg='#6C757D', fg='white', font=self.bold_font, relief=tk.FLAT, padx=20, pady=7, cursor='hand2')
            self.btn_select_folder.pack(side=tk.LEFT)
            
            self.lbl_file_status = tk.Label(upload_container, text='현재 선택된 파일 없음', font=self.default_font, bg='#F8F9FA', fg='#666666')
            self.lbl_file_status.pack(anchor=tk.W, pady=(10, 0))
            
            section2_label = tk.Label(content_frame, text='2. 목차 확인 및 정제', font=self.header_font, bg='white', fg='#333333')
            section2_label.pack(anchor=tk.W, pady=(0, 10))
            
            self.style.configure('Treeview', font=self.default_font, rowheight=30)
            
            tree_container = tk.Frame(content_frame, bg='white')
            tree_container.pack(fill=tk.BOTH, expand=True)
            
            tree_scroll = ttk.Scrollbar(tree_container)
            tree_scroll.pack(side=tk.RIGHT, fill=tk.Y)
            
            self.tree = ttk.Treeview(tree_container, columns=('Page', 'Status'), yscrollcommand=tree_scroll.set, selectmode='none')
            self.tree.heading('#0', text='파일명 / 목차 제목 (Title)', anchor=tk.W)
            self.tree.heading('Page', text='페이지', anchor=tk.CENTER)
            self.tree.heading('Status', text='상태', anchor=tk.CENTER)
            
            self.tree.column('#0', width=420, stretch=True)
            self.tree.column('Page', width=70, anchor=tk.CENTER)
            self.tree.column('Status', width=90, anchor=tk.CENTER)
            self.tree.pack(fill=tk.BOTH, expand=True)
            
            tree_scroll.config(command=self.tree.yview)
            
            bottom_frame = tk.Frame(content_frame, bg='white', pady=15)
            bottom_frame.pack(fill=tk.X)
            
            self.btn_clean = tk.Button(bottom_frame, text='🧹 특수문자 일괄 치환', command=self.clean_toc, bg='#28A745', fg='white', font=self.bold_font, relief=tk.FLAT, padx=20, pady=10, cursor='hand2', state=tk.DISABLED)
            self.btn_clean.pack(side=tk.LEFT)
            
            self.lbl_bookmark_count = tk.Label(bottom_frame, text='추출된 총 목차: 0개', font=self.bold_font, fg='#3D6AF2', bg='white')
            self.lbl_bookmark_count.pack(side=tk.LEFT, padx=20)
            
            actions_frame = tk.Frame(content_frame, bg='white')
            actions_frame.pack(fill=tk.X, pady=(5, 0))
            
            self.btn_save_toc = tk.Button(actions_frame, text='📑 TOC 내보내기', command=self.save_toc, bg='#17A2B8', fg='white', font=self.bold_font, relief=tk.FLAT, padx=15, pady=6, cursor='hand2', state=tk.DISABLED)
            self.btn_save_toc.pack(side=tk.RIGHT)
            
            self.btn_save = tk.Button(actions_frame, text='💾 PDF 저장', command=self.save_pdf, bg='#3D6AF2', fg='white', font=self.bold_font, relief=tk.FLAT, padx=15, pady=6, cursor='hand2', state=tk.DISABLED)
            self.btn_save.pack(side=tk.RIGHT, padx=10)
            
            self.btn_reset = tk.Button(actions_frame, text='🔄 초기화', command=self.reset_all, bg='#FFFFFF', fg='#DC3545', font=self.bold_font, relief=tk.GROOVE, padx=15, pady=5, cursor='hand2')
            self.btn_reset.pack(side=tk.RIGHT)

        def apply_char_replacements(self, text):
            if not text:
                return ('', {})
            
            replaced = re.sub(r'^[\s\n\r\t■\-•▶▷*#]+', '', text.strip())
            replaced = re.sub(r'\s{2,}', ' ', replaced)
            
            char_mapping = [
                ('>', '〉', '> → 〉'), ('＞', '〉', '＞ → 〉'), ('<', '〈', '< → 〈'), 
                ('＜', '〈', '＜ → 〈'), ('·', 'ㆍ', '· → ㆍ'), ('‘', "'", "‘ → '"), 
                ('’', "'", "’ → '"), ('・', 'ㆍ', '・ → ㆍ'), ('‧', 'ㆍ', '‧ → ㆍ'), 
                ('･', 'ㆍ', '･ → ㆍ'), ('｢', '~', '〜'), ('〜', '~', '〜 → ~'), 
                ('․', 'ㆍ', '․ → ㆍ'), ('…', '...', '… → ...'), ('∼', '~', '∼ → ~')
            ]
            
            changes = {}
            for pattern, replacement, label in char_mapping:
                matches = len(re.findall(pattern, replaced))
                if matches > 0:
                    changes[label] = changes.get(label, 0) + matches
                    replaced = re.sub(pattern, replacement, replaced)
            return (replaced, changes)

        def _process_files(self, file_paths):
            self.reset_all()
            self.set_processing_state(True)
            self.lbl_file_status.config(text='파일을 불러오는 중입니다. 잠시만 기다려주세요...', fg='#DC3545')

            def task():
                valid_files = 0
                for path in file_paths:
                    try:
                        doc = fitz.open(path)
                        
                        # [오류 해결] get_toc(True)로 정상 페이지 번호 확보, get_toc(False)로 상세 링크 유지
                        toc_simple = doc.get_toc(True)
                        toc_detailed = doc.get_toc(False)
                        doc.close()

                        if toc_simple:
                            orig_toc = []
                            clean_toc = []
                            for i in range(len(toc_simple)):
                                lvl = int(toc_simple[i][0])
                                title = str(toc_simple[i][1])
                                page = int(toc_simple[i][2])  # 정상적인 1-based 페이지 번호 확보
                                
                                dest_dict = {}
                                # 상세 링크 정보(좌표 등) 보존
                                if i < len(toc_detailed) and len(toc_detailed[i]) > 3:
                                    if isinstance(toc_detailed[i][3], dict):
                                        dest_dict = toc_detailed[i][3].copy()
                                        
                                orig_toc.append([lvl, title, page, dest_dict])
                                clean_toc.append([lvl, title, page])

                            self.pdf_data.append({
                                'path': path, 
                                'filename': os.path.basename(path), 
                                'original_toc': orig_toc, 
                                'cleaned_toc': clean_toc, 
                                'stats': {}
                            })
                            valid_files += 1
                    except Exception as e:
                        print(f'Error processing {path}: {e}')
                    
                    time.sleep(0.01)

                self.root.after(0, self._process_files_complete, valid_files, file_paths)

            threading.Thread(target=task, daemon=True).start()

        def _process_files_complete(self, valid_files, file_paths):
            self.set_processing_state(False)
            if valid_files == 0:
                self.lbl_file_status.config(text='현재 선택된 파일 없음', fg='#666666')
                messagebox.showinfo('알림', '선택한 항목에서 목차(북마크)가 포함된 PDF를 찾을 수 없습니다.')
                return
                
            status_msg = f'{os.path.basename(file_paths[0])}' if valid_files == 1 else f'총 {valid_files}개의 파일 로드됨'
            self.lbl_file_status.config(text=status_msg, fg='#3D6AF2')
            self.render_treeview(is_cleaned=False)

        def load_file(self):
            file_path = filedialog.askopenfilename(title='PDF 파일 선택', filetypes=(('PDF Files', '*.pdf'), ('All Files', '*.*')))
            if file_path:
                self._process_files([file_path])

        def load_folder(self):
            folder_path = filedialog.askdirectory(title='PDF 파일이 있는 폴더 선택')
            if folder_path:
                pdf_files = glob.glob(os.path.join(folder_path, '*.pdf'))
                if not pdf_files:
                    messagebox.showinfo('알림', '해당 폴더에 PDF 파일이 없습니다.')
                else:
                    self._process_files(pdf_files)

        def render_treeview(self, is_cleaned):
            for item in self.tree.get_children():
                self.tree.delete(item)
                
            total_bookmarks = 0
            for data in self.pdf_data:
                file_node = self.tree.insert('', 'end', text=f"📄 {data['filename']}", values=('', ''))
                for i, (lvl, title, page) in enumerate(data['cleaned_toc']):
                    total_bookmarks += 1
                    orig_title = data['original_toc'][i][1]
                    indent = '  ' * (lvl - 1)
                    display_title = f'{indent}├ {title}' if lvl > 1 else title
                    status = '정상'
                    if is_cleaned and orig_title != title:
                        status = '✨ 치환됨'
                    self.tree.insert(file_node, 'end', text=f'  {display_title}', values=(page, status))
                self.tree.item(file_node, open=True)
                
            self.lbl_bookmark_count.config(text=f'추출된 총 목차: {total_bookmarks}개')

        def clean_toc(self):
            if not self.pdf_data:
                return
            
            self.set_processing_state(True)

            def task():
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
                        
                        if i % 100 == 0:
                            time.sleep(0.001)

                self.root.after(0, self._clean_toc_complete)

            threading.Thread(target=task, daemon=True).start()

        def _clean_toc_complete(self):
            self.set_processing_state(False)
            self.render_treeview(is_cleaned=True)
            self.show_summary_modal()

        def show_summary_modal(self):
            if not self.total_stats:
                messagebox.showinfo('치환 완료', '치환할 특수문자가 발견되지 않았습니다. 이미 깨끗합니다!')
            else:
                msg = '✅ 전체 파일 기준 다음 특수문자들이 정제되었습니다:\n\n'
                for label, count in self.total_stats.items():
                    msg += f'- {label} : {count}건\n'
                messagebox.showinfo('일괄 치환 완료 보고서', msg)

        # [최종 보완됨] 예시와 100% 동일하게 모든 최상위 목차 단위로 태그를 열고 닫으며, 하위 목차의 띄어쓰기를 맞춥니다.
        def _write_toc_file(self, data, save_path):
            toc_list = data['cleaned_toc']
            if not toc_list:
                return
                
            min_depth = min((item[0] for item in toc_list))
            toc_content = ''
            current_tag = None
            
            for lvl, title, page in toc_list:
                # 앞뒤 공백 제거하여 순수 텍스트만 추출
                title = title.strip()
                lower_title = title.replace(' ', '').lower()
                
                # 최상위 항목(min_depth)이거나 문서의 가장 첫 번째 항목일 경우
                # 기존 태그를 닫고 새로운 태그를 무조건 엽니다.
                if lvl == min_depth or current_tag is None:
                    if current_tag:
                        toc_content += f'</{current_tag}>\n'
                        
                    # 태그명 결정 (최상위 제목 기준)
                    if '표목차' in lower_title:
                        current_tag = 'table'
                    elif '그림목차' in lower_title or '도목차' in lower_title:
                        current_tag = 'figure'
                    elif '박스목차' in lower_title or '글상자목차' in lower_title:
                        current_tag = 'box'
                    else:
                        current_tag = 'body'
                        
                    # 새 태그 생성 (태그명 바로 뒤에 띄어쓰기 없이 제목, 그 뒤에 공백 1칸 후 페이지번호 삽입)
                    toc_content += f'<{current_tag}>{title} {page}\n'
                
                # 하위 항목일 경우, 현재 열려있는 태그 블록 내부에 레벨 깊이만큼 띄어쓰기 적용
                else:
                    rel_depth = lvl - min_depth
                    indent = ' ' * rel_depth # 레벨 깊이에 맞춘 스페이스 공백 생성
                    toc_content += f'{indent}{title} {page}\n'
                    
            # 제일 마지막 항목까지 처리 후 열려있는 최종 태그 닫기
            if current_tag:
                toc_content += f'</{current_tag}>\n'
                
            with open(save_path, 'w', encoding='utf-8') as f:
                f.write(toc_content)

        def save_toc(self):
            if not self.pdf_data:
                return
                
            if len(self.pdf_data) == 1:
                data = self.pdf_data[0]
                save_path = filedialog.asksaveasfilename(title='TOC 파일 저장', defaultextension='.toc', initialfile=os.path.splitext(data['filename'])[0], filetypes=(('TOC Files', '*.toc'),))
                if not save_path:
                    return
                self.set_processing_state(True)
                def task_single():
                    try:
                        self._write_toc_file(data, save_path)
                        self.root.after(0, lambda: self._save_toc_complete(1, save_path, None, True))
                    except Exception as e:
                        self.root.after(0, lambda e=e: self._save_toc_complete(0, save_path, str(e), True))
                threading.Thread(target=task_single, daemon=True).start()

            else:
                save_dir = filedialog.askdirectory(title='TOC 파일들을 저장할 폴더 선택')
                if not save_dir:
                    return
                self.set_processing_state(True)
                def task_multi():
                    success_count = 0
                    try:
                        for data in self.pdf_data:
                            self._write_toc_file(data, os.path.join(save_dir, f"{os.path.splitext(data['filename'])[0]}.toc"))
                            success_count += 1
                            time.sleep(0.005)
                        self.root.after(0, self._save_toc_complete, success_count, save_dir, None, False)
                    except Exception as e:
                        self.root.after(0, self._save_toc_complete, success_count, save_dir, str(e), False)
                threading.Thread(target=task_multi, daemon=True).start()

        def _save_toc_complete(self, success_count, save_path, error_msg, is_single):
            self.set_processing_state(False)
            if error_msg:
                messagebox.showerror('오류', f'저장 중 오류가 발생했습니다.\n{error_msg}')
            else:
                if is_single:
                    messagebox.showinfo('저장 완료', f'TOC 파일이 성공적으로 저장되었습니다!\n\n경로: {save_path}')
                else:
                    messagebox.showinfo('일괄 저장 완료', f'총 {success_count}개의 TOC 파일이 성공적으로 저장되었습니다!\n\n경로: {save_path}')

        def save_pdf(self):
            if not self.pdf_data:
                return
                
            if len(self.pdf_data) == 1:
                data = self.pdf_data[0]
                save_path = filedialog.asksaveasfilename(title='정제된 PDF 저장', defaultextension='.pdf', initialfile=data['filename'], filetypes=(('PDF Files', '*.pdf'),))
                if not save_path:
                    return
                
                self.set_processing_state(True)
                def task_single():
                    try:
                        doc = fitz.open(data['path'])
                        
                        # [수정] 책갈피 클릭 시 이동하지 않는 문제 완벽 해결!
                        # 원본 책갈피의 상세 링크 정보(dest_dict)를 그대로 유지하면서, 정제된 제목과 펼침 설정만 덮어씌웁니다.
                        new_toc = []
                        for i in range(len(data['original_toc'])):
                            orig_item = data['original_toc'][i]
                            cleaned_item = data['cleaned_toc'][i]
                            
                            lvl = abs(cleaned_item[0])
                            title = cleaned_item[1]
                            page = cleaned_item[2]
                            
                            # 원본에 상세 링크 정보가 있으면 복사하여 유지
                            if len(orig_item) > 3 and isinstance(orig_item[3], dict):
                                dest_dict = orig_item[3].copy()
                            else:
                                dest_dict = {}
                                
                            dest_dict['collapse'] = False
                            new_toc.append([lvl, title, page, dest_dict])

                        # 전체 적용 옵션(collapse=0)과 함께 상세 정보가 포함된 새 목차 덮어쓰기
                        doc.set_toc(new_toc, collapse=0)

                        with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as _tmp:
                            _tmp_path = _tmp.name
                            
                        doc.save(_tmp_path, garbage=1, deflate=True)
                        doc.close()
                        shutil.move(_tmp_path, save_path)
                        
                        self.root.after(0, lambda: self._save_pdf_complete(1, save_path, None, True))
                    except Exception as e:
                        self.root.after(0, lambda e=e: self._save_pdf_complete(0, save_path, str(e), True))
                threading.Thread(target=task_single, daemon=True).start()

            else:
                save_dir = filedialog.askdirectory(title='정제된 파일들을 저장할 폴더 선택')
                if not save_dir:
                    return
                
                self.set_processing_state(True)
                def task_multi():
                    success_count = 0
                    try:
                        for data in self.pdf_data:
                            save_path = os.path.join(save_dir, data['filename'])
                            doc = fitz.open(data['path'])
                            
                            # [수정] 다중 파일 처리 시에도 원본 링크 상세 정보(dest_dict) 완벽 유지
                            new_toc = []
                            for i in range(len(data['original_toc'])):
                                orig_item = data['original_toc'][i]
                                cleaned_item = data['cleaned_toc'][i]
                                
                                lvl = abs(cleaned_item[0])
                                title = cleaned_item[1]
                                page = cleaned_item[2]
                                
                                if len(orig_item) > 3 and isinstance(orig_item[3], dict):
                                    dest_dict = orig_item[3].copy()
                                else:
                                    dest_dict = {}
                                    
                                dest_dict['collapse'] = False
                                new_toc.append([lvl, title, page, dest_dict])

                            doc.set_toc(new_toc, collapse=0)

                            with tempfile.NamedTemporaryFile(suffix='.pdf', delete=False) as _tmp:
                                _tmp_path = _tmp.name
                                
                            doc.save(_tmp_path, garbage=1, deflate=True)
                            doc.close()
                            shutil.move(_tmp_path, save_path)
                            success_count += 1
                            time.sleep(0.01)

                        self.root.after(0, self._save_pdf_complete, success_count, save_dir, None, False)
                    except Exception as e:
                        self.root.after(0, self._save_pdf_complete, success_count, save_dir, str(e), False)
                threading.Thread(target=task_multi, daemon=True).start()

        def _save_pdf_complete(self, success_count, save_path, error_msg, is_single):
            self.set_processing_state(False)
            if error_msg:
                messagebox.showerror('오류', f'저장 중 오류가 발생했습니다.\n{error_msg}')
            else:
                if is_single:
                    messagebox.showinfo('저장 완료', f'성공적으로 저장되었습니다!\n\n경로: {save_path}')
                else:
                    messagebox.showinfo('일괄 저장 완료', f'총 {success_count}개의 파일이 성공적으로 저장되었습니다!\n\n경로: {save_path}')

        def reset_all(self):
            self.pdf_data = []
            self.total_stats = {}
            self.lbl_file_status.config(text='현재 선택된 파일 없음', fg='#666666')
            self.lbl_bookmark_count.config(text='추출된 총 목차: 0개')
            for item in self.tree.get_children():
                self.tree.delete(item)
            self.btn_clean.config(state=tk.DISABLED)
            self.btn_save.config(state=tk.DISABLED)
            self.btn_save_toc.config(state=tk.DISABLED)

    if __name__ == '__main__':
        is_license_valid, license_message = check_startup_license()
        if not is_license_valid:
            import tkinter as _tk
            from tkinter import messagebox as _mb
            _r = _tk.Tk()
            _r.withdraw()
            _mb.showerror(
                "라이선스 오류",
                f"{license_message}\n\n라이선스 파일을 확인 후 다시 실행해 주세요."
            )
            _r.destroy()
            sys.exit(1)
            
        try:
            import ctypes
            from ctypes import windll
            windll.shcore.SetProcessDpiAwareness(1)
            
            # [추가] 작업표시줄 아이콘을 파이썬 기본 아이콘에서 사용자 아이콘으로 분리/변경
            myappid = 'mycompany.pdfcleaner.app.1.0'
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID(myappid)
        except:
            pass
            
        root = tk.Tk()
        
        # [추가] exe 빌드 시 아이콘 경로를 찾기 위한 함수
        def resource_path(relative_path):
            import os, sys
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)

        # [추가] 프로그램 창 왼쪽 위 및 기본 아이콘 설정
        icon_path = resource_path('app_icon.ico')
        import os
        if os.path.exists(icon_path):
            root.iconbitmap(default=icon_path)

        app = PDFCleanerApp(root)
        root.mainloop()

except SystemExit:
    pass
except Exception as e:
    print("\n==================================================")
    print("🚨 프로그램 실행 중 치명적인 오류가 발생했습니다!")
    print("==================================================")
    traceback.print_exc()
    print("==================================================\n")
    input("에러 내용을 확인한 후 엔터 키를 눌러 종료하세요...")
