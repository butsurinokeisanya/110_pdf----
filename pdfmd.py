import os
import PyPDF2
from tkinter import filedialog, Tk, Button, Frame, Label, Entry, Listbox, Scrollbar, MULTIPLE
from tkinterdnd2 import TkinterDnD, DND_FILES
from tkinter import ttk
from ttkthemes import ThemedTk


class PDFTool(TkinterDnD.Tk):  
    def __init__(self):
        super().__init__()

        self.files = []  # ファイルリストを初期化
        self.pdf_info = {}

        self.title("PDF 結合抽出")
        self.geometry("600x700")  # ウィンドウサイズを統一

        # 背景色の設定
        self.configure(bg="#F8F9FA")

        # スタイル設定
        self.style = ttk.Style()
        self.style.configure("TButton", font=("Arial", 16), padding=10)
        self.style.map("TButton",
                       background=[('active', '#007bff'), ('!active', '#007bff')],  # ボタンの色を青色に
                       foreground=[('active', 'white'), ('!active', 'white')])

        self.style.configure("TLabel", font=("Arial", 24), background="#F8F9FA", foreground="black")

        # 上部にソフト名を筆記体で追加
        self.software_name = Label(self, text="PDF 結合抽出", font=("Brush Script MT", 30), bg="#F8F9FA", fg="black", justify="center")
        self.software_name.pack(pady=5)  # 余白を詰める

        # 上部にタイトルを追加
        self.header_label = Label(self, text="PDF抽出 OR 結合\n 1つのPDFならページ数を指定して抽出\n 2つ以上のPDFなら結合", font=("Arial", 18), bg="#F8F9FA", fg="black", justify="center")
        self.header_label.pack(pady=5)  # 余白を詰める

        # 青いファイル選択ボタン（タイトルの下に配置）
        self.select_button = Button(self, text="ファイルを選択", command=self.select_files, bg="#007bff", fg="white", font=("Arial", 16), relief="flat", width=30)
        self.select_button.pack(pady=5)  # 余白を詰める

        # クリアボタンと全部クリアボタンを横に並べるフレーム
        self.clear_buttons_frame = Frame(self, bg="#F8F9FA")
        self.clear_buttons_frame.pack(pady=5)  # 余白を詰める

        # クリアボタンと全部クリアボタンを追加（横並び）
        self.clear_button = Button(self.clear_buttons_frame, text="選択をクリア", command=self.clear_selection, bg="#FF5733", fg="white", font=("Arial", 16), relief="flat", width=15)
        self.clear_button.pack(side="left", padx=5)

        self.clear_all_button = Button(self.clear_buttons_frame, text="全部クリア", command=self.clear_all_files, bg="#FF3333", fg="white", font=("Arial", 16), relief="flat", width=15)
        self.clear_all_button.pack(side="left", padx=5)

        # ドラッグアンドドロップエリアを含むフレーム
        self.drop_frame = Frame(self, bg="#F8F9FA")
        self.drop_frame.pack(pady=5)  # 余白を詰める

        # ドラッグアンドドロップエリアとスクロールバーのフレーム
        self.drop_area_frame = Frame(self.drop_frame, bg="black", relief="solid", borderwidth=2)  # 黒色枠を追加
        self.drop_area_frame.pack(padx=5, pady=5)  # 余白を詰める

        # ドラッグアンドドロップエリアにスクロールバーを追加
        self.scrollbar = Scrollbar(self.drop_area_frame, bg="lightgrey", activebackground="grey")  # スクロールバーを見やすく
        self.scrollbar.pack(side="right", fill="y")

        # ドラッグアンドドロップエリアにファイル名を表示するためのリストボックス
        self.file_listbox = Listbox(self.drop_area_frame, selectmode=MULTIPLE, font=("Arial", 14), bg="#FFFFFF", fg="black", yscrollcommand=self.scrollbar.set, width=40, height=10)
        self.file_listbox.pack(expand=True, fill="both")
        self.scrollbar.config(command=self.file_listbox.yview)

        # ドロップエリアに初期メッセージ
        self.file_listbox.insert('end', "ここにPDFをドラッグ&ドロップ")

        # ドラッグアンドドロップ設定
        self.drop_area_frame.drop_target_register(DND_FILES)
        self.drop_area_frame.dnd_bind('<<Drop>>', self.drop)

        # 抽出ページを指定する入力ボックス
        self.page_entry_label = Label(self, text="ページ番号を指定 (例: 1,3-5)", font=("Arial", 16), bg="#F8F9FA", fg="black")
        self.page_entry_label.pack(pady=5)  # 余白を詰める

        # テキストボックスの枠を黒く
        self.page_entry = Entry(self, font=("Arial", 16), width=30, bg="#e9ecef", fg="black", relief="solid", bd=2)
        self.page_entry.pack(pady=5)  # 余白を詰める

        # 実行ボタンを追加（ファイル数によって動作が変わる）
        self.execute_button = Button(self, text="実行", command=self.execute_action, bg="#007bff", fg="white", font=("Arial", 16), relief="flat", width=30)
        self.execute_button.pack(pady=5)  # 余白を詰める

    def drop(self, event):
        # ドロップされたファイルを処理し、既存ファイル名を保持して表示
        if "ここにPDFをドラッグ&ドロップ" in self.file_listbox.get(0, 'end'):
            self.file_listbox.delete(0)  # 初期メッセージをクリア

        files = self.tk.splitlist(event.data)
        self.add_files(files)

    def select_files(self):
        # エクスプローラからPDFファイルを選択
        files = filedialog.askopenfilenames(filetypes=[("PDF files", "*.pdf")])
        self.add_files(files)

    def add_files(self, files):
        for file in files:
            if file.endswith('.pdf') and file not in self.files:  # 重複を避ける
                pdf_reader = PyPDF2.PdfReader(file)
                num_pages = len(pdf_reader.pages)
                file_name = os.path.basename(file)
                self.files.append(file)
                self.pdf_info[file] = num_pages
                self.file_listbox.insert('end', f"{file_name} - {num_pages} pages")

    def clear_selection(self):
        # 選択されたファイルをクリア
        selected_indices = self.file_listbox.curselection()
        for index in reversed(selected_indices):  # 後ろから削除していくことでインデックスのズレを防ぐ
            file_name = self.file_listbox.get(index).split(' - ')[0]  # 表示されているファイル名
            self.files = [file for file in self.files if os.path.basename(file) != file_name]  # ファイルリストから削除
            self.file_listbox.delete(index)  # リストボックスから削除

        # ファイルがすべてクリアされたら、初期メッセージを再表示
        if len(self.files) == 0:
            self.file_listbox.insert('end', "ここにPDFをドラッグ&ドロップ")

    def clear_all_files(self):
        # 全てのファイルをクリア
        self.files.clear()
        self.file_listbox.delete(0, 'end')
        self.file_listbox.insert('end', "ここにPDFをドラッグ&ドロップ")

    def get_unique_filename(self, directory, base_name):
        """元のファイル名に0001、0002を追加して重複しないファイル名を取得"""
        count = 1
        base, ext = os.path.splitext(base_name)
        while True:
            candidate = f"{base}{count:04}{ext}"
            if not os.path.exists(os.path.join(directory, candidate)):
                return candidate
            count += 1

    def merge_pdfs(self):
        # 2つ以上のファイルを結合
        if len(self.files) < 2:
            return  # 結合するには2つ以上のファイルが必要です

        directory = os.path.dirname(self.files[0])
        original_file_name = os.path.basename(self.files[0])
        output_file = self.get_unique_filename(directory, original_file_name)

        merger = PyPDF2.PdfMerger()

        for pdf in self.files:
            merger.append(pdf)

        # PDFファイルを書き込み、ファイルを閉じる
        with open(os.path.join(directory, output_file), 'wb') as f_out:
            merger.write(f_out)

        # リソースを解放するために閉じる
        merger.close()

        print(f"PDFs merged and saved to {output_file}")

    def extract_pdf(self):
        # 1つのファイルからページを抽出
        if len(self.files) != 1:
            return  # 抽出は1つのファイルのみ対応

        input_pdf = self.files[0]
        if not input_pdf.endswith('.pdf'):
            return

        pages_to_extract = self.page_entry.get()
        if not pages_to_extract:
            return

        pdf_reader = PyPDF2.PdfReader(input_pdf)
        pdf_writer = PyPDF2.PdfWriter()

        try:
            for part in pages_to_extract.split(','):
                if '-' in part:
                    start, end = map(int, part.split('-'))
                    for page_num in range(start - 1, end):
                        pdf_writer.add_page(pdf_reader.pages[page_num])
                else:
                    page_num = int(part) - 1
                    pdf_writer.add_page(pdf_reader.pages[page_num])

            directory = os.path.dirname(input_pdf)
            original_file_name = os.path.basename(input_pdf)
            output_file = self.get_unique_filename(directory, original_file_name)

            with open(os.path.join(directory, output_file), 'wb') as f_out:
                pdf_writer.write(f_out)

            print(f"Pages extracted and saved to {output_file}")

        except Exception as e:
            print(f"Error: {e}")  # エラー処理

    def execute_action(self):
        # ファイルの数に応じて動作を分ける
        if len(self.files) == 1:
            self.extract_pdf()  # 1つのファイルなら抽出
        elif len(self.files) > 1:
            self.merge_pdfs()  # 2つ以上のファイルなら結合
        else:
            print("ファイルが選択されていません。")


if __name__ == "__main__":
    app = PDFTool()
    app.mainloop()
