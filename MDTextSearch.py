import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import fitz  # PyMuPDF
import openpyxl
from openpyxl.styles import Alignment
import re
import os


def search_text_in_pdf(pdf_path, keyword, progress_callback, book_title):
    results = []
    try:
        doc = fitz.open(pdf_path)
        current_chapter = ""
        current_den = ""
        progress_callback(f"検索開始: {book_title}")

        # 柔軟なキーワード正規表現パターン（任意の空白・記号等を許容）
        keyword_flexible = r"\b"
        for char in keyword:
            keyword_flexible += r"[\s\S]*?" + re.escape(char) + r"[\s\S]*?"
        keyword_flexible += r"\b"
        search_pattern_seg = r"(?ms)" + keyword_flexible

        for page_num in range(len(doc)):
            page = doc[page_num]
            text = page.get_text()

            normalized_text = re.sub(r"\s+", " ", text).replace("\n", " ").strip()
            segments = re.split(r"．", normalized_text)

            for i, seg in enumerate(segments):
                seg = seg.strip()
                keyword_match = re.search(search_pattern_seg, seg)

                # -------- 共通：章名（◆〇〇）検出 --------
                chapter_match_seg = re.search(r"\s*◆\s*([^◆]+?)\s*(?:．|$)", seg)
                if chapter_match_seg:
                    new_chapter = chapter_match_seg.group(1).strip()
                    if new_chapter != current_chapter:
                        current_chapter = new_chapter
                        progress_callback(f"　　　篇名更新: {current_chapter}")

                # -------- 『扁鵲倉公伝』固有処理 --------
                if book_title == "扁鵲倉公伝":
                    # ■扁鵲伝 → 伝名更新（早期returnしない）
                    if "■扁鵲伝" in seg and current_den != "扁鵲伝":
                        current_den = "扁鵲伝"
                        progress_callback(f"　　　伝名更新: {current_den}")

                    # ■倉公伝 → 伝名更新
                    elif "■倉公伝" in seg and current_den != "倉公伝":
                        current_den = "倉公伝"
                        progress_callback(f"　　　伝名更新: {current_den}")

                    # ○診藉、○診籍、◎問答 のパターン検出（倉公伝の場合のみ）
                    if current_den == "倉公伝":
                        match = re.search(
                            r"[○◎](診藉|診籍|問答)[①-⑳0-9０-９一二三四五六七八九十百千]+",
                            seg,
                        )
                        if match:
                            chapter_candidate = match.group()
                            if chapter_candidate != current_chapter:
                                current_chapter = chapter_candidate
                                progress_callback(f"　　　篇名更新: {current_chapter}")

                # -------- キーワードヒットした場合、結果保存 --------
                if keyword_match:
                    matched_text = keyword_match.group(0)
                    progress_callback(
                        f"キーワード '{keyword}' を検出 - '{matched_text}'"
                    )
                    results.append(
                        {
                            "page": page_num + 1,
                            "chapter": current_chapter,
                            "den": current_den,
                            "text": seg,
                        }
                    )
                    progress_callback(f"　　　テキスト追加: '{seg[:30]}...'")

            progress_callback(f"検索中: {book_title} - {page_num + 1}/{len(doc)}ページ")

        progress_callback(
            f"検索完了: {book_title} - {len(results)}件ヒット (全{len(doc)}ページ)"
        )
        return results, None

    except Exception as e:
        progress_callback(f"エラー: {book_title} の検索中にエラーが発生しました - {e}")
        return [], e


def center_window(window, width=250, height=250):
    screen_width = window.winfo_screenwidth()
    screen_height = window.winfo_screenheight()

    x = int((screen_width - width) / 2)
    y = int((screen_height - height) / 2)

    window.geometry(f"{width}x{height}+{x}+{y}")


def save_results_to_excel(results, filepath, book_title):
    try:
        workbook = openpyxl.Workbook()
        sheet = workbook.active
        sheet.title = book_title

        # 書籍ごとのヘッダーと列幅設定
        if book_title in ["素問", "霊枢", "難経", "傷寒論", "金匱要略"]:
            headers = ["id", "ページ", "篇名", "検索文字列(正規化)"]
            col_widths = [5, 10, 20, 80]
        elif book_title == "神農本草経":
            headers = ["id", "ページ", "検索文字列(正規化)"]
            col_widths = [5, 10, 80]
        elif book_title == "扁鵲倉公伝":
            headers = ["id", "ページ", "伝名", "篇名", "検索文字列(正規化)"]
            col_widths = [5, 10, 12, 20, 80]
        else:
            # デフォルト（万一該当がなければ）
            headers = ["id", "ページ", "検索文字列(正規化)"]
            col_widths = [5, 10, 80]

        sheet.append(headers)

        for index, result in enumerate(results):
            normalized_text = re.sub(r"\s+", "", result["text"]).strip()

            # 行データの構築
            row = [index + 1, result.get("page", "")]
            if "伝名" in headers:
                row.append(result.get("den", ""))
            if "篇名" in headers:
                row.append(result.get("chapter", ""))
            row.append(normalized_text)
            sheet.append(row)

            # セルの整形（折り返し、左寄せ）
            for col_index in range(1, len(row) + 1):
                cell = sheet.cell(row=index + 2, column=col_index)
                cell.alignment = Alignment(wrap_text=True, horizontal="left")

        # 列幅設定
        for i, width in enumerate(col_widths):
            col_letter = chr(65 + i)  # 'A'〜
            sheet.column_dimensions[col_letter].width = width

        workbook.save(filepath)
        return True, None

    except Exception as e:
        return False, e


def start_search():
    selected_book = book_var.get()
    keyword = keyword_entry.get().strip()
    save_filename = save_filename_entry.get()
    save_folder = save_folder_entry.get()

    progress_text.config(state=tk.NORMAL)
    progress_text.delete(1.0, tk.END)

    if not selected_book:
        messagebox.showerror("エラー", "対象古典を選んでください")
        progress_text.config(state=tk.DISABLED)
        return
    if not keyword:
        messagebox.showerror("エラー", "キーワードを入力してください")
        progress_text.config(state=tk.DISABLED)
        return

    book_files = {
        "素問": "somon.pdf",
        "霊枢": "reisu.pdf",
        "難経": "nangyo.pdf",
        "傷寒論": "shanghanlun.pdf",
        "金匱要略": "jinguiyaolue.pdf",
        "神農本草経": "shennongbencaojing.pdf",
        "扁鵲倉公伝": "henjyaku.pdf",
    }

    pdf_path = book_files.get(selected_book)
    print(f"DEBUG: selected_book = {selected_book}")
    print(f"DEBUG: pdf_path = {pdf_path}")
    print(
        f"DEBUG: os.path.exists(pdf_path) = {os.path.exists(pdf_path)}"
    )  # これがFalseになっているはず

    if not pdf_path or not os.path.exists(pdf_path):
        messagebox.showerror(
            "エラー", f"{selected_book} のPDFファイルが見つかりません。"
        )
        progress_text.config(state=tk.DISABLED)
        return

    results, error = search_text_in_pdf(
        pdf_path, keyword, progress_callback, selected_book
    )

    if error:
        progress_text.insert(tk.END, f"検索エラー: {error}\n")
    elif results:
        if len(results) > 0:
            if not save_filename:
                save_filename = "MD_text"
            filepath = os.path.join(save_folder, f"{save_filename}.xlsx")
            save_success, save_error = save_results_to_excel(
                results, filepath, selected_book
            )
            if save_success:
                progress_text.insert(
                    tk.END, f"検索結果を {filepath} に保存しました。\n"
                )
            else:
                progress_text.insert(tk.END, f"保存エラー: {save_error}\n")
        else:
            progress_text.insert(tk.END, "該当する文字列は、0件でした。\n")
    else:
        progress_text.insert(tk.END, "該当するキーワードは見つかりませんでした。\n")

    progress_text.config(state=tk.DISABLED)


def browse_folder():
    folder_selected = filedialog.askdirectory()
    save_folder_entry.delete(0, tk.END)
    save_folder_entry.insert(0, folder_selected)


def progress_callback(message):
    progress_text.config(state=tk.NORMAL)
    progress_text.insert(tk.END, message + "\n")
    progress_text.see(tk.END)
    root.update()


# GUIの作成
root = tk.Tk()
root.title("医学古典テキスト検索")

# ウィンドウの初期サイズを設定して中央に表示
window_width = 500
window_height = 700
center_window(root, window_width, window_height)

# 説明文
description_label = tk.Label(
    root,
    text="キーワードを入力して、主要古典の該当部分を検索し、Excelファイルに保存します。",
)
description_label.pack(pady=5)

# ラジオボタン
book_var = tk.StringVar()
books = ["素問", "霊枢", "難経", "傷寒論", "金匱要略", "神農本草経", "扁鵲倉公伝"]
book_frame = ttk.LabelFrame(root, text="対象古典")
book_frame.pack(padx=10, pady=5, fill=tk.X)
for book in books:
    ttk.Radiobutton(book_frame, text=book, variable=book_var, value=book).pack(
        anchor=tk.W
    )

# キーワード入力欄
keyword_label = tk.Label(root, text="検索キーワード:")
keyword_label.pack(pady=5)
keyword_entry = tk.Entry(root, width=50)
keyword_entry.pack(padx=10, pady=5)

# 保存ファイル名入力欄
save_filename_label = tk.Label(root, text="保存ファイル名:")
save_filename_label.pack(pady=5)
save_filename_entry = tk.Entry(root, width=50)
save_filename_entry.insert(0, "MD_text")
save_filename_entry.pack(padx=10, pady=5)

# ファイル保存先入力欄
save_folder_label = tk.Label(root, text="ファイル保存先:")
save_folder_label.pack(pady=5)
save_folder_entry = tk.Entry(root, width=50)
save_folder_entry.insert(0, os.path.expanduser("~/Desktop"))
save_folder_entry.pack(padx=10, pady=5)

# 保存先参照ボタン
browse_button = ttk.Button(root, text="参照", command=browse_folder)
browse_button.pack(pady=5)

# 検索開始ボタン
search_button = ttk.Button(root, text="検索開始", command=start_search)
search_button.pack(pady=10)

# 進行状況確認欄
progress_label = tk.Label(root, text="進行状況:")
progress_label.pack(pady=5)
progress_text = scrolledtext.ScrolledText(root, height=10, width=60, state=tk.DISABLED)
progress_text.pack(padx=10, pady=5)

root.mainloop()
