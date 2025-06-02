from pathlib import Path
from PyPDF2 import PdfReader
import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import re
import pandas as pd


def extract_text_from_pdf(
    pdf_path, insert_page_breaks=False, extract_chapters=False, progress_callback=None
):
    try:
        reader = PdfReader(pdf_path)
        all_text = ""
        for page_num, page in enumerate(reader.pages):
            text = page.extract_text()
            if not text:
                continue
            if insert_page_breaks:
                all_text += f"\n--- ページ {page_num + 1} ---\n{text.strip()}\n"
            else:
                all_text += f"{text.strip()}\n"
            if progress_callback and page_num % 10 == 0:
                progress_callback(f"{page_num + 1} ページ目まで処理中...")
        if extract_chapters:
            return all_text, extract_chapters_to_excel(all_text, progress_callback)
        return all_text, None
    except Exception as e:
        if progress_callback:
            progress_callback(f"エラー発生: {e}")
        return "", None


def extract_chapters_to_excel(text, progress_callback=None):
    segments = re.split(r"\s*．\s*", text)
    data = []
    current_chapter = ""
    buffer = []

    for seg in segments:
        chapter_match_seg = re.search(r"\s*◆\s*([^◆]+?)\s*(?:．|$)", seg)
        if chapter_match_seg:
            new_chapter = chapter_match_seg.group(1).strip()
            if current_chapter and buffer:
                data.append(
                    {"篇名": current_chapter, "本文": "．".join(buffer).strip()}
                )
                buffer = []
            current_chapter = new_chapter
            if progress_callback:
                progress_callback(f"　　　篇名検出: {current_chapter}")
        else:
            buffer.append(seg.strip())

    if current_chapter and buffer:
        data.append({"篇名": current_chapter, "本文": "．".join(buffer).strip()})

    return pd.DataFrame(data)


def browse_pdf():
    file_path = filedialog.askopenfilename(filetypes=[("PDF files", "*.pdf")])
    if file_path:
        entry_pdf.delete(0, tk.END)
        entry_pdf.insert(0, file_path)
        default_name = Path(file_path).stem + "_text"
        entry_filename.delete(0, tk.END)
        entry_filename.insert(0, default_name)


def browse_output():
    folder_path = filedialog.askdirectory()
    if folder_path:
        entry_output.delete(0, tk.END)
        entry_output.insert(0, folder_path)


def update_progress(message):
    text_progress.insert(tk.END, message + "\n")
    text_progress.see(tk.END)
    text_progress.update()


def start_extraction():
    pdf_path = entry_pdf.get()
    output_dir = entry_output.get()
    file_name = entry_filename.get().strip()

    if not pdf_path.lower().endswith(".pdf"):
        messagebox.showerror("エラー", "PDFファイルを選んでください")
        return

    output_dir_path = Path(output_dir).expanduser()
    output_dir_path.mkdir(parents=True, exist_ok=True)

    output_text_path = output_dir_path / f"{file_name}.txt"
    output_excel_path = output_dir_path / f"{file_name}.xlsx"

    text_progress.delete("1.0", tk.END)
    update_progress("抽出処理を開始します...")

    insert_page_breaks = var_page_break.get()
    extract_chapter_excel = var_excel.get()

    text, df = extract_text_from_pdf(
        pdf_path, insert_page_breaks, extract_chapter_excel, update_progress
    )

    if not text.strip():
        update_progress("テキスト埋め込み式でないため、抽出できませんでした。")
        return

    with open(output_text_path, "w", encoding="utf-8") as f:
        f.write(text)
        update_progress(f"テキストファイル保存完了: {output_text_path}")

    if extract_chapter_excel and df is not None:
        df.to_excel(output_excel_path, index=False)
        update_progress(f"Excelファイル保存完了: {output_excel_path}")

    update_progress("抽出完了！")


# ---------------------- GUIレイアウト ----------------------
root = tk.Tk()
root.title("PDFテキスト抽出アプリ")

tk.Label(root, text="抽出元PDF").grid(row=0, column=0, sticky="e")
entry_pdf = tk.Entry(root, width=40)
entry_pdf.grid(row=0, column=1, padx=5)
tk.Button(root, text="参照", command=browse_pdf).grid(row=0, column=2)

tk.Label(root, text="保存先フォルダ").grid(row=1, column=0, sticky="e")
entry_output = tk.Entry(root, width=40)
entry_output.insert(0, str(Path.home() / "Desktop"))
entry_output.grid(row=1, column=1, padx=5)
tk.Button(root, text="参照", command=browse_output).grid(row=1, column=2)

tk.Label(root, text="保存ファイル名").grid(row=2, column=0, sticky="e")
entry_filename = tk.Entry(root, width=40)
entry_filename.insert(0, "MD_classics")
entry_filename.grid(row=2, column=1, padx=5)

var_page_break = tk.BooleanVar()
tk.Checkbutton(root, text="ページ区切り挿入", variable=var_page_break).grid(
    row=3, column=1, sticky="w"
)

var_excel = tk.BooleanVar()
tk.Checkbutton(root, text="篇名・本文のExcelファイル出力", variable=var_excel).grid(
    row=4, column=1, sticky="w"
)

tk.Button(root, text="抽出開始", command=start_extraction).grid(
    row=5, column=1, pady=10
)

text_progress = scrolledtext.ScrolledText(root, width=60, height=8)
text_progress.grid(row=6, column=0, columnspan=3, padx=10, pady=5)

root.mainloop()
