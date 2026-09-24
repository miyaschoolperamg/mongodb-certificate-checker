import re
import os
import pdfplumber
from datetime import datetime
from pymongo import MongoClient
from openpyxl import Workbook
import tkinter as tk
from tkinter import filedialog, messagebox
from datetime import datetime



def export_to_excel(output_path="SD_Marks_Report.xlsx",
                     db_name="internship_db", collection_name="SD_Marks"):
    client = MongoClient("mongodb://localhost:27017/")
    collection = client[db_name][collection_name]

    wb = Workbook()
    ws = wb.active
    ws.title = "SD Marks"
    ws.append(["Student Name", "Total Marks Awarded"])

    for doc in collection.find():
        total = sum(cert["score"] for cert in doc.get("certificates", []))
        ws.append([doc["student_name"], total])

    wb.save(output_path)
    client.close()
    print(f"{output_path} を出力しました。")


def read_certificate(pdf_path):
    """PDFから学生名・Certificate名・日付・IDを読み取る。
    画像化されていて文字が取れないPDFの場合は、全項目Noneを返す。"""
    with pdfplumber.open(pdf_path) as pdf:
        text = pdf.pages[0].extract_text()

    if not text:
        # 画像として保存された証明書など、テキストが一切取れない場合
        return {
            "student_name": None,
            "certificate_name": None,
            "certificate_date": None,
            "certificate_id": None,
        }

    lines = [l.strip() for l in text.split("\n") if l.strip()]

    date_match = re.search(r"\d{2}-\d{2}-\d{4}", text)
    cert_date = datetime.strptime(date_match.group(), "%m-%d-%Y") if date_match else None

    id_match = re.search(r"\bMDB[A-Za-z0-9]+\b", text)
    cert_id = id_match.group() if id_match else None

    # 日付の行とID行を除いた残りが「名前」と「Certificate名」
    remaining = [
        l for l in lines
        if l != (date_match.group() if date_match else None) and l != cert_id
    ]

    return {
        "student_name": remaining[0] if len(remaining) > 0 else None,
        "certificate_name": remaining[1] if len(remaining) > 1 else None,
        "certificate_date": cert_date,
        "certificate_id": cert_id,
    }


def parse_folder_name(folder_path):
    """サブフォルダ名から名前部分と提出日時を読み取る。"""
    folder = os.path.basename(folder_path)
    parts = folder.split(" - ")
    name_part = " - ".join(parts[1:-1])
    date_part = parts[-1]
    submission_date = datetime.strptime(date_part, "%d %B %Y %I_%M %p")
    return {
        "folder_name_text": name_part,
        "submission_date": submission_date,
    }


def normalize(s):
    return " ".join(s.lower().split())

def normalize_name(s):
    """名前の比較専用。PDF抽出時の余分な空白(文字間の分割など)を無視する。"""
    return "".join(s.lower().split())


def name_matches(pdf_name, folder_text):
    if not pdf_name:
        return False
    return normalize_name(pdf_name) in normalize_name(folder_text)


def evaluate(cert, folder_info, target_cert_name, target_date):
    reasons = []

    if not name_matches(cert["student_name"], folder_info["folder_name_text"]):
        reasons.append("Student name does not match folder name")

    if cert["certificate_date"] is None:
        reasons.append("Certificate date not found")
    else:
        cert_day = cert["certificate_date"].date()
        if cert_day > folder_info["submission_date"].date():
            reasons.append("Certificate date is after submission date")

    if cert["certificate_name"] is None or normalize(cert["certificate_name"]) != normalize(target_cert_name):
        reasons.append("Certificate name does not match")

    if not cert["certificate_id"]:
        reasons.append("Certificate ID missing")

    if reasons:
        return 0, "; ".join(reasons)

    # ここまで来れば、検証はすべて通っている。期限に間に合ったかどうかだけを見る。
    if cert["certificate_date"].date() > target_date.date():
        return 1, "Valid certificate submitted after target completion date"
    return 2, "Valid certificate submitted on time"

def process_folder(root_path, target_cert_name, target_date):
    """root_path直下の各サブフォルダを処理し、結果のリストを返す。"""
    results = []
    for name in os.listdir(root_path):
        folder_path = os.path.join(root_path, name)
        if not os.path.isdir(folder_path):
            continue

        pdfs = [f for f in os.listdir(folder_path) if f.lower().endswith(".pdf")]
        if len(pdfs) == 0:
            print(f"スキップ: {name}(PDFが0件)")
            continue

        if len(pdfs) > 1:
            # 複数ある場合は、ファイルの更新日時が一番新しいものを使う
            pdfs.sort(
                key=lambda f: os.path.getmtime(os.path.join(folder_path, f)),
                reverse=True,
            )
            print(f"注意: {name} にPDFが{len(pdfs)}件あるため、最新のものを使用: {pdfs[0]}")

        pdf_file = pdfs[0]

        try:
            cert = read_certificate(os.path.join(folder_path, pdf_file))
            folder_info = parse_folder_name(folder_path)
            score, rationale = evaluate(cert, folder_info, target_cert_name, target_date)
        except Exception as e:
            print(f"エラー(想定外): {name}: {e}")
            score, rationale = 0, f"Processing error: {e}"
            cert = {"student_name": None, "certificate_id": None}
            folder_info = {"submission_date": None}

        results.append({
            "student_name": cert["student_name"],
            "certificate_name": target_cert_name,
            "score": score,
            "rationale": rationale,
            "submission_date": folder_info["submission_date"],
            "certificate_number": cert["certificate_id"],
        })
    return results


def save_to_mongodb(results, db_name="internship_db", collection_name="SD_Marks"):
    client = MongoClient("mongodb://localhost:27017/")
    db = client[db_name]
    collection = db[collection_name]

    for r in results:
        entry = {
            "certificate_name": r["certificate_name"],
            "score": r["score"],
            "rationale": r["rationale"],
            "submission_date": r["submission_date"],
            "certificate_number": r["certificate_number"],
        }

        collection.update_one(
            {"student_name": r["student_name"]},
            {"$push": {"certificates": entry}},
            upsert=True,
        )

    client.close()

def browse_folder():
    path = filedialog.askdirectory()
    if path:
        folder_entry.delete(0, tk.END)
        folder_entry.insert(0, path)


def run_process():
    cert_name = cert_entry.get().strip()
    folder_path = folder_entry.get().strip()
    date_str = date_entry.get().strip()

    if not cert_name or not folder_path or not date_str:
        messagebox.showerror("入力エラー", "すべての項目を入力してください。")
        return

    try:
        target_date = datetime.strptime(date_str, "%d/%m/%Y")
    except ValueError:
        messagebox.showerror("入力エラー", "日付は DD/MM/YYYY の形式で入力してください。")
        return

    try:
        results = process_folder(folder_path, cert_name, target_date)
        save_to_mongodb(results)
        export_to_excel()
    except Exception as e:
        messagebox.showerror("エラー", str(e))
        return

    messagebox.showinfo("完了", f"{len(results)} 件処理しました。\nSD_Marks_Report.xlsx を出力しました。")


root = tk.Tk()
root.title("Certificate Checker")
root.geometry("420x200")

tk.Label(root, text="Certificate名").grid(row=0, column=0, sticky="w", padx=10, pady=10)
cert_entry = tk.Entry(root, width=35)
cert_entry.grid(row=0, column=1, padx=10)

tk.Label(root, text="フォルダのパス").grid(row=1, column=0, sticky="w", padx=10, pady=10)
folder_entry = tk.Entry(root, width=28)
folder_entry.grid(row=1, column=1, padx=(10, 0))
tk.Button(root, text="参照", command=browse_folder).grid(row=1, column=2, padx=5)

tk.Label(root, text="Target Completion Date\n(DD/MM/YYYY)").grid(row=2, column=0, sticky="w", padx=10, pady=10)
date_entry = tk.Entry(root, width=35)
date_entry.grid(row=2, column=1, padx=10)

tk.Button(root, text="実行", command=run_process, width=15).grid(row=3, column=1, pady=20)

root.mainloop()

