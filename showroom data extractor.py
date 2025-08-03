import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk
from datetime import datetime

def extract_text_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "".join(page.extract_text() or "" for page in pdf.pages)
        return text
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text from {pdf_path}: {str(e)}")
        return ""

def clean_description(description):
    description = description.replace('%', '')
    description = re.sub(r'(?<!\()\b\d+(\.\d+)?\b(?!\))', '', description)
    return re.sub(r'\s{2,}', ' ', description).strip()

def is_valid_entry(description):
    invalid_keywords = ["Invoice", "State", "Model"]
    return not any(keyword in description for keyword in invalid_keywords)

def extract_info(text):
    line_pattern = r"(\d{5}[A-Za-z\d]{5,9}(-[A-Za-z\d]+)?)\s+([^\n]+)"
    date_pattern = r"(\d{2}/\d{2}/\d{4})"
    matches = re.findall(line_pattern, text)
    date_match = re.search(date_pattern, text)
    date = date_match.group(0) if date_match else "Unknown Date"
    part_numbers = []
    descriptions = []
    for match in matches:
        part_number = match[0]
        description = clean_description(match[2])
        if is_valid_entry(description):
            part_numbers.append(part_number)
            descriptions.append(description)
    return date, part_numbers, descriptions

df_global = pd.DataFrame()

def try_parse_date(date_str):
    try:
        return datetime.strptime(date_str, "%d/%m/%Y")
    except:
        return None

def process_pdfs():
    global df_global
    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files", filetypes=[("PDF files", "*.pdf")]
    )
    if not file_paths:
        messagebox.showwarning("No Files", "No PDF files selected!")
        return
    all_data = []
    for pdf_file in file_paths:
        text = extract_text_from_pdf(pdf_file)
        if text:
            date, part_numbers, descriptions = extract_info(text)
            for part_number, description in zip(part_numbers, descriptions):
                all_data.append([date, part_number, description])
    df_global = pd.DataFrame(all_data, columns=['Date', 'Part Number', 'Description'])
    update_treeview(df_global)
    save_button.config(state="normal")
    messagebox.showinfo("Success", "PDFs processed successfully!")

def update_treeview(df):
    for row in treeview.get_children():
        treeview.delete(row)

    # Parse dates and group by year
    df['Date_Parsed'] = df['Date'].apply(try_parse_date)
    df['Year'] = df['Date_Parsed'].apply(lambda d: d.year if d else "Unknown")
    df = df.sort_values(['Year', 'Date_Parsed'])

    prev_year = None
    for _, row in df.iterrows():
        curr_year = row['Year']
        if curr_year != prev_year:
            # Insert a year separator row for a new year block
            treeview.insert("", "end", values=(f"=== {curr_year} ===", "", ""), tags=("year_sep",))
            prev_year = curr_year
        treeview.insert("", "end", values=(row["Date"], row["Part Number"], row["Description"]))

    # Style for year separators
    treeview.tag_configure('year_sep', background='#d0e0ff', font=('Segoe UI', 10, 'bold'))

def save_to_csv():
    file_path = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
    if file_path:
        data = []
        for iid in treeview.get_children():
            values = treeview.item(iid)["values"]
            # Skip separator rows
            if not str(values[0]).startswith("==="):
                data.append(values)
        df = pd.DataFrame(data, columns=["Date", "Part Number", "Description"])
        df.to_csv(file_path, index=False)
        messagebox.showinfo("Success", f"Data saved to {file_path}")

def sort_treeview(by, reverse=False):
    global df_global
    if df_global.empty:
        return
    df_sorted = df_global.copy()
    if by == "Date":
        df_sorted["Date_Parsed"] = df_sorted["Date"].apply(try_parse_date)
        df_sorted = df_sorted.sort_values("Date_Parsed", ascending=not reverse)
        df_sorted = df_sorted.drop("Date_Parsed", axis=1)
    elif by == "Year":
        df_sorted["Year"] = df_sorted["Date"].apply(lambda x: try_parse_date(x).year if try_parse_date(x) else 0)
        df_sorted = df_sorted.sort_values("Year", ascending=not reverse)
        df_sorted = df_sorted.drop("Year", axis=1)
    else:
        df_sorted = df_sorted.sort_values(by, ascending=not reverse)
    update_treeview(df_sorted)

# GUI setup
root = tk.Tk()
root.title("PDF Data Extractor")
root.geometry("900x600")

# Top process/save
frame = tk.Frame(root)
frame.pack(pady=20)
process_button = tk.Button(frame, text="Select and Process PDFs", command=process_pdfs, width=20)
process_button.grid(row=0, column=0, padx=10)
save_button = tk.Button(frame, text="Save to CSV", command=save_to_csv, width=20, state="disabled")
save_button.grid(row=0, column=1, padx=10)

# Sorting options
sort_frame = tk.LabelFrame(root, text="Sort Options")
sort_frame.pack(pady=5)
options = ["Date", "Year", "Part Number", "Description"]
sort_var = tk.StringVar(value="Date")
def do_sort():
    sort_treeview(sort_var.get(), reverse=False)
def do_sort_reverse():
    sort_treeview(sort_var.get(), reverse=True)
sort_dropdown = ttk.Combobox(sort_frame, textvariable=sort_var, values=options, state="readonly", width=12)
sort_dropdown.grid(row=0, column=0, padx=5)
sort_button = tk.Button(sort_frame, text="Sort Asc", command=do_sort)
sort_button.grid(row=0, column=1, padx=3)
sort_button_desc = tk.Button(sort_frame, text="Sort Desc", command=do_sort_reverse)
sort_button_desc.grid(row=0, column=2, padx=3)

treeview = ttk.Treeview(root, columns=("Date", "Part Number", "Description"), show="headings")
treeview.pack(fill="both", expand=True, padx=20, pady=20)
treeview.heading("Date", text="Date")
treeview.heading("Part Number", text="Part Number")
treeview.heading("Description", text="Description")
treeview.column("Date", width=130)
treeview.column("Part Number", width=200)
treeview.column("Description", width=450)

# Make year separator row stand out
style = ttk.Style(root)
style.configure("Treeview", rowheight=24)

root.mainloop()
