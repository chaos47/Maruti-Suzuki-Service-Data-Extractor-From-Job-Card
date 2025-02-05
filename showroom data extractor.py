import re
import pdfplumber
import pandas as pd
import tkinter as tk
from tkinter import filedialog, messagebox
from tkinter import ttk

# Function to extract text from a PDF using PDFPlumber
def extract_text_from_pdf(pdf_path):
    try:
        with pdfplumber.open(pdf_path) as pdf:
            text = "".join(page.extract_text() or "" for page in pdf.pages)
        return text
    except Exception as e:
        messagebox.showerror("Error", f"Failed to extract text from {pdf_path}: {str(e)}")
        return ""

# Function to clean descriptions
def clean_description(description):
    # Remove all occurrences of the '%' character
    description = description.replace('%', '')
    # Remove standalone numbers and decimals unless inside brackets
    description = re.sub(r'(?<!\()\b\d+(\.\d+)?\b(?!\))', '', description)
    # Remove extra spaces caused by cleaning
    return re.sub(r'\s{2,}', ' ', description).strip()

# Function to filter out invalid rows
def is_valid_entry(description):
    # Keywords to drop if found in the description
    invalid_keywords = ["Invoice", "State", "Model"]
    # Check if description starts with or contains any invalid keywords
    return not any(keyword in description for keyword in invalid_keywords)

# Function to extract part numbers, descriptions, and date
def extract_info(text):
    # Regex for matching part numbers and descriptions
    line_pattern = r"(\d{5}[A-Za-z\d]{5,9}(-[A-Za-z\d]+)?)\s+([^\n]+)"
    date_pattern = r"(\d{2}/\d{2}/\d{4})"  # Date in DD/MM/YYYY format

    # Extract part numbers and descriptions
    matches = re.findall(line_pattern, text)

    # Extract the first date found in the text
    date_match = re.search(date_pattern, text)
    date = date_match.group(0) if date_match else "Unknown Date"

    # Filter and clean descriptions
    part_numbers = []
    descriptions = []
    for match in matches:
        part_number = match[0]
        description = clean_description(match[2])  # Clean the description
        if is_valid_entry(description):  # Only keep valid descriptions
            part_numbers.append(part_number)
            descriptions.append(description)

    return date, part_numbers, descriptions

# Function to process selected PDFs and display results in a table
def process_pdfs():
    # File dialog for selecting PDFs
    file_paths = filedialog.askopenfilenames(
        title="Select PDF Files", filetypes=[("PDF files", "*.pdf")]
    )

    if not file_paths:
        messagebox.showwarning("No Files", "No PDF files selected!")
        return

    all_data = []

    # Process each PDF
    for pdf_file in file_paths:
        text = extract_text_from_pdf(pdf_file)
        if text:
            date, part_numbers, descriptions = extract_info(text)
            # Add the PDF date as a heading
            all_data.append([f"PDF Date: {date}", "", ""])  # Add PDF date as a heading
            for part_number, description in zip(part_numbers, descriptions):
                all_data.append([date, part_number, description])
            # Add separator rows after each PDF's data
            all_data.append(["", "", ""])  # Add 1st blank row
            all_data.append(["", "", ""])  # Add 2nd blank row
            all_data.append(["", "", ""])  # Add 3rd blank row

    # Create DataFrame to display the table
    df = pd.DataFrame(all_data, columns=['Date', 'Part Number', 'Description'])

    # Clear existing table if any
    for row in treeview.get_children():
        treeview.delete(row)

    # Insert new data into the table
    for _, row in df.iterrows():
        treeview.insert("", "end", values=(row["Date"], row["Part Number"], row["Description"]))

    # Enable saving option
    save_button.config(state="normal")
    messagebox.showinfo("Success", "PDFs processed successfully!")

# Function to save the data to a CSV file
def save_to_csv():
    file_path = filedialog.asksaveasfilename(
        defaultextension=".csv", filetypes=[("CSV files", "*.csv")]
    )
    if file_path:
        # Retrieve data from the table (treeview)
        data = [treeview.item(row)["values"] for row in treeview.get_children()]

        # Convert to DataFrame and save as CSV
        df = pd.DataFrame(data, columns=["Date", "Part Number", "Description"])
        df.to_csv(file_path, index=False)
        messagebox.showinfo("Success", f"Data saved to {file_path}")

# Create the main window (root)
root = tk.Tk()
root.title("PDF Data Extractor")
root.geometry("800x600")

# Create a frame for the file selection and buttons
frame = tk.Frame(root)
frame.pack(pady=20)

# Button to process PDFs
process_button = tk.Button(frame, text="Select and Process PDFs", command=process_pdfs, width=20)
process_button.grid(row=0, column=0, padx=10)

# Button to save CSV (disabled until data is processed)
save_button = tk.Button(frame, text="Save to CSV", command=save_to_csv, width=20, state="disabled")
save_button.grid(row=0, column=1, padx=10)

# Create a Treeview widget to display the table
treeview = ttk.Treeview(root, columns=("Date", "Part Number", "Description"), show="headings")
treeview.pack(fill="both", expand=True, padx=20, pady=20)

# Define the headings for the table
treeview.heading("Date", text="Date")
treeview.heading("Part Number", text="Part Number")
treeview.heading("Description", text="Description")

# Set column widths
treeview.column("Date", width=100)
treeview.column("Part Number", width=200)
treeview.column("Description", width=400)

# Start the Tkinter main loop
root.mainloop()