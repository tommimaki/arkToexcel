import tkinter as tk
from tkinter import filedialog
from tkinter import messagebox

# Import your extraction functions here
# from pdf_bot import extract_text_from_pdf, extract_data, write_data_to_excel

# Function to run the extraction process with GUI input


def run_extraction():
    pdf_path = filedialog.askopenfilename(title="Select PDF file",
                                          filetypes=[("PDF files", "*.pdf")])
    if not pdf_path:
        return

    output_path = filedialog.asksaveasfilename(title="Save output as",
                                               defaultextension=".xlsx",
                                               filetypes=[("Excel files", "*.xlsx")])
    if not output_path:
        return

    building_regex = building_regex_var.get()
    floor_regex = floor_regex_var.get()

    if not building_regex or not floor_regex:
        messagebox.showerror(
            "Error", "Please enter regex patterns for buildings and floors.")
        return

    # Update your extraction functions to accept regex patterns as arguments
    # and pass the regex patterns to the functions

    # Run the extraction process
    try:
        text = extract_text_from_pdf(pdf_path)
        data = extract_data(text, building_regex, floor_regex)
        write_data_to_excel(data, output_path)
        messagebox.showinfo(
            "Success", "Data extraction completed successfully!")
    except Exception as e:
        messagebox.showerror(
            "Error", f"An error occurred during extraction: {e}")


# Create the main application window
root = tk.Tk()
root.title("PDF Bot")

# Create input fields for regex patterns
building_regex_label = tk.Label(root, text="Building regex:")
building_regex_label.grid(row=0, column=0, sticky="e")
building_regex_var = tk.StringVar()
building_regex_entry = tk.Entry(root, textvariable=building_regex_var)
building_regex_entry.grid(row=0, column=1)

floor_regex_label = tk.Label(root, text="Floor regex:")
floor_regex_label.grid(row=1, column=0, sticky="e")
floor_regex_var = tk.StringVar()
floor_regex_entry = tk.Entry(root, textvariable=floor_regex_var)
floor_regex_entry.grid(row=1, column=1)

# Create a button to run the extraction process
run_button = tk.Button(root, text="Run extraction", command=run_extraction)
run_button.grid(row=2, column=0, columnspan=2)

# Start the main event loop
root.mainloop()
