import tkinter as tk
from tkinter import filedialog, messagebox, Toplevel
import pandas as pd


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kennametal Data Search")
        self.root.geometry("600x800")
        self.root.resizable(False, False)

        try:
            self.root.wm_iconbitmap('logo.ico')
        except Exception as e:
            print("Icon file not found:", e)

        self.df = None
        self.columns = []
        self.column_vars = {}
        self.selected_columns = []
        self.search_values = {}
        self.search_results = None

        self.create_widgets()

    def create_widgets(self):
        title_label = tk.Label(self.root, text="Kennametal Data Search", font=("Helvetica", 18, "bold"), fg="#333333")
        title_label.pack(pady=20)

        self.upload_button = tk.Button(self.root, text="Upload Excel File", command=self.upload_file, width=25, bg="#4CAF50", fg="white", font=("Helvetica", 12))
        self.upload_button.pack(pady=15)

        self.filter_button = tk.Button(self.root, text="Filter Parameters", command=self.open_checkbox_window, width=25, bg="#2196F3", fg="white", font=("Helvetica", 12))
        self.filter_button.pack(pady=15)

        self.value_entry_frame = tk.LabelFrame(self.root, text="Search Criteria", padx=10, pady=10, font=("Helvetica", 14, "bold"))
        self.value_entry_frame.pack(pady=15, fill=tk.BOTH, expand=True)

        self.search_button = tk.Button(self.root, text="Search", command=self.search_material, width=25, bg="#f57c00", fg="white", font=("Helvetica", 12))
        self.search_button.pack(pady=15)

        self.reset_button = tk.Button(self.root, text="Reset", command=self.reset_search, width=25, bg="#d32f2f", fg="white", font=("Helvetica", 12))
        self.reset_button.pack(pady=15)

        result_label = tk.Label(self.root, text="Results", font=("Helvetica", 16, "bold"))
        result_label.pack(pady=10)

        self.result_text = tk.Text(self.root, height=15, width=70, wrap=tk.WORD, font=("Helvetica", 12), bg="#f1f1f1")
        self.result_text.pack(pady=10)
        self.result_text.config(state=tk.DISABLED)

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.columns = list(self.df.columns)
                self.open_checkbox_window()
                messagebox.showinfo("Success", "Excel file loaded successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load Excel file: {e}")

    def open_checkbox_window(self):
        if not self.columns:
            messagebox.showwarning("File Error", "Please upload an Excel file first.")
            return

        checkbox_window = Toplevel(self.root)
        checkbox_window.title("Select Columns")
        checkbox_window.geometry("300x400")
        checkbox_window.lift()

        canvas = tk.Canvas(checkbox_window)
        scrollbar = tk.Scrollbar(checkbox_window, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        for column in self.columns:
            if column not in self.selected_columns:  # Adjust this condition to filter out unnecessary columns
                var = tk.BooleanVar(value=False)
                checkbox = tk.Checkbutton(scrollable_frame, text=column, variable=var, command=self.update_selected_columns, font=("Helvetica", 12))
                checkbox.pack(anchor=tk.W, pady=2)
                self.column_vars[column] = var

    def update_selected_columns(self):
        self.selected_columns = [col for col, var in self.column_vars.items() if var.get()]
        self.create_entry_fields()

    def create_entry_fields(self):
        for widget in self.value_entry_frame.winfo_children():
            widget.destroy()

        self.entries = {}
        for column in self.selected_columns:
            label = tk.Label(self.value_entry_frame, text=column, font=("Helvetica", 12))
            label.pack(anchor=tk.W, pady=2)
            from_entry = tk.Entry(self.value_entry_frame, font=("Helvetica", 12), bd=2, relief="solid")
            from_entry.pack(anchor=tk.W, fill=tk.X, pady=2)
            to_entry = tk.Entry(self.value_entry_frame, font=("Helvetica", 12), bd=2, relief="solid")
            to_entry.pack(anchor=tk.W, fill=tk.X, pady=2)
            self.entries[column] = (from_entry, to_entry)

    def is_numeric_column(self, column):
        try:
            self.df[column].astype(float) # type: ignore
            return True
        except ValueError:
            return False

    def search_material(self):
        if self.df is not None:
            self.search_values = {}

            for column, (from_entry, to_entry) in self.entries.items():
                from_value = from_entry.get()
                to_value = to_entry.get()
                if from_value.strip() or to_value.strip():
                    self.search_values[column] = (from_value, to_value)

            if self.selected_columns and self.search_values:
                try:
                    result_df = self.df.copy()
                    for column, (from_value, to_value) in self.search_values.items():
                        if self.is_numeric_column(column):
                            if from_value:
                                try:
                                    result_df = result_df[result_df[column].astype(float) >= float(from_value)]
                                except ValueError:
                                    messagebox.showerror("Value Error", f"Column '{column}' contains non-numeric values.")
                                    return
                            if to_value:
                                try:
                                    result_df = result_df[result_df[column].astype(float) <= float(to_value)]
                                except ValueError:
                                    messagebox.showerror("Value Error", f"Column '{column}' contains non-numeric values.")
                                    return
                        else:
                            if from_value:
                                result_df = result_df[result_df[column].astype(str).str.contains(from_value, na=False, case=False)]
                            if to_value:
                                result_df = result_df[result_df[column].astype(str).str.contains(to_value, na=False, case=False)]

                    self.search_results = result_df
                    self.display_results(result_df)
                except KeyError as e:
                    messagebox.showerror("Column Error", f"Column '{e}' does not exist in the DataFrame.")
                except Exception as e:
                    messagebox.showerror("Search Error", f"Error during search: {e}")
            else:
                messagebox.showwarning("Input Error", "Please select at least one column and enter values to search.")
        else:
            messagebox.showwarning("File Error", "Please upload an Excel file first.")

    def display_results(self, result_df):
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete('1.0', tk.END)
        if self.search_values:
            self.result_text.insert(tk.END, "Search Values:\n")
            for col, (from_val, to_val) in self.search_values.items():
                self.result_text.insert(tk.END, f"{col} - From: {from_val}, To: {to_val}\n")
            self.result_text.insert(tk.END, "\n")

        if not result_df.empty:
            result_str = result_df.to_string(index=False)
            num_rows = len(result_df)
            material_numbers = result_df.iloc[:, 1].tolist()
            self.result_text.insert(tk.END, f"Total rows found: {num_rows}\n\n")
            self.result_text.insert(tk.END, "Materials:\n")
            for number in material_numbers:
                self.result_text.insert(tk.END, f"{number}\n")
        else:
            self.result_text.insert(tk.END, "No results found.")
        self.result_text.config(state=tk.DISABLED)

    def reset_search(self):
        self.selected_columns = []
        self.search_values = {}
        self.result_text.config(state=tk.NORMAL)
        self.result_text.delete('1.0', tk.END)
        self.result_text.config(state=tk.DISABLED)
        for widget in self.value_entry_frame.winfo_children():
            widget.destroy()
        for var in self.column_vars.values():
            var.set(False)
        messagebox.showinfo("Reset", "Search criteria and results have been reset.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
