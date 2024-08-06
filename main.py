import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
import os
import re


class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kennametal Data Search")
        self.root.geometry("1000x600")
        self.root.resizable(True, True)

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
        self.main_frame = tk.Frame(self.root)
        self.main_frame.pack(fill=tk.BOTH, expand=True)

        self.left_frame = tk.Frame(self.main_frame)
        self.left_frame.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        self.right_frame = tk.Frame(self.main_frame, width=int(self.root.winfo_screenwidth() * 0.3))
        self.right_frame.pack(side=tk.RIGHT, fill=tk.Y)

        title_label = tk.Label(self.left_frame, text="Kennametal Data Search", font=("Helvetica", 18, "bold"), fg="#333333")
        title_label.pack(pady=20)

        self.upload_button = tk.Button(self.left_frame, text="Upload and Clean Excel File", command=self.upload_and_clean_file, width=30, font=("Helvetica", 12))
        self.upload_button.pack(pady=15)

        self.value_entry_frame = tk.LabelFrame(self.left_frame, text="Search Criteria", padx=10, pady=10, font=("Helvetica", 14, "bold"))
        self.value_entry_frame.pack(pady=15, fill=tk.BOTH, expand=True)

        self.search_button = tk.Button(self.left_frame, text="Search", command=self.search_material, width=25, font=("Helvetica", 12))
        self.search_button.pack(pady=15)

        self.reset_button = tk.Button(self.left_frame, text="Reset", command=self.reset_search, width=25, bg="#d32f2f", font=("Helvetica", 12))
        self.reset_button.pack(pady=15)

        self.results_frame = tk.Frame(self.left_frame)
        self.results_frame.pack(pady=20, fill=tk.BOTH, expand=True)

        self.results_text = tk.Text(self.results_frame, wrap=tk.WORD, font=("Helvetica", 12))
        self.results_text.pack(expand=True, fill=tk.BOTH)

        self.create_column_selection()

    def create_column_selection(self):
        if not self.columns:
            return

        canvas = tk.Canvas(self.right_frame)
        scrollbar = tk.Scrollbar(self.right_frame, orient="vertical", command=canvas.yview)
        self.scrollable_frame = tk.Frame(canvas)

        self.scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=self.scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        canvas.bind_all("<MouseWheel>", lambda event: canvas.yview_scroll(int(-1 * (event.delta / 120)), "units"))

        self.filter_button = tk.Button(self.right_frame, text="Add Filter", command=self.update_selected_columns, width=25, font=("Helvetica", 12))
        self.filter_button.pack(pady=15)

    def upload_and_clean_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            try:
                # Clean and load the Excel file
                self.df = self.clean_excel_file(file_path)
                self.columns = list(self.df.columns)
                self.create_column_selection()
                messagebox.showinfo("Success", f"Data cleaned and loaded from: {os.path.basename(file_path)}")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to process the Excel file: {e}")

    def clean_excel_file(self, file_path):
        try:
            if file_path.endswith('.xls'):
                df = pd.read_excel(file_path, engine='xlrd')
            else:
                df = pd.read_excel(file_path, engine='openpyxl')

            # Clean data
            df = self.clean_data(df)

            # Save cleaned data
            cleaned_file_path = file_path.rsplit('.', 1)[0] + "_cleaned.xlsx"
            df.to_excel(cleaned_file_path, index=False, na_rep='')  # na_rep='' keeps cells empty
            return df
        except Exception as e:
            raise Exception(f"Error cleaning Excel file: {e}")

    def clean_data(self, df):
        for column in df.columns:
            if df[column].dtype == 'object':
                df[column] = df[column].apply(self.clean_value)
        return df

    def clean_value(self, value):
        if isinstance(value, str):
            if re.match(r'^\d', value):
                return re.sub(r'[^\d.]', '', value)
            return value
        return value

    def update_selected_columns(self):
        for widget in self.scrollable_frame.winfo_children():
            widget.destroy()

        for column in self.columns:
            if column not in self.selected_columns:
                var = tk.BooleanVar(value=False)
                checkbox = tk.Checkbutton(self.scrollable_frame, text=column, variable=var, command=self.create_entry_fields, font=("Helvetica", 12))
                checkbox.pack(anchor=tk.W, pady=2)
                self.column_vars[column] = var

    def create_entry_fields(self):
        self.selected_columns = [col for col, var in self.column_vars.items() if var.get()]

        for widget in self.value_entry_frame.winfo_children():
            widget.destroy()

        self.entries = {}
        for column in self.selected_columns:
            if self.is_numeric_column(column):
                frame = tk.Frame(self.value_entry_frame)
                frame.pack(anchor=tk.W, fill=tk.X, pady=2)

                label = tk.Label(frame, text=column, font=("Helvetica", 12))
                label.pack(side=tk.LEFT)

                from_entry = tk.Entry(frame, font=("Helvetica", 12), bd=2, relief="solid", width=10)
                from_entry.pack(side=tk.LEFT, padx=5)

                to_entry = tk.Entry(frame, font=("Helvetica", 12), bd=2, relief="solid", width=10)
                to_entry.pack(side=tk.LEFT, padx=5)

                reset_button = tk.Button(frame, text="Reset Value", command=lambda col=column: self.reset_value(col), font=("Helvetica", 10))
                reset_button.pack(side=tk.LEFT, padx=5)

                remove_button = tk.Button(frame, text="Remove", command=lambda col=column: self.remove_entry(col), font=("Helvetica", 10))
                remove_button.pack(side=tk.LEFT, padx=5)

                self.entries[column] = (from_entry, to_entry)
            else:
                frame = tk.Frame(self.value_entry_frame)
                frame.pack(anchor=tk.W, fill=tk.X, pady=2)

                label = tk.Label(frame, text=column, font=("Helvetica", 12))
                label.pack(side=tk.LEFT)

                dropdown_values = list(self.df[column].dropna().unique())
                dropdown = tk.StringVar()
                dropdown_menu = tk.OptionMenu(frame, dropdown, *dropdown_values)
                dropdown_menu.pack(side=tk.LEFT, padx=5)

                remove_button = tk.Button(frame, text="Remove", command=lambda col=column: self.remove_entry(col), font=("Helvetica", 10))
                remove_button.pack(side=tk.LEFT, padx=5)

                self.entries[column] = dropdown

    def reset_value(self, column):
        if self.is_numeric_column(column):
            from_entry, to_entry = self.entries[column]
            from_entry.delete(0, tk.END)
            to_entry.delete(0, tk.END)
        else:
            dropdown = self.entries[column]
            dropdown.set('')

    def remove_entry(self, column):
        if column in self.entries:
            del self.entries[column]
            self.selected_columns.remove(column)
            self.column_vars[column].set(False)  # Uncheck the corresponding checkbox
            self.create_entry_fields()

    def is_numeric_column(self, column):
        try:
            self.df[column].astype(float)
            return True
        except ValueError:
            return False

    def search_material(self):
        if self.df is not None:
            self.search_values = {}

            for column, entry in self.entries.items():
                if self.is_numeric_column(column):
                    from_value, to_value = entry
                    from_value = from_value.get()
                    to_value = to_value.get()
                    if from_value.strip() or to_value.strip():
                        self.search_values[column] = (from_value, to_value)
                else:
                    dropdown = entry
                    selected_value = dropdown.get()
                    if selected_value:
                        self.search_values[column] = selected_value

            if self.selected_columns and self.search_values:
                try:
                    result_df = self.df.copy()
                    for column, value in self.search_values.items():
                        if self.is_numeric_column(column):
                            from_value, to_value = value
                            from_value = float(from_value) if from_value.strip() else float('-inf')
                            to_value = float(to_value) if to_value.strip() else float('inf')
                            result_df = result_df[(result_df[column].astype(float) >= from_value) & (result_df[column].astype(float) <= to_value)]
                        else:
                            result_df = result_df[result_df[column] == value]

                    self.search_results = result_df
                    self.display_results()
                except Exception as e:
                    messagebox.showerror("Error", f"Search failed: {e}")
            else:
                messagebox.showinfo("Info", "Please select columns and provide search criteria.")
        else:
            messagebox.showinfo("Info", "No data available. Please upload a file first.")

    def display_results(self):
        self.results_text.delete("1.0", tk.END)
        if self.search_results is not None and not self.search_results.empty:
            num_rows = len(self.search_results)
            material_numbers = self.search_results.iloc[:, 1].tolist()  # Assuming material numbers are in the second column
            self.results_text.insert(tk.END, f"Total rows found: {num_rows}\n\nMaterials:\n")
            for number in material_numbers:
                self.results_text.insert(tk.END, f"{number}\n")
        else:
            self.results_text.insert(tk.END, "No results found.")


    def reset_search(self):
        for widget in self.value_entry_frame.winfo_children():
            widget.destroy()
        self.entries = {}
        self.selected_columns = []
        self.search_values = {}
        self.search_results = None
        self.results_text.delete("1.0", tk.END)
        self.update_selected_columns()


if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
