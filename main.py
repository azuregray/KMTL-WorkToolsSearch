import tkinter as tk
from tkinter import filedialog, Toplevel
import pandas as pd

class ExcelApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Kennametal Data Search")
        self.root.geometry("800x600")
        self.root.resizable(True, True)  # Make the main window resizable

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

        self.upload_button = tk.Button(self.left_frame, text="Upload Excel File", command=self.upload_file, width=25, font=("Helvetica", 12))
        self.upload_button.pack(pady=15)

        self.value_entry_frame = tk.LabelFrame(self.left_frame, text="Search Criteria", padx=10, pady=10, font=("Helvetica", 14, "bold"))
        self.value_entry_frame.pack(pady=15, fill=tk.BOTH, expand=True)

        self.search_button = tk.Button(self.left_frame, text="Search", command=self.search_material, width=25, font=("Helvetica", 12))
        self.search_button.pack(pady=15)

        self.reset_button = tk.Button(self.left_frame, text="Reset", command=self.reset_search, width=25, bg="#d32f2f", font=("Helvetica", 12))
        self.reset_button.pack(pady=15)

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

    def upload_file(self):
        file_path = filedialog.askopenfilename(
            title="Select an Excel File",
            filetypes=(("Excel files", "*.xlsx *.xls"), ("All files", "*.*"))
        )
        if file_path:
            try:
                self.df = pd.read_excel(file_path)
                self.columns = list(self.df.columns)
                self.create_column_selection()
            except Exception as e:
                print(f"Failed to load Excel file: {e}")

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

    def reset_value(self, column):
        from_entry, to_entry = self.entries[column]
        from_entry.delete(0, tk.END)
        to_entry.delete(0, tk.END)

    def remove_entry(self, column):
        if column in self.entries:
            del self.entries[column]
            self.selected_columns.remove(column)
            self.column_vars[column].set(False)  # Uncheck the corresponding checkbox
            self.create_entry_fields()

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
                                result_df = result_df[result_df[column].astype(float) >= float(from_value)]
                            if to_value:
                                result_df = result_df[result_df[column].astype(float) <= float(to_value)]
                        else:
                            if from_value:
                                result_df = result_df[result_df[column].astype(str).str.contains(from_value, na=False, case=False)]
                            if to_value:
                                result_df = result_df[result_df[column].astype(str).str.contains(to_value, na=False, case=False)]

                    self.search_results = result_df
                    self.display_results(result_df)
                except KeyError as e:
                    print(f"Column '{e}' does not exist in the DataFrame.")
                except Exception as e:
                    print(f"Error during search: {e}")
            else:
                print("Please select at least one column and enter values to search.")
        else:
            print("Please upload an Excel file first.")

    def display_results(self, result_df):
        results_window = Toplevel(self.root)
        results_window.title("Search Results")
        results_window.geometry("500x400")

        result_text = tk.Text(results_window, wrap=tk.WORD, font=("Helvetica", 12))
        result_text.pack(expand=True, fill=tk.BOTH)

        if not result_df.empty:
            result_str = result_df.to_string(index=False)
            num_rows = len(result_df)
            material_numbers = result_df.iloc[:, 1].tolist()
            result_text.insert(tk.END, f"Total rows found: {num_rows}\n\nMaterials:\n")
            for number in material_numbers:
                result_text.insert(tk.END, f"{number}\n")
        else:
            result_text.insert(tk.END, "No results found.")

    def reset_search(self):
        self.selected_columns = []
        self.search_values = {}
        for widget in self.value_entry_frame.winfo_children():
            widget.destroy()
        for var in self.column_vars.values():
            var.set(False)
        print("Search criteria and results have been reset.")

if __name__ == "__main__":
    root = tk.Tk()
    app = ExcelApp(root)
    root.mainloop()
