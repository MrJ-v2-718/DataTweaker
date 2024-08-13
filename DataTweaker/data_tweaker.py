# MrJ
# CSV Manipulator
# 7/25/2024


import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd


def view_help():
    messagebox.showinfo(
        "Help",
        "Open a CSV file or simply start adding fields to create a new file. "
        "Plenty of useful shortcuts in the File and Edit tabs. "
        "Happy tweaking."
    )


def show_about():
    messagebox.showinfo(
        "About",
        "DataTweaker - A Simple CSV Editor\n"
        f"\tCreated by MrJ\n\t        2024©"
    )


class SimpleCSVEditor:
    def __init__(self, root):
        self.root = root
        self.root.title("DataTweaker")
        self.root.geometry("1024x576")
        self.filepath = None
        self.dataframe = pd.DataFrame()

        self.create_widgets()
        self.update_display()

    def create_widgets(self):
        # Create menu bar
        menu_bar = tk.Menu(self.root)
        file_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="File", menu=file_menu)
        file_menu.add_command(label="New", command=self.new_file, accelerator="Ctrl+N")
        file_menu.add_command(label="Open...", command=self.open_file, accelerator="Ctrl+O")
        file_menu.add_command(label="Save", command=self.save_file, accelerator="Ctrl+S")
        file_menu.add_command(label="Save As...", command=self.save_as_file, accelerator="Ctrl+Shift+S")
        file_menu.add_separator()
        file_menu.add_command(label="Quit", command=self.root.quit, accelerator="Ctrl+Q")

        edit_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Edit", menu=edit_menu)
        edit_menu.add_command(label="Add Field", command=self.add_field, accelerator="Ctrl+F")
        edit_menu.add_command(label="Delete Field", command=self.delete_field, accelerator="Ctrl+Shift+F")
        edit_menu.add_separator()
        edit_menu.add_command(label="Add Row", command=self.add_row_values, accelerator="Ctrl+R")
        edit_menu.add_command(label="Delete Row", command=self.delete_row, accelerator="Ctrl+Shift+R")
        edit_menu.add_separator()
        edit_menu.add_command(label="Add Cell Value", command=self.add_cell_value, accelerator="Ctrl+Shift+C")
        edit_menu.add_command(label="Delete Cell Value", command=self.delete_cell_value, accelerator="Ctrl+Shift+D")

        # Help menu
        help_menu = tk.Menu(menu_bar, tearoff=0)
        menu_bar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="View Help", command=view_help)
        help_menu.add_separator()
        help_menu.add_command(label="About", command=show_about)

        self.root.config(menu=menu_bar)

        # Create line numbers Treeview
        self.line_numbers = ttk.Treeview(
            self.root,
            columns=["Line Number"],
            show="headings",
            height=20,  # Adjust the height to match your Treeview
            selectmode="none"  # Disable row selection
        )
        self.line_numbers.heading("#1", text="Row")
        self.line_numbers.column("#1", width=50, anchor="center")
        self.line_numbers.grid(row=0, column=0, sticky='ns')

        # Create Treeview and scrollbars
        self.tree = ttk.Treeview(self.root, columns=(), show="headings")
        self.tree.grid(row=0, column=1, sticky='nsew')

        self.v_scrollbar = tk.Scrollbar(self.root, orient="vertical", command=self.on_v_scroll)
        self.v_scrollbar.grid(row=0, column=2, sticky='ns')

        self.h_scrollbar = tk.Scrollbar(self.root, orient="horizontal", command=self.tree.xview)
        self.h_scrollbar.grid(row=1, column=1, sticky='ew')

        # Link scrollbars to the treeview
        self.tree.configure(yscrollcommand=self.on_tree_yview, xscrollcommand=self.h_scrollbar.set)

        ttk.Style().configure(
            "Treeview.Heading",
            width=4,
            highlightthickness=0,
            bd=0,
            relief='flat',
            activestyle='none',
            font=("Courier New", 12)
        )

        ttk.Style().configure(
            "Treeview",
            width=4,
            highlightthickness=0,
            bd=0,
            relief='flat',
            activestyle='none',
            font=("Courier New", 12)
        )

        # Configure grid weights to make sure the Treeview expands properly
        self.root.grid_rowconfigure(0, weight=1)
        self.root.grid_columnconfigure(1, weight=1)

        # Ensure the Treeview scrolls with the Line Numbers Treeview
        self.tree.bind("<Configure>", self.sync_scroll)

        # Bind shortcuts
        self.root.bind_all('<Control-n>', self.new_file_shortcut)
        self.root.bind_all('<Control-o>', self.open_file_shortcut)
        self.root.bind_all('<Control-s>', self.save_file_shortcut)
        self.root.bind_all('<Control-Shift-S>', self.save_as_file_shortcut)
        self.root.bind_all('<Control-q>', self.quit_shortcut)

        self.root.bind_all('<Control-f>', self.add_field_shortcut)
        self.root.bind_all('<Control-Shift-F>', self.delete_field_shortcut)

        self.root.bind_all('<Control-r>', self.add_row_shortcut)
        self.root.bind_all('<Control-Shift-R>', self.delete_row_shortcut)

        self.root.bind_all('<Control-Shift-C>', self.add_cell_shortcut)
        self.root.bind_all('<Control-Shift-D>', self.delete_cell_shortcut)

    def new_file_shortcut(self, event):
        self.new_file()

    def open_file_shortcut(self, event):
        self.open_file()

    def save_file_shortcut(self, event):
        self.save_file()

    def save_as_file_shortcut(self, event):
        self.save_as_file()

    def quit_shortcut(self, event):
        self.root.quit()

    def add_field_shortcut(self, event):
        self.add_field()

    def delete_field_shortcut(self, event):
        self.delete_field()

    def add_row_shortcut(self, event):
        self.add_row_values()

    def delete_row_shortcut(self, event):
        self.delete_row()

    def add_cell_shortcut(self, event):
        self.add_cell_value()

    def delete_cell_shortcut(self, event):
        self.delete_cell_value()

    def sync_scroll(self, event=None):
        # Ensure the line numbers match the visible rows
        self.line_numbers.yview_moveto(self.tree.yview()[0])
        # Update line numbers to account for the heading height
        self.update_line_numbers()

    def on_tree_yview(self, *args):
        self.v_scrollbar.set(*args)
        self.line_numbers.yview_moveto(args[0])
        self.sync_scroll()

    def on_v_scroll(self, *args):
        self.tree.yview(*args)
        self.line_numbers.yview(*args)
        self.sync_scroll()

    def open_file(self):
        filepath = filedialog.askopenfilename(filetypes=[("CSV files", "*.csv")])
        if filepath:
            self.filepath = filepath
            self.load_csv()

    def load_csv(self):
        try:
            self.dataframe = pd.read_csv(self.filepath)
            self.update_display()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load file: {e}")

    def save_file(self):
        if self.filepath:
            self.dataframe.to_csv(self.filepath, index=False)
            messagebox.showinfo("Save File", "File saved successfully!")
        else:
            self.save_as_file()

    def save_as_file(self):
        filepath = filedialog.asksaveasfilename(defaultextension=".csv", filetypes=[("CSV files", "*.csv")])
        if filepath:
            self.filepath = filepath
            self.save_file()

    def update_display(self):
        # Replace NaN values with empty strings
        self.dataframe = self.dataframe.fillna("")

        self.tree["columns"] = list(self.dataframe.columns)
        for col in self.tree["columns"]:
            self.tree.heading(col, text=col)

        for row in self.tree.get_children():
            self.tree.delete(row)

        for i, row in self.dataframe.iterrows():
            self.tree.insert("", "end", values=tuple(row))

        self.update_line_numbers()

    def update_line_numbers(self):
        # Clear the Treeview
        for item in self.line_numbers.get_children():
            self.line_numbers.delete(item)

        # Insert line numbers
        for i in range(len(self.dataframe)):
            self.line_numbers.insert("", "end", values=(str(i + 1),))

        # Ensure the line numbers match the visible rows
        self.line_numbers.yview_moveto(self.tree.yview()[0])

    def add_field(self):
        field_name = self.ask_string("Add Field", "Enter new field name:")
        if field_name:
            if field_name in self.dataframe.columns:
                messagebox.showwarning("Field Exists", f"The field '{field_name}' already exists.")
            else:
                self.dataframe[field_name] = ""
                self.update_display()

    def add_row_values(self):
        new_row = {}

        for col in self.dataframe.columns:
            text = self.ask_string(f"Column '{col}'", "Enter row value:")
            if text is not None:
                new_row[col] = text

        if new_row:
            self.dataframe = self.dataframe._append(new_row, ignore_index=True)
            self.update_display()
        else:
            print("No values provided. Row not added.")

    def add_cell_value(self):
        column = self.ask_string("Add Cell Value", "Enter column name:")
        if column in self.dataframe.columns:
            row_index = self.ask_integer("Add Cell Value", "Enter row number:")
            row_index = row_index - 1
            if row_index is not None and 0 <= row_index < len(self.dataframe):
                value = self.ask_string("Add Cell Value", "Enter value:")
                if value is not None:
                    self.dataframe.at[row_index, column] = value
                    self.update_display()
            else:
                messagebox.showerror("Error", "Invalid row number.")
        else:
            messagebox.showerror("Error", "Invalid column name.")

    def delete_cell_value(self):
        column = self.ask_string("Delete Cell Value", "Enter column name:")
        if column in self.dataframe.columns:
            row_index = self.ask_integer("Delete Cell Value", "Enter row number:")
            row_index = row_index - 1
            if row_index is not None and 0 <= row_index < len(self.dataframe):
                self.dataframe.at[row_index, column] = ""
                self.update_display()
            else:
                messagebox.showerror("Error", "Invalid row number.")
        else:
            messagebox.showerror("Error", "Invalid column name.")

    def delete_field(self):
        field_name = self.ask_string("Delete Field", "Enter field name:")
        if field_name in self.dataframe.columns:
            del self.dataframe[field_name]
            self.update_display()
        else:
            messagebox.showerror("Error", "Field not found.")

    def delete_row(self):
        row_index = self.ask_integer("Delete Row", "Enter row number:")
        row_index = row_index - 1
        if row_index is not None and 0 <= row_index < len(self.dataframe):
            self.dataframe = self.dataframe.drop(index=(row_index - 1)).reset_index(drop=True)
            self.update_display()
        else:
            messagebox.showerror("Error", "Invalid row number.")

    def ask_string(self, title, prompt):
        dialog = CustomDialog(self.root, title, prompt, entry_type="string")
        return dialog.result

    def ask_integer(self, title, prompt):
        dialog = CustomDialog(self.root, title, prompt, entry_type="integer")
        return dialog.result

    def new_file(self):
        self.filepath = None
        self.dataframe = pd.DataFrame()
        self.update_display()


class CustomDialog(tk.Toplevel):
    def __init__(self, parent, title, prompt, entry_type="string"):
        super().__init__(parent)
        self.title(title)
        self.transient(parent)  # Make the dialog transient (stays on top of the parent)
        self.grab_set()  # Grab all input events

        self.result = None
        self.entry_type = entry_type

        tk.Label(self, text=prompt).grid(row=0, column=0, padx=10, pady=5, columnspan=2)
        self.entry = tk.Entry(self)
        self.entry.grid(row=1, column=0, padx=10, pady=5, columnspan=2)
        self.entry.focus_set()

        tk.Button(self, text="OK", command=self.ok).grid(row=2, column=0, padx=10, pady=10)
        tk.Button(self, text="Cancel", command=self.cancel).grid(row=2, column=1, padx=10, pady=10)

        self.protocol("WM_DELETE_WINDOW", self.cancel)
        self.wait_window(self)

    def ok(self):
        if self.entry_type == "integer":
            try:
                self.result = int(self.entry.get())
            except ValueError:
                messagebox.showerror("Invalid Input", "Please enter a valid integer.")
                return
        else:
            self.result = self.entry.get()
        self.destroy()

    def cancel(self):
        self.result = None
        self.destroy()


if __name__ == "__main__":
    root = tk.Tk()
    app = SimpleCSVEditor(root)
    root.mainloop()
