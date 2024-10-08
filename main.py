import tkinter as tk
from tkinter import messagebox, ttk, filedialog
import sqlite3
from datetime import datetime
import pandas as pd  # Ensure pandas is installed: `pip install pandas`
import xlsxwriter  # Ensure xlsxwriter is installed: `pip install xlsxwriter`

# Database setup
conn = sqlite3.connect('inventory.db')
c = conn.cursor()
c.execute('''CREATE TABLE IF NOT EXISTS items (
                barcode TEXT PRIMARY KEY,
                name TEXT,
                description TEXT,
                location TEXT,
                count INTEGER DEFAULT 0
            )''')
conn.commit()

# Center window function
def center_window(window):
    window.update_idletasks()
    width = window.winfo_width()
    height = window.winfo_height()
    x = (window.winfo_screenwidth() // 2) - (width // 2)
    y = (window.winfo_screenheight() // 2) - (height // 2)
    window.geometry(f'{width}x{height}+{x}+{y}')

# Main app window
class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Inventory Management")
        self.root.geometry("800x600")

        # Top section for controls
        self.top_frame = tk.Frame(root, height=150)
        self.top_frame.pack(fill=tk.X)

        self.label = tk.Label(self.top_frame, text="Scan Barcode:")
        self.label.pack(side=tk.LEFT, padx=10)

        self.barcode_entry = tk.Entry(self.top_frame)
        self.barcode_entry.pack(side=tk.LEFT, padx=10)

        self.scan_button = tk.Button(self.top_frame, text="Scan", command=self.scan_barcode)
        self.scan_button.pack(side=tk.LEFT, padx=5)

        # Buttons
        self.button_frame = tk.Frame(self.top_frame)
        self.button_frame.pack(side=tk.RIGHT)

        self.count_button = tk.Button(self.button_frame, text="Start Count", command=self.open_count_window)
        self.count_button.pack(side=tk.LEFT, padx=5)

        self.add_item_button = tk.Button(self.button_frame, text="Add New Item", command=self.add_item)
        self.add_item_button.pack(side=tk.LEFT, padx=5)

        self.edit_item_button = tk.Button(self.button_frame, text="Edit Selected Item", command=self.edit_item)
        self.edit_item_button.pack(side=tk.LEFT, padx=5)

        # self.remove_item_button = tk.Button(self.button_frame, text="Remove Selected Item", command=self.remove_selected)
        # self.remove_item_button.pack(side=tk.LEFT, padx=5)

        self.print_button = tk.Button(self.button_frame, text="Print to Spreadsheet", command=self.print_to_spreadsheet)
        self.print_button.pack(side=tk.LEFT, padx=5)

        self.refresh_button = tk.Button(self.button_frame, text="Refresh", command=self.refresh_inventory)
        self.refresh_button.pack(side=tk.LEFT, padx=5)

        self.barcode_entry.bind("<Return>", lambda event: self.scan_barcode())

        # Bottom section for inventory view
        self.bottom_frame = tk.Frame(root)
        self.bottom_frame.pack(fill=tk.BOTH, expand=True)

        # Scrollbar
        self.scrollbar = tk.Scrollbar(self.bottom_frame, orient="vertical")
        self.scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Treeview for displaying inventory
        self.tree = ttk.Treeview(self.bottom_frame, columns=('Barcode', 'Name', 'Description', 'Location', 'Count'),
                                 show='headings', yscrollcommand=self.scrollbar.set)
        self.tree.pack(fill=tk.BOTH, expand=True)
        self.scrollbar.config(command=self.tree.yview)

        for col in ('Barcode', 'Name', 'Description', 'Location', 'Count'):
            self.tree.heading(col, text=col)

        # Load initial data
        self.refresh_inventory()

    def refresh_inventory(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        c.execute("SELECT * FROM items")
        for item in c.fetchall():
            self.tree.insert('', 'end', values=item)

    def scan_barcode(self):
        barcode = self.barcode_entry.get()
        c.execute("SELECT * FROM items WHERE barcode=?", (barcode,))
        item = c.fetchone()
        if item:
            messagebox.showinfo("Item Info", f"Name: {item[1]}\nDescription: {item[2]}\nLocation: {item[3]}\nCount: {item[4]}")
        else:
            messagebox.showerror("Error", "Item not found!")
        self.barcode_entry.delete(0, tk.END)  # Clear barcode entry after scanning

    def print_to_spreadsheet(self):
        current_time = datetime.now().strftime("%m.%d.%y_%H.%M.%S")
        default_filename = f"printer_inventory_{current_time}.xlsx"

        file_path = filedialog.asksaveasfilename(initialfile=default_filename, defaultextension=".xlsx", filetypes=[("Excel files", "*.xlsx")])
        if not file_path:
            return

        items = [(self.tree.item(i, "values")) for i in self.tree.get_children()]
        df = pd.DataFrame(items, columns=['Barcode', 'Name', 'Description', 'Location', 'Count'])

        with pd.ExcelWriter(file_path, engine='xlsxwriter') as writer:
            df.to_excel(writer, index=False, sheet_name='Inventory')

            workbook = writer.book
            worksheet = writer.sheets['Inventory']

            worksheet.set_column('A:E', 20)

            header_format = workbook.add_format({'bold': True, 'border': 1, 'align': 'center'})
            cell_format = workbook.add_format({'border': 1})

            for col_num, value in enumerate(df.columns.values):
                worksheet.write(0, col_num, value, header_format)

            for row in range(1, len(df) + 1):
                worksheet.set_row(row, None, cell_format)

        messagebox.showinfo("Success", f"Inventory saved to {file_path}")

    def open_count_window(self):
        count_window = tk.Toplevel(self.root)
        count_window.title("Count Inventory")
        count_window.attributes('-fullscreen', True)

        scanned_barcodes = {}

        def scan_barcode():
            barcode = barcode_entry.get()
            if barcode:
                c.execute("SELECT * FROM items WHERE barcode=?", (barcode,))
                item = c.fetchone()
                if item:
                    scanned_barcodes[barcode] = scanned_barcodes.get(barcode, 0) + 1
                    scanned_list.insert(tk.END, f"{item[1]} (Scanned {scanned_barcodes[barcode]} times)")
                else:
                    messagebox.showerror("Error", "Item not found!")
                barcode_entry.delete(0, tk.END)

        def update_inventory_counts():
            for barcode, count in scanned_barcodes.items():
                c.execute("UPDATE items SET count = count + ? WHERE barcode=?", (count, barcode))
            conn.commit()
            messagebox.showinfo("Success", "Inventory counts updated")
            self.refresh_inventory()
            count_window.destroy()

        def remove_selected():
            try:
                selected_index = scanned_list.curselection()[0]
                scanned_list.delete(selected_index)
            except IndexError:
                messagebox.showerror("Error", "No item selected to remove!")

        barcode_entry = tk.Entry(count_window)
        barcode_entry.pack(pady=5)
        barcode_entry.focus()

        scan_button = tk.Button(count_window, text="Scan", command=scan_barcode)
        scan_button.pack(pady=5)

        scanned_list = tk.Listbox(count_window, width=50)
        scanned_list.pack(pady=5)

        update_button = tk.Button(count_window, text="Update Counts", command=update_inventory_counts)
        update_button.pack(pady=5)

        remove_button = tk.Button(count_window, text="Remove Selected", command=remove_selected)
        remove_button.pack(pady=5)

        count_window.bind('<Return>', lambda event: scan_barcode())

    def add_item(self):
        add_item_window = tk.Toplevel(self.root)
        add_item_window.title("Add New Item")
        add_item_window.geometry("400x300")

        frame = tk.Frame(add_item_window, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Barcode:").pack(pady=5)
        barcode_entry = tk.Entry(frame)
        barcode_entry.pack(pady=5)

        tk.Label(frame, text="Name:").pack(pady=5)
        name_entry = tk.Entry(frame)
        name_entry.pack(pady=5)

        tk.Label(frame, text="Description:").pack(pady=5)
        description_entry = tk.Entry(frame)
        description_entry.pack(pady=5)

        tk.Label(frame, text="Location:").pack(pady=5)
        location_entry = tk.Entry(frame)
        location_entry.pack(pady=5)

        def save_item():
            barcode = barcode_entry.get()
            name = name_entry.get()
            description = description_entry.get()
            location = location_entry.get()
            if barcode and name and description and location:
                try:
                    c.execute("INSERT INTO items (barcode, name, description, location, count) VALUES (?, ?, ?, ?, 0)",
                              (barcode, name, description, location))
                    conn.commit()
                    messagebox.showinfo("Success", "Item added successfully")
                    add_item_window.destroy()
                    self.refresh_inventory()
                except sqlite3.IntegrityError:
                    messagebox.showerror("Error", "Barcode already exists!")
            else:
                messagebox.showerror("Error", "All fields are required!")

        save_button = tk.Button(frame, text="Save", command=save_item)
        save_button.pack(pady=5)

    def edit_item(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showerror("Error", "No item selected!")
            return

        item_values = self.tree.item(selected_item[0], "values")
        edit_item_window = tk.Toplevel(self.root)
        edit_item_window.title("Edit Item")
        edit_item_window.geometry("400x300")

        frame = tk.Frame(edit_item_window, padx=20, pady=20)
        frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(frame, text="Barcode:").pack(pady=5)
        barcode_entry = tk.Entry(frame)
        barcode_entry.pack(pady=5)
        barcode_entry.insert(0, item_values[0])  # Fill with current barcode
        barcode_entry.config(state='readonly')  # Make barcode uneditable

        tk.Label(frame, text="Name:").pack(pady=5)
        name_entry = tk.Entry(frame)
        name_entry.pack(pady=5)
        name_entry.insert(0, item_values[1])

        tk.Label(frame, text="Description:").pack(pady=5)
        description_entry = tk.Entry(frame)
        description_entry.pack(pady=5)
        description_entry.insert(0, item_values[2])

        tk.Label(frame, text="Location:").pack(pady=5)
        location_entry = tk.Entry(frame)
        location_entry.pack(pady=5)
        location_entry.insert(0, item_values[3])

        def update_item():
            name = name_entry.get()
            description = description_entry.get()
            location = location_entry.get()
            if name and description and location:
                c.execute("UPDATE items SET name=?, description=?, location=? WHERE barcode=?",
                          (name, description, location, item_values[0]))
                conn.commit()
                messagebox.showinfo("Success", "Item updated successfully")
                edit_item_window.destroy()
                self.refresh_inventory()
            else:
                messagebox.showerror("Error", "All fields are required!")

        update_button = tk.Button(frame, text="Update", command=update_item)
        update_button.pack(pady=5)

# Initialize the application
root = tk.Tk()
app = InventoryApp(root)
root.mainloop()
