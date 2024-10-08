import sqlite3
import tkinter as tk
from tkinter import messagebox, scrolledtext
import smtplib
import datetime


# Database connection
def connect_db():
    return sqlite3.connect('inventory.db')


# Function to remove an item by barcode
def remove_item(barcode):
    conn = connect_db()
    cursor = conn.cursor()

    # Query to get item details from the 'items' table
    cursor.execute("SELECT name, description, location, count FROM items WHERE barcode = ?", (barcode,))
    item = cursor.fetchone()

    if item:
        name, description, location, count = item
        count -= 1  # Decrease item count
        if count < 0:  # Ensure count doesn't go negative
            messagebox.showerror("Error", "Item count cannot go below zero.")
            conn.close()
            return

        # Update the database
        cursor.execute("UPDATE items SET count = ? WHERE barcode = ?", (count, barcode))
        conn.commit()

        # Log the removal
        log_removal(name, description, location, count)

        # Check if we need to send a message
        if count <= 2:
            send_message(name, description, count)

        messagebox.showinfo("Success", f"Removed {name} from inventory.")
        conn.close()

        # Update the display
        update_display(f"Removed {name} from {location}. Current count: {count}\n")

    else:
        messagebox.showerror("Error", "Item not found.")


# Log the removal time
def log_removal(name, description, location, count):
    with open("removal_log.txt", "a") as log_file:
        log_entry = f"{datetime.datetime.now():%d/%m/%Y@%H:%M:%S} - Removed {name} ({description}) from {location}. Current count: {count}\n"
        log_file.write(log_entry)


# Function to send a message
def send_message(name, description, count):
    if send_message.sent:  # Prevent sending the same message multiple times
        return
    send_message.sent = True

    msg = f"{name} - {description} is at {count}. Please restock this item."

    try:
        server = smtplib.SMTP('mail.garlynshelton.com', 587)
        server.login('printadmin@garlynshelton.com', '$Helton01!')
        server.sendmail('printadmin@garlynshelton.com', '2544217671@mms.att.net', msg)
        server.quit()
    except Exception as e:
        print("Failed to send message:", e)


send_message.sent = False  # Flag to track if a message has been sent


# Function to update the display text box
def update_display(message):
    display_text.config(state=tk.NORMAL)  # Enable editing
    display_text.delete(1.0, tk.END)  # Clear current text
    display_text.insert(tk.END, message)  # Insert new message
    display_text.config(state=tk.DISABLED)  # Disable editing


# Function to process the barcode input
def process_barcode(event=None):
    barcode = barcode_entry.get().strip()
    if barcode:
        remove_item(barcode)
        barcode_entry.delete(0, tk.END)


# GUI Setup
root = tk.Tk()
root.title("Inventory Management")

barcode_label = tk.Label(root, text="Scan Barcode:")
barcode_label.pack()

barcode_entry = tk.Entry(root)
barcode_entry.pack()
barcode_entry.bind("<Return>", process_barcode)  # Bind the Enter key to the process_barcode function

display_text = scrolledtext.ScrolledText(root, width=50, height=10, state=tk.DISABLED)  # Initially disabled
display_text.pack()

barcode_entry.focus()  # Set focus to the barcode entry field

root.mainloop()
