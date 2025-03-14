import tkinter as tk
from tkinter import font
from tkinter import ttk
import pandas as pd
import os

def create_application():
    # Create the main application window
    app = tk.Tk()
    app.title("Bookstore Management System")

    # Set the window to Windows fullscreen
    app.state("zoomed")

    # Add a custom font for the title
    title_font = font.Font(family="Helvetica", size=20, weight="bold")

    # Function to handle back button click (close the application)
    def handle_back_button():
        app.destroy()

    # Add the "Back" button at the top-right corner
    back_button = tk.Button(app, text="Back", width=10, command=handle_back_button)
    back_button.pack(anchor="ne", padx=10, pady=10)

    # Add the "BMS" label at the top center
    bms_label = tk.Label(app, text="BMS", font=title_font, anchor="center")
    bms_label.pack(pady=(10, 0))

    # Add the "Bookstore Management System" label just below
    subtitle_label = tk.Label(app, text="Bookstore Management System", anchor="center")
    subtitle_label.pack(pady=(5, 20))

    # Input fields (ordered as Book Title, Author, Summary, Price, and Quantity)
    tk.Label(app, text="Book Title:").pack(pady=(10, 0))
    book_title_entry = tk.Entry(app, width=50)
    book_title_entry.pack(pady=(0, 10))

    tk.Label(app, text="Author:").pack(pady=(10, 0))
    author_entry = tk.Entry(app, width=50)
    author_entry.pack(pady=(0, 10))

    tk.Label(app, text="Summary:").pack(pady=(10, 0))
    summary_entry = tk.Entry(app, width=50)
    summary_entry.pack(pady=(0, 10))

    tk.Label(app, text="Price:").pack(pady=(10, 0))
    price_entry = tk.Entry(app, width=50)
    price_entry.pack(pady=(0, 10))

    tk.Label(app, text="Quantity:").pack(pady=(10, 0))
    quantity_entry = tk.Entry(app, width=50)
    quantity_entry.pack(pady=(0, 20))

    # Path to the Excel file
    script_dir = os.path.dirname(os.path.abspath(__file__))
    excel_file = os.path.join(script_dir, "Inventory.xlsx")

    # Function to load and display the inventory
    def load_inventory():
        # Clear existing items in the Treeview
        for row in inventory_tree.get_children():
            inventory_tree.delete(row)

        # Load inventory data
        if os.path.exists(excel_file):
            data = pd.read_excel(excel_file)
        else:
            data = pd.DataFrame(columns=["Item Number", "Quantity", "Price", "Book Title", "Author", "Summary"])

        # Format the Price column to 2 decimal places
        if not data.empty and "Price" in data.columns:
            data["Price"] = data["Price"].apply(lambda x: f"{x:.2f}" if pd.notnull(x) else "")

        # Populate the Treeview with inventory data
        for _, row in data.iterrows():
            inventory_tree.insert("", "end", values=row.tolist())

    def submit_data():
        # Get input values
        book_title = book_title_entry.get()
        author = author_entry.get()
        summary = summary_entry.get()
        price = price_entry.get()
        quantity = quantity_entry.get()

        # Validate inputs
        if not book_title or not author or not summary or not price or not quantity:
            error_label.config(text="All fields are required.", fg="red")
            return

        try:
            price = float(price)
            quantity = int(quantity)
        except ValueError:
            error_label.config(text="Price and Quantity must be valid numbers.", fg="red")
            return

        # Read or create the Excel file
        if os.path.exists(excel_file):
            data = pd.read_excel(excel_file)
        else:
            data = pd.DataFrame(columns=["Item Number", "Quantity", "Price", "Book Title", "Author", "Summary"])

        # Get the next item number
        next_item_number = 1 if data.empty else data["Item Number"].max() + 1

        # Create the new row
        new_row = {
            "Item Number": next_item_number,
            "Quantity": quantity,
            "Price": price,
            "Book Title": book_title,
            "Author": author,
            "Summary": summary
        }
        data = pd.concat([data, pd.DataFrame([new_row])], ignore_index=True)

        # Save back to the Excel file
        data.to_excel(excel_file, index=False)

        # Clear the input fields
        book_title_entry.delete(0, tk.END)
        author_entry.delete(0, tk.END)
        summary_entry.delete(0, tk.END)
        price_entry.delete(0, tk.END)
        quantity_entry.delete(0, tk.END)

        error_label.config(text="Book added successfully!", fg="green")

        # Reload the inventory
        load_inventory()

    def update_inventory():
        # Get the selected item
        selected_item = inventory_tree.focus()
        if not selected_item:
            error_label.config(text="Select an item to update.", fg="red")
            return

        # Get input values
        book_title = book_title_entry.get()
        author = author_entry.get()
        summary = summary_entry.get()
        price = price_entry.get()
        quantity = quantity_entry.get()

        # Validate inputs
        if not book_title or not author or not summary or not price or not quantity:
            error_label.config(text="All fields are required for updating.", fg="red")
            return

        try:
            price = float(price)
            quantity = int(quantity)
        except ValueError:
            error_label.config(text="Price and Quantity must be valid numbers.", fg="red")
            return

        # Load inventory data
        data = pd.read_excel(excel_file)
        item_number = inventory_tree.item(selected_item)["values"][0]

        # Update the selected row
        data.loc[data["Item Number"] == item_number, ["Quantity", "Price", "Book Title", "Author", "Summary"]] = [
            quantity, price, book_title, author, summary
        ]

        # Save back to the Excel file
        data.to_excel(excel_file, index=False)

        error_label.config(text="Inventory updated successfully!", fg="green")

        # Reload the inventory
        load_inventory()

    # Buttons: Add to Inventory and Update Inventory
    button_frame = tk.Frame(app)
    button_frame.pack(pady=(10, 20))

    submit_button = tk.Button(button_frame, text="Add to Inventory", command=submit_data)
    submit_button.grid(row=0, column=0, padx=20)

    update_button = tk.Button(button_frame, text="Update Inventory", command=update_inventory)
    update_button.grid(row=0, column=1, padx=20)

    # Error/Success message
    error_label = tk.Label(app, text="", fg="red")
    error_label.pack()

    # Inventory section
    tk.Label(app, text="Inventory:").pack(pady=(10, 0))

    # Frame for the inventory table with a scrollbar
    inventory_frame = tk.Frame(app)
    inventory_frame.pack(fill="both", expand=True, pady=(10, 20))

    # Treeview for displaying the inventory
    inventory_tree = ttk.Treeview(inventory_frame, columns=["Item Number", "Quantity", "Price", "Book Title", "Author", "Summary"], show="headings", height=10)
    inventory_tree.pack(side="left", fill="both", expand=True)

    # Define column headings
    for col in ["Item Number", "Quantity", "Price", "Book Title", "Author", "Summary"]:
        inventory_tree.heading(col, text=col)
        inventory_tree.column(col, width=150, anchor="center")

    # Adjust column width for the Summary column to stretch
    inventory_tree.column("Summary", width=300, stretch=True)

    # Add scrollbar to the inventory
    scrollbar = ttk.Scrollbar(inventory_frame, orient="vertical", command=inventory_tree.yview)
    scrollbar.pack(side="right", fill="y")
    inventory_tree.configure(yscrollcommand=scrollbar.set)

    # Load inventory initially
    load_inventory()

    # Start the application
    app.mainloop()

if __name__ == "__main__":
    create_application()
