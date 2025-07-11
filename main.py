import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
import os
from datetime import datetime
import csv
import pandas as pd
import sys

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # Get the user's documents folder
        user_docs = os.path.expanduser('~/Documents')
        app_data_dir = os.path.join(user_docs, 'CigarInventory')
        
        # Create the directory if it doesn't exist
        if not os.path.exists(app_data_dir):
            os.makedirs(app_data_dir)
        
        # If it's a data file (json), use the app data directory
        if relative_path.endswith('.json'):
            return os.path.join(app_data_dir, relative_path)
            
        # For other resources, use PyInstaller's temp folder or current directory
        try:
            base_path = sys._MEIPASS
        except Exception:
            base_path = os.path.abspath(".")
        return os.path.join(base_path, relative_path)
    except Exception as e:
        print(f"Error in resource_path: {e}")
        return relative_path

class CigarInventory:
    def __init__(self, root):
        self.root = root
        self.root.title("Cigar Inventory Manager")
        self.root.geometry("1500x900")  # Made taller for calculator
        
        # Set window icon
        try:
            # Get the absolute path to the icon file
            if getattr(sys, 'frozen', False):
                application_path = sys._MEIPASS
            else:
                application_path = os.path.dirname(os.path.abspath(__file__))
            
            icon_path = os.path.join(application_path, 'cigar.ico')
            
            # Only try to set the icon if the file exists
            if os.path.exists(icon_path):
                try:
                    self.root.iconbitmap(icon_path)
                except:
                    # If iconbitmap fails, try alternative method
                    self.root.iconbitmap(default=icon_path)
            else:
                print(f"Icon file not found at: {icon_path}")
                
        except Exception as e:
            print(f"Note: Application will run without custom icon: {e}")
        
        # Data storage
        self.inventory = []
        self.brands = set()
        self.sizes = set()
        self.types = set()
        self.tax_rate = 0.086
        self.checkbox_states = {}
        self.sales_history = []
        self.stored_quantities = {}  # New dictionary to store quantities persistently
        
        # Create notebook for tabs
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Create main frame for inventory tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text='Inventory')
        
        # Create frame for sales history tab
        self.sales_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.sales_frame, text='Sales History')
        
        # Setup UI
        self.setup_inventory_tab()
        self.setup_sales_history_tab()
        
        # Load and display inventory after UI is ready
        self.load_inventory()
        self.refresh_inventory()
        self.refresh_sales_history()
        
        # Ensure proper cleanup
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_inventory_tab(self):
        # Main container
        main_frame = ttk.Frame(self.main_frame)
        main_frame.pack(expand=True, fill='both', padx=5, pady=5)
        
        # Create a horizontal paned window to allow resizing between inventory and buttons
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(expand=True, fill='both')
        
        # Left frame (search and treeview)
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=3)  # Give more weight to the inventory side
        
        # Right frame (buttons and calculators)
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)  # Give less weight to the button side
        
        # Search frame
        search_frame = ttk.Frame(left_frame)
        search_frame.pack(fill='x', pady=(0, 5))
        ttk.Label(search_frame, text="Search:").pack(side='left', padx=5)
        self.search_var = tk.StringVar()
        self.search_var.trace_add('write', self.on_search)
        ttk.Entry(search_frame, textvariable=self.search_var).pack(side='left', fill='x', expand=True)
        
        # Create a frame to hold both treeview and totals
        tree_container = ttk.Frame(left_frame)
        tree_container.pack(expand=True, fill='both')
        
        # Create frame for treeview
        tree_frame = ttk.Frame(tree_container)
        tree_frame.pack(expand=True, fill='both')
        
        # Treeview
        columns = ('select', 'brand', 'cigar', 'size', 'type', 'count', 'price', 'shipping', 
                  'per_stick', 'personal_rating')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
        
        # Add scrollbars
        tree_scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=self.tree.yview)
        tree_xscrollbar = ttk.Scrollbar(tree_frame, orient='horizontal', command=self.tree.xview)
        self.tree.configure(yscrollcommand=tree_scrollbar.set, xscrollcommand=tree_xscrollbar.set)
        
        # Pack scrollbars and treeview
        tree_scrollbar.pack(side='right', fill='y')
        tree_xscrollbar.pack(side='bottom', fill='x')
        self.tree.pack(side='left', expand=True, fill='both')
        
        # Add mouse wheel scrolling for inventory
        def on_mousewheel(event):
            self.tree.yview_scroll(int(-1 * (event.delta / 120)), "units")
        self.tree.bind('<MouseWheel>', on_mousewheel)
        
        # Add totals frame at the bottom of the left frame
        self.setup_totals_frame(tree_container)
        
        # Store checkbox states
        self.checkbox_states = {}
        
        # Sorting setup
        self.sort_column = None
        self.sort_reverse = {}
        
        # Configure headings and columns
        headings = {
            'select': '✓',
            'brand': 'Brand',
            'cigar': 'Cigar',
            'size': 'Size',
            'type': 'Type',
            'count': 'Count',
            'price': 'Price ($)',
            'shipping': 'Shipping ($)',
            'per_stick': 'Price/Stick',
            'personal_rating': 'My Rating'
        }
        
        # Column widths
        widths = {
            'select': 30,  # Reduced from 50
            'brand': 150,  # Reduced from 200
            'cigar': 200,  # Keep this the same as it needs space for names
            'size': 70,   # Reduced from 100
            'type': 100,  # Keep this the same
            'count': 50,  # Reduced from 70
            'price': 80,  # Reduced from 100
            'shipping': 80,  # Reduced from 100
            'per_stick': 80,  # Reduced from 100
            'personal_rating': 70  # Reduced from 100
        }
        
        for col, heading in headings.items():
            self.tree.heading(col, text=heading,
                            command=lambda c=col: self.sort_treeview(c))
            self.tree.column(col, width=widths[col])
        
        # Create a frame for buttons that will reflow
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill='x', pady=5)
        
        # Button definitions
        buttons = [
            ("Save", self.manual_save),
            ("Export", self.export_inventory),
            ("Add New", self.add_new_line),
            ("Remove Selected", self.remove_selected),
            ("Add Brand", self.add_new_brand),
            ("Add Size", self.add_new_size),
            ("Add Type", self.add_new_type)
        ]
        
        # Create buttons in a grid layout that will reflow
        for i, (text, command) in enumerate(buttons):
            button = ttk.Button(button_frame, text=text, command=command, width=15)
            button.grid(row=i//2, column=i%2, padx=2, pady=2, sticky='ew')
        
        # Configure grid columns to be equal width
        button_frame.grid_columnconfigure(0, weight=1)
        button_frame.grid_columnconfigure(1, weight=1)
        
        # Add calculator frames to right frame
        self.setup_calculator_frame(right_frame)
        
        # Bind events
        self.tree.bind('<Button-1>', self.handle_click)
        self.tree.bind('<Double-1>', lambda e: setattr(e, 'double', True) or self.handle_click(e))
        self.tree.bind('<Button-3>', lambda e: self.tree.selection_remove(self.tree.selection()))

    def setup_sales_history_tab(self):
        """Setup the sales history tab."""
        # Create main frame with padding
        main_frame = ttk.Frame(self.sales_frame, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Create treeview for sales history
        columns = ('date', 'brand', 'cigar', 'size', 'price_per_stick', 'quantity', 'total_cost')
        self.sales_tree = ttk.Treeview(main_frame, columns=columns, show='headings')
        
        # Configure columns
        headings = {
            'date': 'Date',
            'brand': 'Brand',
            'cigar': 'Cigar',
            'size': 'Size',
            'price_per_stick': 'Price/Stick',
            'quantity': 'Quantity',
            'total_cost': 'Total Cost'
        }
        
        for col, heading in headings.items():
            self.sales_tree.heading(col, text=heading)
            self.sales_tree.column(col, width=120)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=self.sales_tree.yview)
        self.sales_tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack treeview and scrollbar
        self.sales_tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Add undo button frame
        button_frame = ttk.Frame(self.sales_frame, padding="10")
        button_frame.pack(fill='x')

        def undo_selected_sale():
            selected_items = self.sales_tree.selection()
            if not selected_items:
                messagebox.showwarning("Warning", "Please select a sale to undo")
                return

            if messagebox.askyesno("Confirm Undo", "Are you sure you want to undo the selected sale?"):
                for item in selected_items:
                    values = self.sales_tree.item(item)['values']
                    if values:
                        # Find the sale record
                        sale_date = values[0]
                        cigar_name = values[2]
                        quantity = int(values[5])

                        # Restore inventory count
                        for cigar in self.inventory:
                            if cigar['cigar'] == cigar_name:
                                cigar['count'] = int(cigar['count']) + quantity
                                break

                        # Remove the sale record
                        self.sales_history = [sale for sale in self.sales_history 
                                           if not (sale['date'] == sale_date and 
                                                 sale['cigar'] == cigar_name)]

                # Save changes
                self.save_inventory()
                self.save_sales_history()

                # Refresh displays
                self.refresh_inventory()
                self.refresh_sales_history()
                messagebox.showinfo("Success", "Selected sales have been undone")

        # Add undo button
        ttk.Button(button_frame, text="Undo Selected Sale", 
                  command=undo_selected_sale).pack(side='left', padx=5)

    def add_new_brand(self):
        dialog = tk.Toplevel(self.root)
        dialog.title("Add New Brand")
        dialog.geometry("300x100")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text="Brand Name:").pack(pady=(10,0))
        entry = ttk.Entry(dialog, width=40)
        entry.pack(pady=5)
        entry.focus()
        
        def save_brand():
            brand = entry.get().strip()
            if brand:
                self.brands.add(brand)
                self.save_brands()
                self.refresh_inventory()  # Refresh the display
                messagebox.showinfo("Success", f"Brand '{brand}' added successfully!")
            dialog.destroy()
            
        ttk.Button(dialog, text="Add", command=save_brand).pack(pady=5)
        entry.bind('<Return>', lambda e: save_brand())
        
        # Center the dialog on the screen
        dialog.update_idletasks()
        width = dialog.winfo_width()
        height = dialog.winfo_height()
        x = (dialog.winfo_screenwidth() // 2) - (width // 2)
        y = (dialog.winfo_screenheight() // 2) - (height // 2)
        dialog.geometry(f'{width}x{height}+{x}+{y}')

    def add_new_size(self):
        self._add_new_item("Size", self.sizes, 'cigar_sizes.json')

    def add_new_type(self):
        self._add_new_item("Type", self.types, 'cigar_types.json')

    def _add_new_item(self, item_type, item_set, filename):
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Add New {item_type}")
        dialog.geometry("300x100")
        dialog.transient(self.root)
        dialog.grab_set()
        
        ttk.Label(dialog, text=f"{item_type} Name:").pack(pady=(10,0))
        entry = ttk.Entry(dialog, width=40)
        entry.pack(pady=5)
        entry.focus()
        
        def save_item():
            item = entry.get().strip()
            if item:
                item_set.add(item)
                try:
                    with open(resource_path(filename), 'w') as f:
                        json.dump(list(item_set), f, indent=2)
                except Exception as e:
                    messagebox.showerror("Error", f"Failed to save {item_type.lower()}: {str(e)}")
            dialog.destroy()
            
        ttk.Button(dialog, text="Add", command=save_item).pack(pady=5)
        entry.bind('<Return>', lambda e: save_item())

    def handle_click(self, event):
        region = self.tree.identify_region(event.x, event.y)
        if region != "cell":
            return
            
        column = self.tree.identify_column(event.x)
        item = self.tree.identify_row(event.y)
        
        if not item:
            return
            
        # Fix column number calculation - remove the '#' and convert to int
        col_num = int(column.replace('#', '')) - 1
        col_name = self.tree['columns'][col_num]
        
        # Handle checkbox column click
        if column == '#1':  # First column (checkbox)
            cigar_name = self.tree.item(item)['values'][2]  # Cigar name is now at index 2
            self.checkbox_states[cigar_name] = not self.checkbox_states.get(cigar_name, False)
            self.refresh_inventory()
            self.update_selected_cigars_display()  # Update the selected cigars display
            return
            
        # Don't clear other selections if Shift or Control is held
        if not (event.state & 0x4):  # Check if Control is held
            if not (event.state & 0x1):  # Check if Shift is held
                self.tree.selection_set(item)
        
        if col_name not in ['select', 'per_stick']:
            x, y, w, h = self.tree.bbox(item, column)
            
            # Clear any existing popups first
            for widget in self.tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
            
            # Only show editors on double-click
            if hasattr(event, 'double') and event.double:
                print(f"Double-click detected on column: {col_name}")
                if col_name == 'brand':
                    self.show_dropdown(item, col_name, self.brands, x, y, w, h)
                elif col_name == 'size':
                    self.show_dropdown(item, col_name, self.sizes, x, y, w, h)
                elif col_name == 'type':
                    self.show_dropdown(item, col_name, self.types, x, y, w, h)
                elif col_name == 'personal_rating':
                    print("Showing rating dropdown")
                    self.show_rating_dropdown(item, col_name, x, y, w, h)
                elif col_name == 'count':
                    self.show_spinbox(item, col_name, x, y, w, h)
                elif col_name in ['price', 'shipping']:
                    self.show_price_entry(item, col_name, x, y, w, h)
                else:  # cigar name field
                    self.show_text_entry(item, col_name, x, y, w, h)

    def show_dropdown(self, item, column, values, x, y, w, h):
        # Destroy any existing popups
        for widget in self.tree.winfo_children():
            if isinstance(widget, ttk.Frame):
                widget.destroy()
                
        frame = ttk.Frame(self.tree)
        
        # Convert values to list if it's a range object
        if isinstance(values, range):
            values = [str(i) for i in values]
        else:
            values = sorted(list(values))
            
        values = [''] + values  # Add empty option
        
        combo = ttk.Combobox(frame, values=values, width=w//10)
        current_value = self.tree.set(item, column)
        combo.set(current_value)
        combo.pack(expand=True, fill='both')
        
        def save_value(event=None):
            try:
                value = combo.get().strip()
                
                # Get all values from the tree item
                item_values = self.tree.item(item)['values']
                if not item_values:
                    return
                
                # Get cigar name (index 2 because of checkbox)
                cigar_name = item_values[2]
                
                # Find the cigar in inventory and update it
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        # Update the value
                        cigar[column] = value
                        
                        # Add to appropriate set if not empty
                        if value:
                            if column == 'brand':
                                self.brands.add(value)
                                self.save_brands()
                            elif column == 'size':
                                self.sizes.add(value)
                                self.save_sets('cigar_sizes.json', self.sizes)
                            elif column == 'type':
                                self.types.add(value)
                                self.save_sets('cigar_types.json', self.types)
                        break
                
                # Save changes
                self.save_inventory()
                
                # Update display
                self.tree.set(item, column, value)
                
            except Exception as e:
                print(f"Error saving value: {e}")
            finally:
                frame.destroy()
        
        combo.bind('<<ComboboxSelected>>', save_value)
        combo.bind('<Return>', save_value)
        combo.bind('<FocusOut>', save_value)
        combo.bind('<Escape>', lambda e: frame.destroy())
        
        frame.place(x=x, y=y, width=w, height=h)
        combo.focus()

    def show_rating_dropdown(self, item, column, x, y, w, h):
        """Show a simple dropdown for ratings 1-10."""
        print("Entering show_rating_dropdown method")
        
        # Get the cigar name and current rating
        cigar_values = self.tree.item(item)['values']
        if not cigar_values:
            print("No cigar values found")
            return
            
        cigar_name = cigar_values[2]  # Index 2 for cigar name
        current_rating = self.tree.set(item, column)
        print(f"Current rating: {current_rating}")
        
        # Create frame and combobox
        frame = ttk.Frame(self.tree)
        
        # Create list of ratings 1-10
        ratings = [''] + [str(i) for i in range(1, 11)]
        
        # Create and configure combobox
        combo = ttk.Combobox(frame, values=ratings, width=w//10)
        combo.set(current_rating if current_rating != 'N/A' else '')
        combo.pack(expand=True, fill='both')
        
        def save_rating(event=None):
            try:
                print("Saving rating...")
                # Get selected rating
                rating = combo.get()
                print(f"Selected rating: {rating}")
                
                # Update inventory
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        cigar[column] = int(rating) if rating else None
                        print(f"Updated rating for {cigar_name}: {cigar[column]}")
                        break
                
                # Save to file
                self.save_inventory()
                
                # Update display
                self.tree.set(item, column, rating if rating else 'N/A')
                
            except Exception as e:
                print(f"Error saving rating: {e}")
            finally:
                frame.destroy()
        
        # Bind events
        combo.bind('<<ComboboxSelected>>', save_rating)
        combo.bind('<Return>', save_rating)
        combo.bind('<Escape>', lambda e: frame.destroy())
        combo.bind('<FocusOut>', save_rating)
        
        # Position frame and focus
        frame.place(x=x, y=y, width=w, height=h)
        combo.focus()
        print("Rating dropdown setup complete")

    def show_spinbox(self, item, column, x, y, w, h):
        current_value = self.tree.set(item, column)
        if current_value == 'N/A':
            current_value = '0'
            
        frame = ttk.Frame(self.tree)
        
        # Configure spinbox based on column
        if column == 'count':
            from_, to = 0, 999
        elif column == 'personal_rating':
            from_, to = 1, 10
        else:  # overall_rating
            from_, to = 1, 100
            
        spinbox = ttk.Spinbox(
            frame,
            from_=from_,
            to=to,
            width=w//10,
            justify='center'
        )
        spinbox.set(current_value)
        spinbox.pack(expand=True, fill='both')
        spinbox.focus()

        def save_value(event=None, manual_entry=False, should_close=False):
            try:
                value = int(spinbox.get())
                
                # Get all values from the tree item
                item_values = self.tree.item(item)['values']
                if not item_values:
                    return
                
                # Get cigar name (index 2 because of checkbox)
                cigar_name = item_values[2]
                
                # Find the cigar in inventory and update it
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        # Always update the count
                        cigar[column] = value
                        
                        # Only recalculate price_per_stick if this was a manual entry (typing)
                        if column == 'count' and manual_entry and value > 0:
                            price = float(cigar.get('price', 0))
                            shipping = float(cigar.get('shipping', 0))
                            cigar['price_per_stick'] = self.calculate_price_per_stick(price, shipping, value)
                            # Update the price_per_stick display in the tree
                            self.tree.set(item, 'per_stick', f"${cigar['price_per_stick']:.2f}")
                        break
                
                # Save changes
                self.save_inventory()
                
                # Update display and totals
                self.tree.set(item, column, str(value))
                self.update_inventory_totals()  # Update totals after count change
                
            except ValueError:
                # If invalid value, restore previous value
                self.tree.set(item, column, current_value)
            finally:
                if should_close:
                    frame.destroy()

        # Create a command for spinbox arrows that saves without closing
        def on_arrow():
            save_value(manual_entry=False, should_close=False)
        spinbox.configure(command=on_arrow)
        
        # Bind events that should save and close
        spinbox.bind('<Return>', lambda e: save_value(e, manual_entry=True, should_close=True))
        spinbox.bind('<FocusOut>', lambda e: save_value(e, manual_entry=False, should_close=True))
        spinbox.bind('<Escape>', lambda e: frame.destroy())
        
        # Bind event for manual entry that saves but doesn't close
        spinbox.bind('<KeyRelease>', lambda e: save_value(e, manual_entry=True, should_close=False))
        
        frame.place(x=x, y=y, width=w, height=h)

    def show_price_entry(self, item, column, x, y, w, h):
        current_value = self.tree.set(item, column).replace('$', '')
        
        frame = ttk.Frame(self.tree)
        entry = ttk.Entry(frame, justify='right')
        entry.insert(0, current_value)
        entry.pack(expand=True, fill='both')
        entry.focus()
        entry.select_range(0, tk.END)

        def save_value(event=None):
            try:
                value = float(entry.get())
                
                # Get all values from the tree item
                item_values = self.tree.item(item)['values']
                if not item_values:
                    return
                
                # Get cigar name (index 2 because of checkbox)
                cigar_name = item_values[2]
                
                # Find the cigar in inventory and update it
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        cigar[column] = value
                        # Recalculate price per stick only when price or shipping is manually changed
                        if column in ['price', 'shipping'] and cigar['count'] > 0:
                            cigar['price_per_stick'] = self.calculate_price_per_stick(
                                cigar['price'],
                                cigar['shipping'],
                                cigar['count']
                            )
                        break
                
                # Save changes
                self.save_inventory()
                
                # Update display and totals
                self.tree.set(item, column, f"${value:.2f}")
                self.update_inventory_totals()  # Update totals after price change
                
            except ValueError:
                # If invalid value, restore previous value
                self.tree.set(item, column, f"${float(current_value):.2f}")
            finally:
                frame.destroy()

        entry.bind('<Return>', save_value)
        entry.bind('<FocusOut>', save_value)
        entry.bind('<Escape>', lambda e: frame.destroy())
        
        frame.place(x=x, y=y, width=w, height=h)

    def show_text_entry(self, item, column, x, y, w, h):
        current_value = self.tree.set(item, column)
        
        frame = ttk.Frame(self.tree)
        entry = ttk.Entry(frame)
        entry.insert(0, current_value)
        entry.pack(expand=True, fill='both')
        entry.focus()
        entry.select_range(0, tk.END)

        def save_value(event=None):
            try:
                value = entry.get().strip()
                if value:  # Only save if there's a value
                    # Get all values from the tree item
                    item_values = self.tree.item(item)['values']
                    if not item_values:
                        return
                    
                    # Get cigar name (index 2 because of checkbox)
                    cigar_name = item_values[2]
                    
                    # Find the cigar in inventory and update it
                    for cigar in self.inventory:
                        if cigar['cigar'] == cigar_name:
                            # Update the value
                            cigar[column] = value
                            break
                    
                    # Save changes
                    self.save_inventory()
                    
                    # Update display
                    self.tree.set(item, column, value)
                
            except Exception as e:
                print(f"Error saving value: {e}")
            finally:
                frame.destroy()
            
        entry.bind('<Return>', save_value)
        entry.bind('<FocusOut>', save_value)
        entry.bind('<Escape>', lambda e: frame.destroy())
        
        frame.place(x=x, y=y, width=w, height=h)

    def calculate_price_per_stick(self, price, shipping, count, original_quantity=None):
        """Calculate price per stick including tax and shipping."""
        try:
            price = float(price)
            shipping = float(shipping)
            count = int(count)
            
            if count <= 0:
                return 0
                
            # Calculate base price per stick
            price_per_stick = price / count
            
            # Add tax to the base price
            price_with_tax = price_per_stick * (1 + self.tax_rate)
            
            # If original_quantity is provided (new cigars), use it for shipping
            # Otherwise (existing cigars), use current count
            shipping_quantity = original_quantity if original_quantity is not None else count
            shipping_per_stick = shipping / shipping_quantity
            
            # Total cost per stick
            total_per_stick = price_with_tax + shipping_per_stick
            
            return total_per_stick
            
        except (ValueError, TypeError):
            return 0
        
    def on_search(self, *args):
        self.refresh_inventory()
    
    def remove_selected(self):
        selected_items = self.tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select items to remove")
            return
        
        # Get list of cigars to be removed
        cigars_to_remove = []
        for item in selected_items:
            values = self.tree.item(item)['values']
            if values:
                cigar_name = values[2]  # Index 2 for cigar name
                cigars_to_remove.append(cigar_name)
        
        # Confirm deletion with user
        confirm_msg = "Are you sure you want to remove these cigars?\n\n"
        confirm_msg += "\n".join(cigars_to_remove)
        if messagebox.askyesno("Confirm Removal", confirm_msg):
            # Remove from checkbox states if present
            for cigar_name in cigars_to_remove:
                if cigar_name in self.checkbox_states:
                    del self.checkbox_states[cigar_name]
                
            # Remove from inventory
            self.inventory = [x for x in self.inventory if x['cigar'] not in cigars_to_remove]
            
            # Save changes
            self.save_inventory()
            
            # Refresh display
            self.refresh_inventory()
            
            # Update totals
            self.update_inventory_totals()
            self.update_selected_cigars_display()
            
            messagebox.showinfo("Success", f"Removed {len(cigars_to_remove)} cigar(s) from inventory")

    def update_count(self):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an item to update")
            return
            
        item = self.tree.item(selected_item[0])
        name = item['values'][1]
        current_count = item['values'][2]
        
        # Create update dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Count")
        dialog.geometry("300x150")
        
        ttk.Label(dialog, text=f"Current count: {current_count}").pack(pady=10)
        count_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=count_var).pack(pady=10)
        
        def update():
            try:
                new_count = int(count_var.get())
                if new_count < 0:
                    raise ValueError("Count cannot be negative")
                    
                for cigar in self.inventory:
                    if cigar['cigar'] == name:
                        cigar['count'] = new_count
                        cigar['price_per_stick'] = self.calculate_price_per_stick(
                            cigar['price'], cigar['shipping'], new_count)
                        break
                        
                self.save_inventory()
                self.refresh_inventory()
                dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                
        ttk.Button(dialog, text="Update", command=update).pack(pady=10)
    
    def update_rating(self, rating_type):
        selected_item = self.tree.selection()
        if not selected_item:
            messagebox.showwarning("Warning", "Please select an item to update")
            return
            
        item = self.tree.item(selected_item[0])
        name = item['values'][1]
        current_rating = item['values'][5] if rating_type == 'personal' else item['values'][6]
        
        # Create update dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Rating")
        dialog.geometry("300x150")
        
        ttk.Label(dialog, text=f"Current {rating_type} rating: {current_rating}").pack(pady=10)
        rating_var = tk.StringVar()
        ttk.Entry(dialog, textvariable=rating_var).pack(pady=10)
        
        def update():
            try:
                new_rating = float(rating_var.get())
                if not (1 <= new_rating <= 10 if rating_type == 'personal' else 100):
                    raise ValueError(f"{rating_type.capitalize()} rating must be between 1 and 10" if rating_type == 'personal' else "Overall rating must be between 1 and 100")
                    
                for cigar in self.inventory:
                    if cigar['cigar'] == name:
                        if rating_type == 'personal':
                            cigar['personal_rating'] = new_rating
                        else:
                            cigar['overall_rating'] = new_rating
                        break
                        
                self.save_inventory()
                self.refresh_inventory()
                dialog.destroy()
                
            except ValueError as e:
                messagebox.showerror("Error", str(e))
                
        ttk.Button(dialog, text="Update", command=update).pack(pady=10)
        
    def refresh_inventory(self):
        try:
            # Store current selections by brand and cigar name
            selected_items = []
            for item in self.tree.selection():
                values = self.tree.item(item)['values']
                if values and len(values) >= 3:  # Make sure we have enough values
                    selected_items.append((values[1], values[2]))  # brand, cigar
            
            # Clear current items
            self.tree.delete(*self.tree.get_children())
            
            search_term = self.search_var.get().lower()
            
            # Sort inventory if sort column is set
            if self.sort_column and self.sort_column != 'select':
                def sort_key(x):
                    value = x.get(self.sort_column, '')
                    if self.sort_column == 'per_stick':
                        return float(x.get('price_per_stick', 0))
                    elif self.sort_column in ['count']:
                        return int(value) if str(value).isdigit() else 0
                    elif self.sort_column in ['price', 'shipping']:
                        try:
                            return float(str(value).replace('$', ''))
                        except ValueError:
                            return 0.0
                    elif self.sort_column == 'personal_rating':
                        return int(value) if value is not None else -1
                    return str(value).lower()
                
                sorted_inventory = sorted(
                    self.inventory,
                    key=sort_key,
                    reverse=self.sort_reverse.get(self.sort_column, False)
                )
            else:
                # Sort by brand and cigar name by default
                sorted_inventory = sorted(
                    self.inventory,
                    key=lambda x: (x.get('brand', '').lower(), x.get('cigar', '').lower())
                )
            
            # Display items
            for cigar in sorted_inventory:
                if search_term and search_term not in cigar.get('cigar', '').lower() and search_term not in cigar.get('brand', '').lower():
                    continue
                    
                # Format rating properly
                personal_rating = str(cigar.get('personal_rating')) if cigar.get('personal_rating') is not None else "N/A"
                
                # Get checkbox state
                is_selected = self.checkbox_states.get(cigar.get('cigar', ''), False)
                
                # For existing cigars (without original_quantity), use current calculation
                # For new cigars (with original_quantity), use new calculation
                if 'original_quantity' in cigar:
                    price_per_stick = self.calculate_price_per_stick(
                        cigar.get('price', 0),
                        cigar.get('shipping', 0),
                        cigar.get('count', 0),
                        cigar.get('original_quantity')
                    )
                else:
                    # Use existing calculation for old inventory
                    price_per_stick = self.calculate_price_per_stick(
                        cigar.get('price', 0),
                        cigar.get('shipping', 0),
                        cigar.get('count', 0)
                    )
                
                values = [
                    '☒' if is_selected else '☐',
                    cigar.get('brand', ''),
                    cigar.get('cigar', ''),
                    cigar.get('size', ''),
                    cigar.get('type', ''),
                    str(cigar.get('count', 0)),
                    f"${cigar.get('price', 0):.2f}",
                    f"${cigar.get('shipping', 0):.2f}",
                    f"${price_per_stick:.2f}",
                    personal_rating
                ]
                
                item_id = self.tree.insert('', 'end', values=values)
                
                # Restore selection if this was a previously selected item
                if (values[1], values[2]) in selected_items:
                    self.tree.selection_add(item_id)
                    self.tree.see(item_id)  # Make sure the item is visible
            
            # Update totals
            self.update_order_total()
            self.update_inventory_totals()
                
        except Exception as e:
            print(f"Error refreshing display: {str(e)}")

    def save_inventory(self):
        try:
            with open(resource_path('cigar_inventory.json'), 'w') as f:
                json.dump(self.inventory, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inventory: {str(e)}")
            
    def load_inventory(self):
        try:
            # Load sales history first
            self.load_sales_history()
            
            # Load brands first
            try:
                with open(resource_path('cigar_brands.json'), 'r') as f:
                    self.brands = set(json.load(f))
            except FileNotFoundError:
                self.brands = set()
            
            # Load sizes
            try:
                with open(resource_path('cigar_sizes.json'), 'r') as f:
                    self.sizes = set(json.load(f))
            except FileNotFoundError:
                self.sizes = set()
                
            # Load types
            try:
                with open(resource_path('cigar_types.json'), 'r') as f:
                    self.types = set(json.load(f))
            except FileNotFoundError:
                self.types = set()
            
            # Load inventory
            with open(resource_path('cigar_inventory.json'), 'r') as f:
                self.inventory = json.load(f)
                # Add existing values to sets and ensure proper data structure
                for cigar in self.inventory:
                    if 'brand' not in cigar: cigar['brand'] = ''
                    if 'size' not in cigar: cigar['size'] = ''
                    if 'type' not in cigar: cigar['type'] = ''
                    if 'personal_rating' not in cigar: cigar['personal_rating'] = None
                    
                    if cigar['brand']: self.brands.add(cigar['brand'])
                    if cigar['size']: self.sizes.add(cigar['size'])
                    if cigar['type']: self.types.add(cigar['type'])
                    
                    # Only calculate price_per_stick if it doesn't exist
                    if 'price_per_stick' not in cigar:
                        cigar['price_per_stick'] = self.calculate_price_per_stick(
                            cigar['price'], 
                            cigar['shipping'], 
                            cigar['count']
                        )
        except FileNotFoundError:
            self.inventory = []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load inventory: {str(e)}")
            self.inventory = []

    def export_inventory(self):
        if not self.inventory:
            messagebox.showwarning("Warning", "No data to export")
            return
            
        # Ask user for export format
        export_format = tk.StringVar(value="csv")
        dialog = tk.Toplevel(self.root)
        dialog.title("Export Inventory")
        dialog.geometry("300x150")
        
        ttk.Label(dialog, text="Choose export format:").pack(pady=10)
        ttk.Radiobutton(dialog, text="CSV", variable=export_format, value="csv").pack()
        ttk.Radiobutton(dialog, text="Excel", variable=export_format, value="excel").pack()
        
        def do_export():
            fmt = export_format.get()
            if fmt == "csv":
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".csv",
                    filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
                )
                if file_path:
                    with open(file_path, 'w', newline='') as f:
                        writer = csv.DictWriter(f, fieldnames=[
                            'brand', 'cigar', 'size', 'type', 'count', 'price', 'shipping', 'price_per_stick',
                            'personal_rating'
                        ])
                        writer.writeheader()
                        writer.writerows(self.inventory)
            else:
                file_path = filedialog.asksaveasfilename(
                    defaultextension=".xlsx",
                    filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
                )
                if file_path:
                    df = pd.DataFrame(self.inventory)
                    df.to_excel(file_path, index=False)
            
            dialog.destroy()
            messagebox.showinfo("Success", "Inventory exported successfully!")
            
        ttk.Button(dialog, text="Export", command=do_export).pack(pady=10)

    def on_closing(self):
        try:
            self.save_inventory()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inventory: {str(e)}")
        finally:
            self.root.destroy()

    def save_brands(self):
        try:
            with open(resource_path('cigar_brands.json'), 'w') as f:
                json.dump(list(self.brands), f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save brands: {str(e)}")

    def add_new_line(self):
        """Add a new line to the inventory."""
        new_cigar = {
            'brand': '',
            'cigar': f'New Cigar {len(self.inventory) + 1}',  # Add number to make unique
            'size': '',
            'type': '',
            'count': 0,
            'price': 0.0,
            'shipping': 0.0,
            'price_per_stick': 0.0,
            'personal_rating': None
        }
        
        # Calculate initial price per stick using count=1 to avoid division by zero
        new_cigar['price_per_stick'] = self.calculate_price_per_stick(
            new_cigar['price'],
            new_cigar['shipping'],
            1  # Always use 1 for initial calculation
        )
        
        self.inventory.append(new_cigar)
        self.save_inventory()
        self.refresh_inventory()
        
        # Select the new item
        last_item = self.tree.get_children()[-1]
        self.tree.selection_add(last_item)  # Changed from selection_set to selection_add
        self.tree.see(last_item)

    def sort_treeview(self, col):
        """Sort treeview by column."""
        if col == 'select':  # Don't sort the checkbox column
            return
        
        try:
            # Store current selections by brand and cigar name
            selected_items = []
            for item in self.tree.selection():
                values = self.tree.item(item)['values']
                if values:
                    selected_items.append((values[1], values[2]))  # brand, cigar
            
            # Toggle sort direction
            if not hasattr(self, 'sort_reverse'):
                self.sort_reverse = {}
            self.sort_reverse[col] = not self.sort_reverse.get(col, False)
            
            # Create a list of tuples (sort_key, original_item)
            items_to_sort = []
            for item in self.inventory:
                if col == 'per_stick':
                    # Handle price_per_stick sorting
                    try:
                        key = float(str(item.get('price_per_stick', '0')).replace('$', '').replace(',', ''))
                    except ValueError:
                        key = 0.0
                elif col in ['count', 'personal_rating']:
                    try:
                        key = float(str(item.get(col, '0')).replace('N/A', '-1'))
                    except ValueError:
                        key = -1
                elif col in ['price', 'shipping']:
                    try:
                        key = float(str(item.get(col, '0')).replace('$', '').replace(',', ''))
                    except ValueError:
                        key = 0.0
                else:
                    key = str(item.get(col, '')).lower()
                
                items_to_sort.append((key, item))
            
            # Sort the items
            items_to_sort.sort(key=lambda x: x[0], reverse=self.sort_reverse[col])
            
            # Update inventory with sorted items
            self.inventory = [item[1] for item in items_to_sort]
            
            # Clear and repopulate the tree
            self.tree.delete(*self.tree.get_children())
            
            # Repopulate with sorted data
            for item in self.inventory:
                values = [
                    '☒' if self.checkbox_states.get(item.get('cigar', ''), False) else '☐',
                    item.get('brand', ''),
                    item.get('cigar', ''),
                    item.get('size', ''),
                    item.get('type', ''),
                    str(item.get('count', '0')),
                    f"${float(item.get('price', '0')):.2f}",
                    f"${float(item.get('shipping', '0')):.2f}",
                    f"${float(item.get('price_per_stick', '0')):.2f}",
                    item.get('personal_rating', 'N/A')
                ]
                
                iid = self.tree.insert('', 'end', values=values)
                
                # Restore selection if this was previously selected
                if (values[1], values[2]) in selected_items:
                    self.tree.selection_add(iid)
                    self.tree.see(iid)
            
            # Update column header
            self.tree.heading(col, text=f"{col.title()} {'↓' if self.sort_reverse[col] else '↑'}")
            
            # Update displays
            self.update_inventory_totals()
            self.update_selected_cigars_display()
            
        except Exception as e:
            print(f"Sort error: {str(e)}")

    def save_sets(self, filename, data_set):
        """Save a set to a JSON file."""
        try:
            with open(resource_path(filename), 'w') as f:
                json.dump(list(data_set), f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save {filename}: {str(e)}")

    def setup_calculator_frame(self, parent_frame):
        # Shipping Calculator frame
        calc_frame = ttk.LabelFrame(parent_frame, text="Shipping Calculator", padding=10)
        calc_frame.pack(fill='x', pady=10)

        # Input fields
        input_frame = ttk.Frame(calc_frame)
        input_frame.pack(fill='x', pady=5)

        ttk.Label(input_frame, text="Shipping Cost ($):").pack(fill='x', pady=2)
        self.ship_cost = ttk.Entry(input_frame, width=15)
        self.ship_cost.pack(fill='x', pady=2)

        ttk.Label(input_frame, text="Total Cigars:").pack(fill='x', pady=2)
        self.total_cigars = ttk.Entry(input_frame, width=15)
        self.total_cigars.pack(fill='x', pady=2)

        ttk.Button(input_frame, text="Calculate", command=self.calculate_shipping).pack(fill='x', pady=5)

        # Results labels
        self.per_stick = ttk.Label(calc_frame, text="Per Stick (Total): $0.00")
        self.per_stick.pack(fill='x', pady=2)
        
        self.five_pack = ttk.Label(calc_frame, text="5-Pack (per stick): $0.00")
        self.five_pack.pack(fill='x', pady=2)
        
        self.ten_pack = ttk.Label(calc_frame, text="10-Pack (per stick): $0.00")
        self.ten_pack.pack(fill='x', pady=2)

        # Build Your Pack frame
        order_frame = ttk.LabelFrame(parent_frame, text="Build Your Pack", padding=10)
        order_frame.pack(fill='x', pady=10)

        # Create main container for pack list and totals
        pack_container = ttk.Frame(order_frame)
        pack_container.pack(fill='both', expand=True)

        # Create scrollable frame for selected cigars
        pack_canvas = tk.Canvas(pack_container)
        pack_scrollbar = ttk.Scrollbar(pack_container, orient="vertical", command=pack_canvas.yview)
        
        # Create a frame to hold both the cigars list and the totals
        self.selected_cigars_frame = ttk.Frame(pack_canvas)

        # Configure canvas
        pack_canvas.configure(yscrollcommand=pack_scrollbar.set)
        
        # Pack scrollbar and canvas
        pack_scrollbar.pack(side="right", fill="y")
        pack_canvas.pack(side="left", fill="both", expand=True)
        
        # Create window in canvas for the frame
        canvas_frame = pack_canvas.create_window((0, 0), window=self.selected_cigars_frame, anchor="nw", width=pack_canvas.winfo_reqwidth())

        # Configure canvas scrolling
        def configure_scroll_region(event):
            pack_canvas.configure(scrollregion=pack_canvas.bbox("all"))
            # Update the window width when frame changes
            pack_canvas.itemconfig(canvas_frame, width=pack_canvas.winfo_width())
            
        self.selected_cigars_frame.bind("<Configure>", configure_scroll_region)
        pack_canvas.bind('<Configure>', lambda e: pack_canvas.itemconfig(canvas_frame, width=e.width))

        # Add mouse wheel scrolling for pack list
        def on_pack_mousewheel(event):
            pack_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        pack_canvas.bind('<MouseWheel>', on_pack_mousewheel)
        self.selected_cigars_frame.bind('<MouseWheel>', lambda e: pack_canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        # Set fixed height for pack list
        pack_canvas.configure(height=300)  # Increased height to accommodate totals

        # Dictionary to store quantity spinboxes
        self.quantity_spinboxes = {}

        # Create frame for total labels and sell button at the bottom of the scrollable area
        self.totals_frame = ttk.Frame(self.selected_cigars_frame)
        
        # Add separator above totals
        ttk.Separator(self.selected_cigars_frame, orient='horizontal').pack(fill='x', pady=10)
        
        # Pack the totals frame at the bottom
        self.totals_frame.pack(fill='x', pady=5, side='bottom')

        self.order_total_label = ttk.Label(self.totals_frame, text="Selected Singles Total: $0.00", font=('TkDefaultFont', 10, 'bold'))
        self.order_total_label.pack(fill='x', pady=2)

        self.order_count_label = ttk.Label(self.totals_frame, text="Selected Singles: 0")
        self.order_count_label.pack(fill='x', pady=2)

        # Add Sell Selected button
        ttk.Button(self.totals_frame, text="Sell Selected", command=self.sell_selected, width=20).pack(pady=5)

    def calculate_shipping(self):
        try:
            shipping = float(self.ship_cost.get() or 0)
            total_cigars = int(self.total_cigars.get() or 0)
            
            # Calculate shipping costs per stick for different quantities
            per_stick_cost = shipping / total_cigars if total_cigars > 0 else shipping
            five_pack = shipping / 5
            ten_pack = shipping / 10
            
            # Update labels with more detailed information
            self.per_stick.config(text=f"Per Stick ({total_cigars} total): ${per_stick_cost:.2f}")
            self.five_pack.config(text=f"5-Pack (per stick): ${five_pack:.2f}")
            self.ten_pack.config(text=f"10-Pack (per stick): ${ten_pack:.2f}")
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numbers for shipping cost and total cigars")
        except ZeroDivisionError:
            messagebox.showerror("Error", "Total cigars must be greater than 0")

    def update_order_total(self):
        total_price = 0
        total_cigars = 0
        
        for cigar in self.inventory:
            cigar_name = cigar.get('cigar', '')
            if self.checkbox_states.get(cigar_name, False):
                try:
                    # Get quantity from spinbox
                    quantity = int(self.quantity_spinboxes[cigar_name].get())
                    # Get price per stick
                    price_per_stick = float(cigar.get('price_per_stick', 0))
                    
                    # Calculate total for this cigar
                    total_price += price_per_stick * quantity
                    total_cigars += quantity
                    
                except (ValueError, KeyError):
                    # If there's an error with the spinbox, assume quantity of 1
                    total_price += float(cigar.get('price_per_stick', 0))
                    total_cigars += 1
        
        # Update the labels
        self.order_total_label.config(text=f"Selected Singles Total: ${total_price:.2f}")
        self.order_count_label.config(text=f"Selected Singles: {total_cigars}")

    def setup_totals_frame(self, parent_frame):
        """Setup frame for displaying inventory totals."""
        totals_frame = ttk.LabelFrame(parent_frame, text="Inventory Totals", padding=10)
        totals_frame.pack(fill='x', padx=10, pady=5, side='bottom')

        # Create labels for totals in a grid layout
        self.total_count_label = ttk.Label(totals_frame, text="Total Count: 0", anchor='center')
        self.total_count_label.pack(side='left', expand=True, padx=10)

        self.total_value_label = ttk.Label(totals_frame, text="Total Value: $0.00", anchor='center')
        self.total_value_label.pack(side='left', expand=True, padx=10)

        self.avg_shipping_label = ttk.Label(totals_frame, text="Avg Shipping: $0.00", anchor='center')
        self.avg_shipping_label.pack(side='left', expand=True, padx=10)

        self.avg_price_stick_label = ttk.Label(totals_frame, text="Avg Price/Stick: $0.00", anchor='center')
        self.avg_price_stick_label.pack(side='left', expand=True, padx=10)

    def update_inventory_totals(self):
        """Update the inventory totals display."""
        total_count = 0
        total_value = 0.0
        total_items = 0  # Count all items, not just those with stock

        for cigar in self.inventory:
            count = int(cigar.get('count', 0))
            if count > 0:
                # Calculate total count
                total_count += count
                
                # Calculate total value using price_per_stick × count
                price_per_stick = float(cigar.get('price_per_stick', 0))
                total_value += (price_per_stick * count)
                total_items += 1

        # Calculate average shipping for items with stock
        avg_shipping = sum(float(cigar.get('shipping', 0)) for cigar in self.inventory if int(cigar.get('count', 0)) > 0) / total_items if total_items > 0 else 0
        
        # Calculate average price per stick
        avg_price_stick = total_value / total_count if total_count > 0 else 0

        # Update labels with formatted values
        self.total_count_label.config(text=f"Total Count: {total_count}")
        self.total_value_label.config(text=f"Total Value: ${total_value:.2f}")
        self.avg_shipping_label.config(text=f"Avg Shipping: ${avg_shipping:.2f}")
        self.avg_price_stick_label.config(text=f"Avg Price/Stick: ${avg_price_stick:.2f}")

    def update_selected_cigars_display(self):
        """Update the display of selected cigars and their quantity controls."""
        # Store current quantities before clearing
        current_quantities = {}
        for cigar_name, spinbox in self.quantity_spinboxes.items():
            try:
                current_quantities[cigar_name] = spinbox.get()
            except:
                pass

        # Clear existing widgets except totals frame
        for widget in self.selected_cigars_frame.winfo_children():
            if widget != self.totals_frame:
                widget.destroy()
        
        # Clear the spinboxes dictionary
        self.quantity_spinboxes.clear()

        # Add selected cigars with quantity controls
        for cigar in self.inventory:
            cigar_name = cigar.get('cigar', '')
            if self.checkbox_states.get(cigar_name, False):
                # Create frame for this cigar
                cigar_frame = ttk.Frame(self.selected_cigars_frame)
                cigar_frame.pack(fill='x', pady=2)

                # Add cigar name label
                name_label = ttk.Label(cigar_frame, text=f"{cigar.get('brand', '')} - {cigar_name}")
                name_label.pack(side='left', padx=5)

                # Add quantity spinbox
                # Use stored quantity if it exists, otherwise default to 1
                initial_quantity = current_quantities.get(cigar_name, "1")
                
                def make_spinbox_update(cigar_name):
                    def update_wrapper(*args):
                        self.update_order_total()
                    return update_wrapper

                spinbox = ttk.Spinbox(
                    cigar_frame,
                    from_=1,
                    to=cigar.get('count', 99),
                    width=5,
                    command=make_spinbox_update(cigar_name)
                )
                spinbox.set(initial_quantity)
                spinbox.pack(side='right', padx=5)
                
                # Bind the spinbox to update totals when value is changed manually
                spinbox.bind('<KeyRelease>', make_spinbox_update(cigar_name))
                
                # Store the spinbox reference
                self.quantity_spinboxes[cigar_name] = spinbox

                # Add quantity label
                ttk.Label(cigar_frame, text="Qty:").pack(side='right', padx=2)

        # Ensure totals frame stays at the bottom
        self.totals_frame.pack_forget()
        ttk.Separator(self.selected_cigars_frame, orient='horizontal').pack(fill='x', pady=10)
        self.totals_frame.pack(fill='x', pady=5, side='bottom')

        # Update order total
        self.update_order_total()

    def show_sale_confirmation(self, sale_records, selected_cigars):
        """Show a nicely formatted sale confirmation dialog with undo option."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Sale")
        dialog.geometry("400x500")
        dialog.transient(self.root)
        dialog.grab_set()

        # Center the dialog relative to the main window
        dialog.update_idletasks()
        dialog_width = dialog.winfo_width()
        dialog_height = dialog.winfo_height()
        parent_x = self.root.winfo_x()
        parent_y = self.root.winfo_y()
        parent_width = self.root.winfo_width()
        parent_height = self.root.winfo_height()
        
        x = parent_x + (parent_width - dialog_width) // 2
        y = parent_y + (parent_height - dialog_height) // 2
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        # Main frame with padding
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Sale details in a treeview
        columns = ('cigar', 'quantity', 'price', 'total')
        tree = ttk.Treeview(main_frame, columns=columns, show='headings', height=10)
        
        # Configure columns
        tree.heading('cigar', text='Cigar')
        tree.heading('quantity', text='Qty')
        tree.heading('price', text='Price/Stick')
        tree.heading('total', text='Total')
        
        tree.column('cigar', width=150)
        tree.column('quantity', width=50)
        tree.column('price', width=80)
        tree.column('total', width=80)

        # Add scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Calculate grand total
        grand_total = 0.0
        
        # Add sale records to tree
        for record in sale_records:
            price_per_stick = float(record['price_per_stick'])
            quantity = int(record['quantity'])
            total = price_per_stick * quantity
            grand_total += total
            
            values = (
                f"{record['brand']} - {record['cigar']}",
                str(quantity),
                f"${price_per_stick:.2f}",
                f"${total:.2f}"
            )
            tree.insert('', 'end', values=values)

        # Total frame
        total_frame = ttk.Frame(dialog, padding="10")
        total_frame.pack(fill='x', pady=(0, 10))
        
        # Grand total label
        total_label = ttk.Label(total_frame, text=f"Total: ${grand_total:.2f}", 
                              font=('TkDefaultFont', 12, 'bold'))
        total_label.pack(side='right')

        # Button frame
        button_frame = ttk.Frame(dialog, padding="10")
        button_frame.pack(fill='x')

        def undo_sale():
            if messagebox.askyesno("Confirm Undo", "Are you sure you want to undo this sale?"):
                # Restore inventory counts
                for cigar_name, quantity in selected_cigars:
                    for cigar in self.inventory:
                        if cigar['cigar'] == cigar_name:
                            cigar['count'] = int(cigar['count']) + quantity
                            break
                
                # Remove sale records
                for record in sale_records:
                    self.sales_history.remove(record)
                
                # Save changes
                self.save_inventory()
                self.save_sales_history()
                
                # Refresh displays
                self.refresh_inventory()
                self.refresh_sales_history()
                
                dialog.destroy()
                messagebox.showinfo("Success", "Sale has been undone")

        # Add buttons
        ttk.Button(button_frame, text="Print", command=lambda: print("Print functionality not implemented"), 
                  width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Undo Sale", command=undo_sale, 
                  width=15).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Close", command=dialog.destroy, 
                  width=15).pack(side='right', padx=5)

    def sell_selected(self):
        """Process the sale of selected cigars."""
        selected_cigars = []
        sale_records = []
        sale_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Process each selected cigar
        for cigar in self.inventory:
            cigar_name = cigar.get('cigar', '')
            if self.checkbox_states.get(cigar_name, False):
                try:
                    quantity = int(self.quantity_spinboxes[cigar_name].get())
                    if quantity < 1:
                        messagebox.showwarning("Warning", f"Purchase quantity for {cigar_name} must be at least 1")
                        continue
                except (ValueError, KeyError):
                    messagebox.showwarning("Warning", f"Invalid purchase quantity for {cigar_name}")
                    continue

                current_count = int(cigar.get('count', 0))
                if current_count >= quantity:  # Check if we have enough stock
                    # Calculate total cost for this sale
                    price_per_stick = float(cigar.get('price_per_stick', 0))
                    total_cost = price_per_stick * quantity

                    # Create sale record
                    sale_record = {
                        'date': sale_date,
                        'brand': cigar.get('brand', ''),
                        'cigar': cigar_name,
                        'size': cigar.get('size', ''),
                        'price_per_stick': price_per_stick,
                        'quantity': quantity,
                        'total_cost': total_cost
                    }
                    sale_records.append(sale_record)
                    self.sales_history.append(sale_record)
                    
                    # Decrease the count by quantity
                    cigar['count'] = current_count - quantity
                    selected_cigars.append((cigar_name, quantity))
                else:
                    messagebox.showwarning("Warning", f"Not enough stock for {cigar_name}. Only {current_count} available.")
        
        if selected_cigars:
            # Save both inventory and sales history
            self.save_inventory()
            self.save_sales_history()
            
            # Clear checkboxes and update displays
            self.checkbox_states = {}
            self.quantity_spinboxes.clear()
            
            # Refresh displays
            self.refresh_inventory()
            self.refresh_sales_history()
            self.update_selected_cigars_display()
            self.update_inventory_totals()
            
            # Show sale confirmation dialog
            self.show_sale_confirmation(sale_records, selected_cigars)
        else:
            messagebox.showwarning("Warning", "No cigars selected for sale or insufficient stock")

    def show_sales_history(self):
        """Display the sales history in a new window."""
        history_window = tk.Toplevel(self.root)
        history_window.title("Sales History")
        history_window.geometry("800x600")
        
        # Create treeview for sales history
        columns = ('date', 'brand', 'cigar', 'size', 'price_per_stick')
        tree = ttk.Treeview(history_window, columns=columns, show='headings')
        
        # Configure columns
        headings = {
            'date': 'Date',
            'brand': 'Brand',
            'cigar': 'Cigar',
            'size': 'Size',
            'price_per_stick': 'Price/Stick'
        }
        
        for col, heading in headings.items():
            tree.heading(col, text=heading)
            tree.column(col, width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(history_window, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack widgets
        tree.pack(side='left', expand=True, fill='both')
        scrollbar.pack(side='right', fill='y')
        
        # Load and display sales history
        try:
            self.load_sales_history()
            for sale in reversed(self.sales_history):  # Show newest first
                values = (
                    sale['date'],
                    sale['brand'],
                    sale['cigar'],
                    sale['size'],
                    f"${sale['price_per_stick']:.2f}"
                )
                tree.insert('', 'end', values=values)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sales history: {str(e)}")

    def save_sales_history(self):
        """Save sales history to JSON file."""
        try:
            with open(resource_path('sales_history.json'), 'w') as f:
                json.dump(self.sales_history, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sales history: {str(e)}")

    def load_sales_history(self):
        """Load sales history from JSON file."""
        try:
            with open(resource_path('sales_history.json'), 'r') as f:
                self.sales_history = json.load(f)
                
            # Update old format records to new format
            for sale in self.sales_history:
                if 'quantity' not in sale:
                    sale['quantity'] = 1
                if 'total_cost' not in sale:
                    sale['total_cost'] = float(sale['price_per_stick'])
                    
        except FileNotFoundError:
            self.sales_history = []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sales history: {str(e)}")
            self.sales_history = []

    def refresh_sales_history(self):
        """Refresh the sales history display."""
        # Clear current display
        for item in self.sales_tree.get_children():
            self.sales_tree.delete(item)
            
        # Add all sales records
        for sale in reversed(self.sales_history):  # Show newest first
            try:
                # Ensure all required fields exist
                quantity = sale.get('quantity', 1)
                price_per_stick = float(sale.get('price_per_stick', 0))
                total_cost = sale.get('total_cost', price_per_stick * quantity)
                
                values = (
                    sale.get('date', ''),
                    sale.get('brand', ''),
                    sale.get('cigar', ''),
                    sale.get('size', ''),
                    f"${price_per_stick:.2f}",
                    str(quantity),
                    f"${total_cost:.2f}"
                )
                self.sales_tree.insert('', 'end', values=values)
            except Exception as e:
                print(f"Error displaying sale record: {e}")

    def manual_save(self):
        """Manually save all data and show confirmation."""
        try:
            # Save inventory
            self.save_inventory()
            
            # Save brands, sizes, and types
            self.save_brands()
            self.save_sets('cigar_sizes.json', self.sizes)
            self.save_sets('cigar_types.json', self.types)
            
            # Save sales history
            self.save_sales_history()
            
            messagebox.showinfo("Success", "All data has been saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")

def main():
    try:
        root = tk.Tk()
        app = CigarInventory(root)
        root.mainloop()
    except KeyboardInterrupt:
        print("Application terminated by user")
    except Exception as e:
        print(f"An error occurred: {str(e)}")
    finally:
        try:
            root.destroy()
        except:
            pass

if __name__ == "__main__":
    main()