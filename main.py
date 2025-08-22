import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import json
import os
from datetime import datetime
import csv
import pandas as pd
import sys
import uuid  # Add this import for generating unique transaction IDs

# Try to import modern themes
try:
    from tkinter import ttk
    from tkinter.ttk import Style
    MODERN_THEME_AVAILABLE = True
except ImportError:
    MODERN_THEME_AVAILABLE = False

def resource_path(relative_path):
    """Get absolute path to resource, works for dev and for PyInstaller"""
    try:
        # For non-JSON files (like icons), use the application directory
        if not relative_path.endswith('.json'):
            try:
                base_path = sys._MEIPASS
            except Exception:
                base_path = os.path.abspath(".")
            return os.path.join(base_path, relative_path)
        
        # For JSON files, we'll use the instance's data directory
        # This is a temporary fallback - the actual method will be called from the instance
        return os.path.join(os.path.abspath("."), relative_path)
            
    except Exception as e:
        print(f"Error in resource_path: {e}")
        return relative_path

class CigarInventory:
    def __init__(self, root):
        self.root = root
        self.root.title("Cigar Inventory Manager")
        self.root.geometry("1500x900")  # Made taller for calculator
        
        # Initialize data directory (start with project directory for safety)
        self.data_directory = os.path.abspath(".")
        self.humidor_name = "Default"
        
        # Apply modern styling
        self.setup_modern_theme()
        
        # Create menu bar
        self.setup_menu_bar()
        
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
        self.resupply_history = []
        self.stored_quantities = {}  # New dictionary to store quantities persistently
        
        # Create main container
        main_container = ttk.Frame(root)
        main_container.pack(expand=True, fill='both', padx=10, pady=5)
        
        # Add file location bar at the top
        self.setup_file_location_bar(main_container)
        
        # Create notebook for tabs with modern styling
        self.notebook = ttk.Notebook(main_container, style='Modern.TNotebook')
        self.notebook.pack(expand=True, fill='both', pady=(5, 0))
        
        # Create main frame for inventory tab
        self.main_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.main_frame, text='Inventory')
        
        # Create frame for resupply tab
        self.resupply_tab_frame = ttk.Frame(self.notebook)
        self.notebook.add(self.resupply_tab_frame, text='Resupply')
        
        # Setup UI
        self.setup_inventory_tab()
        self.setup_resupply_tab()
        
        # Update location display
        self.update_location_display()
        
        # Load and display inventory after UI is ready
        self.load_inventory()
        self.refresh_inventory()
        self.refresh_resupply_dropdowns()
        
        # Ensure proper cleanup
        self.root.protocol("WM_DELETE_WINDOW", self.on_closing)

    def setup_modern_theme(self):
        """Apply modern styling to make the app look more contemporary."""
        # Configure the style
        style = ttk.Style()
        
        # Set the theme to a more modern one if available
        available_themes = style.theme_names()
        modern_themes = ['vista', 'xpnative', 'winnative', 'clam']
        
        for theme in modern_themes:
            if theme in available_themes:
                style.theme_use(theme)
                break
        
        # Modern color scheme
        bg_color = "#f8f9fa"  # Light gray background
        accent_color = "#007bff"  # Modern blue
        text_color = "#212529"  # Dark gray text
        border_color = "#dee2e6"  # Light border
        
        # Configure root window
        self.root.configure(bg=bg_color)
        
        # Configure modern button style
        style.configure('Modern.TButton',
                       background='#007bff',
                       foreground='white',
                       borderwidth=0,
                       focuscolor='none',
                       padding=(10, 8))
        
        style.map('Modern.TButton',
                 background=[('active', '#0056b3'),
                           ('pressed', '#004085')])
        
        # Configure modern label frame style  
        style.configure('Modern.TLabelframe',
                       background=bg_color,
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Modern.TLabelframe.Label',
                       background=bg_color,
                       foreground=accent_color,
                       font=('Segoe UI', 10, 'bold'))
        
        # Configure modern treeview style
        style.configure('Modern.Treeview',
                       background='white',
                       foreground=text_color,
                       rowheight=28,
                       fieldbackground='white',
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Modern.Treeview.Heading',
                       background='#e9ecef',
                       foreground=text_color,
                       font=('Segoe UI', 9, 'bold'),
                       borderwidth=1,
                       relief='solid')
        
        # Configure modern entry style
        style.configure('Modern.TEntry',
                       borderwidth=1,
                       relief='solid',
                       padding=(8, 6))
        
        # Configure modern notebook style
        style.configure('Modern.TNotebook',
                       background=bg_color,
                       borderwidth=0)
        
        style.configure('Modern.TNotebook.Tab',
                       background='#e9ecef',
                       foreground=text_color,
                       padding=(20, 10),
                       font=('Segoe UI', 10))
        
        style.map('Modern.TNotebook.Tab',
                 background=[('selected', 'white'),
                           ('active', '#f8f9fa')])

    def setup_menu_bar(self):
        """Create a professional menu bar with File, Tools, Help menus."""
        # Create menu bar
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # File menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="File", menu=file_menu)
        
        # Add file operations
        file_menu.add_command(label="New Humidor...", command=self.new_humidor, accelerator="Ctrl+N")
        file_menu.add_command(label="Load Humidor...", command=self.load_humidor, accelerator="Ctrl+O")
        file_menu.add_separator()
        file_menu.add_command(label="Change Data Directory...", command=self.change_data_directory)
        file_menu.add_separator()
        file_menu.add_command(label="Backup Data...", command=self.backup_data, accelerator="Ctrl+B")
        file_menu.add_command(label="Export Inventory...", command=self.export_inventory, accelerator="Ctrl+E")
        file_menu.add_separator()
        file_menu.add_command(label="Save", command=self.manual_save, accelerator="Ctrl+S")
        file_menu.add_separator()
        file_menu.add_command(label="Exit", command=self.on_closing, accelerator="Alt+F4")
        
        # Tools menu
        tools_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Tools", menu=tools_menu)
        tools_menu.add_command(label="Shipping Calculator", command=self.show_shipping_calculator)
        tools_menu.add_separator()
        tools_menu.add_command(label="Sales History...", command=self.show_sales_history_window)
        tools_menu.add_command(label="Resupply History...", command=self.show_resupply_history_window)
        tools_menu.add_separator()
        tools_menu.add_command(label="Add New Brand...", command=self.add_new_brand)
        tools_menu.add_command(label="Add New Size...", command=self.add_new_size)
        tools_menu.add_command(label="Add New Type...", command=self.add_new_type)
        
        # Help menu (for future expansion)
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="About Cigar Inventory Manager", command=self.show_about)
        
        # Bind keyboard shortcuts
        self.root.bind('<Control-n>', lambda e: self.new_humidor())
        self.root.bind('<Control-o>', lambda e: self.load_humidor())
        self.root.bind('<Control-s>', lambda e: self.manual_save())
        self.root.bind('<Control-b>', lambda e: self.backup_data())
        self.root.bind('<Control-e>', lambda e: self.export_inventory())

    def show_shipping_calculator(self):
        """Show a standalone shipping calculator dialog."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Shipping Calculator")
        dialog.geometry("400x480")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (200)
        y = (dialog.winfo_screenheight() // 2) - (240)
        dialog.geometry(f"400x480+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        # Header
        ttk.Label(main_frame, text="Shipping Calculator", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=(0, 20))
        
        # Input fields
        input_frame = ttk.LabelFrame(main_frame, text="Order Details", padding="15")
        input_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Label(input_frame, text="Total Shipping Cost ($):").pack(anchor='w', pady=(0, 5))
        ship_cost_var = tk.StringVar()
        ship_cost_entry = ttk.Entry(input_frame, textvariable=ship_cost_var, width=20)
        ship_cost_entry.pack(fill='x', pady=(0, 10))
        
        ttk.Label(input_frame, text="Total Cigars in Order:").pack(anchor='w', pady=(0, 5))
        total_cigars_var = tk.StringVar()
        total_cigars_entry = ttk.Entry(input_frame, textvariable=total_cigars_var, width=20)
        total_cigars_entry.pack(fill='x', pady=(0, 10))
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Results", padding="15")
        results_frame.pack(fill='x', pady=(0, 15))
        
        # Result labels
        per_stick_label = ttk.Label(results_frame, text="Per Stick Shipping: $0.00")
        per_stick_label.pack(anchor='w', pady=2)
        
        five_pack_label = ttk.Label(results_frame, text="5-Pack Total Shipping: $0.00")
        five_pack_label.pack(anchor='w', pady=2)
        
        ten_pack_label = ttk.Label(results_frame, text="10-Pack Total Shipping: $0.00")
        ten_pack_label.pack(anchor='w', pady=2)
        
        def calculate_shipping():
            try:
                shipping = float(ship_cost_var.get() or 0)
                total_cigars = int(total_cigars_var.get() or 0)
                
                if total_cigars <= 0:
                    messagebox.showerror("Error", "Total cigars must be greater than 0")
                    return
                
                # Calculate shipping cost per stick
                per_stick_cost = shipping / total_cigars
                
                # Calculate total shipping costs for different quantities
                five_pack_total = per_stick_cost * 5
                ten_pack_total = per_stick_cost * 10
                
                # Update labels
                per_stick_label.config(text=f"Per Stick Shipping: ${per_stick_cost:.2f}")
                five_pack_label.config(text=f"5-Pack Total Shipping: ${five_pack_total:.2f}")
                ten_pack_label.config(text=f"10-Pack Total Shipping: ${ten_pack_total:.2f}")
                
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numbers")
            except ZeroDivisionError:
                messagebox.showerror("Error", "Total cigars must be greater than 0")
        
        # Calculate button
        ttk.Button(input_frame, text="Calculate", 
                  command=calculate_shipping).pack(pady=(5, 0))
        
        # Close button
        ttk.Button(main_frame, text="Close", command=dialog.destroy, width=15).pack(pady=(10, 0))
        
        # Focus on first field
        ship_cost_entry.focus()

    def setup_resupply_tab(self):
        """Setup the resupply tab for adding inventory."""
        # Main container
        main_frame = ttk.Frame(self.resupply_tab_frame)
        main_frame.pack(expand=True, fill='both', padx=10, pady=10)
        
        # Header
        ttk.Label(main_frame, text="Resupply Inventory", 
                 font=('TkDefaultFont', 14, 'bold')).pack(pady=(0, 20))
        
        # Create a horizontal paned window
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(expand=True, fill='both')
        
        # Left frame for order building
        left_frame = ttk.Frame(paned_window)
        paned_window.add(left_frame, weight=2)
        
        # Right frame for order totals and processing
        right_frame = ttk.Frame(paned_window)
        paned_window.add(right_frame, weight=1)
        
        # === LEFT FRAME: Order Building ===
        # Order totals frame
        totals_frame = ttk.LabelFrame(left_frame, text="Order Totals", padding="10")
        totals_frame.pack(fill='x', pady=(0, 15))
        
        # Create grid for order totals
        ttk.Label(totals_frame, text="Total Order Shipping ($):").grid(row=0, column=0, sticky='w', padx=(0, 10))
        self.resupply_total_shipping_var = tk.StringVar()
        self.resupply_total_shipping_entry = ttk.Entry(totals_frame, textvariable=self.resupply_total_shipping_var, width=15)
        self.resupply_total_shipping_entry.grid(row=0, column=1, padx=(0, 20))
        
        # Tax rate setting
        ttk.Label(totals_frame, text="Tax Rate (%):").grid(row=0, column=2, sticky='w', padx=(0, 10))
        self.resupply_tax_rate_var = tk.StringVar(value=str(self.tax_rate * 100))
        self.resupply_tax_rate_entry = ttk.Entry(totals_frame, textvariable=self.resupply_tax_rate_var, width=10)
        self.resupply_tax_rate_entry.grid(row=0, column=3, padx=(0, 20))
        
        # Total tax display
        ttk.Label(totals_frame, text="Total Tax ($):").grid(row=0, column=4, sticky='w', padx=(0, 10))
        self.resupply_total_tax_label = ttk.Label(totals_frame, text="$0.00", font=('TkDefaultFont', 9, 'bold'))
        self.resupply_total_tax_label.grid(row=0, column=5)
        
        # Cigars in order frame
        cigars_frame = ttk.LabelFrame(left_frame, text="Cigars in Order", padding="10")
        cigars_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        # Create treeview for cigars in order
        columns = ('brand', 'cigar', 'size', 'type', 'count', 'price', 'calc_shipping', 'calc_tax', 'price_per_stick')
        self.resupply_cigars_tree = ttk.Treeview(cigars_frame, columns=columns, show='headings', height=12)
        
        # Configure columns
        self.resupply_cigars_tree.heading('brand', text='Brand')
        self.resupply_cigars_tree.heading('cigar', text='Cigar')
        self.resupply_cigars_tree.heading('size', text='Size')
        self.resupply_cigars_tree.heading('type', text='Type')
        self.resupply_cigars_tree.heading('count', text='Count')
        self.resupply_cigars_tree.heading('price', text='Price')
        self.resupply_cigars_tree.heading('calc_shipping', text='Shipping')
        self.resupply_cigars_tree.heading('calc_tax', text='Tax')
        self.resupply_cigars_tree.heading('price_per_stick', text='Price/Stick')
        
        self.resupply_cigars_tree.column('brand', width=100)
        self.resupply_cigars_tree.column('cigar', width=120)
        self.resupply_cigars_tree.column('size', width=70)
        self.resupply_cigars_tree.column('type', width=80)
        self.resupply_cigars_tree.column('count', width=60)
        self.resupply_cigars_tree.column('price', width=80)
        self.resupply_cigars_tree.column('calc_shipping', width=80)
        self.resupply_cigars_tree.column('calc_tax', width=70)
        self.resupply_cigars_tree.column('price_per_stick', width=90)
        
        # Add scrollbar
        resupply_cigars_scrollbar = ttk.Scrollbar(cigars_frame, orient='vertical', command=self.resupply_cigars_tree.yview)
        self.resupply_cigars_tree.configure(yscrollcommand=resupply_cigars_scrollbar.set)
        
        # Pack tree and scrollbar
        self.resupply_cigars_tree.pack(side='left', fill='both', expand=True)
        resupply_cigars_scrollbar.pack(side='right', fill='y')
        
        # Initialize order storage
        self.current_resupply_order = []
        
        # === RIGHT FRAME: Add Item and Controls ===
        # Add item frame
        add_frame = ttk.LabelFrame(right_frame, text="Add Item to Order", padding="15")
        add_frame.pack(fill='x', pady=(0, 15))
        
        # Input fields
        ttk.Label(add_frame, text="Brand:").pack(anchor='w', pady=(0, 5))
        self.resupply_brand_var = tk.StringVar()
        self.resupply_brand_combo = ttk.Combobox(add_frame, textvariable=self.resupply_brand_var, 
                                               values=sorted(list(self.brands)), width=25)
        self.resupply_brand_combo.pack(fill='x', pady=(0, 10))
        
        ttk.Label(add_frame, text="Cigar:").pack(anchor='w', pady=(0, 5))
        self.resupply_cigar_var = tk.StringVar()
        existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
        self.resupply_cigar_combo = ttk.Combobox(add_frame, textvariable=self.resupply_cigar_var, 
                                               values=existing_cigars, width=25)
        self.resupply_cigar_combo.pack(fill='x', pady=(0, 10))
        
        ttk.Label(add_frame, text="Size:").pack(anchor='w', pady=(0, 5))
        self.resupply_size_var = tk.StringVar()
        self.resupply_size_combo = ttk.Combobox(add_frame, textvariable=self.resupply_size_var, 
                                              values=sorted(list(self.sizes)), width=25)
        self.resupply_size_combo.pack(fill='x', pady=(0, 10))
        
        ttk.Label(add_frame, text="Type:").pack(anchor='w', pady=(0, 5))
        self.resupply_type_var = tk.StringVar()
        self.resupply_type_combo = ttk.Combobox(add_frame, textvariable=self.resupply_type_var, 
                                              values=sorted(list(self.types)), width=25)
        self.resupply_type_combo.pack(fill='x', pady=(0, 10))
        
        ttk.Label(add_frame, text="Count:").pack(anchor='w', pady=(0, 5))
        self.resupply_count_var = tk.StringVar()
        self.resupply_count_entry = ttk.Entry(add_frame, textvariable=self.resupply_count_var, width=25)
        self.resupply_count_entry.pack(fill='x', pady=(0, 10))
        
        ttk.Label(add_frame, text="Price ($):").pack(anchor='w', pady=(0, 5))
        self.resupply_price_var = tk.StringVar()
        self.resupply_price_entry = ttk.Entry(add_frame, textvariable=self.resupply_price_var, width=25)
        self.resupply_price_entry.pack(fill='x', pady=(0, 15))
        
        # Add to order button
        ttk.Button(add_frame, text="Add to Order", command=self.add_to_resupply_order, 
                  width=25).pack(pady=(0, 10))
        
        # Order management buttons
        management_frame = ttk.LabelFrame(right_frame, text="Order Management", padding="15")
        management_frame.pack(fill='x', pady=(0, 15))
        
        ttk.Button(management_frame, text="Remove Selected", 
                  command=self.remove_from_resupply_order, width=25).pack(pady=(0, 5))
        ttk.Button(management_frame, text="Clear Order", 
                  command=self.clear_resupply_order, width=25).pack(pady=(0, 5))
        ttk.Button(management_frame, text="Process Order", 
                  command=self.process_resupply_order, width=25).pack(pady=(0, 5))
        
        # Order summary
        summary_frame = ttk.LabelFrame(right_frame, text="Order Summary", padding="15")
        summary_frame.pack(fill='x')
        
        self.resupply_summary_items_label = ttk.Label(summary_frame, text="Total Items: 0")
        self.resupply_summary_items_label.pack(anchor='w', pady=2)
        
        self.resupply_summary_shipping_label = ttk.Label(summary_frame, text="Total Shipping Allocated: $0.00")
        self.resupply_summary_shipping_label.pack(anchor='w', pady=2)
        
        self.resupply_summary_tax_label = ttk.Label(summary_frame, text="Total Tax Calculated: $0.00")
        self.resupply_summary_tax_label.pack(anchor='w', pady=2)
        
        # Bind events for automatic calculation
        self.resupply_total_shipping_var.trace('w', lambda *args: self.calculate_resupply_costs())
        self.resupply_tax_rate_var.trace('w', lambda *args: self.calculate_resupply_costs())
        
        # Bind brand change to update cigar dropdown
        self.resupply_brand_var.trace('w', self.update_resupply_cigar_dropdown)

    def refresh_resupply_dropdowns(self):
        """Refresh the resupply tab dropdown values after data is loaded."""
        # Update brand dropdown
        self.resupply_brand_combo.config(values=sorted(list(self.brands)))
        
        # Update cigar dropdown with all existing cigars
        existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
        self.resupply_cigar_combo.config(values=existing_cigars)
        
        # Update size dropdown
        self.resupply_size_combo.config(values=sorted(list(self.sizes)))
        
        # Update type dropdown  
        self.resupply_type_combo.config(values=sorted(list(self.types)))

    def show_sales_history_window(self):
        """Show sales history in a separate window."""
        window = tk.Toplevel(self.root)
        window.title("Sales History")
        window.geometry("1200x700")
        window.transient(self.root)
        
        # Center the window
        window.update_idletasks()
        x = (window.winfo_screenwidth() // 2) - (600)
        y = (window.winfo_screenheight() // 2) - (350)
        window.geometry(f"1200x700+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Create a PanedWindow for resizable layout
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(fill='both', expand=True)

        # Left frame for transaction list
        left_frame = ttk.LabelFrame(paned_window, text="Sales Transactions", padding="5")
        paned_window.add(left_frame, weight=1)

        # Right frame for transaction details and actions
        right_frame = ttk.LabelFrame(paned_window, text="Transaction Details & Returns", padding="5")
        paned_window.add(right_frame, weight=1)

        # === LEFT FRAME: Transaction List ===
        trans_columns = ('date', 'items', 'total')
        transaction_tree = ttk.Treeview(left_frame, columns=trans_columns, show='headings', height=15)
        
        transaction_tree.heading('date', text='Date & Time')
        transaction_tree.heading('items', text='Items Sold')
        transaction_tree.heading('total', text='Total Value')
        
        transaction_tree.column('date', width=150)
        transaction_tree.column('items', width=80)
        transaction_tree.column('total', width=100)

        trans_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=transaction_tree.yview)
        transaction_tree.configure(yscrollcommand=trans_scrollbar.set)
        
        transaction_tree.pack(side='left', fill='both', expand=True)
        trans_scrollbar.pack(side='right', fill='y')

        # === RIGHT FRAME: Transaction Details ===
        detail_columns = ('cigar', 'brand', 'size', 'quantity', 'price_per_stick', 'total_cost')
        detail_tree = ttk.Treeview(right_frame, columns=detail_columns, show='headings', height=10)
        
        detail_tree.heading('cigar', text='Cigar')
        detail_tree.heading('brand', text='Brand')
        detail_tree.heading('size', text='Size')
        detail_tree.heading('quantity', text='Qty')
        detail_tree.heading('price_per_stick', text='Price/Stick')
        detail_tree.heading('total_cost', text='Total')
        
        detail_tree.column('cigar', width=120)
        detail_tree.column('brand', width=100)
        detail_tree.column('size', width=70)
        detail_tree.column('quantity', width=50)
        detail_tree.column('price_per_stick', width=80)
        detail_tree.column('total_cost', width=80)

        detail_scrollbar = ttk.Scrollbar(right_frame, orient='vertical', command=detail_tree.yview)
        detail_tree.configure(yscrollcommand=detail_scrollbar.set)
        
        detail_tree.pack(side='top', fill='both', expand=True)
        detail_scrollbar.pack(side='right', fill='y')

        # Local variables for this window
        current_transaction_id = None
        transaction_id_map = {}

        def refresh_local_display():
            """Refresh the local window display."""
            # Clear transaction display
            for item in transaction_tree.get_children():
                transaction_tree.delete(item)
            detail_tree.delete(*detail_tree.get_children())
            
            # Reload and redisplay transactions
            self.load_sales_history()
            
            # Group sales by transaction_id
            transactions = {}
            for sale in self.sales_history:
                transaction_id = sale.get('transaction_id', 'unknown')
                if transaction_id not in transactions:
                    transactions[transaction_id] = []
                transactions[transaction_id].append(sale)
            
            # Clear old mapping
            transaction_id_map.clear()
            
            # Display transactions (newest first)
            sorted_transactions = sorted(
                transactions.items(), 
                key=lambda x: x[1][0].get('date', ''), 
                reverse=True
            )
            
            for transaction_id, sales in sorted_transactions:
                try:
                    total_items = sum(int(sale.get('quantity', 1)) for sale in sales)
                    total_value = sum(float(sale.get('total_cost', 0)) for sale in sales)
                    date = sales[0].get('date', 'Unknown')
                    
                    values = (date, str(total_items), f"${total_value:.2f}")
                    item_id = transaction_tree.insert('', 'end', values=values)
                    transaction_id_map[item_id] = transaction_id
                except Exception as e:
                    print(f"Error displaying transaction {transaction_id}: {e}")
        
        def refresh_transaction_details():
            """Refresh details for selected transaction."""
            detail_tree.delete(*detail_tree.get_children())
            if not current_transaction_id:
                return
            
            for sale in self.sales_history:
                if sale.get('transaction_id') == current_transaction_id:
                    values = (
                        sale.get('cigar', 'Unknown'),
                        sale.get('brand', 'Unknown'),
                        sale.get('size', 'N/A'),
                        str(sale.get('quantity', 1)),
                        f"${float(sale.get('price_per_stick', 0)):.2f}",
                        f"${float(sale.get('total_cost', 0)):.2f}"
                    )
                    detail_tree.insert('', 'end', values=values)

        def return_selected_items_local():
            """Handle partial return of selected items from the current transaction."""
            selected_items = detail_tree.selection()
            if not selected_items:
                messagebox.showwarning("Warning", "Please select items to return.")
                return

            if not current_transaction_id:
                messagebox.showwarning("Warning", "No transaction selected.")
                return

            # Collect information about selected items
            items_data = []
            for item_id in selected_items:
                values = detail_tree.item(item_id)['values']
                if values:
                    cigar_name = values[0]
                    brand = values[1]
                    current_qty = int(values[3])
                    items_data.append((cigar_name, brand, current_qty))

            if not items_data:
                return

            # Show single dialog for all return quantities
            return_quantities = self.ask_multiple_return_quantities(items_data)
            
            if not return_quantities:
                return  # User cancelled or no items to return

            # Process returns directly (no confirmation dialog)
            total_items = 0
            for cigar_name, (brand, return_qty, max_qty) in return_quantities.items():
                # Add back to inventory
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        cigar['count'] = int(cigar['count']) + return_qty
                        break

                # Update or remove sales records
                for sale in self.sales_history:
                    if (sale.get('transaction_id') == current_transaction_id and 
                        sale.get('cigar') == cigar_name):
                        
                        current_sale_qty = int(sale.get('quantity', 1))
                        if return_qty >= current_sale_qty:
                            # Remove entire sale record
                            self.sales_history.remove(sale)
                        else:
                            # Reduce quantity and recalculate cost
                            new_qty = current_sale_qty - return_qty
                            sale['quantity'] = new_qty
                            price_per_stick = float(sale.get('price_per_stick', 0))
                            sale['total_cost'] = price_per_stick * new_qty
                        break
                
                total_items += return_qty

            # Save changes and refresh displays
            self.save_inventory()
            self.save_sales_history()
            self.refresh_inventory()
            self.update_inventory_totals()
            
            # Refresh both this window and main application
            refresh_local_display()
            
            # Show simple success message
            messagebox.showinfo("Success", f"Successfully returned {total_items} items.")

        def return_entire_transaction_local():
            """Handle returning the entire current transaction."""
            if not current_transaction_id:
                messagebox.showwarning("Warning", "No transaction selected.")
                return

            # Get all items in the transaction
            transaction_items = []
            total_items = 0
            total_value = 0.0
            
            for sale in self.sales_history:
                if sale.get('transaction_id') == current_transaction_id:
                    quantity = int(sale.get('quantity', 1))
                    total_cost = float(sale.get('total_cost', 0))
                    transaction_items.append((sale.get('cigar'), sale.get('brand'), quantity))
                    total_items += quantity
                    total_value += total_cost

            if not transaction_items:
                messagebox.showwarning("Warning", "No items found in this transaction.")
                return

            # Get transaction date for confirmation
            transaction_date = "Unknown"
            for sale in self.sales_history:
                if sale.get('transaction_id') == current_transaction_id:
                    transaction_date = sale.get('date', 'Unknown')
                    break

            # Single confirmation for entire transaction
            confirm_msg = f"Are you sure you want to return the entire transaction?\n\n"
            confirm_msg += f"Date: {transaction_date}\n"
            confirm_msg += f"Total items: {total_items}\n"
            confirm_msg += f"Total value: ${total_value:.2f}\n\n"
            confirm_msg += "Items to return:\n"
            for cigar_name, brand, quantity in transaction_items:
                confirm_msg += f"• {quantity}x {brand} - {cigar_name}\n"

            if messagebox.askyesno("Confirm Return", confirm_msg):
                # Return all items to inventory
                for cigar_name, _, quantity in transaction_items:
                    for cigar in self.inventory:
                        if cigar['cigar'] == cigar_name:
                            cigar['count'] = int(cigar['count']) + quantity
                            break

                # Remove all sales records for this transaction
                self.sales_history = [sale for sale in self.sales_history 
                                    if sale.get('transaction_id') != current_transaction_id]

                # Save changes and refresh displays
                self.save_inventory()
                self.save_sales_history()
                self.refresh_inventory()
                self.update_inventory_totals()
                
                # Refresh both this window and main application
                refresh_local_display()
                
                messagebox.showinfo("Success", f"Successfully returned entire transaction ({total_items} items).")

        # === ACTION BUTTONS ===
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill='x', pady=(10, 0))

        # Return buttons
        ttk.Button(button_frame, text="Return Selected Item(s)", 
                  command=return_selected_items_local, width=25).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Return Entire Transaction", 
                  command=return_entire_transaction_local, width=25).pack(side='left', padx=5)
        
        def on_transaction_select(event):
            """Handle transaction selection."""
            nonlocal current_transaction_id
            selected_item = transaction_tree.selection()
            if not selected_item:
                current_transaction_id = None
                detail_tree.delete(*detail_tree.get_children())
                return
            
            selected_item = selected_item[0]
            current_transaction_id = transaction_id_map.get(selected_item)
            refresh_transaction_details()
        
        # Bind selection event
        transaction_tree.bind('<<TreeviewSelect>>', on_transaction_select)
        
        # Load and display sales history
        self.load_sales_history()  # Make sure data is loaded
        
        # Group sales by transaction_id
        transactions = {}
        for sale in self.sales_history:
            transaction_id = sale.get('transaction_id', 'unknown')
            if transaction_id not in transactions:
                transactions[transaction_id] = []
            transactions[transaction_id].append(sale)
        
        # Display transactions (newest first)
        sorted_transactions = sorted(
            transactions.items(), 
            key=lambda x: x[1][0].get('date', ''), 
            reverse=True
        )
        
        for transaction_id, sales in sorted_transactions:
            try:
                total_items = sum(int(sale.get('quantity', 1)) for sale in sales)
                total_value = sum(float(sale.get('total_cost', 0)) for sale in sales)
                date = sales[0].get('date', 'Unknown')
                
                values = (date, str(total_items), f"${total_value:.2f}")
                item_id = transaction_tree.insert('', 'end', values=values)
                transaction_id_map[item_id] = transaction_id
            except Exception as e:
                print(f"Error displaying transaction {transaction_id}: {e}")

    def show_resupply_history_window(self):
        """Show resupply history in a separate window."""
        window = tk.Toplevel(self.root)
        window.title("Resupply History")
        window.geometry("1200x700")
        window.transient(self.root)
        
        # Center the window
        window.update_idletasks()
        x = (window.winfo_screenwidth() // 2) - (600)
        y = (window.winfo_screenheight() // 2) - (350)
        window.geometry(f"1200x700+{x}+{y}")
        
        # Create main frame
        main_frame = ttk.Frame(window, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Create a PanedWindow for resizable layout
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(fill='both', expand=True)

        # Left frame for resupply list
        left_frame = ttk.LabelFrame(paned_window, text="Resupply Orders", padding="5")
        paned_window.add(left_frame, weight=1)

        # Right frame for order details
        right_frame = ttk.LabelFrame(paned_window, text="Order Details", padding="5")
        paned_window.add(right_frame, weight=1)

        # === LEFT FRAME: Resupply Orders List ===
        resupply_columns = ('date', 'items', 'total_cost', 'total_shipping')
        resupply_tree = ttk.Treeview(left_frame, columns=resupply_columns, show='headings', height=15)
        
        resupply_tree.heading('date', text='Date & Time')
        resupply_tree.heading('items', text='Items')
        resupply_tree.heading('total_cost', text='Total Cost')
        resupply_tree.heading('total_shipping', text='Shipping')
        
        resupply_tree.column('date', width=150)
        resupply_tree.column('items', width=80)
        resupply_tree.column('total_cost', width=100)
        resupply_tree.column('total_shipping', width=100)

        resupply_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=resupply_tree.yview)
        resupply_tree.configure(yscrollcommand=resupply_scrollbar.set)
        
        resupply_tree.pack(side='left', fill='both', expand=True)
        resupply_scrollbar.pack(side='right', fill='y')

        # === RIGHT FRAME: Order Details ===
        detail_columns = ('brand', 'cigar', 'size', 'type', 'quantity', 'price', 'shipping', 'total_cost')
        resupply_detail_tree = ttk.Treeview(right_frame, columns=detail_columns, show='headings', height=15)
        
        resupply_detail_tree.heading('brand', text='Brand')
        resupply_detail_tree.heading('cigar', text='Cigar')
        resupply_detail_tree.heading('size', text='Size')
        resupply_detail_tree.heading('type', text='Type')
        resupply_detail_tree.heading('quantity', text='Qty')
        resupply_detail_tree.heading('price', text='Price')
        resupply_detail_tree.heading('shipping', text='Ship+Tax')
        resupply_detail_tree.heading('total_cost', text='Total')
        
        resupply_detail_tree.column('brand', width=100)
        resupply_detail_tree.column('cigar', width=120)
        resupply_detail_tree.column('size', width=70)
        resupply_detail_tree.column('type', width=80)
        resupply_detail_tree.column('quantity', width=50)
        resupply_detail_tree.column('price', width=80)
        resupply_detail_tree.column('shipping', width=80)
        resupply_detail_tree.column('total_cost', width=80)

        resupply_detail_scrollbar = ttk.Scrollbar(right_frame, orient='vertical', command=resupply_detail_tree.yview)
        resupply_detail_tree.configure(yscrollcommand=resupply_detail_scrollbar.set)
        
        resupply_detail_tree.pack(side='top', fill='both', expand=True)
        resupply_detail_scrollbar.pack(side='right', fill='y')

        # Local variables for this window
        current_resupply_id = None
        resupply_id_map = {}

        def delete_resupply_order():
            """Delete the selected resupply order and remove cigars from inventory."""
            if not current_resupply_id:
                messagebox.showwarning("Warning", "No order selected.")
                return

            # Get order details for confirmation
            order_items = []
            total_items = 0
            total_cost = 0.0
            order_date = "Unknown"
            
            for resupply in self.resupply_history:
                if resupply.get('order_id') == current_resupply_id:
                    quantity = int(resupply.get('quantity', 1))
                    cost = float(resupply.get('total_cost', 0))
                    brand = resupply.get('brand')
                    cigar_name = resupply.get('cigar')
                    size = resupply.get('size')
                    order_items.append((brand, cigar_name, size, quantity))
                    total_items += quantity
                    total_cost += cost
                    if order_date == "Unknown":
                        order_date = resupply.get('date', 'Unknown')

            if not order_items:
                messagebox.showwarning("Warning", "No items found in this order.")
                return

            # Check inventory availability before confirming
            inventory_issues = []
            for brand, cigar_name, size, quantity in order_items:
                # Find matching cigar in inventory
                matching_cigar = None
                for cigar in self.inventory:
                    if (cigar.get('brand', '').lower() == brand.lower() and
                        cigar.get('cigar', '').lower() == cigar_name.lower() and
                        cigar.get('size', '').lower() == size.lower()):
                        matching_cigar = cigar
                        break
                
                if not matching_cigar:
                    inventory_issues.append(f"• {brand} - {cigar_name} ({size}) not found in inventory")
                elif int(matching_cigar.get('count', 0)) < quantity:
                    current_count = int(matching_cigar.get('count', 0))
                    inventory_issues.append(f"• {brand} - {cigar_name} ({size}): need {quantity}, only have {current_count}")

            confirm_msg = f"Are you sure you want to delete this resupply order?\n\n"
            confirm_msg += f"Date: {order_date}\n"
            confirm_msg += f"Total items: {total_items}\n"
            confirm_msg += f"Total cost: ${total_cost:.2f}\n\n"
            
            if inventory_issues:
                confirm_msg += "⚠️ INVENTORY WARNINGS:\n"
                for issue in inventory_issues:
                    confirm_msg += f"{issue}\n"
                confirm_msg += "\nThese items will be skipped during removal.\n\n"
            
            confirm_msg += "🔄 This will REMOVE these cigars from your inventory:\n"
            for brand, cigar_name, size, quantity in order_items:
                confirm_msg += f"• {quantity}x {brand} - {cigar_name} ({size})\n"

            if messagebox.askyesno("Confirm Delete & Remove", confirm_msg):
                # Remove cigars from inventory
                removed_items = []
                skipped_items = []
                
                for brand, cigar_name, size, quantity in order_items:
                    # Find matching cigar in inventory
                    for cigar in self.inventory:
                        if (cigar.get('brand', '').lower() == brand.lower() and
                            cigar.get('cigar', '').lower() == cigar_name.lower() and
                            cigar.get('size', '').lower() == size.lower()):
                            
                            current_count = int(cigar.get('count', 0))
                            if current_count >= quantity:
                                # Remove the quantity
                                cigar['count'] = current_count - quantity
                                removed_items.append(f"{quantity}x {brand} - {cigar_name} ({size})")
                                
                                # If count reaches 0, optionally remove the cigar entirely
                                # (keeping it with 0 count for now to preserve purchase history)
                            else:
                                skipped_items.append(f"{brand} - {cigar_name} ({size}): insufficient quantity")
                            break
                    else:
                        # Cigar not found in inventory
                        skipped_items.append(f"{brand} - {cigar_name} ({size}): not found in inventory")

                # Remove all resupply records for this order
                self.resupply_history = [resupply for resupply in self.resupply_history 
                                       if resupply.get('order_id') != current_resupply_id]

                # Save changes
                self.save_inventory()
                self.save_resupply_history()
                
                # Refresh displays
                refresh_resupply_local_display()
                self.refresh_inventory()
                self.update_inventory_totals()
                
                # Show detailed success message
                success_msg = "Resupply order deleted and cigars removed from inventory.\n\n"
                
                if removed_items:
                    success_msg += "✅ Removed from inventory:\n"
                    for item in removed_items:
                        success_msg += f"• {item}\n"
                
                if skipped_items:
                    success_msg += "\n⚠️ Skipped (not enough inventory):\n"
                    for item in skipped_items:
                        success_msg += f"• {item}\n"
                
                messagebox.showinfo("Success", success_msg.strip())

        def refresh_resupply_local_display():
            """Refresh the local resupply window display."""
            # Clear current display
            for item in resupply_tree.get_children():
                resupply_tree.delete(item)
            resupply_detail_tree.delete(*resupply_detail_tree.get_children())
            
            # Reload and redisplay orders
            self.load_resupply_history()
            
            # Group resupply records by order_id
            orders = {}
            for resupply in self.resupply_history:
                order_id = resupply.get('order_id', 'unknown')
                if order_id not in orders:
                    orders[order_id] = []
                orders[order_id].append(resupply)
            
            # Clear old mapping
            resupply_id_map.clear()
            
            # Display orders (newest first)
            sorted_orders = sorted(
                orders.items(), 
                key=lambda x: x[1][0].get('date', ''), 
                reverse=True
            )
            
            for order_id, resupplies in sorted_orders:
                try:
                    total_items = sum(int(resupply.get('quantity', 1)) for resupply in resupplies)
                    total_cost = sum(float(resupply.get('total_cost', 0)) for resupply in resupplies)
                    total_shipping = sum(float(resupply.get('shipping_tax', 0)) for resupply in resupplies)
                    date = resupplies[0].get('date', 'Unknown')
                    
                    values = (date, str(total_items), f"${total_cost:.2f}", f"${total_shipping:.2f}")
                    item_id = resupply_tree.insert('', 'end', values=values)
                    resupply_id_map[item_id] = order_id
                except Exception as e:
                    print(f"Error displaying resupply order {order_id}: {e}")

        # === ACTION BUTTONS ===
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill='x', pady=(10, 0))

        def return_selected_resupply_items():
            """Handle partial return of selected items from the current resupply order."""
            selected_items = resupply_detail_tree.selection()
            if not selected_items:
                messagebox.showwarning("Warning", "Please select items to return.")
                return

            if not current_resupply_id:
                messagebox.showwarning("Warning", "No order selected.")
                return

            # Collect information about selected items
            items_data = []
            for item_id in selected_items:
                values = resupply_detail_tree.item(item_id)['values']
                if values:
                    brand = values[0]
                    cigar_name = values[1] 
                    size = values[2]
                    current_qty = int(values[4])
                    display_name = f"{brand} - {cigar_name} ({size})"
                    items_data.append((display_name, brand, current_qty))

            if not items_data:
                return

            # Show dialog for return quantities
            return_quantities = self.ask_multiple_return_quantities(items_data)
            
            if not return_quantities:
                return  # User cancelled or no items to return

            # Process returns directly (remove from inventory)
            total_items = 0
            returned_items = []
            skipped_items = []
            
            for display_name, (brand_unused, return_qty, max_qty) in return_quantities.items():
                # Parse the display name to get individual components
                # Format: "Brand - Cigar (Size)"
                try:
                    # Extract brand, cigar, and size from display name
                    parts = display_name.split(' - ', 1)
                    if len(parts) >= 2:
                        brand_part = parts[0]
                        remainder = parts[1]
                        if '(' in remainder and ')' in remainder:
                            cigar_part = remainder.split(' (')[0]
                            size_part = remainder.split(' (')[1].replace(')', '')
                        else:
                            cigar_part = remainder
                            size_part = ''
                    else:
                        continue  # Skip malformed display names
                    
                    # Find matching cigar in inventory
                    inventory_found = False
                    for cigar in self.inventory:
                        if (cigar.get('brand', '').lower() == brand_part.lower() and
                            cigar.get('cigar', '').lower() == cigar_part.lower() and
                            cigar.get('size', '').lower() == size_part.lower()):
                            
                            current_inventory = int(cigar.get('count', 0))
                            if current_inventory >= return_qty:
                                # Remove from inventory
                                cigar['count'] = current_inventory - return_qty
                                returned_items.append(f"{return_qty}x {display_name}")
                                inventory_found = True
                            else:
                                skipped_items.append(f"{display_name}: only {current_inventory} in inventory, can't return {return_qty}")
                                inventory_found = True
                            break
                    
                    if not inventory_found:
                        skipped_items.append(f"{display_name}: not found in inventory")
                    
                    # Update resupply history record
                    for resupply in self.resupply_history:
                        if (resupply.get('order_id') == current_resupply_id and 
                            resupply.get('brand', '').lower() == brand_part.lower() and
                            resupply.get('cigar', '').lower() == cigar_part.lower() and
                            resupply.get('size', '').lower() == size_part.lower()):
                            
                            current_resupply_qty = int(resupply.get('quantity', 1))
                            if return_qty >= current_resupply_qty:
                                # Remove entire resupply record
                                self.resupply_history.remove(resupply)
                            else:
                                # Reduce quantity and recalculate cost
                                new_qty = current_resupply_qty - return_qty
                                resupply['quantity'] = new_qty
                                price_per_unit = float(resupply.get('price', 0)) / current_resupply_qty
                                shipping_per_unit = float(resupply.get('shipping_tax', 0)) / current_resupply_qty
                                total_per_unit = float(resupply.get('total_cost', 0)) / current_resupply_qty
                                
                                resupply['price'] = price_per_unit * new_qty
                                resupply['shipping_tax'] = shipping_per_unit * new_qty
                                resupply['total_cost'] = total_per_unit * new_qty
                            break
                    
                    total_items += return_qty
                    
                except Exception as e:
                    skipped_items.append(f"{display_name}: processing error - {str(e)}")

            # Save changes and refresh displays
            self.save_inventory()
            self.save_resupply_history()
            self.refresh_inventory()
            self.update_inventory_totals()
            
            # Refresh both this window and main application
            refresh_resupply_local_display()
            
            # Show detailed success message
            success_msg = f"Processed return of {total_items} items.\n\n"
            
            if returned_items:
                success_msg += "✅ Returned (removed from inventory):\n"
                for item in returned_items:
                    success_msg += f"• {item}\n"
            
            if skipped_items:
                success_msg += "\n⚠️ Skipped:\n"
                for item in skipped_items:
                    success_msg += f"• {item}\n"
            
            messagebox.showinfo("Return Complete", success_msg.strip())

        # Action buttons
        ttk.Button(button_frame, text="Return Selected Item(s)", 
                  command=return_selected_resupply_items, width=25).pack(side='left', padx=5)

        # Delete order button
        ttk.Button(button_frame, text="Delete Order & Remove Cigars", 
                  command=delete_resupply_order, width=30).pack(side='left', padx=5)
        
        def refresh_resupply_details():
            """Refresh details for selected resupply order."""
            resupply_detail_tree.delete(*resupply_detail_tree.get_children())
            if not current_resupply_id:
                return

            for resupply in self.resupply_history:
                if resupply.get('order_id') == current_resupply_id:
                    values = (
                        resupply.get('brand', 'Unknown'),
                        resupply.get('cigar', 'Unknown'),
                        resupply.get('size', 'N/A'),
                        resupply.get('type', 'N/A'),
                        str(resupply.get('quantity', 1)),
                        f"${float(resupply.get('price', 0)):.2f}",
                        f"${float(resupply.get('shipping_tax', 0)):.2f}",
                        f"${float(resupply.get('total_cost', 0)):.2f}"
                    )
                    resupply_detail_tree.insert('', 'end', values=values)
        
        def on_resupply_select(event):
            """Handle resupply order selection."""
            nonlocal current_resupply_id
            selected_item = resupply_tree.selection()
            if not selected_item:
                current_resupply_id = None
                resupply_detail_tree.delete(*resupply_detail_tree.get_children())
                return

            selected_item = selected_item[0]
            current_resupply_id = resupply_id_map.get(selected_item)
            refresh_resupply_details()
        
        # Bind selection event
        resupply_tree.bind('<<TreeviewSelect>>', on_resupply_select)
        
        # Load and display resupply history
        self.load_resupply_history()  # Make sure data is loaded
        
        # Group resupply records by order_id
        orders = {}
        for resupply in self.resupply_history:
            order_id = resupply.get('order_id', 'unknown')
            if order_id not in orders:
                orders[order_id] = []
            orders[order_id].append(resupply)
        
        # Display orders (newest first)
        sorted_orders = sorted(
            orders.items(), 
            key=lambda x: x[1][0].get('date', ''), 
            reverse=True
        )
        
        for order_id, resupplies in sorted_orders:
            try:
                total_items = sum(int(resupply.get('quantity', 1)) for resupply in resupplies)
                total_cost = sum(float(resupply.get('total_cost', 0)) for resupply in resupplies)
                total_shipping = sum(float(resupply.get('shipping_tax', 0)) for resupply in resupplies)
                date = resupplies[0].get('date', 'Unknown')
                
                values = (date, str(total_items), f"${total_cost:.2f}", f"${total_shipping:.2f}")
                item_id = resupply_tree.insert('', 'end', values=values)
                resupply_id_map[item_id] = order_id
            except Exception as e:
                print(f"Error displaying resupply order {order_id}: {e}")

    def add_to_resupply_order(self):
        """Add a cigar to the current resupply order."""
        try:
            brand = self.resupply_brand_var.get().strip()
            cigar_name = self.resupply_cigar_var.get().strip()
            size = self.resupply_size_var.get().strip()
            type_name = self.resupply_type_var.get().strip()
            count = int(self.resupply_count_var.get())
            price = float(self.resupply_price_var.get())
            
            if not cigar_name or count <= 0 or price < 0:
                messagebox.showwarning("Warning", "Please enter valid cigar details.")
                return
            
            # Add to brands/sizes/types if new
            if brand and brand not in self.brands:
                self.brands.add(brand)
            if size and size not in self.sizes:
                self.sizes.add(size)
            if type_name and type_name not in self.types:
                self.types.add(type_name)
            
            # Create cigar object
            new_cigar = {
                'brand': brand,
                'cigar': cigar_name,
                'size': size,
                'type': type_name,
                'count': count,
                'price': price,
                'proportional_shipping': 0,
                'proportional_tax': 0,
                'price_per_stick': 0
            }
            
            self.current_resupply_order.append(new_cigar)
            
            # Clear input fields
            self.resupply_brand_var.set('')
            self.resupply_cigar_var.set('')
            self.resupply_size_var.set('')
            self.resupply_type_var.set('')
            self.resupply_count_var.set('')
            self.resupply_price_var.set('')
            
            # Recalculate and refresh
            self.calculate_resupply_costs()
            
            # Focus back to brand field
            self.resupply_brand_combo.focus()
            
        except ValueError:
            messagebox.showerror("Error", "Please enter valid numeric values.")

    def calculate_resupply_costs(self):
        """Calculate proportional shipping and tax for all cigars in resupply order."""
        try:
            total_shipping = float(self.resupply_total_shipping_var.get() or 0)
            tax_rate_percent = float(self.resupply_tax_rate_var.get() or 0)
            tax_rate = tax_rate_percent / 100
            
            # Update the humidor's tax rate
            self.tax_rate = tax_rate
            
            # Calculate total cigars
            total_cigars = sum(cigar['count'] for cigar in self.current_resupply_order)
            
            if total_cigars == 0:
                for cigar in self.current_resupply_order:
                    cigar['proportional_shipping'] = 0
                    cigar['proportional_tax'] = 0
                    cigar['price_per_stick'] = 0
                self.resupply_total_tax_label.config(text="$0.00")
                self.refresh_resupply_cigars_display()
                return
            
            total_tax_amount = 0
            shipping_per_cigar = total_shipping / total_cigars
            
            # Update each cigar with proportional costs
            for cigar in self.current_resupply_order:
                cigar_count = cigar['count']
                base_price = cigar['price']
                
                # Calculate proportional shipping
                cigar['proportional_shipping'] = shipping_per_cigar * cigar_count
                
                # Calculate tax
                base_price_per_stick = base_price / cigar_count if cigar_count > 0 else 0
                tax_per_stick = base_price_per_stick * tax_rate
                cigar['proportional_tax'] = tax_per_stick * cigar_count
                
                total_tax_amount += cigar['proportional_tax']
                
                # Calculate price per stick
                total_cost = base_price + cigar['proportional_tax'] + cigar['proportional_shipping']
                cigar['price_per_stick'] = total_cost / cigar_count if cigar_count > 0 else 0
            
            # Update displays
            self.resupply_total_tax_label.config(text=f"${total_tax_amount:.2f}")
            self.refresh_resupply_cigars_display()
            self.update_resupply_summary()
            
        except (ValueError, ZeroDivisionError):
            for cigar in self.current_resupply_order:
                cigar['proportional_shipping'] = 0
                cigar['proportional_tax'] = 0
                cigar['price_per_stick'] = 0
            self.resupply_total_tax_label.config(text="$0.00")
            self.refresh_resupply_cigars_display()

    def refresh_resupply_cigars_display(self):
        """Refresh the resupply cigars treeview display."""
        # Clear existing items
        for item in self.resupply_cigars_tree.get_children():
            self.resupply_cigars_tree.delete(item)
        
        # Add all cigars
        for cigar in self.current_resupply_order:
            values = (
                cigar.get('brand', ''),
                cigar.get('cigar', ''),
                cigar.get('size', ''),
                cigar.get('type', ''),
                str(cigar.get('count', 0)),
                f"${cigar.get('price', 0):.2f}",
                f"${cigar.get('proportional_shipping', 0):.2f}",
                f"${cigar.get('proportional_tax', 0):.2f}",
                f"${cigar.get('price_per_stick', 0):.2f}"
            )
            self.resupply_cigars_tree.insert('', 'end', values=values)

    def update_resupply_summary(self):
        """Update the resupply summary display."""
        total_items = sum(cigar.get('count', 0) for cigar in self.current_resupply_order)
        total_allocated_shipping = sum(cigar.get('proportional_shipping', 0) for cigar in self.current_resupply_order)
        total_calculated_tax = sum(cigar.get('proportional_tax', 0) for cigar in self.current_resupply_order)
        
        self.resupply_summary_items_label.config(text=f"Total Items: {total_items}")
        self.resupply_summary_shipping_label.config(text=f"Total Shipping Allocated: ${total_allocated_shipping:.2f}")
        self.resupply_summary_tax_label.config(text=f"Total Tax Calculated: ${total_calculated_tax:.2f}")

    def update_resupply_cigar_dropdown(self, *args):
        """Update cigar dropdown based on selected brand."""
        selected_brand = self.resupply_brand_var.get().strip()
        if selected_brand:
            filtered_cigars = sorted(list(set(
                cigar.get('cigar', '') for cigar in self.inventory 
                if cigar.get('cigar', '') and cigar.get('brand', '').lower() == selected_brand.lower()
            )))
            self.resupply_cigar_combo.config(values=filtered_cigars)
            current_cigar = self.resupply_cigar_var.get()
            if current_cigar and current_cigar not in filtered_cigars:
                self.resupply_cigar_var.set('')
            if len(filtered_cigars) == 1:
                self.resupply_cigar_var.set(filtered_cigars[0])
        else:
            existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
            self.resupply_cigar_combo.config(values=existing_cigars)

    def remove_from_resupply_order(self):
        """Remove selected item from resupply order."""
        selected = self.resupply_cigars_tree.selection()
        if not selected:
            messagebox.showwarning("Warning", "Please select an item to remove.")
            return
        
        selected_item = selected[0]
        item_index = self.resupply_cigars_tree.index(selected_item)
        
        if 0 <= item_index < len(self.current_resupply_order):
            self.current_resupply_order.pop(item_index)
            self.calculate_resupply_costs()

    def clear_resupply_order(self):
        """Clear the entire resupply order."""
        if self.current_resupply_order and messagebox.askyesno("Confirm", "Clear entire order?"):
            self.current_resupply_order.clear()
            self.calculate_resupply_costs()

    def process_resupply_order(self):
        """Process the resupply order and add to inventory."""
        if not self.current_resupply_order:
            messagebox.showwarning("Warning", "Please add items to the order.")
            return
        
        try:
            added_count = 0
            combined_details = []
            order_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            order_id = str(uuid.uuid4())
            
            for cigar_data in self.current_resupply_order:
                # Create resupply record
                resupply_record = {
                    'order_id': order_id,
                    'date': order_date,
                    'brand': cigar_data['brand'],
                    'cigar': cigar_data['cigar'],
                    'size': cigar_data['size'],
                    'type': cigar_data['type'],
                    'quantity': cigar_data['count'],
                    'price': cigar_data['price'],
                    'shipping_tax': cigar_data['proportional_shipping'] + cigar_data['proportional_tax'],
                    'total_cost': cigar_data['price_per_stick'] * cigar_data['count']
                }
                self.resupply_history.append(resupply_record)
                
                # Check for duplicates
                existing_cigar = self.check_for_duplicate_cigar(
                    cigar_data['brand'], cigar_data['cigar'], cigar_data['size'])
                
                if existing_cigar:
                    old_count = existing_cigar.get('count', 0)
                    old_price_per_stick = existing_cigar.get('price_per_stick', 0)
                    
                    equiv_shipping = cigar_data['proportional_shipping'] + cigar_data['proportional_tax']
                    
                    self.combine_cigar_purchases(
                        existing_cigar, 
                        cigar_data['count'], 
                        cigar_data['price'], 
                        equiv_shipping
                    )
                    
                    new_count = existing_cigar.get('count', 0)
                    new_price_per_stick = existing_cigar.get('price_per_stick', 0)
                    
                    combined_details.append({
                        'name': f"{cigar_data['brand']} - {cigar_data['cigar']}",
                        'added_count': cigar_data['count'],
                        'old_count': old_count,
                        'new_count': new_count,
                        'old_price': old_price_per_stick,
                        'new_price': new_price_per_stick
                    })
                else:
                    # Add as new cigar
                    equiv_shipping = cigar_data['proportional_shipping'] + cigar_data['proportional_tax']
                    
                    new_cigar = {
                        'brand': cigar_data['brand'],
                        'cigar': cigar_data['cigar'],
                        'size': cigar_data['size'],
                        'type': cigar_data['type'],
                        'count': cigar_data['count'],
                        'price': cigar_data['price'],
                        'shipping': equiv_shipping,
                        'price_per_stick': cigar_data['price_per_stick'],
                        'personal_rating': None,
                        'original_quantity': cigar_data['count']  # Store original quantity for cost basis
                    }
                    
                    self.inventory.append(new_cigar)
                    added_count += 1
            
            # Save all changes
            self.save_inventory()
            self.save_resupply_history()
            self.save_brands()
            self.save_sets('cigar_sizes.json', self.sizes)
            self.save_sets('cigar_types.json', self.types)
            self.save_humidor_settings()
            
            # Clear order and refresh
            self.current_resupply_order.clear()
            self.calculate_resupply_costs()
            self.refresh_inventory()
            self.update_inventory_totals()
            self.refresh_resupply_dropdowns()
            
            # Show success message
            message = f"Order processed successfully!\n\n"
            if added_count > 0:
                message += f"Added {added_count} new cigar type(s)\n\n"
            if combined_details:
                message += "Combined with existing inventory:\n"
                for detail in combined_details:
                    message += f"• {detail['name']}: {detail['old_count']} → {detail['new_count']} total\n"
            
            messagebox.showinfo("Success", message.strip())
            
        except Exception as e:
            messagebox.showerror("Error", f"Error processing order: {str(e)}")
        
    def show_about(self):
        """Show about dialog."""
        about_text = """Cigar Inventory Manager
        
A comprehensive tool for managing your cigar collection.

Features:
• Inventory management with size-based duplicate detection
• Sales and resupply tracking
• Shipping cost calculator
• Multiple humidor support
• Data backup and export

Version: 2.0
"""
        messagebox.showinfo("About", about_text)

    def setup_file_location_bar(self, parent):
        """Setup the file location display and management bar."""
        # File location frame with modern styling
        location_frame = ttk.LabelFrame(parent, text="Data Location & Humidor Management", 
                                       padding="5", style='Modern.TLabelframe')
        location_frame.pack(fill='x', pady=(0, 5))
        
        # Current location display
        location_info_frame = ttk.Frame(location_frame)
        location_info_frame.pack(fill='x', pady=(0, 5))
        
        ttk.Label(location_info_frame, text="Current Humidor:").pack(side='left', padx=(0, 5))
        self.humidor_label = ttk.Label(location_info_frame, text=self.humidor_name, 
                                      font=('Segoe UI', 9, 'bold'), foreground='#007bff')
        self.humidor_label.pack(side='left', padx=(0, 10))
        
        ttk.Label(location_info_frame, text="Data Directory:").pack(side='left', padx=(0, 5))
        self.location_label = ttk.Label(location_info_frame, text=self.data_directory, 
                                       font=('Segoe UI', 8), foreground='#6c757d')
        self.location_label.pack(side='left', fill='x', expand=True)
        
        # Note: File management buttons are now in the File menu
        
        # Status indicator
        self.status_label = ttk.Label(location_info_frame, text="Ready", 
                                     font=('Segoe UI', 8), foreground='#28a745')
        self.status_label.pack(side='right', padx=(5, 0))

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
        ttk.Entry(search_frame, textvariable=self.search_var, style='Modern.TEntry').pack(side='left', fill='x', expand=True)
        
        # Create a frame to hold both treeview and totals
        tree_container = ttk.Frame(left_frame)
        tree_container.pack(expand=True, fill='both')
        
        # Create frame for treeview
        tree_frame = ttk.Frame(tree_container)
        tree_frame.pack(expand=True, fill='both')
        
        # Treeview with modern styling
        columns = ('select', 'brand', 'cigar', 'size', 'type', 'count', 'price', 'shipping', 
                  'per_stick', 'personal_rating')
        self.tree = ttk.Treeview(tree_frame, columns=columns, show='headings', style='Modern.Treeview')
        
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
        
        # Button definitions - core inventory management only
        buttons = [
            ("Add New", self.add_new_line),
            ("Resupply", self.resupply_order),
            ("Remove Selected", self.remove_selected)
        ]
        
        # Create buttons in a single column layout
        for i, (text, command) in enumerate(buttons):
            button = ttk.Button(button_frame, text=text, command=command, width=20)
            button.pack(fill='x', pady=2)
        
        # Configure the frame to expand properly
        button_frame.pack_propagate(False)
        
        # Add sales frame to right frame
        self.setup_sales_frame(right_frame)
        
        # Bind events
        self.tree.bind('<Button-1>', self.handle_click)
        self.tree.bind('<Double-1>', lambda e: setattr(e, 'double', True) or self.handle_click(e))
        self.tree.bind('<Button-3>', lambda e: self.tree.selection_remove(self.tree.selection()))

    def setup_sales_history_in_frame(self, parent_frame):
        """Setup the enhanced sales history window with transaction grouping and partial returns."""
        # Create main frame with padding
        main_frame = ttk.Frame(parent_frame, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Create a PanedWindow for resizable layout
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(fill='both', expand=True)

        # Left frame for transaction list
        left_frame = ttk.LabelFrame(paned_window, text="Sales Transactions", padding="5")
        paned_window.add(left_frame, weight=1)

        # Right frame for transaction details and actions
        right_frame = ttk.LabelFrame(paned_window, text="Transaction Details & Returns", padding="5")
        paned_window.add(right_frame, weight=1)

        # === LEFT FRAME: Transaction List ===
        # Treeview for transactions (grouped by transaction_id)
        trans_columns = ('date', 'items', 'total')
        self.transaction_tree = ttk.Treeview(left_frame, columns=trans_columns, show='headings', height=15)
        
        # Configure transaction columns
        self.transaction_tree.heading('date', text='Date & Time')
        self.transaction_tree.heading('items', text='Items Sold')
        self.transaction_tree.heading('total', text='Total Value')
        
        self.transaction_tree.column('date', width=150)
        self.transaction_tree.column('items', width=80)
        self.transaction_tree.column('total', width=100)

        # Add scrollbar for transactions
        trans_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.transaction_tree.yview)
        self.transaction_tree.configure(yscrollcommand=trans_scrollbar.set)
        
        # Pack transaction tree and scrollbar
        self.transaction_tree.pack(side='left', fill='both', expand=True)
        trans_scrollbar.pack(side='right', fill='y')

        # === RIGHT FRAME: Transaction Details ===
        # Treeview for individual items in selected transaction
        detail_columns = ('cigar', 'brand', 'size', 'quantity', 'price_per_stick', 'total_cost')
        self.detail_tree = ttk.Treeview(right_frame, columns=detail_columns, show='headings', height=10)
        
        # Configure detail columns
        self.detail_tree.heading('cigar', text='Cigar')
        self.detail_tree.heading('brand', text='Brand')
        self.detail_tree.heading('size', text='Size')
        self.detail_tree.heading('quantity', text='Qty')
        self.detail_tree.heading('price_per_stick', text='Price/Stick')
        self.detail_tree.heading('total_cost', text='Total')
        
        self.detail_tree.column('cigar', width=120)
        self.detail_tree.column('brand', width=100)
        self.detail_tree.column('size', width=70)
        self.detail_tree.column('quantity', width=50)
        self.detail_tree.column('price_per_stick', width=80)
        self.detail_tree.column('total_cost', width=80)

        # Add scrollbar for details
        detail_scrollbar = ttk.Scrollbar(right_frame, orient='vertical', command=self.detail_tree.yview)
        self.detail_tree.configure(yscrollcommand=detail_scrollbar.set)
        
        # Pack detail tree and scrollbar
        self.detail_tree.pack(side='top', fill='both', expand=True)
        detail_scrollbar.pack(side='right', fill='y')

        # === ACTION BUTTONS ===
        button_frame = ttk.Frame(right_frame)
        button_frame.pack(fill='x', pady=(10, 0))

        # Return buttons
        ttk.Button(button_frame, text="Return Selected Item(s)", 
                  command=self.return_selected_items, width=25).pack(side='left', padx=5)
        
        ttk.Button(button_frame, text="Return Entire Transaction", 
                  command=self.return_entire_transaction, width=25).pack(side='left', padx=5)

        # === BIND EVENTS ===
        # When a transaction is selected, show its details
        self.transaction_tree.bind('<<TreeviewSelect>>', self.on_transaction_select)

        # Store current transaction for detail view
        self.current_transaction_id = None

    def setup_resupply_history_in_frame(self, parent_frame):
        """Setup the resupply history window to track all purchases."""
        # Create main frame with padding
        main_frame = ttk.Frame(parent_frame, padding="10")
        main_frame.pack(fill='both', expand=True)

        # Create a PanedWindow for resizable layout
        paned_window = ttk.PanedWindow(main_frame, orient='horizontal')
        paned_window.pack(fill='both', expand=True)

        # Left frame for resupply list
        left_frame = ttk.LabelFrame(paned_window, text="Resupply Orders", padding="5")
        paned_window.add(left_frame, weight=1)

        # Right frame for order details
        right_frame = ttk.LabelFrame(paned_window, text="Order Details", padding="5")
        paned_window.add(right_frame, weight=1)

        # === LEFT FRAME: Resupply Orders List ===
        # Treeview for resupply orders (grouped by order_id)
        resupply_columns = ('date', 'items', 'total_cost', 'total_shipping')
        self.resupply_tree = ttk.Treeview(left_frame, columns=resupply_columns, show='headings', height=15)
        
        # Configure resupply columns
        self.resupply_tree.heading('date', text='Date & Time')
        self.resupply_tree.heading('items', text='Items')
        self.resupply_tree.heading('total_cost', text='Total Cost')
        self.resupply_tree.heading('total_shipping', text='Shipping')
        
        self.resupply_tree.column('date', width=150)
        self.resupply_tree.column('items', width=80)
        self.resupply_tree.column('total_cost', width=100)
        self.resupply_tree.column('total_shipping', width=100)

        # Add scrollbar for resupply orders
        resupply_scrollbar = ttk.Scrollbar(left_frame, orient='vertical', command=self.resupply_tree.yview)
        self.resupply_tree.configure(yscrollcommand=resupply_scrollbar.set)
        
        # Pack resupply tree and scrollbar
        self.resupply_tree.pack(side='left', fill='both', expand=True)
        resupply_scrollbar.pack(side='right', fill='y')

        # === RIGHT FRAME: Order Details ===
        # Treeview for individual items in selected resupply order
        detail_columns = ('brand', 'cigar', 'size', 'type', 'quantity', 'price', 'shipping', 'total_cost')
        self.resupply_detail_tree = ttk.Treeview(right_frame, columns=detail_columns, show='headings', height=15)
        
        # Configure detail columns
        self.resupply_detail_tree.heading('brand', text='Brand')
        self.resupply_detail_tree.heading('cigar', text='Cigar')
        self.resupply_detail_tree.heading('size', text='Size')
        self.resupply_detail_tree.heading('type', text='Type')
        self.resupply_detail_tree.heading('quantity', text='Qty')
        self.resupply_detail_tree.heading('price', text='Price')
        self.resupply_detail_tree.heading('shipping', text='Ship+Tax')
        self.resupply_detail_tree.heading('total_cost', text='Total')
        
        self.resupply_detail_tree.column('brand', width=100)
        self.resupply_detail_tree.column('cigar', width=120)
        self.resupply_detail_tree.column('size', width=70)
        self.resupply_detail_tree.column('type', width=80)
        self.resupply_detail_tree.column('quantity', width=50)
        self.resupply_detail_tree.column('price', width=80)
        self.resupply_detail_tree.column('shipping', width=80)
        self.resupply_detail_tree.column('total_cost', width=80)

        # Add scrollbar for details
        resupply_detail_scrollbar = ttk.Scrollbar(right_frame, orient='vertical', command=self.resupply_detail_tree.yview)
        self.resupply_detail_tree.configure(yscrollcommand=resupply_detail_scrollbar.set)
        
        # Pack detail tree and scrollbar
        self.resupply_detail_tree.pack(side='top', fill='both', expand=True)
        resupply_detail_scrollbar.pack(side='right', fill='y')

        # === BIND EVENTS ===
        # When a resupply order is selected, show its details
        self.resupply_tree.bind('<<TreeviewSelect>>', self.on_resupply_select)

        # Store current resupply order for detail view
        self.current_resupply_id = None

    def on_resupply_select(self, event):
        """Handle selection of a resupply order in the order list."""
        selected_item = self.resupply_tree.selection()
        if not selected_item:
            self.current_resupply_id = None
            self.resupply_detail_tree.delete(*self.resupply_detail_tree.get_children())
            return

        selected_item = selected_item[0]  # Get the first selected item
        # Get resupply_id from our mapping
        self.current_resupply_id = getattr(self, 'resupply_id_map', {}).get(selected_item)

        # Refresh the resupply details
        self.refresh_resupply_details()

    def refresh_resupply_details(self):
        """Refresh the details view for the currently selected resupply order."""
        # Clear existing details
        self.resupply_detail_tree.delete(*self.resupply_detail_tree.get_children())

        if not self.current_resupply_id:
            return

        # Load and display details for the selected resupply order
        for resupply in self.resupply_history:
            if resupply.get('order_id') == self.current_resupply_id:
                # Get all the resupply details
                brand = resupply.get('brand', 'Unknown')
                cigar_name = resupply.get('cigar', 'Unknown')
                size = resupply.get('size', 'N/A')
                cigar_type = resupply.get('type', 'N/A')
                quantity = int(resupply.get('quantity', 1))
                price = float(resupply.get('price', 0))
                shipping_tax = float(resupply.get('shipping_tax', 0))
                total_cost = float(resupply.get('total_cost', 0))
                
                values = (
                    brand,
                    cigar_name,
                    size,
                    cigar_type,
                    str(quantity),
                    f"${price:.2f}",
                    f"${shipping_tax:.2f}",
                    f"${total_cost:.2f}"
                )
                self.resupply_detail_tree.insert('', 'end', values=values)

    def refresh_resupply_history(self):
        """Refresh the resupply history display with order grouping."""
        # Clear current display
        for item in self.resupply_tree.get_children():
            self.resupply_tree.delete(item)
            
        # Clear detail view if no order is selected
        if not hasattr(self, 'current_resupply_id') or not self.current_resupply_id:
            for item in self.resupply_detail_tree.get_children():
                self.resupply_detail_tree.delete(item)
        
        # Group resupply records by order_id
        orders = {}
        for resupply in self.resupply_history:
            order_id = resupply.get('order_id', 'unknown')
            if order_id not in orders:
                orders[order_id] = []
            orders[order_id].append(resupply)
        
        # Store order ID mapping for later retrieval
        if not hasattr(self, 'resupply_id_map'):
            self.resupply_id_map = {}
        
        # Display orders (newest first)
        sorted_orders = sorted(
            orders.items(), 
            key=lambda x: x[1][0].get('date', ''), 
            reverse=True
        )
        
        for order_id, resupplies in sorted_orders:
            try:
                # Calculate order totals
                total_items = sum(int(resupply.get('quantity', 1)) for resupply in resupplies)
                total_cost = sum(float(resupply.get('total_cost', 0)) for resupply in resupplies)
                total_shipping = sum(float(resupply.get('shipping_tax', 0)) for resupply in resupplies)
                date = resupplies[0].get('date', 'Unknown')
                
                # Use the date from the first resupply in the order
                date = resupplies[0].get('date', 'Unknown')
                
                # Create order summary
                values = (
                    date,
                    str(total_items),
                    f"${total_cost:.2f}",
                    f"${total_shipping:.2f}"
                )
                
                # Insert the item and store the order_id in our mapping
                item_id = self.resupply_tree.insert('', 'end', values=values)
                self.resupply_id_map[item_id] = order_id
                
            except Exception as e:
                print(f"Error displaying resupply order {order_id}: {e}")
        
        # If we have a current order selected, refresh its details
        if hasattr(self, 'current_resupply_id') and self.current_resupply_id:
            self.refresh_resupply_details()

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
                self.refresh_resupply_dropdowns()  # Refresh resupply dropdowns
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
                    with open(self.get_data_file_path(filename), 'w') as f:
                        json.dump(list(item_set), f, indent=2)
                    self.refresh_resupply_dropdowns()  # Refresh resupply dropdowns
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
            
            # Update only the checkbox display for this specific row
            current_values = list(self.tree.item(item)['values'])
            current_values[0] = '☒' if self.checkbox_states[cigar_name] else '☐'
            self.tree.item(item, values=current_values)
            
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
                current_cigar_name = item_values[2]
                current_brand = item_values[1]
                
                # Find the cigar in inventory and update it
                current_cigar = None
                for cigar in self.inventory:
                    if cigar['cigar'] == current_cigar_name:
                        current_cigar = cigar
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
                
                # Check for duplicates after updating brand
                if current_cigar and column == 'brand':
                    # Get the updated brand, cigar name, and size
                    updated_brand = current_cigar.get('brand', '')
                    updated_cigar_name = current_cigar.get('cigar', '')
                    updated_size = current_cigar.get('size', '')
                    
                    # Check if this combination now matches another cigar
                    duplicate_cigar = self.check_for_duplicate_cigar_excluding_current(
                        updated_brand, updated_cigar_name, updated_size, current_cigar_name)
                    
                    if duplicate_cigar:
                        # Offer to combine
                        self.handle_automatic_duplicate_detection(current_cigar, duplicate_cigar)
                        return  # Don't continue with normal save process
                
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
                    current_cigar_name = item_values[2]
                    current_brand = item_values[1]
                    
                    # Find the cigar in inventory and update it
                    current_cigar = None
                    for cigar in self.inventory:
                        if cigar['cigar'] == current_cigar_name:
                            current_cigar = cigar
                            # Update the value
                            cigar[column] = value
                            break
                    
                    # Check for duplicates after updating brand, cigar name, or size
                    if current_cigar and column in ['brand', 'cigar', 'size']:
                        # Get the updated brand, cigar name, and size
                        updated_brand = current_cigar.get('brand', '')
                        updated_cigar_name = current_cigar.get('cigar', '')
                        updated_size = current_cigar.get('size', '')
                        
                        # Check if this combination now matches another cigar
                        duplicate_cigar = self.check_for_duplicate_cigar_excluding_current(
                            updated_brand, updated_cigar_name, updated_size, current_cigar_name)
                        
                        if duplicate_cigar:
                            # Offer to combine
                            self.handle_automatic_duplicate_detection(current_cigar, duplicate_cigar)
                            return  # Don't continue with normal save process
                    
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
                        # Update original_quantity to preserve cost basis
                        cigar['original_quantity'] = new_count
                        cigar['price_per_stick'] = self.calculate_price_per_stick(
                            cigar['price'], cigar['shipping'], new_count, new_count)
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
                
                # All inventory items now have original_quantity field
                # Use original_quantity for shipping distribution to preserve cost basis
                price_per_stick = self.calculate_price_per_stick(
                    cigar.get('price', 0),
                    cigar.get('shipping', 0),
                    cigar.get('count', 0),
                    cigar.get('original_quantity', cigar.get('count', 0))
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
            with open(self.get_data_file_path('cigar_inventory.json'), 'w') as f:
                json.dump(self.inventory, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save inventory: {str(e)}")
            
    def load_inventory(self):
        try:
            # Load sales and resupply history first
            self.load_sales_history()
            self.load_resupply_history()
            self.load_humidor_settings()
            
            # Load brands first
            try:
                with open(self.get_data_file_path('cigar_brands.json'), 'r') as f:
                    self.brands = set(json.load(f))
            except FileNotFoundError:
                self.brands = set()
            
            # Load sizes
            try:
                with open(self.get_data_file_path('cigar_sizes.json'), 'r') as f:
                    self.sizes = set(json.load(f))
            except FileNotFoundError:
                self.sizes = set()
                
            # Load types
            try:
                with open(self.get_data_file_path('cigar_types.json'), 'r') as f:
                    self.types = set(json.load(f))
            except FileNotFoundError:
                self.types = set()
            
            # Load inventory
            try:
                with open(self.get_data_file_path('cigar_inventory.json'), 'r') as f:
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
                        
                        # Add original_quantity field for existing inventory to preserve cost basis
                        if 'original_quantity' not in cigar:
                            cigar['original_quantity'] = cigar.get('count', 0)
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
            with open(self.get_data_file_path('cigar_brands.json'), 'w') as f:
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
            'personal_rating': None,
            'original_quantity': 0  # Store the original quantity for cost basis
        }
        
        # Calculate initial price per stick using count=1 to avoid division by zero
        new_cigar['price_per_stick'] = self.calculate_price_per_stick(
            new_cigar['price'],
            new_cigar['shipping'],
            1,  # Always use 1 for initial calculation
            1   # Use 1 as original quantity for initial calculation
        )
        
        self.inventory.append(new_cigar)
        self.save_inventory()
        self.refresh_inventory()
        
        # Select the new item
        last_item = self.tree.get_children()[-1]
        self.tree.selection_add(last_item)  # Changed from selection_set to selection_add
        self.tree.see(last_item)

    def resupply_order(self):
        """Create a resupply order window for adding multiple cigars from one order."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Resupply")
        dialog.geometry("1000x800")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (500)
        y = (dialog.winfo_screenheight() // 2) - (400)
        dialog.geometry(f"1000x800+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill='both', expand=True)
        
        # Header
        ttk.Label(main_frame, text="Resupply", 
                 font=('TkDefaultFont', 14, 'bold')).pack(pady=(0, 15))
        
        # Single column layout
        main_content_frame = ttk.Frame(main_frame)
        main_content_frame.pack(fill='both', expand=True)
        
        # Order totals frame
        totals_frame = ttk.LabelFrame(main_content_frame, text="Order Totals", padding="10")
        totals_frame.pack(fill='x', pady=(0, 15))
        
        # Create grid for order totals
        ttk.Label(totals_frame, text="Total Order Shipping ($):").grid(row=0, column=0, sticky='w', padx=(0, 10))
        total_shipping_var = tk.StringVar()
        total_shipping_entry = ttk.Entry(totals_frame, textvariable=total_shipping_var, width=15)
        total_shipping_entry.grid(row=0, column=1, padx=(0, 20))
        
        # Tax rate setting
        ttk.Label(totals_frame, text="Tax Rate (%):").grid(row=0, column=2, sticky='w', padx=(0, 10))
        tax_rate_var = tk.StringVar(value=str(self.tax_rate * 100))  # Convert to percentage
        tax_rate_entry = ttk.Entry(totals_frame, textvariable=tax_rate_var, width=10)
        tax_rate_entry.grid(row=0, column=3, padx=(0, 20))
        
        # Total tax display (calculated)
        ttk.Label(totals_frame, text="Total Tax ($):").grid(row=0, column=4, sticky='w', padx=(0, 10))
        total_tax_label = ttk.Label(totals_frame, text="$0.00", font=('TkDefaultFont', 9, 'bold'))
        total_tax_label.grid(row=0, column=5)
        
        # Cigars frame
        cigars_frame = ttk.LabelFrame(main_content_frame, text="Cigars in Order", padding="10")
        cigars_frame.pack(fill='both', expand=True, pady=(0, 15))
        
        # Create treeview for cigars
        columns = ('brand', 'cigar', 'size', 'type', 'count', 'price', 'calc_shipping', 'calc_tax', 'price_per_stick')
        cigars_tree = ttk.Treeview(cigars_frame, columns=columns, show='headings', height=12)
        
        # Configure columns
        cigars_tree.heading('brand', text='Brand')
        cigars_tree.heading('cigar', text='Cigar')
        cigars_tree.heading('size', text='Size')
        cigars_tree.heading('type', text='Type')
        cigars_tree.heading('count', text='Count')
        cigars_tree.heading('price', text='Price')
        cigars_tree.heading('calc_shipping', text='Shipping')
        cigars_tree.heading('calc_tax', text='Tax')
        cigars_tree.heading('price_per_stick', text='Price/Stick')
        
        cigars_tree.column('brand', width=100)
        cigars_tree.column('cigar', width=120)
        cigars_tree.column('size', width=70)
        cigars_tree.column('type', width=80)
        cigars_tree.column('count', width=60)
        cigars_tree.column('price', width=80)
        cigars_tree.column('calc_shipping', width=80)
        cigars_tree.column('calc_tax', width=70)
        cigars_tree.column('price_per_stick', width=90)
        
        # Add scrollbar
        cigars_scrollbar = ttk.Scrollbar(cigars_frame, orient='vertical', command=cigars_tree.yview)
        cigars_tree.configure(yscrollcommand=cigars_scrollbar.set)
        
        # Pack tree and scrollbar
        cigars_tree.pack(side='left', fill='both', expand=True)
        cigars_scrollbar.pack(side='right', fill='y')
        
        # Add instruction label
        instruction_label = ttk.Label(cigars_frame, text="Double-click any field to edit after adding to order • Cigar dropdown filters by brand", 
                                     font=('TkDefaultFont', 8), foreground='gray')
        instruction_label.pack(pady=(5, 0))
        
        # Store cigars data
        order_cigars = []
        
        # Bind double-click event for editing cigars in the order
        def on_cigar_double_click(event):
            region = cigars_tree.identify_region(event.x, event.y)
            if region != "cell":
                return
                
            column = cigars_tree.identify_column(event.x)
            item = cigars_tree.identify_row(event.y)
            
            if not item:
                return
                
            # Get the index of the selected item
            item_index = cigars_tree.index(item)
            if item_index >= len(order_cigars):
                return
                
            # Get column number and name
            col_num = int(column.replace('#', '')) - 1
            col_name = cigars_tree['columns'][col_num]
            
            # Don't allow editing calculated fields
            if col_name in ['calc_shipping', 'calc_tax', 'price_per_stick']:
                return
                
            # Get bounding box for the cell
            x, y, w, h = cigars_tree.bbox(item, column)
            
            # Show appropriate editor
            if col_name == 'brand':
                show_resupply_dropdown(item, item_index, col_name, sorted(list(self.brands)), x, y, w, h)
            elif col_name == 'cigar':
                show_resupply_cigar_dropdown(item, item_index, col_name, x, y, w, h)
            elif col_name == 'size':
                show_resupply_dropdown(item, item_index, col_name, sorted(list(self.sizes)), x, y, w, h)
            elif col_name == 'type':
                show_resupply_dropdown(item, item_index, col_name, sorted(list(self.types)), x, y, w, h)
            elif col_name == 'count':
                show_resupply_spinbox(item, item_index, col_name, x, y, w, h)
            elif col_name == 'price':
                show_resupply_price_entry(item, item_index, col_name, x, y, w, h)
        
        cigars_tree.bind('<Double-1>', on_cigar_double_click)
        
        def show_resupply_dropdown(item, item_index, column, values, x, y, w, h):
            # Destroy any existing popups
            for widget in cigars_tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
                    
            frame = ttk.Frame(cigars_tree)
            
            values = [''] + values  # Add empty option
            
            combo = ttk.Combobox(frame, values=values, width=w//10)
            current_value = order_cigars[item_index].get(column, '')
            combo.set(current_value)
            combo.pack(expand=True, fill='both')
            
            def save_value(event=None):
                try:
                    value = combo.get().strip()
                    order_cigars[item_index][column] = value
                    
                    # Add to appropriate set if not empty
                    if value:
                        if column == 'brand':
                            self.brands.add(value)
                        elif column == 'size':
                            self.sizes.add(value)
                        elif column == 'type':
                            self.types.add(value)
                    
                    # Recalculate and refresh
                    calculate_proportional_costs()
                    
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
        
        def show_resupply_cigar_dropdown(item, item_index, column, x, y, w, h):
            # Destroy any existing popups
            for widget in cigars_tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
                    
            frame = ttk.Frame(cigars_tree)
            
            # Get the brand of the current order item to filter cigars
            current_brand = order_cigars[item_index].get('brand', '').strip()
            
            # Get existing cigar names for dropdown
            if current_brand:
                # Filter cigars by selected brand (case-insensitive)
                filtered_cigars = sorted(list(set(
                    cigar.get('cigar', '') for cigar in self.inventory 
                    if cigar.get('cigar', '') and cigar.get('brand', '').lower() == current_brand.lower()
                )))
                cigar_values = [''] + filtered_cigars
            else:
                # Show all cigars if no brand selected
                existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
                cigar_values = [''] + existing_cigars
            
            combo = ttk.Combobox(frame, values=cigar_values, width=w//10)
            current_value = order_cigars[item_index].get(column, '')
            combo.set(current_value)
            combo.pack(expand=True, fill='both')
            
            def save_value(event=None):
                try:
                    value = combo.get().strip()
                    if value:  # Only save if there's a value
                        order_cigars[item_index][column] = value
                        # Recalculate and refresh
                        calculate_proportional_costs()
                    
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
        
        def show_resupply_text_entry(item, item_index, column, x, y, w, h):
            # Destroy any existing popups
            for widget in cigars_tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
                    
            frame = ttk.Frame(cigars_tree)
            entry = ttk.Entry(frame)
            current_value = order_cigars[item_index].get(column, '')
            entry.insert(0, current_value)
            entry.pack(expand=True, fill='both')
            entry.focus()
            entry.select_range(0, tk.END)

            def save_value(event=None):
                try:
                    value = entry.get().strip()
                    if value:  # Only save if there's a value
                        order_cigars[item_index][column] = value
                        # Recalculate and refresh
                        calculate_proportional_costs()
                    
                except Exception as e:
                    print(f"Error saving value: {e}")
                finally:
                    frame.destroy()
                
            entry.bind('<Return>', save_value)
            entry.bind('<FocusOut>', save_value)
            entry.bind('<Escape>', lambda e: frame.destroy())
            
            frame.place(x=x, y=y, width=w, height=h)
        
        def show_resupply_spinbox(item, item_index, column, x, y, w, h):
            # Destroy any existing popups
            for widget in cigars_tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
                    
            frame = ttk.Frame(cigars_tree)
            
            current_value = order_cigars[item_index].get(column, 0)
            
            spinbox = ttk.Spinbox(
                frame,
                from_=1,
                to=999,
                width=w//10,
                justify='center'
            )
            spinbox.set(str(current_value))
            spinbox.pack(expand=True, fill='both')
            spinbox.focus()

            def save_value(event=None):
                try:
                    value = int(spinbox.get())
                    order_cigars[item_index][column] = value
                    # Recalculate and refresh
                    calculate_proportional_costs()
                    
                except ValueError:
                    # If invalid value, restore previous value
                    spinbox.set(str(current_value))
                except Exception as e:
                    print(f"Error saving value: {e}")
                finally:
                    frame.destroy()

            spinbox.bind('<Return>', save_value)
            spinbox.bind('<FocusOut>', save_value)
            spinbox.bind('<Escape>', lambda e: frame.destroy())
            
            frame.place(x=x, y=y, width=w, height=h)
        
        def show_resupply_price_entry(item, item_index, column, x, y, w, h):
            # Destroy any existing popups
            for widget in cigars_tree.winfo_children():
                if isinstance(widget, ttk.Frame):
                    widget.destroy()
                    
            frame = ttk.Frame(cigars_tree)
            entry = ttk.Entry(frame, justify='right')
            current_value = order_cigars[item_index].get(column, 0)
            entry.insert(0, str(current_value))
            entry.pack(expand=True, fill='both')
            entry.focus()
            entry.select_range(0, tk.END)

            def save_value(event=None):
                try:
                    value = float(entry.get())
                    order_cigars[item_index][column] = value
                    # Recalculate and refresh
                    calculate_proportional_costs()
                    
                except ValueError:
                    # If invalid value, restore previous value
                    entry.insert(0, str(current_value))
                except Exception as e:
                    print(f"Error saving value: {e}")
                finally:
                    frame.destroy()

            entry.bind('<Return>', save_value)
            entry.bind('<FocusOut>', save_value)
            entry.bind('<Escape>', lambda e: frame.destroy())
            
            frame.place(x=x, y=y, width=w, height=h)
        
        # Add cigar frame
        add_frame = ttk.Frame(cigars_frame)
        add_frame.pack(fill='x', pady=(10, 0))
        
        # Input fields for new cigar
        input_frame = ttk.Frame(add_frame)
        input_frame.pack(fill='x')
        
        # Create input fields in vertical layout
        row = 0
        
        ttk.Label(input_frame, text="Brand:").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        brand_var = tk.StringVar()
        brand_combo = ttk.Combobox(input_frame, textvariable=brand_var, 
                                   values=sorted(list(self.brands)), width=20)
        brand_combo.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        row += 1
        
        ttk.Label(input_frame, text="Cigar:").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        cigar_var = tk.StringVar()
        # Get existing cigar names for dropdown
        existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
        cigar_combo = ttk.Combobox(input_frame, textvariable=cigar_var, 
                                   values=existing_cigars, width=20)
        cigar_combo.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        row += 1
        
        def update_cigar_dropdown(*args):
            """Update cigar dropdown based on selected brand."""
            selected_brand = brand_var.get().strip()
            if selected_brand:
                # Filter cigars by selected brand (case-insensitive)
                filtered_cigars = sorted(list(set(
                    cigar.get('cigar', '') for cigar in self.inventory 
                    if cigar.get('cigar', '') and cigar.get('brand', '').lower() == selected_brand.lower()
                )))
                cigar_combo.config(values=filtered_cigars)
                # Clear current selection if it doesn't match the brand
                current_cigar = cigar_var.get()
                if current_cigar and current_cigar not in filtered_cigars:
                    cigar_var.set('')
                # If only one cigar matches, auto-select it
                if len(filtered_cigars) == 1:
                    cigar_var.set(filtered_cigars[0])
            else:
                # Show all cigars if no brand selected
                cigar_combo.config(values=existing_cigars)
        
        # Add trace to brand combobox to update cigar dropdown
        brand_var.trace('w', update_cigar_dropdown)
        
        ttk.Label(input_frame, text="Size:").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        size_var = tk.StringVar()
        size_combo = ttk.Combobox(input_frame, textvariable=size_var, 
                                  values=sorted(list(self.sizes)), width=20)
        size_combo.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        row += 1
        
        ttk.Label(input_frame, text="Type:").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        type_var = tk.StringVar()
        type_combo = ttk.Combobox(input_frame, textvariable=type_var, 
                                  values=sorted(list(self.types)), width=20)
        type_combo.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        row += 1
        
        ttk.Label(input_frame, text="Count:").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        count_var = tk.StringVar()
        count_entry = ttk.Entry(input_frame, textvariable=count_var, width=20)
        count_entry.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        row += 1
        
        ttk.Label(input_frame, text="Price ($):").grid(row=row, column=0, sticky='w', padx=(0, 5), pady=(0, 2))
        price_var = tk.StringVar()
        price_entry = ttk.Entry(input_frame, textvariable=price_var, width=20)
        price_entry.grid(row=row, column=1, padx=(0, 10), pady=(0, 2), sticky='ew')
        
        # Configure column weights for proper stretching
        input_frame.grid_columnconfigure(1, weight=1)
        
        # Add to Order button (positioned right after input fields)
        add_button_frame = ttk.Frame(add_frame)
        add_button_frame.pack(fill='x', pady=(10, 0))
        
        add_button = ttk.Button(add_button_frame, text="Add to Order", 
                              width=20)
        add_button.pack()
        
        # Add separator above other buttons
        ttk.Separator(add_frame, orient='horizontal').pack(fill='x', pady=(10, 0))
        
        # Add remaining buttons under the separator
        button_frame = ttk.Frame(add_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        # Store button references for later assignment
        remove_button = ttk.Button(button_frame, text="Remove Selected", 
                                  width=20)
        remove_button.pack(pady=(0, 5))
        
        process_button = ttk.Button(button_frame, text="Process Order", 
                                   width=20)
        process_button.pack(pady=(0, 5))
        
        cancel_button = ttk.Button(button_frame, text="Cancel", 
                                  width=20)
        cancel_button.pack()
        
        def calculate_proportional_costs():
            """Calculate proportional shipping and tax for all cigars."""
            try:
                total_shipping = float(total_shipping_var.get() or 0)
                tax_rate_percent = float(tax_rate_var.get() or 0)
                tax_rate = tax_rate_percent / 100  # Convert percentage to decimal
                
                # Update the humidor's tax rate
                self.tax_rate = tax_rate
                
                # Calculate total cigars and base prices
                total_cigars = sum(cigar['count'] for cigar in order_cigars)
                
                if total_cigars == 0:
                    # Clear all proportional costs if no cigars
                    for cigar in order_cigars:
                        cigar['proportional_shipping'] = 0
                        cigar['proportional_tax'] = 0
                        cigar['price_per_stick'] = 0
                    total_tax_label.config(text="$0.00")
                    refresh_cigars_display()
                    return
                
                # Calculate total tax based on base prices and tax rate
                total_tax_amount = 0
                shipping_per_cigar = total_shipping / total_cigars
                
                # Update each cigar with proportional costs
                for cigar in order_cigars:
                    cigar_count = cigar['count']
                    base_price = cigar['price']
                    
                    # Calculate proportional shipping
                    cigar['proportional_shipping'] = shipping_per_cigar * cigar_count
                    
                    # Calculate tax based on base price and tax rate
                    base_price_per_stick = base_price / cigar_count if cigar_count > 0 else 0
                    tax_per_stick = base_price_per_stick * tax_rate
                    cigar['proportional_tax'] = tax_per_stick * cigar_count
                    
                    # Add to total tax amount
                    total_tax_amount += cigar['proportional_tax']
                    
                    # Calculate price per stick (base price + tax + proportional shipping) / count
                    total_cost = base_price + cigar['proportional_tax'] + cigar['proportional_shipping']
                    cigar['price_per_stick'] = total_cost / cigar_count if cigar_count > 0 else 0
                
                # Update total tax display
                total_tax_label.config(text=f"${total_tax_amount:.2f}")
                
                # Refresh the display
                refresh_cigars_display()
                
            except (ValueError, ZeroDivisionError):
                # Clear calculations on error
                for cigar in order_cigars:
                    cigar['proportional_shipping'] = 0
                    cigar['proportional_tax'] = 0
                    cigar['price_per_stick'] = 0
                total_tax_label.config(text="$0.00")
                refresh_cigars_display()
        
        def refresh_cigars_display():
            """Refresh the cigars treeview display."""
            # Clear existing items
            for item in cigars_tree.get_children():
                cigars_tree.delete(item)
            
            # Add all cigars
            for cigar in order_cigars:
                values = (
                    cigar.get('brand', ''),
                    cigar.get('cigar', ''),
                    cigar.get('size', ''),
                    cigar.get('type', ''),
                    str(cigar.get('count', 0)),
                    f"${cigar.get('price', 0):.2f}",
                    f"${cigar.get('proportional_shipping', 0):.2f}",
                    f"${cigar.get('proportional_tax', 0):.2f}",
                    f"${cigar.get('price_per_stick', 0):.2f}"
                )
                cigars_tree.insert('', 'end', values=values)
        
        def add_cigar():
            """Add a cigar to the order."""
            try:
                brand = brand_var.get().strip()
                cigar_name = cigar_var.get().strip()
                size = size_var.get().strip()
                type_name = type_var.get().strip()
                count = int(count_var.get())
                price = float(price_var.get())
                
                if not cigar_name or count <= 0 or price < 0:
                    messagebox.showwarning("Warning", "Please enter valid cigar details.")
                    return
                
                # Add to brands/sizes/types if new
                needs_refresh = False
                if brand and brand not in self.brands:
                    self.brands.add(brand)
                    needs_refresh = True
                if size and size not in self.sizes:
                    self.sizes.add(size)
                    needs_refresh = True
                if type_name and type_name not in self.types:
                    self.types.add(type_name)
                    needs_refresh = True
                
                # Refresh dropdowns if new items were added
                if needs_refresh:
                    self.refresh_resupply_dropdowns()
                
                # Create cigar object
                new_cigar = {
                    'brand': brand,
                    'cigar': cigar_name,
                    'size': size,
                    'type': type_name,
                    'count': count,
                    'price': price,
                    'proportional_shipping': 0,
                    'proportional_tax': 0,
                    'price_per_stick': 0
                }
                
                order_cigars.append(new_cigar)
                
                # Clear input fields
                brand_var.set('')
                cigar_var.set('')
                size_var.set('')
                type_var.set('')
                count_var.set('')
                price_var.set('')
                
                # Recalculate and refresh
                calculate_proportional_costs()
                
                # Focus back to brand field
                brand_combo.focus()
                
            except ValueError:
                messagebox.showerror("Error", "Please enter valid numeric values.")
        
        def remove_selected_cigar():
            """Remove selected cigar from the order."""
            selected = cigars_tree.selection()
            if not selected:
                messagebox.showwarning("Warning", "Please select a cigar to remove.")
                return
            
            # Get the index of the selected item
            selected_item = selected[0]
            item_index = cigars_tree.index(selected_item)
            
            # Remove from order_cigars
            if 0 <= item_index < len(order_cigars):
                order_cigars.pop(item_index)
                calculate_proportional_costs()
        
        # Add summary frame between cigars and buttons
        summary_frame = ttk.LabelFrame(main_content_frame, text="Order Summary", padding="10")
        summary_frame.pack(fill='x', pady=(0, 15))
        
        # Create summary labels
        summary_shipping_label = ttk.Label(summary_frame, text="Total Shipping Allocated: $0.00")
        summary_shipping_label.pack(anchor='w', pady=2)
        
        summary_tax_label = ttk.Label(summary_frame, text="Total Tax Calculated: $0.00")
        summary_tax_label.pack(anchor='w', pady=2)
        
        summary_items_label = ttk.Label(summary_frame, text="Total Items: 0")
        summary_items_label.pack(anchor='w', pady=2)
        
        def update_summary():
            """Update the summary display with current totals."""
            total_allocated_shipping = sum(cigar.get('proportional_shipping', 0) for cigar in order_cigars)
            total_calculated_tax = sum(cigar.get('proportional_tax', 0) for cigar in order_cigars)
            total_items = sum(cigar.get('count', 0) for cigar in order_cigars)
            
            summary_shipping_label.config(text=f"Total Shipping Allocated: ${total_allocated_shipping:.2f}")
            summary_tax_label.config(text=f"Total Tax Calculated: ${total_calculated_tax:.2f}")
            summary_items_label.config(text=f"Total Items: {total_items}")
        
        # Update the calculate_proportional_costs function to also update summary
        original_calculate = calculate_proportional_costs
        def calculate_proportional_costs():
            original_calculate()
            update_summary()
        
        # Bind calculation updates to total fields
        total_shipping_var.trace('w', lambda *args: calculate_proportional_costs())
        tax_rate_var.trace('w', lambda *args: calculate_proportional_costs())
        
        def process_order():
            """Process the entire order and add to inventory."""
            if not order_cigars:
                messagebox.showwarning("Warning", "Please add at least one cigar to the order.")
                return
            
            try:
                added_count = 0
                combined_details = []
                order_date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                order_id = str(uuid.uuid4())  # Generate unique order ID
                
                for cigar_data in order_cigars:
                    # Create resupply record for history
                    resupply_record = {
                        'order_id': order_id,
                        'date': order_date,
                        'brand': cigar_data['brand'],
                        'cigar': cigar_data['cigar'],
                        'size': cigar_data['size'],
                        'type': cigar_data['type'],
                        'quantity': cigar_data['count'],
                        'price': cigar_data['price'],
                        'shipping_tax': cigar_data['proportional_shipping'] + cigar_data['proportional_tax'],
                        'total_cost': cigar_data['price_per_stick'] * cigar_data['count']
                    }
                    self.resupply_history.append(resupply_record)
                    
                    # Check for duplicates
                    existing_cigar = self.check_for_duplicate_cigar(
                        cigar_data['brand'], cigar_data['cigar'], cigar_data['size'])
                    
                    if existing_cigar:
                        # Capture old values before combining
                        old_count = existing_cigar.get('count', 0)
                        old_price_per_stick = existing_cigar.get('price_per_stick', 0)
                        
                        # Calculate the equivalent "shipping" value for the existing calculation method
                        equiv_shipping = cigar_data['proportional_shipping'] + cigar_data['proportional_tax']
                        
                        self.combine_cigar_purchases(
                            existing_cigar, 
                            cigar_data['count'], 
                            cigar_data['price'], 
                            equiv_shipping
                        )
                        
                        # Capture new values after combining
                        new_count = existing_cigar.get('count', 0)
                        new_price_per_stick = existing_cigar.get('price_per_stick', 0)
                        
                        # Store combination details
                        combined_details.append({
                            'name': f"{cigar_data['brand']} - {cigar_data['cigar']}",
                            'added_count': cigar_data['count'],
                            'old_count': old_count,
                            'new_count': new_count,
                            'old_price': old_price_per_stick,
                            'new_price': new_price_per_stick
                        })
                    else:
                        # Add as new cigar
                        # Calculate equivalent shipping for storage
                        equiv_shipping = cigar_data['proportional_shipping'] + cigar_data['proportional_tax']
                        price_per_stick = cigar_data['price_per_stick']
                        
                        new_cigar = {
                            'brand': cigar_data['brand'],
                            'cigar': cigar_data['cigar'],
                            'size': cigar_data['size'],
                            'type': cigar_data['type'],
                            'count': cigar_data['count'],
                            'price': cigar_data['price'],
                            'shipping': equiv_shipping,  # Store combined shipping+tax as shipping
                            'price_per_stick': price_per_stick,
                            'personal_rating': None
                        }
                        
                        self.inventory.append(new_cigar)
                        added_count += 1
                
                # Save all changes
                self.save_inventory()
                self.save_resupply_history()
                self.save_brands()
                self.save_sets('cigar_sizes.json', self.sizes)
                self.save_sets('cigar_types.json', self.types)
                self.save_humidor_settings()
                
                # Refresh displays
                self.refresh_inventory()
                self.refresh_resupply_history()
                self.update_inventory_totals()
                
                dialog.destroy()
                
                # Show detailed success message
                message = f"Order processed successfully!\n\n"
                
                if added_count > 0:
                    message += f"Added {added_count} new cigar type(s) to inventory\n\n"
                
                if combined_details:
                    message += f"Combined with existing inventory:\n"
                    for detail in combined_details:
                        message += f"• {detail['name']}:\n"
                        message += f"  Added {detail['added_count']} cigars ({detail['old_count']} → {detail['new_count']} total)\n"
                        message += f"  Price per stick: ${detail['old_price']:.2f} → ${detail['new_price']:.2f}\n\n"
                
                messagebox.showinfo("Success", message.strip())
                
            except Exception as e:
                messagebox.showerror("Error", f"Error processing order: {str(e)}")
        
        def cancel_order():
            """Cancel the order."""
            dialog.destroy()
        
        # Assign commands to buttons
        add_button.config(command=add_cigar)
        remove_button.config(command=remove_selected_cigar)
        process_button.config(command=process_order)
        cancel_button.config(command=cancel_order)
        
        # Bind Enter key to add cigar on the last field
        price_entry.bind('<Return>', lambda e: add_cigar())
        
        # Focus on first field
        total_shipping_entry.focus()

    def check_for_duplicate_cigar(self, brand, cigar_name, size):
        """Check if a cigar with the same brand, name, and size already exists."""
        for existing_cigar in self.inventory:
            if (existing_cigar.get('brand', '').lower() == brand.lower() and 
                existing_cigar.get('cigar', '').lower() == cigar_name.lower() and
                existing_cigar.get('size', '').lower() == size.lower()):
                return existing_cigar
        return None

    def check_for_duplicate_cigar_excluding_current(self, brand, cigar_name, size, current_cigar_name):
        """Check if a cigar with the same brand, name, and size already exists, excluding the current one being edited."""
        for existing_cigar in self.inventory:
            if (existing_cigar.get('cigar', '') != current_cigar_name and
                existing_cigar.get('brand', '').lower() == brand.lower() and 
                existing_cigar.get('cigar', '').lower() == cigar_name.lower() and
                existing_cigar.get('size', '').lower() == size.lower()):
                return existing_cigar
        return None

    def handle_automatic_duplicate_detection(self, current_cigar, duplicate_cigar):
        """Handle when a duplicate is automatically detected during editing."""
        # Check if both cigars have valid data for combining
        current_count = int(current_cigar.get('count', 0))
        current_price = float(current_cigar.get('price', 0))
        current_shipping = float(current_cigar.get('shipping', 0))
        
        duplicate_count = int(duplicate_cigar.get('count', 0))
        
        # If the current cigar has no count/price data, just merge the names
        if current_count == 0 and current_price == 0 and current_shipping == 0:
            # Simple case: just remove the empty current cigar and update the duplicate
            brand = current_cigar.get('brand', '')
            cigar_name = current_cigar.get('cigar', '')
            
            # Remove the current (empty) cigar
            self.inventory.remove(current_cigar)
            
            # Update the duplicate cigar's name if needed
            if duplicate_cigar.get('brand', '') != brand:
                duplicate_cigar['brand'] = brand
            if duplicate_cigar.get('cigar', '') != cigar_name:
                duplicate_cigar['cigar'] = cigar_name
                
            # Save and refresh
            self.save_inventory()
            self.refresh_inventory()
            messagebox.showinfo("Merged", f"Merged with existing cigar: {brand} - {cigar_name}")
            return
            
        # Complex case: both cigars have data, offer to combine them
        result = self.offer_combine_cigars(duplicate_cigar, current_count, current_price, current_shipping)
        
        if result == "combine":
            # Combine the cigars
            self.combine_cigar_purchases(duplicate_cigar, current_count, current_price, current_shipping)
            
            # Remove the current cigar since it's been combined
            self.inventory.remove(current_cigar)
            
            # Save and refresh
            self.save_inventory()
            self.refresh_inventory()
            self.update_inventory_totals()
            
            messagebox.showinfo("Combined", "Cigars have been combined successfully!")
            
        elif result == "separate":
            # Keep them separate - restore the original names to avoid confusion
            # We need to modify the names to make them unique
            current_cigar['cigar'] = current_cigar['cigar'] + " (2)"
            
            # Save and refresh
            self.save_inventory()
            self.refresh_inventory()
            
            messagebox.showinfo("Kept Separate", "Cigars kept as separate entries. Added '(2)' to distinguish them.")
            
        else:  # cancel
            # Restore original values - this is tricky since we already changed them
            # For now, we'll just refresh to let the user try again
            self.refresh_inventory()
            messagebox.showinfo("Cancelled", "Changes cancelled. Please try again.")

    def offer_combine_cigars(self, existing_cigar, new_count, new_price, new_shipping):
        """Offer to combine a new purchase with an existing cigar."""
        brand = existing_cigar.get('brand', '')
        cigar_name = existing_cigar.get('cigar', '')
        current_count = existing_cigar.get('count', 0)
        current_price_per_stick = existing_cigar.get('price_per_stick', 0)
        
        # Calculate what the new combined price per stick would be following existing style
        new_price_per_stick = self.calculate_price_per_stick(new_price, new_shipping, new_count)
        
        if current_count > 0:
            # Get the current total price and shipping (backing out from existing data)
            current_price = existing_cigar.get('price', 0)
            current_shipping = existing_cigar.get('shipping', 0)
            
            # Combine totals (matching existing calculation style)
            combined_count = current_count + new_count
            combined_price = current_price + new_price  # Total price for combined order
            combined_shipping = current_shipping + new_shipping  # Total shipping for combined order
            
            # Calculate combined price per stick using existing method
            combined_price_per_stick = self.calculate_price_per_stick(
                combined_price, 
                combined_shipping, 
                combined_count
            )
        else:
            combined_count = new_count
            combined_price_per_stick = new_price_per_stick
        
        # Create confirmation dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Duplicate Cigar Detected")
        dialog.geometry("450x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (225)
        y = (dialog.winfo_screenheight() // 2) - (150)
        dialog.geometry(f"450x300+{x}+{y}")
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill='both', expand=True)
        
        # Header
        ttk.Label(main_frame, text="Duplicate Cigar Found", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=(0, 10))
        
        # Info text
        info_text = f"A cigar with the same name already exists:\n\n"
        info_text += f"Brand: {brand}\n"
        info_text += f"Cigar: {cigar_name}\n\n"
        info_text += f"Current: {current_count} cigars at ${current_price_per_stick:.2f}/stick\n"
        info_text += f"New: {new_count} cigars at ${new_price_per_stick:.2f}/stick\n\n"
        info_text += f"Combined would be: {combined_count} cigars at ${combined_price_per_stick:.2f}/stick\n\n"
        info_text += "Would you like to combine them?"
        
        info_label = ttk.Label(main_frame, text=info_text, wraplength=400, justify='left')
        info_label.pack(pady=(0, 20))
        
        # Result variable
        result = tk.StringVar(value="")
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x', pady=(10, 0))
        
        def combine_cigars():
            """Combine the cigars."""
            result.set("combine")
            dialog.destroy()
        
        def keep_separate():
            """Keep them as separate entries."""
            result.set("separate")
            dialog.destroy()
        
        def cancel_action():
            """Cancel the operation."""
            result.set("cancel")
            dialog.destroy()
        
        ttk.Button(button_frame, text="Combine", command=combine_cigars, 
                  width=12).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="Keep Separate", command=keep_separate, 
                  width=12).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="Cancel", command=cancel_action, 
                  width=12).pack(side='left')
        
        # Wait for user response
        dialog.wait_window()
        return result.get()



    def combine_cigar_purchases(self, existing_cigar, new_count, new_price, new_shipping):
        """Combine new purchase with existing cigar using weighted averages."""
        # Get current values
        old_count = int(existing_cigar.get('count', 0))
        old_price = float(existing_cigar.get('price', 0))
        old_shipping = float(existing_cigar.get('shipping', 0))
        
        # Calculate combined values following the existing calculation style
        if old_count > 0:
            # Combine the totals (matching the existing style)
            combined_count = old_count + new_count
            combined_price = old_price + new_price  # Total price for combined order
            combined_shipping = old_shipping + new_shipping  # Total shipping for combined order
            
            # Calculate the new price per stick using the existing calculation method
            combined_price_per_stick = self.calculate_price_per_stick(
                combined_price, 
                combined_shipping, 
                combined_count
            )
            
            # Update the existing cigar to match the existing data structure
            existing_cigar['count'] = combined_count
            existing_cigar['price'] = combined_price  # Total price for all cigars
            existing_cigar['shipping'] = combined_shipping  # Total shipping for all cigars
            existing_cigar['price_per_stick'] = combined_price_per_stick
            
            # Update original_quantity to reflect the total combined quantity for cost basis
            existing_cigar['original_quantity'] = combined_count
            
            # Store the original purchase info for reference
            if 'purchase_history' not in existing_cigar:
                existing_cigar['purchase_history'] = []
            
            existing_cigar['purchase_history'].append({
                'date': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'count': new_count,
                'price': new_price,
                'shipping': new_shipping,
                'price_per_stick': self.calculate_price_per_stick(new_price, new_shipping, new_count)
            })
            
        else:
            # If no existing count, just set the new values
            existing_cigar['count'] = new_count
            existing_cigar['price'] = new_price
            existing_cigar['shipping'] = new_shipping
            existing_cigar['price_per_stick'] = self.calculate_price_per_stick(new_price, new_shipping, new_count)
            existing_cigar['original_quantity'] = new_count  # Set original quantity for cost basis

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
            with open(self.get_data_file_path(filename), 'w') as f:
                json.dump(list(data_set), f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save {filename}: {str(e)}")

    def setup_sales_frame(self, parent_frame):
        # Sale frame
        order_frame = ttk.LabelFrame(parent_frame, text="Sale", padding=10)
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

        self.order_total_label = ttk.Label(self.totals_frame, text="Sale Total: $0.00", font=('TkDefaultFont', 10, 'bold'))
        self.order_total_label.pack(fill='x', pady=2)

        self.order_count_label = ttk.Label(self.totals_frame, text="Cigar Count: 0")
        self.order_count_label.pack(fill='x', pady=2)

        # Add Process Sale button
        ttk.Button(self.totals_frame, text="Process Sale", command=self.sell_selected, width=20).pack(pady=5)



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
        self.order_total_label.config(text=f"Sale Total: ${total_price:.2f}")
        self.order_count_label.config(text=f"Cigar Count: {total_cigars}")

    def setup_totals_frame(self, parent_frame):
        """Setup frame for displaying inventory totals."""
        totals_frame = ttk.LabelFrame(parent_frame, text="Inventory Totals", 
                                     padding=10, style='Modern.TLabelframe')
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
        transaction_id = str(uuid.uuid4())  # Generate unique transaction ID
        
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
                        'transaction_id': transaction_id,  # Add transaction ID
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
            with open(self.get_data_file_path('sales_history.json'), 'w') as f:
                json.dump(self.sales_history, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save sales history: {str(e)}")

    def load_sales_history(self):
        """Load sales history from JSON file."""
        try:
            with open(self.get_data_file_path('sales_history.json'), 'r') as f:
                self.sales_history = json.load(f)
                
            # Update old format records to new format
            for sale in self.sales_history:
                if 'quantity' not in sale:
                    sale['quantity'] = 1
                if 'total_cost' not in sale:
                    sale['total_cost'] = float(sale['price_per_stick'])
                # Add transaction_id to older records that don't have it
                if 'transaction_id' not in sale:
                    # Generate a unique transaction ID based on date and cigar name
                    # This groups sales that happened at the same time
                    sale['transaction_id'] = str(uuid.uuid5(uuid.NAMESPACE_DNS, 
                                                          f"{sale.get('date', '')}-{sale.get('cigar', '')}"))
                    
        except FileNotFoundError:
            self.sales_history = []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load sales history: {str(e)}")
            self.sales_history = []

    def save_resupply_history(self):
        """Save resupply history to JSON file."""
        try:
            with open(self.get_data_file_path('resupply_history.json'), 'w') as f:
                json.dump(self.resupply_history, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save resupply history: {str(e)}")

    def load_resupply_history(self):
        """Load resupply history from JSON file."""
        try:
            with open(self.get_data_file_path('resupply_history.json'), 'r') as f:
                self.resupply_history = json.load(f)
                    
        except FileNotFoundError:
            self.resupply_history = []
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load resupply history: {str(e)}")
            self.resupply_history = []

    def save_humidor_settings(self):
        """Save humidor-specific settings like tax rate."""
        try:
            settings = {
                'tax_rate': self.tax_rate,
                'humidor_name': self.humidor_name
            }
            with open(self.get_data_file_path('humidor_settings.json'), 'w') as f:
                json.dump(settings, f, indent=2)
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save humidor settings: {str(e)}")

    def load_humidor_settings(self):
        """Load humidor-specific settings like tax rate."""
        try:
            with open(self.get_data_file_path('humidor_settings.json'), 'r') as f:
                settings = json.load(f)
                self.tax_rate = settings.get('tax_rate', 0.086)  # Default to 8.6%
                # Update humidor name if saved
                saved_name = settings.get('humidor_name')
                if saved_name:
                    self.humidor_name = saved_name
        except FileNotFoundError:
            # Use defaults if file doesn't exist
            self.tax_rate = 0.086  # Default 8.6%
        except Exception as e:
            messagebox.showerror("Error", f"Failed to load humidor settings: {str(e)}")
            self.tax_rate = 0.086  # Default 8.6%

    def manual_save(self):
        """Manually save all data and show confirmation."""
        try:
            # Save inventory
            self.save_inventory()
            
            # Save brands, sizes, and types
            self.save_brands()
            self.save_sets('cigar_sizes.json', self.sizes)
            self.save_sets('cigar_types.json', self.types)
            
            # Save sales and resupply history
            self.save_sales_history()
            self.save_resupply_history()
            
            # Save humidor settings
            self.save_humidor_settings()
            
            messagebox.showinfo("Success", "All data has been saved successfully!")
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save data: {str(e)}")

    def on_transaction_select(self, event):
        """Handle selection of a transaction in the transaction list."""
        selected_item = self.transaction_tree.selection()
        if not selected_item:
            self.current_transaction_id = None
            self.detail_tree.delete(*self.detail_tree.get_children())
            return

        selected_item = selected_item[0]  # Get the first selected item
        # Get transaction_id from our mapping
        self.current_transaction_id = self.transaction_id_map.get(selected_item)

        # Refresh the transaction details
        self.refresh_transaction_details()

    def refresh_transaction_details(self):
        """Refresh the details view for the currently selected transaction."""
        # Clear existing details
        self.detail_tree.delete(*self.detail_tree.get_children())

        if not self.current_transaction_id:
            return

        # Load and display details for the selected transaction
        for sale in self.sales_history:
            if sale.get('transaction_id') == self.current_transaction_id:
                # Get all the sale details
                cigar_name = sale.get('cigar', 'Unknown')
                brand = sale.get('brand', 'Unknown')
                size = sale.get('size', 'N/A')
                quantity = int(sale.get('quantity', 1))
                price_per_stick = float(sale.get('price_per_stick', 0))
                total_cost = float(sale.get('total_cost', 0))
                
                values = (
                    cigar_name,
                    brand,
                    size,
                    str(quantity),
                    f"${price_per_stick:.2f}",
                    f"${total_cost:.2f}"
                )
                self.detail_tree.insert('', 'end', values=values)

    def return_selected_items(self):
        """Handle partial return of selected items from the current transaction."""
        selected_items = self.detail_tree.selection()
        if not selected_items:
            messagebox.showwarning("Warning", "Please select items to return.")
            return

        if not self.current_transaction_id:
            messagebox.showwarning("Warning", "No transaction selected.")
            return

        # Collect information about selected items
        items_data = []
        for item_id in selected_items:
            values = self.detail_tree.item(item_id)['values']
            if values:
                cigar_name = values[0]
                brand = values[1]
                current_qty = int(values[3])
                items_data.append((cigar_name, brand, current_qty))

        if not items_data:
            return

        # Show single dialog for all return quantities
        return_quantities = self.ask_multiple_return_quantities(items_data)
        
        if not return_quantities:
            return  # User cancelled or no items to return

        # Process the returns
        self.process_returns(return_quantities, self.current_transaction_id)

    def ask_multiple_return_quantities(self, items_data):
        """Show a single dialog to set return quantities for multiple items."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Return Quantities")
        dialog.geometry("700x500")  # Made wider for side-by-side layout
        dialog.transient(self.root)
        dialog.grab_set()

        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (350)  # Half of 700
        y = (dialog.winfo_screenheight() // 2) - (250)  # Half of 500
        dialog.geometry(f"700x500+{x}+{y}")

        # Main frame
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill='both', expand=True)

        # Header
        ttk.Label(main_frame, text="Set Return Quantities", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Set the quantity to return for each item:").pack(pady=(0, 15))

        # Create horizontal layout: items on left, buttons on right
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True)

        # Left side: Scrollable frame for items
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 15))

        canvas = tk.Canvas(left_frame)
        scrollbar = ttk.Scrollbar(left_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas)

        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )

        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        # Pack canvas and scrollbar
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        # Store spinbox references
        quantity_vars = {}

        # Create quantity controls for each item
        for i, (cigar_name, brand, max_qty) in enumerate(items_data):
            item_frame = ttk.LabelFrame(scrollable_frame, text=f"{brand} - {cigar_name}", padding="10")
            item_frame.pack(fill='x', pady=5, padx=5)

            # Current available quantity
            ttk.Label(item_frame, text=f"Sold quantity: {max_qty}").pack(anchor='w')

            # Quantity selection
            qty_frame = ttk.Frame(item_frame)
            qty_frame.pack(fill='x', pady=(10, 0))

            ttk.Label(qty_frame, text="Return quantity:").pack(side='left')
            
            var = tk.IntVar(value=max_qty)  # Default to returning all
            quantity_vars[cigar_name] = var
            
            spinbox = ttk.Spinbox(qty_frame, from_=0, to=max_qty, 
                                 textvariable=var, width=10)
            spinbox.pack(side='right')

        # Right side: Button panel
        button_panel = ttk.Frame(content_frame)
        button_panel.pack(side='right', fill='y')

        # Results storage
        result = {}

        def confirm_quantities():
            # Collect return quantities
            for cigar_name, brand, max_qty in items_data:
                return_qty = quantity_vars[cigar_name].get()
                if return_qty > 0:
                    result[cigar_name] = (brand, return_qty, max_qty)
            dialog.destroy()

        def cancel_quantities():
            dialog.destroy()

        def select_all():
            """Set all quantities to maximum."""
            for cigar_name, _, max_qty in items_data:
                quantity_vars[cigar_name].set(max_qty)

        def clear_all():
            """Set all quantities to 0."""
            for cigar_name, _, _ in items_data:
                quantity_vars[cigar_name].set(0)

        # Stack buttons vertically on the right
        ttk.Button(button_panel, text="Select All", command=select_all, 
                  width=15).pack(pady=(0, 10), fill='x')
        
        ttk.Button(button_panel, text="Confirm", command=confirm_quantities, 
                  width=15).pack(pady=(0, 10), fill='x')
        
        ttk.Button(button_panel, text="Clear All", command=clear_all, 
                  width=15).pack(pady=(0, 10), fill='x')
        
        ttk.Button(button_panel, text="Cancel", command=cancel_quantities, 
                  width=15).pack(fill='x')

        # Add mouse wheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
        canvas.bind("<MouseWheel>", on_mousewheel)

        dialog.wait_window()
        return result

    def process_returns(self, return_quantities, current_transaction_id=None):
        """Process the returns and show confirmation."""
        if not return_quantities:
            return

        # Create return list for confirmation
        return_list = []
        for cigar_name, (brand, return_qty, max_qty) in return_quantities.items():
            return_list.append((cigar_name, brand, return_qty, max_qty))

        # Show confirmation dialog
        self.show_return_confirmation(return_list, current_transaction_id)

    def show_return_confirmation(self, return_list, current_transaction_id=None):
        """Show a single confirmation dialog for all items being returned."""
        # Create confirmation dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Confirm Return")
        dialog.geometry("650x450")  # Made wider
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set minimum size to prevent buttons from being hidden
        dialog.minsize(550, 350)

        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (325)  # Half of 650
        y = (dialog.winfo_screenheight() // 2) - (225)  # Half of 450
        dialog.geometry(f"650x450+{x}+{y}")

        # Main frame
        main_frame = ttk.Frame(dialog, padding="15")
        main_frame.pack(fill='both', expand=True)

        # Header
        ttk.Label(main_frame, text="Confirm Return", 
                 font=('TkDefaultFont', 12, 'bold')).pack(pady=(0, 10))
        
        ttk.Label(main_frame, text="Are you sure you want to return the following items?").pack(pady=(0, 15))

        # Create horizontal layout: tree on left, buttons on right
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill='both', expand=True)
        
        # Configure content_frame grid weights to ensure proper resizing
        content_frame.grid_columnconfigure(0, weight=3)  # Tree gets more space
        content_frame.grid_columnconfigure(1, weight=1)  # Button panel gets fixed space
        content_frame.grid_rowconfigure(0, weight=1)     # Allow vertical expansion

        # Left side: Items list in a treeview
        tree_frame = ttk.Frame(content_frame)
        tree_frame.grid(row=0, column=0, sticky='nsew', padx=(0, 15))

        columns = ('item', 'quantity')
        tree = ttk.Treeview(tree_frame, columns=columns, show='headings', height=12)
        
        tree.heading('item', text='Item')
        tree.heading('quantity', text='Return Qty')
        
        tree.column('item', width=400)  # Made wider
        tree.column('quantity', width=100)

        # Add scrollbar for tree
        scrollbar = ttk.Scrollbar(tree_frame, orient='vertical', command=tree.yview)
        tree.configure(yscrollcommand=scrollbar.set)
        
        # Pack tree and scrollbar
        tree.pack(side='left', fill='both', expand=True)
        scrollbar.pack(side='right', fill='y')

        # Populate the tree with return items
        total_items = 0
        for cigar_name, brand, return_qty, _ in return_list:
            values = (f"{brand} - {cigar_name}", str(return_qty))
            tree.insert('', 'end', values=values)
            total_items += return_qty

        # Right side: Buttons and summary stacked vertically
        button_panel = ttk.Frame(content_frame)
        button_panel.grid(row=0, column=1, sticky='nsew', padx=(0, 0))

        # Total summary at top of button panel
        summary_frame = ttk.LabelFrame(button_panel, text="Summary", padding="10")
        summary_frame.pack(fill='x', pady=(0, 20))
        
        ttk.Label(summary_frame, text=f"Total items to return:", 
                 font=('TkDefaultFont', 9)).pack()
        ttk.Label(summary_frame, text=f"{total_items}", 
                 font=('TkDefaultFont', 11, 'bold')).pack()

        # Button frame - always stays visible at bottom
        button_frame = ttk.Frame(button_panel)
        button_frame.pack(fill='x', side='bottom', pady=(20, 0))

        def confirm_return():
            # Process each return
            for cigar_name, brand, return_qty, original_qty in return_list:
                # Add back to inventory
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        cigar['count'] = int(cigar['count']) + return_qty
                        break

                # Update or remove sales records
                for sale in self.sales_history:
                    if (sale.get('transaction_id') == current_transaction_id and 
                        sale.get('cigar') == cigar_name):
                        
                        current_sale_qty = int(sale.get('quantity', 1))
                        if return_qty >= current_sale_qty:
                            # Remove entire sale record
                            self.sales_history.remove(sale)
                        else:
                            # Reduce quantity and recalculate cost
                            new_qty = current_sale_qty - return_qty
                            sale['quantity'] = new_qty
                            price_per_stick = float(sale.get('price_per_stick', 0))
                            sale['total_cost'] = price_per_stick * new_qty
                        break

            # Save changes and refresh displays
            self.save_inventory()
            self.save_sales_history()
            self.refresh_inventory()
            self.update_inventory_totals()
            
            dialog.destroy()
            messagebox.showinfo("Success", f"Successfully returned {total_items} items.")

        def cancel_return():
            dialog.destroy()

        # Stack buttons vertically with proper spacing
        ttk.Button(button_frame, text="Confirm Return", command=confirm_return, 
                  width=18).pack(pady=(0, 10), fill='x')
        ttk.Button(button_frame, text="Cancel", command=cancel_return, 
                  width=18).pack(fill='x')

    def return_entire_transaction(self):
        """Handle returning the entire current transaction."""
        if not self.current_transaction_id:
            messagebox.showwarning("Warning", "No transaction selected.")
            return

        # Get all items in the transaction
        transaction_items = []
        total_items = 0
        total_value = 0.0
        
        for sale in self.sales_history:
            if sale.get('transaction_id') == self.current_transaction_id:
                quantity = int(sale.get('quantity', 1))
                total_cost = float(sale.get('total_cost', 0))
                transaction_items.append((sale.get('cigar'), sale.get('brand'), quantity))
                total_items += quantity
                total_value += total_cost

        if not transaction_items:
            messagebox.showwarning("Warning", "No items found in this transaction.")
            return

        # Get transaction date for confirmation
        transaction_date = "Unknown"
        for sale in self.sales_history:
            if sale.get('transaction_id') == self.current_transaction_id:
                transaction_date = sale.get('date', 'Unknown')
                break

        # Single confirmation for entire transaction
        confirm_msg = f"Are you sure you want to return the entire transaction?\n\n"
        confirm_msg += f"Date: {transaction_date}\n"
        confirm_msg += f"Total items: {total_items}\n"
        confirm_msg += f"Total value: ${total_value:.2f}\n\n"
        confirm_msg += "Items to return:\n"
        for cigar_name, brand, quantity in transaction_items:
            confirm_msg += f"• {quantity}x {brand} - {cigar_name}\n"

        if messagebox.askyesno("Confirm Return", confirm_msg):
            # Return all items to inventory
            for cigar_name, _, quantity in transaction_items:
                for cigar in self.inventory:
                    if cigar['cigar'] == cigar_name:
                        cigar['count'] = int(cigar['count']) + quantity
                        break

            # Remove all sales records for this transaction
            self.sales_history = [sale for sale in self.sales_history 
                                if sale.get('transaction_id') != self.current_transaction_id]

            # Save changes and refresh displays
            self.save_inventory()
            self.save_sales_history()
            self.refresh_inventory()
            self.update_inventory_totals()
            
            # Refresh both this window and main application
            refresh_local_display()
            
            messagebox.showinfo("Success", f"Successfully returned entire transaction ({total_items} items).")

    def ask_return_quantity(self, item_name, max_quantity):
        """Ask user how many items to return."""
        dialog = tk.Toplevel(self.root)
        dialog.title("Return Quantity")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Set a reasonable minimum size
        dialog.minsize(400, 250)

        result = tk.IntVar(value=max_quantity)

        # Main frame with padding
        main_frame = ttk.Frame(dialog, padding="20")
        main_frame.pack(fill='both', expand=True)

        # Item name label
        ttk.Label(main_frame, text=f"Return quantity for:", font=('TkDefaultFont', 9)).pack(pady=(0, 5))
        ttk.Label(main_frame, text=item_name, font=('TkDefaultFont', 9, 'bold')).pack(pady=(0, 15))
        
        # Maximum quantity label
        ttk.Label(main_frame, text=f"Maximum: {max_quantity}").pack(pady=(0, 15))
        
        # Quantity input frame
        qty_frame = ttk.Frame(main_frame)
        qty_frame.pack(pady=(0, 20))
        
        ttk.Label(qty_frame, text="Quantity:").pack(side='left', padx=(0, 10))
        spinbox = ttk.Spinbox(qty_frame, from_=0, to=max_quantity, 
                             textvariable=result, width=10)
        spinbox.pack(side='left')

        # Button frame with adequate spacing
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(20, 0))

        def confirm():
            dialog.destroy()

        def cancel():
            result.set(0)
            dialog.destroy()

        ttk.Button(button_frame, text="Return", command=confirm, width=12).pack(side='left', padx=(0, 20))
        ttk.Button(button_frame, text="Cancel", command=cancel, width=12).pack(side='left')

        # Center the dialog after setting up all widgets
        dialog.update_idletasks()
        dialog_width = dialog.winfo_reqwidth()
        dialog_height = dialog.winfo_reqheight()
        
        # Ensure minimum size
        dialog_width = max(dialog_width, 400)
        dialog_height = max(dialog_height, 250)
        
        # Center on screen
        x = (dialog.winfo_screenwidth() // 2) - (dialog_width // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog_height // 2)
        
        dialog.geometry(f"{dialog_width}x{dialog_height}+{x}+{y}")

        spinbox.focus()
        dialog.wait_window()
        
        return result.get()

    def get_data_file_path(self, filename):
        """Get the full path for a data file in the current data directory."""
        return os.path.join(self.data_directory, filename)

    def update_location_display(self):
        """Update the location display labels."""
        self.location_label.config(text=self.data_directory)
        self.humidor_label.config(text=self.humidor_name)
        # Update window title to include humidor name
        self.root.title(f"Cigar Inventory Manager - {self.humidor_name}")

    def change_data_directory(self):
        """Allow user to change the data directory."""
        new_directory = filedialog.askdirectory(
            title="Select Data Directory",
            initialdir=self.data_directory
        )
        
        if new_directory:
            self.data_directory = new_directory
            # Try to determine humidor name from directory
            dir_name = os.path.basename(new_directory)
            if dir_name and dir_name != ".":
                self.humidor_name = dir_name
            else:
                self.humidor_name = "Custom Location"
            
            self.update_location_display()
            self.status_label.config(text="Directory changed - Loading data...", foreground='orange')
            
            # Reload data from new directory
            try:
                self.load_inventory()
                self.refresh_inventory()
                self.refresh_sales_history()
                self.refresh_resupply_dropdowns()
                self.status_label.config(text="Data loaded successfully", foreground='green')
            except Exception as e:
                self.status_label.config(text=f"Error loading data: {str(e)}", foreground='red')
                messagebox.showerror("Error", f"Failed to load data from new directory: {str(e)}")

    def new_humidor(self):
        """Create a new humidor in a new directory."""
        # Ask for humidor name
        humidor_name = tk.simpledialog.askstring(
            "New Humidor",
            "Enter name for the new humidor:",
            initialvalue="My Humidor"
        )
        
        if not humidor_name:
            return
            
        # Ask for directory location
        parent_directory = filedialog.askdirectory(
            title="Select Parent Directory for New Humidor",
            initialdir=os.path.dirname(self.data_directory)
        )
        
        if not parent_directory:
            return
            
        # Create new directory
        new_directory = os.path.join(parent_directory, humidor_name.replace(" ", "_"))
        
        try:
            if not os.path.exists(new_directory):
                os.makedirs(new_directory)
            
            # Set new directory and name
            self.data_directory = new_directory
            self.humidor_name = humidor_name
            self.update_location_display()
            
            # Clear current data
            self.inventory = []
            self.brands = set()
            self.sizes = set()
            self.types = set()
            self.sales_history = []
            self.checkbox_states = {}
            
            # Save empty data files
            self.save_inventory()
            self.save_brands()
            self.save_sets('cigar_sizes.json', self.sizes)
            self.save_sets('cigar_types.json', self.types)
            self.save_sales_history()
            
            # Refresh displays
            self.refresh_inventory()
            self.refresh_sales_history()
            self.refresh_resupply_dropdowns()
            
            self.status_label.config(text=f"New humidor '{humidor_name}' created", foreground='green')
            messagebox.showinfo("Success", f"New humidor '{humidor_name}' created successfully!")
            
        except Exception as e:
            self.status_label.config(text=f"Error creating humidor: {str(e)}", foreground='red')
            messagebox.showerror("Error", f"Failed to create new humidor: {str(e)}")

    def load_humidor(self):
        """Load an existing humidor from a directory."""
        new_directory = filedialog.askdirectory(
            title="Select Humidor Directory",
            initialdir=os.path.dirname(self.data_directory)
        )
        
        if new_directory:
            # Check if directory contains cigar inventory files
            inventory_file = os.path.join(new_directory, 'cigar_inventory.json')
            
            if not os.path.exists(inventory_file):
                if messagebox.askyesno("No Data Found", 
                    "This directory doesn't contain cigar inventory data. Create new humidor here?"):
                    # Get humidor name
                    dir_name = os.path.basename(new_directory)
                    humidor_name = tk.simpledialog.askstring(
                        "Humidor Name",
                        "Enter name for this humidor:",
                        initialvalue=dir_name if dir_name != "." else "New Humidor"
                    )
                    
                    if humidor_name:
                        self.data_directory = new_directory
                        self.humidor_name = humidor_name
                        self.update_location_display()
                        
                        # Clear and save empty data
                        self.inventory = []
                        self.brands = set()
                        self.sizes = set()
                        self.types = set()
                        self.sales_history = []
                        self.checkbox_states = {}
                        
                        self.save_inventory()
                        self.save_brands()
                        self.save_sets('cigar_sizes.json', self.sizes)
                        self.save_sets('cigar_types.json', self.types)
                        self.save_sales_history()
                        
                        self.refresh_inventory()
                        self.refresh_sales_history()
                        self.refresh_resupply_dropdowns()
                        
                        self.status_label.config(text=f"New humidor created", foreground='green')
                return
            
            # Load existing humidor
            self.data_directory = new_directory
            dir_name = os.path.basename(new_directory)
            self.humidor_name = dir_name if dir_name != "." else "Loaded Humidor"
            
            self.update_location_display()
            self.status_label.config(text="Loading humidor data...", foreground='orange')
            
            try:
                self.load_inventory()
                self.refresh_inventory()
                self.refresh_sales_history()
                self.refresh_resupply_dropdowns()
                self.status_label.config(text=f"Humidor '{self.humidor_name}' loaded", foreground='green')
            except Exception as e:
                self.status_label.config(text=f"Error loading data: {str(e)}", foreground='red')
                messagebox.showerror("Error", f"Failed to load humidor data: {str(e)}")

    def backup_data(self):
        """Create a backup of the current humidor data."""
        if not self.inventory and not self.sales_history:
            messagebox.showwarning("Warning", "No data to backup.")
            return
            
        # Ask for backup location
        backup_file = filedialog.asksaveasfilename(
            title="Save Backup As",
            defaultextension=".zip",
            filetypes=[("Zip files", "*.zip"), ("All files", "*.*")],
            initialvalue=f"{self.humidor_name}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip"
        )
        
        if backup_file:
            try:
                import zipfile
                import glob
                
                with zipfile.ZipFile(backup_file, 'w', zipfile.ZIP_DEFLATED) as zipf:
                    # Add all JSON files from the data directory
                    json_files = glob.glob(os.path.join(self.data_directory, "*.json"))
                    for file_path in json_files:
                        filename = os.path.basename(file_path)
                        zipf.write(file_path, filename)
                    
                    # Add a metadata file with humidor info
                    metadata = {
                        "humidor_name": self.humidor_name,
                        "data_directory": self.data_directory,
                        "backup_date": datetime.now().isoformat(),
                        "inventory_count": len(self.inventory),
                        "sales_count": len(self.sales_history)
                    }
                    
                    import json
                    zipf.writestr("backup_metadata.json", json.dumps(metadata, indent=2))
                
                self.status_label.config(text="Backup created successfully", foreground='green')
                messagebox.showinfo("Success", f"Backup created: {backup_file}")
                
            except Exception as e:
                self.status_label.config(text=f"Backup failed: {str(e)}", foreground='red')
                messagebox.showerror("Error", f"Failed to create backup: {str(e)}")

    def refresh_sales_history(self):
        """Refresh the sales history display with transaction grouping."""
        # Check if the sales history UI has been created
        if not hasattr(self, 'transaction_tree') or not hasattr(self, 'detail_tree'):
            return  # UI not created yet, skip refresh
            
        # Clear current display
        for item in self.transaction_tree.get_children():
            self.transaction_tree.delete(item)
            
        # Clear detail view if no transaction is selected
        if not hasattr(self, 'current_transaction_id') or not self.current_transaction_id:
            for item in self.detail_tree.get_children():
                self.detail_tree.delete(item)
        
        # Group sales by transaction_id
        transactions = {}
        for sale in self.sales_history:
            transaction_id = sale.get('transaction_id', 'unknown')
            if transaction_id not in transactions:
                transactions[transaction_id] = []
            transactions[transaction_id].append(sale)
        
        # Store transaction ID mapping for later retrieval
        if not hasattr(self, 'transaction_id_map'):
            self.transaction_id_map = {}
        
        # Display transactions (newest first)
        sorted_transactions = sorted(
            transactions.items(), 
            key=lambda x: x[1][0].get('date', ''), 
            reverse=True
        )
        
        for transaction_id, sales in sorted_transactions:
            try:
                # Calculate transaction totals
                total_items = sum(int(sale.get('quantity', 1)) for sale in sales)
                total_value = sum(float(sale.get('total_cost', 0)) for sale in sales)
                
                # Use the date from the first sale in the transaction
                date = sales[0].get('date', 'Unknown')
                
                # Create transaction summary
                values = (
                    date,
                    str(total_items),
                    f"${total_value:.2f}"
                )
                
                # Insert the item and store the transaction_id in our mapping
                item_id = self.transaction_tree.insert('', 'end', values=values)
                self.transaction_id_map[item_id] = transaction_id
                
            except Exception as e:
                print(f"Error displaying transaction {transaction_id}: {e}")
        
        # If we have a current transaction selected, refresh its details
        if hasattr(self, 'current_transaction_id') and self.current_transaction_id:
            self.refresh_transaction_details()

    # Add new method after the setup_resupply_tab method
    def refresh_resupply_dropdowns(self):
        """Refresh the resupply tab dropdown values after data is loaded."""
        # Update brand dropdown
        self.resupply_brand_combo.config(values=sorted(list(self.brands)))
        
        # Update cigar dropdown with all existing cigars
        existing_cigars = sorted(list(set(cigar.get('cigar', '') for cigar in self.inventory if cigar.get('cigar', ''))))
        self.resupply_cigar_combo.config(values=existing_cigars)
        
        # Update size dropdown
        self.resupply_size_combo.config(values=sorted(list(self.sizes)))
        
        # Update type dropdown  
        self.resupply_type_combo.config(values=sorted(list(self.types)))

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