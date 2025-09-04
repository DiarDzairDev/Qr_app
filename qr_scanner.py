# -*- coding: utf-8 -*-
"""
Python QR Code Scanner and Generator Application
Supports barcode scanner (douchette) input and Excel integration
"""

import pandas as pd
import qrcode
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk, ImageDraw
import os
from dataclasses import dataclass, asdict

@dataclass
class ProductData:
    """Data structure for product information"""
    ID_Produit: str = ""
    Marque: str = ""
    Designation_Reference: str = ""
    Serial_Number: str = ""
    Couleur: str = ""
    Matricule: str = ""
    Magasin: str = ""
    Photo: str = ""

class QRScannerApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("QR Code Scanner & Generator")
        self.root.geometry("1280x720")
        
        # Configure encoding for French characters
        self.root.option_add('*Font', 'TkDefaultFont')
        
        # Data storage
        self.products_data = []
        self.excel_file = None
        
        # Scanner state
        self.scanning = False
        self.scan_buffer = ""
        self.scan_timer = None  # Timer for auto-processing scanned data
        
        self.setup_ui()
        self.setup_scanner_listener()
    
    def setup_ui(self):
        """Setup the user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        
        # Title
        title_label = ttk.Label(main_frame, text="QR Code Scanner & Generator", 
                               font=('Arial', 16, 'bold'))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))
        
        # File operations frame
        file_frame = ttk.LabelFrame(main_frame, text="File Operations", padding="10")
        file_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(file_frame, text="Load Excel File", 
                  command=self.load_excel_file).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(file_frame, text="Save Excel File", 
                  command=self.save_excel_file).grid(row=0, column=1, padx=5)
        ttk.Button(file_frame, text="Clear All Data", 
                  command=self.clear_all_data).grid(row=0, column=2, padx=5)
        
        self.file_label = ttk.Label(file_frame, text="No file loaded")
        self.file_label.grid(row=1, column=0, columnspan=3, pady=(10, 0))
        
        # Scanner frame
        scanner_frame = ttk.LabelFrame(main_frame, text="QR Code Scanner", padding="10")
        scanner_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        # Scanner input field
        ttk.Label(scanner_frame, text="Scanner Input:").grid(row=0, column=0, sticky=tk.NW, pady=(5, 0))
        
        # Create a frame for the text widget and scrollbar
        text_frame = ttk.Frame(scanner_frame)
        text_frame.grid(row=0, column=1, padx=(10, 0), sticky=(tk.W, tk.E, tk.N, tk.S))
        text_frame.columnconfigure(0, weight=1)
        
        # Text widget for multi-line input
        self.scanner_entry = tk.Text(text_frame, width=60, height=4, font=('Courier', 10), wrap='word')
        self.scanner_entry.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Scrollbar for the text widget
        scanner_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.scanner_entry.yview)
        scanner_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.scanner_entry.configure(yscrollcommand=scanner_scrollbar.set)
        
        # Bind events (note: Text widget uses different event handling)
        self.scanner_entry.bind('<KeyRelease>', self.on_scanner_input)
        
        scanner_frame.columnconfigure(1, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(scanner_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="Manual Input", 
                  command=self.manual_input).grid(row=0, column=2, padx=5)
        
        # Status
        self.status_label = ttk.Label(scanner_frame, text="Ready to scan...", 
                                     foreground="green")
        self.status_label.grid(row=2, column=0, columnspan=2, pady=(10, 0))
        
        # Data display frame
        data_frame = ttk.LabelFrame(main_frame, text="Product Data", padding="10")
        data_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        main_frame.rowconfigure(3, weight=1)
        
        # Search and filter frame
        search_frame = ttk.Frame(data_frame)
        search_frame.grid(row=0, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        search_frame.columnconfigure(1, weight=1)
        
        # Search functionality
        ttk.Label(search_frame, text="Search:").grid(row=0, column=0, padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.grid(row=0, column=1, padx=(0, 10), sticky=(tk.W, tk.E))
        
        # Filter by field
        ttk.Label(search_frame, text="Filter by:").grid(row=0, column=2, padx=(10, 5))
        self.filter_field = tk.StringVar(value="All Fields")
        filter_combo = ttk.Combobox(search_frame, textvariable=self.filter_field, width=15, state="readonly")
        filter_combo['values'] = ("All Fields", "ID_Produit", "Marque", "Designation_Reference", 
                                  "Serial_Number", "Couleur", "Matricule", "Magasin")
        filter_combo.grid(row=0, column=3, padx=(0, 10))
        filter_combo.bind('<<ComboboxSelected>>', self.on_filter_change)
        
        # CRUD buttons
        crud_frame = ttk.Frame(data_frame)
        crud_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(crud_frame, text="Edit Selected", 
                  command=self.edit_selected_record).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(crud_frame, text="Delete Selected", 
                  command=self.delete_selected_record).grid(row=0, column=1, padx=5)
        ttk.Button(crud_frame, text="Duplicate Selected", 
                  command=self.duplicate_selected_record).grid(row=0, column=2, padx=5)
        
        # Treeview for data display
        columns = ('ID_Produit', 'Marque', 'Designation_Reference', 'Serial_Number', 
                  'Couleur', 'Matricule', 'Magasin', 'Photo')
        
        self.tree = ttk.Treeview(data_frame, columns=columns, show='headings', height=10)
        
        # Configure column headings and widths with sorting
        for col in columns:
            self.tree.heading(col, text=col.replace('_', ' '), 
                            command=lambda c=col: self.sort_column(c, False))
            self.tree.column(col, width=120, minwidth=80)
        
        # Enable editing on double-click
        self.tree.bind('<Double-1>', self.on_item_double_click)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=2, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=2, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=3, column=0, sticky=(tk.W, tk.E))
        
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(2, weight=1)
        
        # QR Generation frame
        qr_frame = ttk.LabelFrame(main_frame, text="QR Code Generation", padding="10")
        qr_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(qr_frame, text="Generate QR from Selected Row", 
                  command=self.generate_qr_from_selection).grid(row=0, column=0, padx=(0, 5))
        
        # Focus scanner input
        self.scanner_entry.focus_set()
    
    def setup_scanner_listener(self):
        """Setup scanner input detection"""
        self.scanner_entry.focus_set()
    
    def ignore_enter_key(self, event):
        """Ignore Enter key press to prevent accidental processing"""
        return "break"  # This prevents the default Enter key behavior
    
    def on_scanner_input(self, event):
        """Handle scanner input in real-time"""
        current_text = self.scanner_entry.get("1.0", tk.END).strip()
        
        # Ensure proper encoding for French characters
        if current_text:
            try:
                # Try to encode/decode to ensure proper UTF-8 handling
                current_text.encode('utf-8').decode('utf-8')
            except (UnicodeDecodeError, UnicodeEncodeError):
                # If there are encoding issues, try to fix them
                current_text = current_text.encode('utf-8', errors='ignore').decode('utf-8')
        
        # Cancel any existing timer
        if self.scan_timer:
            self.root.after_cancel(self.scan_timer)
            self.scan_timer = None
        
        if len(current_text) > 0:
            self.status_label.config(text="Receiving data...", foreground="orange")
            # Set timer to auto-process after 250ms of no input
            self.scan_timer = self.root.after(250, self.auto_process_scan)
        else:
            self.status_label.config(text="Ready to scan...", foreground="green")
    
    def auto_process_scan(self):
        """Automatically process scan after delay"""
        current_text = self.scanner_entry.get("1.0", tk.END).strip()
        if current_text and not self.scanning:
            self.scanning = True
            self.process_scanned_data(None)
    
    def process_scanned_data(self, event=None):
        """Process the scanned QR code data"""
        qr_data = self.scanner_entry.get("1.0", tk.END).strip()
        
        if not qr_data:
            messagebox.showwarning("No Data", "No QR code data to process!")
            return
        
        try:
            # Parse the QR code data
            product_data = self.parse_qr_data(qr_data)
            
            # Add to data list
            self.products_data.append(product_data)
            
            # Update display
            self.update_tree_display()
            
            # Clear scanner input
            self.clear_scanner_input()
            
            self.status_label.config(text=f"Successfully added product: {product_data.ID_Produit}", 
                                   foreground="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process QR data: {str(e)}")
            self.status_label.config(text="Error processing scan", foreground="red")
        
        finally:
            self.scanning = False
    
    def parse_qr_data(self, qr_data: str) -> ProductData:
        """Parse QR code data with backward compatibility"""
        product = ProductData()
        
        # Check if data contains line breaks (structured format)
        if '\n' in qr_data or '\r' in qr_data:
            # Split by lines and clean up
            lines = qr_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
            lines = [line.strip() for line in lines if line.strip()]
            
            # Check if first line starts with asterisk (legacy format)
            if lines and lines[0].startswith('*') and lines[0].endswith('*'):
                # Legacy format with asterisks
                # 1st line: *ID* (between asterisks)
                # 2nd line: Designation/Reference  
                # 3rd line: Marque
                # 4th line: Couleur
                # 5th line: Magasin (Unité...)
                # 6th line: Serial Number
                
                legacy_field_mapping = [
                    'ID_Produit',        # *VMSDZ06CUKI191698* -> extract content between asterisks
                    'Designation_Reference',  # MOTOCYCLE CUKI -I-
                    'Marque',            # CUKI
                    'Couleur',           # bleu nuit/ blanc
                    'Magasin',           # Unité Oued-Ghir
                    'Serial_Number'      # CUKI I 06/2025
                ]
                for i, line in enumerate(lines):
                    if i < len(legacy_field_mapping):
                        field_name = legacy_field_mapping[i]
                        if i == 0:  # First line with asterisks
                            # Extract content between asterisks
                            if line.startswith('*') and line.endswith('*'):
                                setattr(product, field_name, line[1:-1])  # Remove asterisks
                            else:
                                setattr(product, field_name, line)
                        else:
                            setattr(product, field_name, line)
                
                # Matricule is empty in legacy format (as specified)
                product.Matricule = ""
                
            else:
                # New structured format - split by lines
                field_mapping = [
                    'ID_Produit',
                    'Marque',
                    'Designation_Reference',
                    'Serial_Number',
                    'Couleur',
                    'Matricule',
                    'Magasin',
                    'Photo'
                ]
                
                for i, line in enumerate(lines):
                    if i < len(field_mapping):
                        setattr(product, field_mapping[i], line)
        else:
            # Scanner input (single line) - parse using generic field detection
            product = self.parse_scanner_data_generic(qr_data)
        
        return product
    
    def test_legacy_parsing(self):
        """Test the legacy format parsing with sample data"""
        # Sample data from scan.txt
        sample_legacy_data = """*VMSDZ06CUKI191698*
MOTOCYCLE CUKI -I-
CUKI
bleu nuit/ blanc
Unitª Oued-Ghir
CUKI I 06/2025"""
        
        try:
            product = self.parse_qr_data(sample_legacy_data)
            print("Legacy parsing test:")
            print(f"ID_Produit: {product.ID_Produit}")
            print(f"Designation_Reference: {product.Designation_Reference}")
            print(f"Marque: {product.Marque}")
            print(f"Couleur: {product.Couleur}")
            print(f"Matricule: {product.Matricule}")
            print(f"Magasin: {product.Magasin}")
            print(f"Serial_Number: {product.Serial_Number}")
            return product
        except Exception as e:
            print(f"Error in legacy parsing: {e}")
            return None
    
    def parse_scanner_data_generic(self, scanner_data: str) -> ProductData:
        """Parse scanner data - handles line breaks"""
        product = ProductData()
        
        # Ensure proper encoding for French characters
        if isinstance(scanner_data, bytes):
            scanner_data = scanner_data.decode('utf-8', errors='ignore')
        
        # Check if the data contains any line break characters
        if '\n' in scanner_data or '\r' in scanner_data:
            # Parse normally using line breaks
            lines = scanner_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
            lines = [line.strip() for line in lines if line.strip()]
            
            # Map lines to fields based on position (generic order)
            field_mapping = [
                'ID_Produit',
                'Marque',
                'Designation_Reference',
                'Serial_Number',
                'Couleur',
                'Matricule',
                'Magasin',
                'Photo'
            ]
            
            for i, line in enumerate(lines):
                if i < len(field_mapping):
                    setattr(product, field_mapping[i], line)
        else:
            # No line breaks found - treat entire input as ID_Produit only
            product.ID_Produit = scanner_data.strip()
        
        return product
    
    def generate_qr_data(self, product_data: ProductData) -> str:
        """Generate QR code data in structured format"""
        qr_lines = [
            product_data.ID_Produit,
            product_data.Marque,
            product_data.Designation_Reference,
            product_data.Serial_Number,
            product_data.Couleur,
            product_data.Matricule,
            product_data.Magasin,
            product_data.Photo
        ]
        
        return '\n'.join(qr_lines)
    
    def clear_scanner_input(self):
        """Clear the scanner input field"""
        # Cancel any pending scan timer
        if self.scan_timer:
            self.root.after_cancel(self.scan_timer)
            self.scan_timer = None
            
        self.scanner_entry.delete("1.0", tk.END)
        self.scanner_entry.focus_set()
        self.status_label.config(text="Ready to scan...", foreground="green")
    
    def manual_input(self):
        """Allow manual input of product data via form"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Manual Product Input")
        dialog.geometry("500x600")
        dialog.resizable(True, True)
        
        # Main frame with scrollbar
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title
        ttk.Label(main_frame, text="Enter Product Information:", 
                 font=('Arial', 12, 'bold')).pack(pady=(0, 20))
        
        # Create entry fields
        fields = {}
        field_labels = {
            'ID_Produit': 'ID Produit:',
            'Marque': 'Marque:',
            'Designation_Reference': 'Désignation/Référence:',
            'Serial_Number': 'Numéro de Série:',
            'Couleur': 'Couleur:',
            'Matricule': 'Matricule:',
            'Magasin': 'Magasin:',
            'Photo': 'Photo:'
        }
        
        # Create form fields
        for field, label in field_labels.items():
            frame = ttk.Frame(main_frame)
            frame.pack(fill='x', pady=5)
            
            ttk.Label(frame, text=label, width=20).pack(side='left')
            entry = ttk.Entry(frame, width=40)
            entry.pack(side='left', fill='x', expand=True, padx=(10, 0))
            fields[field] = entry
        
        # Pre-fill with sample data
        sample_data = {
            'ID_Produit': 'VMSDZ06CUKI191858',
            'Marque': 'CUKI',
            'Designation_Reference': 'MOTOCYCLE CUKI -I-',
            'Serial_Number': 'CUKI I 06/2025',
            'Couleur': 'bleu nuit/ blanc',
            'Matricule': 'Unité Oued-Ghir',
            'Magasin': '',
            'Photo': ''
        }
        
        for field, value in sample_data.items():
            fields[field].insert(0, value)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        def add_product():
            # Create ProductData from form inputs
            product = ProductData()
            for field, entry in fields.items():
                setattr(product, field, entry.get().strip())
            
            # Add to data list
            self.products_data.append(product)
            
            # Update display
            self.update_tree_display()
            
            # Show success message
            messagebox.showinfo("Succès", f"Produit ajouté: {product.ID_Produit}")
            dialog.destroy()
        
        def clear_form():
            for entry in fields.values():
                entry.delete(0, tk.END)
        
        ttk.Button(button_frame, text="Ajouter Produit", 
                  command=add_product).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Effacer Formulaire", 
                  command=clear_form).pack(side='left', padx=5)
        ttk.Button(button_frame, text="Annuler", 
                  command=dialog.destroy).pack(side='left', padx=5)
    
    def update_tree_display(self):
        """Update the treeview with current data, applying search and filter"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Get search text and filter field
        search_text = self.search_var.get().lower() if hasattr(self, 'search_var') else ""
        filter_field = self.filter_field.get() if hasattr(self, 'filter_field') else "All Fields"
        
        # Filter and search products
        filtered_products = []
        for i, product in enumerate(self.products_data):
            # Apply search filter
            if search_text:
                product_values = [
                    product.ID_Produit or "",
                    product.Marque or "",
                    product.Designation_Reference or "",
                    product.Serial_Number or "",
                    product.Couleur or "",
                    product.Matricule or "",
                    product.Magasin or "",
                    product.Photo or ""
                ]
                
                # Check if search text is in the filtered field or all fields
                if filter_field == "All Fields":
                    if not any(search_text in str(value).lower() for value in product_values):
                        continue
                else:
                    # Map filter field to product attribute
                    field_mapping = {
                        "ID_Produit": product.ID_Produit,
                        "Marque": product.Marque,
                        "Designation_Reference": product.Designation_Reference,
                        "Serial_Number": product.Serial_Number,
                        "Couleur": product.Couleur,
                        "Matricule": product.Matricule,
                        "Magasin": product.Magasin
                    }
                    
                    field_value = field_mapping.get(filter_field, "")
                    if search_text not in str(field_value or "").lower():
                        continue
            
            filtered_products.append((i, product))
        
        # Add filtered products to tree
        for original_index, product in filtered_products:
            values = (
                product.ID_Produit or "",
                product.Marque or "",
                product.Designation_Reference or "",
                product.Serial_Number or "",
                product.Couleur or "",
                product.Matricule or "",
                product.Magasin or "",
                product.Photo or ""
            )
            self.tree.insert('', 'end', values=values, tags=(str(original_index),))
    
    def generate_qr_from_selection(self):
        """Generate QR code from selected row"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("No Selection", "Please select a row to generate QR code")
            return
        
        # Get selected item index
        item = selection[0]
        tags = self.tree.item(item, 'tags')
        if tags:
            index = int(tags[0])
            product = self.products_data[index]
            # Generate and display QR code
            self.show_qr_code(product)
    
    def show_qr_code(self, product_data: ProductData):
        """Display QR code with logo and rounded pixels in a new window"""
        qr_data = self.generate_qr_data(product_data)
        
        # Create QR code with higher error correction for logo overlay
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_H,  # High error correction for logo
            box_size=10,
            border=4,
        )
        qr.add_data(qr_data)
        qr.make(fit=True)
        
        # Get the QR code matrix
        qr_matrix = qr.get_matrix()
        
        # Create custom QR code with rounded pixels
        qr_image = self.create_rounded_qr_image(qr_matrix, box_size=10, border=4)
        
        # Try to add logo to QR code
        try:
            logo_path = os.path.join(os.path.dirname(__file__), 'logo_diardzair.jpg')
            if os.path.exists(logo_path):
                # Load and process logo
                logo = Image.open(logo_path)
                
                # Convert logo to RGB if needed
                if logo.mode in ('P', 'PA'):  # Palette mode with transparency
                    logo = logo.convert('RGBA')
                elif logo.mode != 'RGBA' and 'transparency' in logo.info:
                    logo = logo.convert('RGBA')
                
                # Calculate logo size (about 10-15% of QR code size)
                qr_width, qr_height = qr_image.size
                logo_size = min(qr_width, qr_height) // 5  # About 20% of QR code for better visibility
                
                # Resize logo maintaining aspect ratio
                logo.thumbnail((logo_size, logo_size), Image.Resampling.LANCZOS)
                
                # Create a white background for the logo with rounded corners effect
                logo_bg_size = logo_size + 10  # Less padding for better proportions
                logo_bg = Image.new('RGB', (logo_bg_size, logo_bg_size), 'white')
                
                # Calculate position to center logo on white background
                logo_x = (logo_bg_size - logo.size[0]) // 2
                logo_y = (logo_bg_size - logo.size[1]) // 2
                
                # Paste logo onto white background
                if logo.mode == 'RGBA':
                    # Handle RGBA logos properly
                    logo_bg.paste(logo, (logo_x, logo_y), logo)
                else:
                    # For RGB or other modes, convert to RGB first
                    if logo.mode != 'RGB':
                        logo = logo.convert('RGB')
                    logo_bg.paste(logo, (logo_x, logo_y))
                
                # Calculate position to center logo on QR code
                pos_x = (qr_width - logo_bg_size) // 2
                pos_y = (qr_height - logo_bg_size) // 2
                
                # Paste logo with background onto QR code
                qr_image.paste(logo_bg, (pos_x, pos_y))
                
                
            else:
                print(f"Logo file not found at: {logo_path}")
                
        except Exception as e:
            print(f"Error adding logo to QR code: {e}")
            print("Continuing without logo...")
            # Continue without logo if there's an error
        
        # Display in new window with scrollable canvas
        qr_window = tk.Toplevel(self.root)
        qr_window.title(f"QR Code - {product_data.ID_Produit}")
        qr_window.geometry("600x700")
        qr_window.resizable(True, True)
        
        # Create main frame
        main_frame = ttk.Frame(qr_window)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Create canvas and scrollbar for scrollable content
        canvas = tk.Canvas(main_frame)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
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
        
        # Convert PIL image to PhotoImage
        photo = ImageTk.PhotoImage(qr_image)
        
        # QR Code display in scrollable frame
        qr_label = ttk.Label(scrollable_frame, image=photo)
        qr_label.image = photo  # Keep a reference
        qr_label.pack(pady=20)
        
        # Product information frame
        info_frame = ttk.LabelFrame(scrollable_frame, text="Informations du Produit", padding="10")
        info_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # Display product information
        product_info = [
            ("ID Produit:", product_data.ID_Produit),
            ("Marque:", product_data.Marque),
            ("Désignation/Référence:", product_data.Designation_Reference),
            ("Numéro de Série:", product_data.Serial_Number),
            ("Couleur:", product_data.Couleur),
            ("Matricule:", product_data.Matricule),
            ("Magasin:", product_data.Magasin)
        ]
        
        for i, (label_text, value) in enumerate(product_info):
            if value:  # Only show non-empty fields
                info_label = ttk.Label(info_frame, text=f"{label_text} {value}", font=('Arial', 10))
                info_label.pack(anchor='w', pady=2)
        
        # Buttons frame (fixed at bottom of window)
        button_frame = ttk.Frame(qr_window)
        button_frame.pack(side='bottom', pady=10)
        
        # Save button
        def save_qr():
            filename = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )
            if filename:
                qr_image.save(filename)
                messagebox.showinfo("Saved", f"QR code saved as {filename}")
        
        # Print button
        def print_qr():
            self.print_qr_code(qr_image, product_data)
        
        ttk.Button(button_frame, text="Save QR Code", command=save_qr).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Print QR Code", command=print_qr).pack(side=tk.LEFT, padx=5)
        
        # Enable mouse wheel scrolling
        def on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        
        canvas.bind("<MouseWheel>", on_mousewheel)
    
    def create_rounded_qr_image(self, matrix, box_size=10, border=4, corner_radius=2):
        """Create a QR code image with rounded pixels"""
        from PIL import ImageDraw
        
        # Calculate dimensions
        width = height = len(matrix) * box_size + 2 * border * box_size
        
        # Create image
        image = Image.new('RGB', (width, height), 'white')
        draw = ImageDraw.Draw(image)
        
        # Draw rounded rectangles for each black pixel
        for row in range(len(matrix)):
            for col in range(len(matrix[row])):
                if matrix[row][col]:  # Black pixel
                    # Calculate position
                    x1 = col * box_size + border * box_size
                    y1 = row * box_size + border * box_size
                    x2 = x1 + box_size
                    y2 = y1 + box_size
                    
                    # Draw rounded rectangle
                    self.draw_rounded_rectangle(draw, (x1, y1, x2, y2), corner_radius, 'black')
        
        return image
    
    def draw_rounded_rectangle(self, draw, bbox, corner_radius, fill):
        """Draw a rounded rectangle"""
        x1, y1, x2, y2 = bbox
        
        # Ensure corner radius is not too large
        corner_radius = min(corner_radius, (x2 - x1) // 2, (y2 - y1) // 2)
        
        # Draw the rounded rectangle using multiple shapes
        # Main rectangle (center)
        draw.rectangle([x1 + corner_radius, y1, x2 - corner_radius, y2], fill=fill)
        draw.rectangle([x1, y1 + corner_radius, x2, y2 - corner_radius], fill=fill)
        
        # Four corners
        draw.pieslice([x1, y1, x1 + 2 * corner_radius, y1 + 2 * corner_radius], 180, 270, fill=fill)
        draw.pieslice([x2 - 2 * corner_radius, y1, x2, y1 + 2 * corner_radius], 270, 360, fill=fill)
        draw.pieslice([x1, y2 - 2 * corner_radius, x1 + 2 * corner_radius, y2], 90, 180, fill=fill)
        draw.pieslice([x2 - 2 * corner_radius, y2 - 2 * corner_radius, x2, y2], 0, 90, fill=fill)
    
    def print_qr_code(self, qr_image, product_data):
        """Print only the QR code and open Windows print dialog directly"""
        try:
            import tempfile
            import subprocess
            import platform
            import os
            
            # Create a high-resolution version of just the QR code for printing
            # Scale up the QR code for better print quality
            print_size = 600  # Larger size for better print quality
            qr_print_image = qr_image.resize((print_size, print_size), Image.Resampling.LANCZOS)
            
            # Save to temporary file with high DPI
            with tempfile.NamedTemporaryFile(suffix='.png', delete=False) as temp_file:
                temp_path = temp_file.name
                qr_print_image.save(temp_path, 'PNG', dpi=(300, 300))  # High DPI for printing
            
            # Print based on operating system
            system = platform.system()
            
            if system == "Windows":
                try:
                    # Try to use Windows print dialog directly
                    import win32print
                    import win32ui
                    import win32con
                    from PIL import ImageWin
                    
                    # Get default printer
                    printer_name = win32print.GetDefaultPrinter()
                    
                    # Create printer device context
                    hDC = win32ui.CreateDC()
                    hDC.CreatePrinterDC(printer_name)
                    
                    # Start print job
                    hDC.StartDoc(f"QR Code - {product_data.ID_Produit}")
                    hDC.StartPage()
                    
                    # Get printer capabilities
                    printable_area = hDC.GetDeviceCaps(win32con.HORZRES), hDC.GetDeviceCaps(win32con.VERTRES)
                    printer_size = hDC.GetDeviceCaps(win32con.HORZSIZE), hDC.GetDeviceCaps(win32con.VERTSIZE)
                    
                    # Calculate position to center QR code
                    qr_size_mm = 50  # 50mm QR code
                    qr_size_pixels = int((qr_size_mm / printer_size[0]) * printable_area[0])
                    
                    x = (printable_area[0] - qr_size_pixels) // 2
                    y = (printable_area[1] - qr_size_pixels) // 2
                    
                    # Print the QR code
                    dib = ImageWin.Dib(qr_print_image)
                    dib.draw(hDC.GetHandleOutput(), (x, y, x + qr_size_pixels, y + qr_size_pixels))
                    
                    # End print job
                    hDC.EndPage()
                    hDC.EndDoc()
                    hDC.DeleteDC()
                    
                    messagebox.showinfo("Impression", "QR code envoyé à l'imprimante avec succès!")
                    
                except ImportError:
                    # Fallback: Open print dialog through default image viewer
                    # Use mspaint to open print dialog directly
                    subprocess.run(['mspaint', '/p', temp_path], check=False)
                    
                except Exception as print_error:
                    print(f"Direct print failed: {print_error}")
                    # Fallback: Open with default viewer
                    os.startfile(temp_path, 'print')
                    
            elif system == "Darwin":  # macOS
                # Use lpr command for direct printing
                subprocess.run(['lpr', temp_path], check=True)
                messagebox.showinfo("Impression", "QR code envoyé à l'imprimante!")
                
            else:  # Linux
                # Use lp command for direct printing
                subprocess.run(['lp', temp_path], check=True)
                messagebox.showinfo("Impression", "QR code envoyé à l'imprimante!")
                
        except Exception as e:
            messagebox.showerror("Erreur d'impression", f"Impossible d'imprimer: {str(e)}")
    
    def create_printable_qr_image(self, qr_image, product_data):
        """Create a printable image with QR code and product information"""
        try:
            from PIL import ImageFont
            
            # Create a larger canvas for printing (A4-like proportions)
            canvas_width = 800
            canvas_height = 1000
            canvas = Image.new('RGB', (canvas_width, canvas_height), 'white')
            
            # Resize QR code for printing (larger)
            qr_size = 400
            qr_resized = qr_image.resize((qr_size, qr_size), Image.Resampling.LANCZOS)
            
            # Position QR code on canvas
            qr_x = (canvas_width - qr_size) // 2
            qr_y = 100
            canvas.paste(qr_resized, (qr_x, qr_y))
            
            # Add text information below QR code
            draw = ImageDraw.Draw(canvas)
            
            # Try to use a system font, fallback to default
            try:
                title_font = ImageFont.truetype("arial.ttf", 24)
                text_font = ImageFont.truetype("arial.ttf", 16)
            except:
                try:
                    title_font = ImageFont.truetype("calibri.ttf", 24)
                    text_font = ImageFont.truetype("calibri.ttf", 16)
                except:
                    # Fallback to default font
                    title_font = ImageFont.load_default()
                    text_font = ImageFont.load_default()
            
            # Title
            title = "QR Code - Product Information"
            title_bbox = draw.textbbox((0, 0), title, font=title_font)
            title_width = title_bbox[2] - title_bbox[0]
            title_x = (canvas_width - title_width) // 2
            draw.text((title_x, 30), title, fill='black', font=title_font)
            
            # Product information
            info_start_y = qr_y + qr_size + 50
            line_height = 30
            
            product_info = [
                f"ID Produit: {product_data.ID_Produit}",
                f"Marque: {product_data.Marque}",
                f"Désignation: {product_data.Designation_Reference}",
                f"Numéro de Série: {product_data.Serial_Number}",
                f"Couleur: {product_data.Couleur}",
                f"Matricule: {product_data.Matricule}",
                f"Magasin: {product_data.Magasin}"
            ]
            
            for i, info_line in enumerate(product_info):
                if info_line.split(': ')[1]:  # Only show non-empty fields
                    y_pos = info_start_y + (i * line_height)
                    draw.text((50, y_pos), info_line, fill='black', font=text_font)
            
            # Add timestamp
            from datetime import datetime
            timestamp = f"Généré le: {datetime.now().strftime('%d/%m/%Y à %H:%M')}"
            timestamp_y = canvas_height - 50
            draw.text((50, timestamp_y), timestamp, fill='gray', font=text_font)
            
            return canvas
            
        except Exception as e:
            print(f"Error creating printable image: {e}")
            # Return the original QR image if there's an error
            return qr_image
    
    def load_excel_file(self):
        """Load data from Excel file"""
        filename = filedialog.askopenfilename(
            filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
        )
        if filename:
            try:
                df = pd.read_excel(filename)
                
                # Mapping between Excel column names and ProductData fields
                column_mapping = {
                    'ID_Produit': 'ID_Produit',
                    'Marque': 'Marque',
                    'Designation/Reference': 'Designation_Reference',
                    'Designation_Reference': 'Designation_Reference',
                    'Serial Number': 'Serial_Number',
                    'Serial_Number': 'Serial_Number',
                    'Couleur': 'Couleur',
                    'Matricule': 'Matricule',
                    'Magasin': 'Magasin',
                    'Photo': 'Photo'
                }
                
                # Convert dataframe to ProductData objects
                self.products_data = []
                for _, row in df.iterrows():
                    product = ProductData()
                    
                    # Map Excel columns to ProductData fields
                    for excel_col, product_field in column_mapping.items():
                        if excel_col in df.columns:
                            value = row[excel_col]
                            setattr(product, product_field, str(value) if pd.notna(value) else "")
                    
                    self.products_data.append(product)
                
                self.excel_file = filename
                self.update_tree_display()
                self.file_label.config(text=f"Loaded: {os.path.basename(filename)}")
                messagebox.showinfo("Success", f"Loaded {len(self.products_data)} records")
                
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load file: {str(e)}")
    
    def save_excel_file(self):
        """Save data to Excel file"""
        # If no file is currently loaded, ask for filename
        if not self.excel_file:
            filename = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                title="Save Excel File As..."
            )
            if not filename:
                return  # User cancelled
            self.excel_file = filename
        
        # Check if there's data to save
        if not self.products_data:
            messagebox.showwarning("No Data", "No product data to save!")
            return
        
        try:
            # Convert ProductData objects to dictionary list
            data_list = [asdict(product) for product in self.products_data]
            
            # Create DataFrame
            df = pd.DataFrame(data_list)
            
            # Rename columns to match Excel headers
            df.columns = [
                'ID_Produit', 'Marque', 'Designation/Reference', 
                'Serial Number', 'Couleur', 'Matricule', 
                'Magasin', 'Photo'
            ]
            
            # Save to Excel
            df.to_excel(self.excel_file, index=False)
            
            # Update file label
            self.file_label.config(text=f"File: {os.path.basename(self.excel_file)}")
            
            messagebox.showinfo("Success", f"Data saved to {os.path.basename(self.excel_file)}")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to save file: {str(e)}")

    def clear_all_data(self):
        """Clear all data from the application"""
        if not self.products_data:
            messagebox.showinfo("Info", "Aucune donnée à effacer")
            return
        
        # Confirm with user
        result = messagebox.askyesno(
            "Confirmation", 
            f"Êtes-vous sûr de vouloir effacer toutes les données?\n"
            f"Cela supprimera {len(self.products_data)} produits.",
            icon='warning'
        )
        
        if result:
            # Clear all data
            self.products_data = []
            self.excel_file = None
            
            # Update UI
            self.update_tree_display()
            self.file_label.config(text="No file loaded")
            self.clear_scanner_input()
            
            messagebox.showinfo("Succès", "Toutes les données ont été effacées")

    def edit_selected_record(self):
        """Edit the selected record in the tree"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Sélection", "Veuillez sélectionner un produit à modifier")
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        # Find the index in products_data
        index = self.find_product_index_by_values(values)
        if index is None:
            messagebox.showerror("Erreur", "Impossible de trouver le produit sélectionné")
            return
        
        # Open edit dialog
        self.open_edit_dialog(index)
    
    def delete_selected_record(self):
        """Delete the selected record from the tree"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Sélection", "Veuillez sélectionner un produit à supprimer")
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        # Confirm deletion
        result = messagebox.askyesno(
            "Confirmation", 
            f"Êtes-vous sûr de vouloir supprimer ce produit?\n\n"
            f"ID: {values[0]}\n"
            f"Marque: {values[1]}\n"
            f"Référence: {values[2]}",
            icon='warning'
        )
        
        if result:
            # Find the index in products_data
            index = self.find_product_index_by_values(values)
            if index is not None:
                del self.products_data[index]
                self.update_tree_display()
                messagebox.showinfo("Succès", "Produit supprimé avec succès")
            else:
                messagebox.showerror("Erreur", "Impossible de trouver le produit à supprimer")
    
    def duplicate_selected_record(self):
        """Duplicate the selected record"""
        selection = self.tree.selection()
        if not selection:
            messagebox.showwarning("Sélection", "Veuillez sélectionner un produit à dupliquer")
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        # Find the index in products_data
        index = self.find_product_index_by_values(values)
        if index is None:
            messagebox.showerror("Erreur", "Impossible de trouver le produit sélectionné")
            return
        
        # Create a copy of the product
        original_product = self.products_data[index]
        duplicate_product = ProductData(
            ID_Produit=f"{original_product.ID_Produit}_copy",
            Marque=original_product.Marque,
            Designation_Reference=original_product.Designation_Reference,
            Serial_Number=original_product.Serial_Number,
            Couleur=original_product.Couleur,
            Matricule=original_product.Matricule,
            Magasin=original_product.Magasin,
            Photo=original_product.Photo
        )
        
        # Add to products_data
        self.products_data.append(duplicate_product)
        self.update_tree_display()
        messagebox.showinfo("Succès", "Produit dupliqué avec succès")
    
    def find_product_index_by_values(self, values):
        """Find the index of a product in products_data by tree values"""
        for i, product in enumerate(self.products_data):
            if (product.ID_Produit == values[0] and 
                product.Marque == values[1] and 
                product.Designation_Reference == values[2] and
                product.Serial_Number == values[3]):
                return i
        return None
    
    def open_edit_dialog(self, index):
        """Open a dialog to edit a product"""
        product = self.products_data[index]
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title("Modifier le produit")
        dialog.geometry("400x300")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Entry fields
        fields = [
            ("ID Produit:", product.ID_Produit),
            ("Marque:", product.Marque),
            ("Désignation/Référence:", product.Designation_Reference),
            ("Serial Number:", product.Serial_Number),
            ("Couleur:", product.Couleur),
            ("Matricule:", product.Matricule),
            ("Magasin:", product.Magasin),
            ("Photo:", product.Photo)
        ]
        
        entries = {}
        for i, (label, value) in enumerate(fields):
            ttk.Label(dialog, text=label).grid(row=i, column=0, sticky=tk.W, padx=10, pady=5)
            entry = ttk.Entry(dialog, width=30)
            entry.insert(0, value or "")
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[label] = entry
        
        # Buttons
        button_frame = ttk.Frame(dialog)
        button_frame.grid(row=len(fields), column=0, columnspan=2, pady=20)
        
        def save_changes():
            try:
                # Update the product
                self.products_data[index] = ProductData(
                    ID_Produit=entries["ID Produit:"].get(),
                    Marque=entries["Marque:"].get(),
                    Designation_Reference=entries["Désignation/Référence:"].get(),
                    Serial_Number=entries["Serial Number:"].get(),
                    Couleur=entries["Couleur:"].get(),
                    Matricule=entries["Matricule:"].get(),
                    Magasin=entries["Magasin:"].get(),
                    Photo=entries["Photo:"].get()
                )
                
                self.update_tree_display()
                dialog.destroy()
                messagebox.showinfo("Succès", "Produit modifié avec succès")
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la modification: {str(e)}")
        
        ttk.Button(button_frame, text="Sauvegarder", command=save_changes).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Annuler", command=dialog.destroy).pack(side=tk.LEFT, padx=5)
    
    def on_search_change(self, *args):
        """Handle search text changes"""
        self.update_tree_display()
    
    def on_filter_change(self, event=None):
        """Handle filter changes"""
        self.update_tree_display()
    
    def sort_column(self, col, reverse):
        """Sort tree contents by column"""
        # Get all items and their values
        data = [(self.tree.set(child, col), child) for child in self.tree.get_children('')]
        
        # Sort the data
        try:
            # Try to sort as numbers if possible
            data.sort(key=lambda x: float(x[0]) if x[0].replace('.', '', 1).isdigit() else x[0], reverse=reverse)
        except:
            # Fall back to string sorting
            data.sort(key=lambda x: x[0], reverse=reverse)
        
        # Rearrange items in sorted positions
        for index, (val, child) in enumerate(data):
            self.tree.move(child, '', index)
        
        # Reverse sort next time
        self.tree.heading(col, command=lambda: self.sort_column(col, not reverse))
    
    def on_item_double_click(self, event):
        """Handle double-click on tree item"""
        selection = self.tree.selection()
        if selection:
            item = selection[0]
            values = self.tree.item(item, 'values')
            index = self.find_product_index_by_values(values)
            if index is not None:
                self.open_edit_dialog(index)
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    """Main function to run the application"""
    app = QRScannerApp()
    app.run()

if __name__ == "__main__":
    main()