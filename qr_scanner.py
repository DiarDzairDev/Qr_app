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

# --- Auto-Updater Dependencies ---
import requests
import threading
import time

# --- Configuration for Updater ---
# 1. Define the current version of the running application
CURRENT_VERSION = "1.0.0" 
# 2. Define the URL where the latest version number is stored (e.g., a raw file on GitHub)
#    IMPORTANT: Replace this with the actual URL to a plain text file containing ONLY the latest version number (e.g., "1.0.1")
REMOTE_VERSION_URL = "https://raw.githubusercontent.com/username/repo/main/VERSION.txt" # Placeholder URL

# --- Utility Functions for Updater ---

def version_to_tuple(version_str):
    """Converts a version string (e.g., '1.0.5') to a tuple of integers (1, 0, 5) for comparison."""
    try:
        return tuple(map(int, version_str.split('.')))
    except ValueError:
        print(f"Error parsing version string: {version_str}")
        return (0, 0, 0) # Fallback


@dataclass
class ProductData:
    """Data structure for product information - Entrée type"""
    Reference: str = ""
    Fournisseur: str = ""
    Designation: str = ""
    Num_Chasse: str = ""
    Couleur: str = ""
    Lot: str = ""
    Magasin: str = ""
    Relation: str = ""

@dataclass
class SortieData:
    """Data structure for sortie information - Sortie type"""
    Date: str = ""
    Heure: str = ""
    DESIGNATION: str = "MOTOS"
    N_CHASSIS: str = ""
    ID_CLIENT: str = ""
    NOM_PRENOM: str = ""
    WILAYA: str = ""

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
        self.data_type = "Entrée"  # Can be "Entrée" or "Sortie"

        # Updater variables
        self.remote_version = None
        self.update_dialog = None
        self.update_status_label = None
        self.update_version_label = None
        self.update_button_in_dialog = None

        
        # Scanner state
        self.scanning = False
        self.scan_buffer = ""
        self.scan_timer = None  # Timer for auto-processing scanned data
        
        # Static client list for Sortie type
        self.clients = [
            {"ID_CLIENT": "C001", "NOM_PRENOM": "Ahmed Benali", "WILAYA": "Alger"},
            {"ID_CLIENT": "C002", "NOM_PRENOM": "Fatima Boumediene", "WILAYA": "Oran"},
            {"ID_CLIENT": "C003", "NOM_PRENOM": "Mohamed Kherroubi", "WILAYA": "Constantine"},
            {"ID_CLIENT": "C004", "NOM_PRENOM": "Amina Cherif", "WILAYA": "Blida"},
            {"ID_CLIENT": "C005", "NOM_PRENOM": "Karim Hadj", "WILAYA": "Batna"},
            {"ID_CLIENT": "C006", "NOM_PRENOM": "Leila Mansouri", "WILAYA": "Béjaïa"},
            {"ID_CLIENT": "C007", "NOM_PRENOM": "Omar Zenati", "WILAYA": "Tizi Ouzou"},
            {"ID_CLIENT": "C008", "NOM_PRENOM": "Nadia Bencheikh", "WILAYA": "Sétif"},
            {"ID_CLIENT": "C009", "NOM_PRENOM": "Youssef Brahimi", "WILAYA": "Annaba"},
            {"ID_CLIENT": "C010", "NOM_PRENOM": "Salima Benaissa", "WILAYA": "Mostaganem"}
        ]
        
        # 58 Algerian Wilayas
        self.wilayas = [
            "Adrar", "Chlef", "Laghouat", "Oum El Bouaghi", "Batna", "Béjaïa", "Biskra", "Béchar",
            "Blida", "Bouira", "Tamanrasset", "Tébessa", "Tlemcen", "Tiaret", "Tizi Ouzou", "Alger",
            "Djelfa", "Jijel", "Sétif", "Saïda", "Skikda", "Sidi Bel Abbès", "Annaba", "Guelma",
            "Constantine", "Médéa", "Mostaganem", "M'Sila", "Mascara", "Ouargla", "Oran", "El Bayadh",
            "Illizi", "Bordj Bou Arréridj", "Boumerdès", "El Tarf", "Tindouf", "Tissemsilt", "El Oued",
            "Khenchela", "Souk Ahras", "Tipaza", "Mila", "Aïn Defla", "Naâma", "Aïn Témouchent",
            "Ghardaïa", "Relizane", "Timimoun", "Bordj Badji Mokhtar", "Ouled Djellal", "Béni Abbès",
            "In Salah", "In Guezzam", "Touggourt", "Djanet", "El M'Ghair", "El Meniaa"
        ]
        
        self.setup_ui()
        self.setup_scanner_listener()
        self.setup_updater_menu() # Setup the new Help menu item
    
    # --- New Updater UI Setup ---


    def setup_updater_menu(self):
        """Adds the 'Check for Updates' menu item."""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)

        # Existing menus (File, Edit, etc. would go here)
        
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Aide", menu=help_menu)
        
        # Add the update command
        help_menu.add_command(label="Vérifier les Mises à Jour...", command=self.start_check_thread)
        
        
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
        
        # Data type selection frame
        type_frame = ttk.LabelFrame(main_frame, text="Type de Données", padding="10")
        type_frame.grid(row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Label(type_frame, text="Type:").grid(row=0, column=0, padx=(0, 5))
        self.data_type_var = tk.StringVar(value="Entrée")
        data_type_combo = ttk.Combobox(type_frame, textvariable=self.data_type_var, width=15, state="readonly")
        data_type_combo['values'] = ("Entrée", "Sortie")
        data_type_combo.grid(row=0, column=1, padx=(0, 10))
        data_type_combo.bind('<<ComboboxSelected>>', self.on_data_type_change)
        
        # File operations frame
        file_frame = ttk.LabelFrame(main_frame, text="File Operations", padding="10")
        file_frame.grid(row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
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
        scanner_frame.grid(row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
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
        data_frame.grid(row=4, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10))
        main_frame.rowconfigure(4, weight=1)
        
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
        self.filter_combo = ttk.Combobox(search_frame, textvariable=self.filter_field, width=15, state="readonly")
        self.filter_combo['values'] = ("All Fields", "Num_Chasse", "Fournisseur", "Designation", 
                                      "Reference", "Couleur", "Lot", "Magasin")
        self.filter_combo.grid(row=0, column=3, padx=(0, 10))
        self.filter_combo.bind('<<ComboboxSelected>>', self.on_filter_change)
        
        # CRUD buttons
        crud_frame = ttk.Frame(data_frame)
        crud_frame.grid(row=1, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(crud_frame, text="Edit Selected", 
                  command=self.edit_selected_record).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(crud_frame, text="Delete Selected", 
                  command=self.delete_selected_record).grid(row=0, column=1, padx=5)
        ttk.Button(crud_frame, text="Duplicate Selected", 
                  command=self.duplicate_selected_record).grid(row=0, column=2, padx=5)
        
        # Treeview for data display (Num_Chasse first, then Reference)
        columns = ('Num_Chasse', 'Fournisseur', 'Designation', 'Reference', 
                  'Couleur', 'Lot', 'Magasin', 'Relation')
        
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
        qr_frame.grid(row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        
        ttk.Button(qr_frame, text="Generate QR from Selected Row", 
                  command=self.generate_qr_from_selection).grid(row=0, column=0, padx=(0, 5))
        
        # Minimal skeleton to ensure the app runs
        main_label = ttk.Label(self.root, text="QR Scanner App", font=('Inter', 24, 'bold'))
        # main_label.pack(pady=50)
        version_label = ttk.Label(self.root, text=f"Version Actuelle: {CURRENT_VERSION}", font=('Inter', 10))
        # version_label.pack(pady=10)
        
        
        # Focus scanner input
        self.scanner_entry.focus_set()
    
    def setup_scanner_listener(self):
        """Setup scanner input detection"""
        self.scanner_entry.focus_set()
    
    def on_data_type_change(self, event=None):
        """Handle data type change between Entrée and Sortie"""
        self.data_type = self.data_type_var.get()
        
        # Clear existing data when switching types
        if self.products_data:
            result = messagebox.askyesno(
                "Changement de Type", 
                f"Changer vers le type '{self.data_type}' effacera toutes les données actuelles.\n"
                f"Voulez-vous continuer?",
                icon='warning'
            )
            if result:
                self.products_data = []
                self.update_tree_display()
            else:
                # Revert the selection
                old_type = "Sortie" if self.data_type == "Entrée" else "Entrée"
                self.data_type_var.set(old_type)
                self.data_type = old_type
                return
        
        # Update the UI components based on data type
        self.setup_dynamic_ui()
        self.update_tree_display()
    
    def setup_dynamic_ui(self):
        """Setup UI components based on current data type"""
        if self.data_type == "Entrée":
            # Update treeview columns for Entrée (Num_Chasse first, then Reference)
            columns = ('Num_Chasse', 'Fournisseur', 'Designation', 'Reference', 
                      'Couleur', 'Lot', 'Magasin', 'Relation')
            filter_values = ("All Fields", "Num_Chasse", "Fournisseur", "Designation", 
                           "Reference", "Couleur", "Lot", "Magasin")
        else:  # Sortie
            # Update treeview columns for Sortie
            columns = ('Date', 'Heure', 'DESIGNATION', 'N_CHASSIS', 
                      'ID_CLIENT', 'NOM_PRENOM', 'WILAYA')
            filter_values = ("All Fields", "Date", "Heure", "DESIGNATION", 
                           "N_CHASSIS", "ID_CLIENT", "NOM_PRENOM", "WILAYA")
        
        # Update treeview columns
        self.tree.configure(columns=columns)
        
        # Clear existing column configurations
        for col in self.tree['columns']:
            self.tree.heading(col, text="")
        
        # Configure new column headings
        for col in columns:
            self.tree.heading(col, text=col.replace('_', ' '), 
                            command=lambda c=col: self.sort_column(c, False))
            self.tree.column(col, width=120, minwidth=80)
        
        # Update filter dropdown values
        if hasattr(self, 'filter_combo'):
            self.filter_combo['values'] = filter_values
            self.filter_field.set("All Fields")
    
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
            
            # For Sortie type, open client selection dialog
            if self.data_type == "Sortie":
                selected_client = self.open_client_selection_dialog()
                if selected_client:
                    # Update the product_data with selected client info
                    product_data.ID_CLIENT = selected_client["ID_CLIENT"]
                    product_data.NOM_PRENOM = selected_client["NOM_PRENOM"]
                    product_data.WILAYA = selected_client["WILAYA"]
                else:
                    # User cancelled client selection
                    self.status_label.config(text="Scan cancelled - no client selected", foreground="orange")
                    self.clear_scanner_input()
                    return
            
            # Add to data list
            self.products_data.append(product_data)
            
            # Update display
            self.update_tree_display()
            
            # Clear scanner input
            self.clear_scanner_input()

            # self.status_label.config(text=f"Successfully added product: {product_data.DESIGNATION}", 
            #                        foreground="green")
            
        except Exception as e:
            messagebox.showerror("Error", f"Failed to process QR data: {str(e)}")
            self.status_label.config(text="Error processing scan", foreground="red")
        
        finally:
            self.scanning = False
    
    def parse_qr_data(self, qr_data: str):
        """Parse QR code data with backward compatibility"""
        if self.data_type == "Entrée":
            product = ProductData()
        else:  # Sortie
            from datetime import datetime
            now = datetime.now()
            product = SortieData()
            product.Date = now.strftime("%d/%m/%Y")
            product.Heure = now.strftime("%H:%M")
            product.DESIGNATION = "MOTOS"
        
        if self.data_type == "Entrée":
            # Check if data contains line breaks (structured format)
            if '\n' in qr_data or '\r' in qr_data:
                # Split by lines and clean up
                lines = qr_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
                lines = [line.strip() for line in lines if line.strip()]
                
                # Check if first line starts with asterisk (legacy format)
                if lines and lines[0].startswith('*') and lines[0].endswith('*'):
                    # Legacy format with asterisks
                    legacy_field_mapping = [
                        'Reference',     # *VMSDZ06CUKI191698* -> extract content between asterisks
                        'Designation',   # MOTOCYCLE CUKI -I-
                        'Fournisseur',   # CUKI
                        'Couleur',       # bleu nuit/ blanc
                        'Magasin',       # Unité Oued-Ghir
                        'Num_Chasse'     # CUKI I 06/2025
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
                    
                    # Lot is empty in legacy format (as specified)
                    product.Lot = ""
                    
                else:
                    # New structured format - split by lines
                    field_mapping = [
                        'Reference',
                        'Fournisseur',
                        'Designation',
                        'Num_Chasse',
                        'Couleur',
                        'Lot',
                        'Magasin',
                        'Relation'
                    ]
                    
                    for i, line in enumerate(lines):
                        if i < len(field_mapping):
                            setattr(product, field_mapping[i], line)
            else:
                # Single line input - check for concatenated format
                # Handle "*VMSDZ06CUKI191858*MOTOCYCLE CUKI -II-CUKI" format
                if '*' in qr_data:
                    # Extract everything between first pair of asterisks as Reference
                    import re
                    # Find first asterisk pair
                    asterisk_match = re.search(r'\*([^*]+)\*', qr_data)
                    if asterisk_match:
                        product.Reference = asterisk_match.group(1)
                        # Extract remaining text after the closing asterisk
                        remaining_text = qr_data[asterisk_match.end():].strip()
                        if remaining_text:
                            # Split remaining text by common separators or patterns
                            # Look for pattern: "MOTOCYCLE CUKI -II-CUKI" or similar
                            parts = []
                            if ' -' in remaining_text and '-' in remaining_text:
                                # Pattern like "MOTOCYCLE CUKI -II-CUKI"
                                # Split on the dash pattern
                                dash_parts = remaining_text.split('-')
                                if len(dash_parts) >= 2:
                                    # First part before dash is designation
                                    product.Designation = dash_parts[0].strip()
                                    # Last part after last dash could be fournisseur
                                    if len(dash_parts) > 2:
                                        product.Fournisseur = dash_parts[-1].strip()
                            else:
                                # Simple text - use as designation
                                product.Designation = remaining_text
                else:
                    # No asterisks - treat as simple reference
                    product.Reference = qr_data.strip()
        
        else:  # Sortie type
            # For Sortie, extract chassis number and designation from QR data
            if '\n' in qr_data or '\r' in qr_data:
                lines = qr_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
                lines = [line.strip() for line in lines if line.strip()]
                
                # Check if first line starts with asterisk (legacy format)
                if lines and lines[0].startswith('*') and lines[0].endswith('*'):
                    # Extract chassis number from between asterisks
                    product.N_CHASSIS = lines[0][1:-1]  # Remove asterisks
                    
                    # Extract designation from second line if available
                    if len(lines) > 1:
                        product.DESIGNATION = lines[1]
                else:
                    # Try to find chassis number from various positions
                    for line in lines:
                        if line and not line.startswith('*'):
                            # Use the first meaningful line as chassis number
                            product.N_CHASSIS = line
                            break
            else:
                # Single line - handle concatenated format for Sortie
                # Handle "*VMSDZ06CUKI191858*MOTOCYCLE CUKI -II-CUKI" format
                if '*' in qr_data:
                    import re
                    # Find first asterisk pair for chassis number
                    asterisk_match = re.search(r'\*([^*]+)\*', qr_data)
                    if asterisk_match:
                        product.N_CHASSIS = asterisk_match.group(1)
                        # Extract remaining text after the closing asterisk for designation
                        remaining_text = qr_data[asterisk_match.end():].strip()
                        if remaining_text:
                            product.DESIGNATION = remaining_text
                else:
                    # No asterisks - treat entire string as chassis number
                    product.N_CHASSIS = qr_data.strip()
        
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
            print(f"Reference: {product.Reference}")
            print(f"Designation: {product.Designation}")
            print(f"Fournisseur: {product.Fournisseur}")
            print(f"Couleur: {product.Couleur}")
            print(f"Lot: {product.Lot}")
            print(f"Magasin: {product.Magasin}")
            print(f"Num_Chasse: {product.Num_Chasse}")
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
                'Reference',
                'Fournisseur',
                'Designation',
                'Num_Chasse',
                'Couleur',
                'Lot',
                'Magasin',
                'Relation'
            ]
            
            for i, line in enumerate(lines):
                if i < len(field_mapping):
                    setattr(product, field_mapping[i], line)
        else:
            # No line breaks found - treat entire input as Reference only
            product.Reference = scanner_data.strip()
        
        return product
    
    def generate_qr_data(self, product_data) -> str:
        """Generate QR code data in structured format based on data type"""
        if isinstance(product_data, ProductData):
            # For Entrée type - generate legacy format with asterisks
            qr_lines = [
                f"*{product_data.Reference}*",
                product_data.Designation,
                product_data.Fournisseur,
                product_data.Couleur,
                product_data.Magasin,
                product_data.Num_Chasse
            ]
        elif isinstance(product_data, SortieData):
            # For Sortie type - generate format with chassis in asterisks
            qr_lines = [
                f"*{product_data.N_CHASSIS}*",
                product_data.DESIGNATION,
                product_data.ID_CLIENT,
                product_data.NOM_PRENOM,
                product_data.WILAYA,
                f"{product_data.Date} {product_data.Heure}"
            ]
        else:
            # Fallback - assume it's Entrée format
            qr_lines = [
                getattr(product_data, 'Reference', ''),
                getattr(product_data, 'Fournisseur', ''),
                getattr(product_data, 'Designation', ''),
                getattr(product_data, 'Num_Chasse', ''),
                getattr(product_data, 'Couleur', ''),
                getattr(product_data, 'Lot', ''),
                getattr(product_data, 'Magasin', ''),
                getattr(product_data, 'Relation', '')
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
        dialog.title(f"Manual {self.data_type} Input")
        dialog.geometry("500x600")
        dialog.resizable(True, True)
        
        # Main frame with scrollbar
        main_frame = ttk.Frame(dialog)
        main_frame.pack(fill='both', expand=True, padx=10, pady=10)
        
        # Title
        ttk.Label(main_frame, text=f"Enter {self.data_type} Information:", 
                 font=('Arial', 12, 'bold')).pack(pady=(0, 20))
        
        # Create entry fields based on data type
        fields = {}
        if self.data_type == "Entrée":
            field_labels = {
                'Reference': 'Référence:',
                'Fournisseur': 'Fournisseur:',
                'Designation': 'Désignation:',
                'Num_Chasse': 'Numéro de Châsse:',
                'Couleur': 'Couleur:',
                'Lot': 'Lot:',
                'Magasin': 'Magasin:',
                'Relation': 'Relation:'
            }
        else:  # Sortie
            field_labels = {
                'Date': 'Date:',
                'Heure': 'Heure:',
                'DESIGNATION': 'DESIGNATION:',
                'N_CHASSIS': 'N° CHASSIS:',
                'ID_CLIENT': 'ID-CLIENT:',
                'NOM_PRENOM': 'NOM & PRENOM:',
                'WILAYA': 'WILAYA:'
            }
        
        # Create form fields
        for field, label in field_labels.items():
            frame = ttk.Frame(main_frame)
            frame.pack(fill='x', pady=5)
            
            ttk.Label(frame, text=label, width=20).pack(side='left')
            entry = ttk.Entry(frame, width=40)
            entry.pack(side='left', fill='x', expand=True, padx=(10, 0))
            fields[field] = entry
        
        # Pre-fill with sample data based on type
        if self.data_type == "Entrée":
            sample_data = {
                'Reference': 'VMSDZ06CUKI191858',
                'Fournisseur': 'CUKI',
                'Designation': 'MOTOCYCLE CUKI -I-',
                'Num_Chasse': 'CUKI I 06/2025',
                'Couleur': 'bleu nuit/ blanc',
                'Lot': '',
                'Magasin': 'Unité Oued-Ghir',
                'Relation': ''
            }
        else:  # Sortie
            from datetime import datetime
            now = datetime.now()
            sample_data = {
                'Date': now.strftime("%d/%m/%Y"),
                'Heure': now.strftime("%H:%M"),
                'DESIGNATION': 'MOTOS',
                'N_CHASSIS': '',
                'ID_CLIENT': '',
                'NOM_PRENOM': '',
                'WILAYA': ''
            }
        
        for field, value in sample_data.items():
            if field in fields:  # Only insert if field exists
                fields[field].insert(0, value)
        
        # Buttons frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        def add_product():
            # Crée l'objet selon le type
            if self.data_type == "Entrée":
                product = ProductData()
                for field, entry in fields.items():
                    setattr(product, field, entry.get().strip())
            else:  # Sortie
                # For Sortie type, first get basic info from form
                product = SortieData()
                for field, entry in fields.items():
                    if field in ['Date', 'Heure', 'DESIGNATION', 'N_CHASSIS']:
                        setattr(product, field, entry.get().strip())
                
                # Open client selection dialog for client info
                selected_client = self.open_client_selection_dialog()
                if selected_client:
                    product.ID_CLIENT = selected_client["ID_CLIENT"]
                    product.NOM_PRENOM = selected_client["NOM_PRENOM"]
                    product.WILAYA = selected_client["WILAYA"]
                else:
                    # User cancelled client selection
                    messagebox.showinfo("Annulé", "Ajout de sortie annulé - aucun client sélectionné")
                    return
            
            # Ajoute à la liste
            self.products_data.append(product)
            self.update_tree_display()
            # Message de succès
            if self.data_type == "Entrée":
                messagebox.showinfo("Succès", f"Produit ajouté: {product.Reference}")
            else:
                messagebox.showinfo("Succès", f"Sortie ajoutée: {product.N_CHASSIS}")
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
        
        # Filter and search products based on data type
        filtered_products = []
        for i, product in enumerate(self.products_data):
            # Get product values for search
            if self.data_type == "Entrée":
                product_values = [
                    str(getattr(product, 'Reference', "") or ""),
                    str(getattr(product, 'Fournisseur', "") or ""),
                    str(getattr(product, 'Designation', "") or ""),
                    str(getattr(product, 'Num_Chasse', "") or ""),
                    str(getattr(product, 'Couleur', "") or ""),
                    str(getattr(product, 'Lot', "") or ""),
                    str(getattr(product, 'Magasin', "") or ""),
                    str(getattr(product, 'Relation', "") or "")
                ]
                # Map filter field to product attribute for Entrée
                field_mapping = {
                    "Reference": str(getattr(product, 'Reference', "") or ""),
                    "Fournisseur": str(getattr(product, 'Fournisseur', "") or ""),
                    "Designation": str(getattr(product, 'Designation', "") or ""),
                    "Num_Chasse": str(getattr(product, 'Num_Chasse', "") or ""),
                    "Couleur": str(getattr(product, 'Couleur', "") or ""),
                    "Lot": str(getattr(product, 'Lot', "") or ""),
                    "Magasin": str(getattr(product, 'Magasin', "") or ""),
                    "Relation": str(getattr(product, 'Relation', "") or "")
                }
            else:  # Sortie
                product_values = [
                    str(getattr(product, 'Date', "") or ""),
                    str(getattr(product, 'Heure', "") or ""),
                    str(getattr(product, 'DESIGNATION', "") or ""),
                    str(getattr(product, 'N_CHASSIS', "") or ""),
                    str(getattr(product, 'ID_CLIENT', "") or ""),
                    str(getattr(product, 'NOM_PRENOM', "") or ""),
                    str(getattr(product, 'WILAYA', "") or "")
                ]
                # Map filter field to product attribute for Sortie
                field_mapping = {
                    "Date": str(getattr(product, 'Date', "") or ""),
                    "Heure": str(getattr(product, 'Heure', "") or ""),
                    "DESIGNATION": str(getattr(product, 'DESIGNATION', "") or ""),
                    "N_CHASSIS": str(getattr(product, 'N_CHASSIS', "") or ""),
                    "ID_CLIENT": str(getattr(product, 'ID_CLIENT', "") or ""),
                    "NOM_PRENOM": str(getattr(product, 'NOM_PRENOM', "") or ""),
                    "WILAYA": str(getattr(product, 'WILAYA', "") or "")
                }
            
            # Apply search filter
            include_product = True
            if search_text:
                if filter_field == "All Fields":
                    # Search in all fields
                    include_product = any(search_text in value.lower() for value in product_values)
                else:
                    # Search in specific field
                    field_value = field_mapping.get(filter_field, "")
                    include_product = search_text in field_value.lower()
            
            if include_product:
                filtered_products.append((i, product))
        
        # Add filtered products to tree (use filtered_products instead of all products)
        for original_index, product in filtered_products:
            if self.data_type == "Entrée":
                values = (
                    str(getattr(product, 'Num_Chasse', "") or ""),
                    str(getattr(product, 'Fournisseur', "") or ""),
                    str(getattr(product, 'Designation', "") or ""),
                    str(getattr(product, 'Reference', "") or ""),
                    str(getattr(product, 'Couleur', "") or ""),
                    str(getattr(product, 'Lot', "") or ""),
                    str(getattr(product, 'Magasin', "") or ""),
                    str(getattr(product, 'Relation', "") or "")
                )
            else:  # Sortie
                values = (
                    str(getattr(product, 'Date', "") or ""),
                    str(getattr(product, 'Heure', "") or ""),
                    str(getattr(product, 'DESIGNATION', "") or ""),
                    str(getattr(product, 'N_CHASSIS', "") or ""),
                    str(getattr(product, 'ID_CLIENT', "") or ""),
                    str(getattr(product, 'NOM_PRENOM', "") or ""),
                    str(getattr(product, 'WILAYA', "") or "")
                )
            # Use original index for proper identification
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
    
    def show_qr_code(self, product_data):
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
        
        # Set title based on data type
        if isinstance(product_data, ProductData):
            window_title = f"QR Code - {product_data.Reference}"
        elif isinstance(product_data, SortieData):
            window_title = f"QR Code - {product_data.N_CHASSIS}"
        else:
            window_title = "QR Code"
            
        qr_window.title(window_title)
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
        info_frame = ttk.LabelFrame(scrollable_frame, text="Informations", padding="10")
        info_frame.pack(fill='x', padx=20, pady=(0, 20))
        
        # Display product information based on data type
        if isinstance(product_data, ProductData):
            # For Entrée type
            product_info = [
                ("Référence:", product_data.Reference),
                ("Fournisseur:", product_data.Fournisseur),
                ("Désignation:", product_data.Designation),
                ("Numéro de Châsse:", product_data.Num_Chasse),
                ("Couleur:", product_data.Couleur),
                ("Lot:", product_data.Lot),
                ("Magasin:", product_data.Magasin),
                ("Relation:", product_data.Relation)
            ]
        elif isinstance(product_data, SortieData):
            # For Sortie type
            product_info = [
                ("Date:", product_data.Date),
                ("Heure:", product_data.Heure),
                ("Désignation:", product_data.DESIGNATION),
                ("N° Châssis:", product_data.N_CHASSIS),
                ("ID Client:", product_data.ID_CLIENT),
                ("Nom & Prénom:", product_data.NOM_PRENOM),
                ("Wilaya:", product_data.WILAYA)
            ]
        else:
            # Fallback for unknown type
            product_info = []
        
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
                    hDC.StartDoc(f"QR Code - {product_data.Reference}")
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
                f"Référence: {product_data.Reference}",
                f"Fournisseur: {product_data.Fournisseur}",
                f"Désignation: {product_data.Designation}",
                f"Numéro de Châsse: {product_data.Num_Chasse}",
                f"Couleur: {product_data.Couleur}",
                f"Lot: {product_data.Lot}",
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
                self.products_data = []
                if self.data_type == "Entrée":
                    column_mapping = {
                        'Reference': 'Reference',
                        'ID_Produit': 'Reference',
                        'Fournisseur': 'Fournisseur',
                        'Marque': 'Fournisseur',
                        'Designation/Reference': 'Designation',
                        'Designation_Reference': 'Designation',
                        'Designation': 'Designation',
                        'Num_Chasse': 'Num_Chasse',
                        'Serial Number': 'Num_Chasse',
                        'Serial_Number': 'Num_Chasse',
                        'Couleur': 'Couleur',
                        'Matricule': 'Lot',
                        'Lot': 'Lot',
                        'Magasin': 'Magasin',
                        'Photo': 'Relation',
                        'Relation': 'Relation'
                    }
                    for _, row in df.iterrows():
                        product = ProductData()
                        for excel_col, product_field in column_mapping.items():
                            if excel_col in df.columns:
                                value = row[excel_col]
                                setattr(product, product_field, str(value) if pd.notna(value) else "")
                        self.products_data.append(product)
                else:
                    # Sortie: logique intelligente pour trouver les colonnes et données
                    sortie_columns = ['Date', 'Heure', 'DESIGNATION', 'N_CHASSIS', 'ID_CLIENT', 'NOM_PRENOM', 'WILAYA']
                    
                    # Première tentative: vérifier si les colonnes sont directement dans le DataFrame
                    found = all(col in df.columns for col in sortie_columns)
                    
                    if found:
                        # Colonnes trouvées directement - charger les données en ignorant les titres
                        for _, row in df.iterrows():
                            # Ignorer les lignes qui contiennent le titre "SORTIE LIVRAISON"
                            row_values = [str(val) if pd.notna(val) else "" for val in row.values]
                            if any("SORTIE" in val.upper() for val in row_values if isinstance(val, str)):
                                continue  # Ignorer les lignes de titre
                            
                            # Vérifier que la ligne contient des données valides
                            if all(val == "" or val == "nan" for val in row_values):
                                continue  # Ignorer les lignes vides
                            
                            product = SortieData()
                            for col in sortie_columns:
                                setattr(product, col, str(row[col]) if pd.notna(row[col]) else "")
                            self.products_data.append(product)
                    else:
                        # Colonnes non trouvées - chercher dans le contenu du fichier
                        # Relire le fichier sans en-tête pour analyser le contenu
                        try:
                            # Lire le fichier ligne par ligne pour trouver l'en-tête des colonnes
                            raw_df = pd.read_excel(filename, header=None)
                            header_row_index = None
                            
                            # Chercher la ligne qui contient toutes nos colonnes attendues
                            for index, row in raw_df.iterrows():
                                row_values = [str(val).strip() if pd.notna(val) else "" for val in row.values]
                                # Vérifier si cette ligne contient nos colonnes (au moins 4 sur 7)
                                matches = sum(1 for col in sortie_columns if col in row_values)
                                if matches >= 4:  # Au moins 4 colonnes trouvées
                                    header_row_index = index
                                    break
                            
                            if header_row_index is not None:
                                # Relire le fichier avec l'en-tête trouvé
                                df_with_header = pd.read_excel(filename, header=header_row_index)
                                
                                # Vérifier que nous avons maintenant les bonnes colonnes
                                found_columns = [col for col in sortie_columns if col in df_with_header.columns]
                                
                                if len(found_columns) >= 4:  # Au moins 4 colonnes trouvées
                                    for _, row in df_with_header.iterrows():
                                        # Ignorer les lignes vides ou qui contiennent le titre
                                        row_values = [str(val) if pd.notna(val) else "" for val in row.values]
                                        if any("SORTIE" in val.upper() for val in row_values if isinstance(val, str)):
                                            continue  # Ignorer les lignes de titre
                                        
                                        # Vérifier que la ligne contient des données valides
                                        if all(val == "" or val == "nan" for val in row_values):
                                            continue  # Ignorer les lignes vides
                                        
                                        product = SortieData()
                                        for col in sortie_columns:
                                            if col in df_with_header.columns:
                                                value = row[col]
                                                setattr(product, col, str(value) if pd.notna(value) else "")
                                        self.products_data.append(product)
                                else:
                                    raise ValueError("Colonnes Sortie non trouvées dans le fichier")
                            else:
                                raise ValueError("En-tête des colonnes Sortie non trouvé")
                                
                        except Exception as search_error:
                            # Si on ne trouve pas les colonnes, créer une nouvelle table
                            print(f"Erreur lors de la recherche des colonnes: {search_error}")
                            df = pd.DataFrame(columns=sortie_columns)
                            # Ajouter le titre au centre
                            title_row = ["" for _ in sortie_columns]
                            title_row[int(len(sortie_columns)/2)] = "SORTIE LIVRAISON JOURNALIERE"
                            
                            # Créer un DataFrame avec le titre et les en-têtes
                            title_df = pd.DataFrame([title_row], columns=sortie_columns)
                            empty_row = pd.DataFrame([["" for _ in sortie_columns]], columns=sortie_columns)
                            final_df = pd.concat([title_df, empty_row, df], ignore_index=True)
                            
                            final_df.to_excel(filename, index=False)
                            messagebox.showinfo("Info", "Table de sortie créée dans le fichier Excel.")
                            return
                
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
            # Convert data objects to dictionary list
            data_list = [asdict(product) for product in self.products_data]
            
            # Create DataFrame
            df = pd.DataFrame(data_list)
            
            # Set appropriate column headers based on data type
            if self.data_type == "Entrée":
                # Rename columns for Entrée type (8 columns)
                df.columns = [
                    'Reference', 'Fournisseur', 'Designation', 
                    'Num_Chasse', 'Couleur', 'Lot', 
                    'Magasin', 'Relation'
                ]
                # Save normally for Entrée
                df.to_excel(self.excel_file, index=False)
            else:  # Sortie
                # For Sortie type (7 columns)
                df.columns = [
                    'Date', 'Heure', 'DESIGNATION', 
                    'N_CHASSIS', 'ID_CLIENT', 'NOM_PRENOM', 
                    'WILAYA'
                ]
                
                # Use ExcelWriter to write title separately from table
                with pd.ExcelWriter(self.excel_file, engine='openpyxl') as writer:
                    # Write the DataFrame starting from row 3 (leaving rows 1-2 for title)
                    df.to_excel(writer, index=False, startrow=2, header=True, sheet_name='Sheet1')
                    
                    # Get the workbook and worksheet to add formatting
                    workbook = writer.book
                    worksheet = workbook['Sheet1']
                    
                    # Make title bold and centered
                    from openpyxl.styles import Font, Alignment
                    title_font = Font(bold=True, size=14)
                    center_alignment = Alignment(horizontal='center', vertical='center')
                    
                    # Merge cells for title (A1:G1 to span all columns)
                    worksheet.merge_cells('A1:G1')
                    worksheet['A1'] = "SORTIE LIVRAISON JOURNALIERE"
                    worksheet['A1'].font = title_font
                    worksheet['A1'].alignment = center_alignment
                    
                    # Apply formatting to headers
                    header_font = Font(bold=True)
                    for col_num, column_title in enumerate(df.columns, 1):
                        cell = worksheet.cell(row=3, column=col_num)
                        cell.font = header_font
            
            # For Entrée type, save normally
            if self.data_type == "Entrée":
                pass  # Already saved above
            
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
            messagebox.showwarning("Sélection", f"Veuillez sélectionner un {self.data_type.lower()} à supprimer")
            return
        
        item = selection[0]
        values = self.tree.item(item, 'values')
        
        # Confirm deletion with appropriate fields based on data type
        if self.data_type == "Entrée":
            confirm_message = (
                f"Êtes-vous sûr de vouloir supprimer ce produit?\n\n"
                f"Numéro de Châsse: {values[0]}\n"
                f"Fournisseur: {values[1]}\n"
                f"Désignation: {values[2]}\n"
                f"Référence: {values[3]}"
            )
        else:  # Sortie
            confirm_message = (
                f"Êtes-vous sûr de vouloir supprimer cette sortie?\n\n"
                f"Date: {values[0]}\n"
                f"Heure: {values[1]}\n"
                f"N° Châssis: {values[3]}"
            )
        
        result = messagebox.askyesno("Confirmation", confirm_message, icon='warning')
        
        if result:
            # Find the index in products_data
            index = self.find_product_index_by_values(values)
            if index is not None:
                del self.products_data[index]
                self.update_tree_display()
                messagebox.showinfo("Succès", f"{self.data_type} supprimé avec succès")
            else:
                messagebox.showerror("Erreur", f"Impossible de trouver l'{self.data_type.lower()} à supprimer")
    
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
        
        # Create a copy of the product based on data type
        original_product = self.products_data[index]
        
        if self.data_type == "Entrée" and isinstance(original_product, ProductData):
            duplicate_product = ProductData(
                Reference=f"{original_product.Reference}_copy",
                Fournisseur=original_product.Fournisseur,
                Designation=original_product.Designation,
                Num_Chasse=original_product.Num_Chasse,
                Couleur=original_product.Couleur,
                Lot=original_product.Lot,
                Magasin=original_product.Magasin,
                Relation=original_product.Relation
            )
        elif self.data_type == "Sortie" and isinstance(original_product, SortieData):
            duplicate_product = SortieData(
                Date=original_product.Date,
                Heure=original_product.Heure,
                DESIGNATION=original_product.DESIGNATION,
                N_CHASSIS=f"{original_product.N_CHASSIS}_copy",
                ID_CLIENT=original_product.ID_CLIENT,
                NOM_PRENOM=original_product.NOM_PRENOM,
                WILAYA=original_product.WILAYA
            )
        else:
            messagebox.showerror("Erreur", "Type de données incompatible")
            return
        
        # Add to products_data
        self.products_data.append(duplicate_product)
        self.update_tree_display()
        messagebox.showinfo("Succès", "Enregistrement dupliqué avec succès")
    
    def find_product_index_by_values(self, values):
        """Find the index of a product in products_data by tree values"""
        for i, product in enumerate(self.products_data):
            if self.data_type == "Entrée":
                if (getattr(product, 'Num_Chasse', '') == values[0] and 
                    getattr(product, 'Fournisseur', '') == values[1] and 
                    getattr(product, 'Designation', '') == values[2] and
                    getattr(product, 'Reference', '') == values[3]):
                    return i
            else:  # Sortie
                if (getattr(product, 'Date', '') == values[0] and 
                    getattr(product, 'Heure', '') == values[1] and 
                    getattr(product, 'DESIGNATION', '') == values[2] and
                    getattr(product, 'N_CHASSIS', '') == values[3]):
                    return i
        return None
    
    def open_edit_dialog(self, index):
        """Open a dialog to edit a product"""
        product = self.products_data[index]
        
        # Create dialog window
        dialog = tk.Toplevel(self.root)
        dialog.title(f"Modifier {self.data_type}")
        dialog.geometry("400x350")
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Create entry fields based on data type
        if self.data_type == "Entrée":
            fields = [
                ("Référence:", getattr(product, 'Reference', '')),
                ("Fournisseur:", getattr(product, 'Fournisseur', '')),
                ("Désignation:", getattr(product, 'Designation', '')),
                ("Numéro de Châsse:", getattr(product, 'Num_Chasse', '')),
                ("Couleur:", getattr(product, 'Couleur', '')),
                ("Lot:", getattr(product, 'Lot', '')),
                ("Magasin:", getattr(product, 'Magasin', '')),
                ("Relation:", getattr(product, 'Relation', ''))
            ]
        else:  # Sortie
            fields = [
                ("Date:", getattr(product, 'Date', '')),
                ("Heure:", getattr(product, 'Heure', '')),
                ("DESIGNATION:", getattr(product, 'DESIGNATION', '')),
                ("N° CHASSIS:", getattr(product, 'N_CHASSIS', '')),
                ("ID-CLIENT:", getattr(product, 'ID_CLIENT', '')),
                ("NOM & PRENOM:", getattr(product, 'NOM_PRENOM', '')),
                ("WILAYA:", getattr(product, 'WILAYA', ''))
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
                # Update the product based on data type
                if self.data_type == "Entrée":
                    self.products_data[index] = ProductData(
                        Reference=entries["Référence:"].get(),
                        Fournisseur=entries["Fournisseur:"].get(),
                        Designation=entries["Désignation:"].get(),
                        Num_Chasse=entries["Numéro de Châsse:"].get(),
                        Couleur=entries["Couleur:"].get(),
                        Lot=entries["Lot:"].get(),
                        Magasin=entries["Magasin:"].get(),
                        Relation=entries["Relation:"].get()
                    )
                else:  # Sortie
                    self.products_data[index] = SortieData(
                        Date=entries["Date:"].get(),
                        Heure=entries["Heure:"].get(),
                        DESIGNATION=entries["DESIGNATION:"].get(),
                        N_CHASSIS=entries["N° CHASSIS:"].get(),
                        ID_CLIENT=entries["ID-CLIENT:"].get(),
                        NOM_PRENOM=entries["NOM & PRENOM:"].get(),
                        WILAYA=entries["WILAYA:"].get()
                    )
                
                self.update_tree_display()
                dialog.destroy()
                messagebox.showinfo("Succès", f"{self.data_type} modifié avec succès")
                
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
    
    def open_client_selection_dialog(self):
        """Open dialog to select a client for Sortie type"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Sélection du Client")
        dialog.geometry("700x600")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        selected_client = None
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Title
        ttk.Label(main_frame, text="Sélection du Client", 
                 font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        # Search frame
        search_frame = ttk.LabelFrame(main_frame, text="Recherche Client", padding="10")
        search_frame.pack(fill='x', pady=(0, 10))
        
        # Search by ID
        ttk.Label(search_frame, text="Rechercher par ID:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        search_id_var = tk.StringVar()
        search_id_entry = ttk.Entry(search_frame, textvariable=search_id_var, width=20)
        search_id_entry.grid(row=0, column=1, padx=(0, 20), sticky=(tk.W, tk.E))
        
        # Search by Name
        ttk.Label(search_frame, text="Rechercher par Nom:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        search_name_var = tk.StringVar()
        search_name_entry = ttk.Entry(search_frame, textvariable=search_name_var, width=30)
        search_name_entry.grid(row=0, column=3, padx=(0, 10), sticky=(tk.W, tk.E))
        
        search_frame.columnconfigure(1, weight=1)
        search_frame.columnconfigure(3, weight=2)
        
        # Client list frame
        list_frame = ttk.LabelFrame(main_frame, text="Liste des Clients", padding="10")
        list_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Treeview for client list
        client_columns = ('ID', 'Nom & Prénom', 'Wilaya')
        client_tree = ttk.Treeview(list_frame, columns=client_columns, show='headings', height=12)
        
        # Configure column headings
        for col in client_columns:
            client_tree.heading(col, text=col)
            client_tree.column(col, width=150, minwidth=100)
        
        # Scrollbars for client tree
        v_scrollbar = ttk.Scrollbar(list_frame, orient="vertical", command=client_tree.yview)
        h_scrollbar = ttk.Scrollbar(list_frame, orient="horizontal", command=client_tree.xview)
        client_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout for client tree
        client_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        list_frame.columnconfigure(0, weight=1)
        list_frame.rowconfigure(0, weight=1)
        
        # Manual entry frame
        manual_frame = ttk.LabelFrame(main_frame, text="Nouveau Client (Entrée Manuelle)", padding="10")
        manual_frame.pack(fill='x', pady=(0, 10))
        
        # Manual entry fields
        ttk.Label(manual_frame, text="ID Client:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        manual_id_var = tk.StringVar()
        manual_id_entry = ttk.Entry(manual_frame, textvariable=manual_id_var, width=15)
        manual_id_entry.grid(row=0, column=1, padx=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Label(manual_frame, text="Nom & Prénom:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        manual_name_var = tk.StringVar()
        manual_name_entry = ttk.Entry(manual_frame, textvariable=manual_name_var, width=25)
        manual_name_entry.grid(row=0, column=3, padx=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Label(manual_frame, text="Wilaya:").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        manual_wilaya_var = tk.StringVar()
        manual_wilaya_combo = ttk.Combobox(manual_frame, textvariable=manual_wilaya_var, 
                                          values=self.wilayas, state="readonly", width=15)
        manual_wilaya_combo.grid(row=0, column=5, padx=(0, 10), sticky=(tk.W, tk.E))
        
        manual_frame.columnconfigure(1, weight=1)
        manual_frame.columnconfigure(3, weight=2)
        manual_frame.columnconfigure(5, weight=1)
        
        # Function to update client list based on search
        def update_client_list():
            # Clear existing items
            for item in client_tree.get_children():
                client_tree.delete(item)
            
            search_id = search_id_var.get().lower()
            search_name = search_name_var.get().lower()
            
            # Filter clients based on search criteria
            for client in self.clients:
                include_client = True
                
                if search_id and search_id not in client["ID_CLIENT"].lower():
                    include_client = False
                
                if search_name and search_name not in client["NOM_PRENOM"].lower():
                    include_client = False
                
                if include_client:
                    client_tree.insert('', 'end', values=(
                        client["ID_CLIENT"],
                        client["NOM_PRENOM"],
                        client["WILAYA"]
                    ), tags=(str(self.clients.index(client)),))
        
        # Bind search events
        search_id_var.trace('w', lambda *args: update_client_list())
        search_name_var.trace('w', lambda *args: update_client_list())
        
        # Initial population of client list
        update_client_list()
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        def select_from_list():
            nonlocal selected_client
            selection = client_tree.selection()
            if not selection:
                messagebox.showwarning("Sélection", "Veuillez sélectionner un client dans la liste")
                return
            
            item = selection[0]
            tags = client_tree.item(item, 'tags')
            if tags:
                client_index = int(tags[0])
                selected_client = self.clients[client_index]
                dialog.destroy()
        
        def use_manual_entry():
            nonlocal selected_client
            if not manual_id_var.get() or not manual_name_var.get() or not manual_wilaya_var.get():
                messagebox.showwarning("Entrée Manuelle", 
                                     "Veuillez remplir tous les champs pour le nouveau client")
                return
            
            # Create new client entry
            new_client = {
                "ID_CLIENT": manual_id_var.get().strip(),
                "NOM_PRENOM": manual_name_var.get().strip(),
                "WILAYA": manual_wilaya_var.get()
            }
            
            # Check if ID already exists
            existing_ids = [client["ID_CLIENT"] for client in self.clients]
            if new_client["ID_CLIENT"] in existing_ids:
                messagebox.showwarning("ID Existant", 
                                     f"L'ID client '{new_client['ID_CLIENT']}' existe déjà")
                return
            
            # Add to clients list and use it
            self.clients.append(new_client)
            selected_client = new_client
            dialog.destroy()
        
        def cancel_selection():
            nonlocal selected_client
            selected_client = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="Sélectionner Client", 
                  command=select_from_list).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Utiliser Nouveau Client", 
                  command=use_manual_entry).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Annuler", 
                  command=cancel_selection).pack(side=tk.LEFT, padx=5)
        
        # Double-click to select client
        def on_double_click(event):
            select_from_list()
        
        client_tree.bind('<Double-1>', on_double_click)
        
        # Focus on search field
        search_id_entry.focus_set()
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return selected_client
    def start_check_thread(self):
        """Initializes the modal dialog and starts the version check in a separate thread."""
        
        # 1. Create Modal Dialog
        self.update_dialog = tk.Toplevel(self.root)
        self.update_dialog.title("Vérification des mises à jour")
        self.update_dialog.geometry("350x180")
        self.update_dialog.transient(self.root) # Make it stay above the main window
        self.update_dialog.grab_set() # Make it modal (block interaction with main window)
        self.update_dialog.protocol("WM_DELETE_WINDOW", lambda: None) # Disable closing by X button during check

        dialog_frame = ttk.Frame(self.update_dialog, padding=20)
        dialog_frame.pack(expand=True, fill='both')

        # Status Label
        self.update_status_label = ttk.Label(dialog_frame, text="Connexion au serveur...", font=('Inter', 10), foreground='orange')
        self.update_status_label.pack(pady=10)
        
        # Remote Version Label
        self.update_version_label = ttk.Label(dialog_frame, text="", font=('Inter', 12, 'bold'))
        self.update_version_label.pack(pady=5)
        
        # Action Button
        self.update_button_in_dialog = ttk.Button(dialog_frame, text="Annuler", command=self.update_dialog.destroy)
        self.update_button_in_dialog.pack(pady=10)
        
        # 2. Start Network Check
        self.update_button_in_dialog.config(state=tk.DISABLED) # Disable button while checking
        threading.Thread(target=self.check_for_update, daemon=True).start()

    def check_for_update(self):
        """Fetches the remote version from the server (runs in a separate thread)."""
        print(f"Checking remote URL: {REMOTE_VERSION_URL}")
        try:
            # Simulate network latency
            time.sleep(1.0)
            response = requests.get(REMOTE_VERSION_URL, timeout=10)
            response.raise_for_status() 

            remote_version = response.text.strip()
            
            # Schedule the UI update back on the main Tkinter thread
            self.root.after(0, lambda: self.compare_versions_and_update_ui(remote_version))

        except requests.exceptions.RequestException as e:
            error_msg = f"Network error: {e.__class__.__name__}"
            print(error_msg)
            # Schedule the error message update back on the main Tkinter thread
            self.root.after(0, lambda: self.handle_check_error(error_msg))

    def handle_check_error(self, message):
        """Updates UI in case of network error (runs on main thread)."""
        if self.update_dialog and self.update_dialog.winfo_exists():
            self.update_version_label.config(text="Échec")
            self.update_status_label.config(text="Erreur de connexion au serveur.", foreground='red')
            self.update_button_in_dialog.config(text="Fermer", state=tk.NORMAL, command=self.update_dialog.destroy)
            messagebox.showerror("Vérification Échouée", f"Impossible de vérifier la mise à jour.\nDétails: {message.split(':')[-1].strip()}")

    def compare_versions_and_update_ui(self, remote_version_str):
        """Compares versions and updates the UI accordingly (runs on main thread)."""
        if not (self.update_dialog and self.update_dialog.winfo_exists()):
            return # Dialog was closed prematurely
            
        self.remote_version = remote_version_str
        self.update_version_label.config(text=f"Remote: V{remote_version_str}", foreground='#0a84ff')
        
        current = version_to_tuple(CURRENT_VERSION)
        remote = version_to_tuple(remote_version_str)

        if remote > current:
            self.update_status_label.config(text=f"Mise à jour disponible: {remote_version_str}", foreground='#dc3545')
            self.update_button_in_dialog.config(text="Télécharger et Installer", state=tk.NORMAL, command=self.download_and_install)
            self.update_dialog.protocol("WM_DELETE_WINDOW", self.update_dialog.destroy) # Allow closing
        else:
            self.update_status_label.config(text=f"Vous utilisez la dernière version ({CURRENT_VERSION}).", foreground='green')
            self.update_button_in_dialog.config(text="Fermer", state=tk.NORMAL, command=self.update_dialog.destroy)
            self.update_dialog.protocol("WM_DELETE_WINDOW", self.update_dialog.destroy) # Allow closing
            # Automatically close success dialog after a few seconds
            self.root.after(3000, self.update_dialog.destroy)

    def download_and_install(self):
        """Initiates the simulated download process."""
        if self.update_dialog and self.update_dialog.winfo_exists():
            self.update_button_in_dialog.config(state=tk.DISABLED, text="Téléchargement en cours...")
            self.update_status_label.config(text=f"Début du téléchargement de V{self.remote_version}...", foreground='blue')
            # Start download simulation in a new thread
            threading.Thread(target=self._simulate_download, daemon=True).start()

    def _simulate_download(self):
        """Simulates downloading and installing the new executable/files."""
        # --- In a real app, replace this with actual file download and installation logic ---
        time.sleep(3) # Simulate a 3-second download/install process
        # --- End of real logic replacement ---
        
        # Schedule the success message back on the main Tkinter thread
        self.root.after(0, self._show_update_success)

    def _show_update_success(self):
        """Shows final success message and prompts for restart."""
        if self.update_dialog and self.update_dialog.winfo_exists():
            self.update_status_label.config(text=f"Mise à jour V{self.remote_version} installée!", foreground='darkgreen')
            self.update_button_in_dialog.config(text="Redémarrer l'Application", state=tk.NORMAL, command=self.restart_app)
        
        messagebox.showinfo(
            "Mise à Jour Terminée",
            f"Version {self.remote_version} a été téléchargée et installée. L'application va maintenant redémarrer."
        )

    def restart_app(self):
        """Simulates application restart."""
        if self.update_dialog and self.update_dialog.winfo_exists():
            self.update_dialog.destroy()
            
        print("Application shutting down to restart...")
        # In a real scenario, this is where you would launch the new version executable 
        # using 'subprocess.Popen' and then immediately exit the current application.
        
        messagebox.showinfo("Simulation de Redémarrage", "L'application va maintenant redémarrer.")
        self.root.destroy()
    
    # --- Existing Methods (Omitted for brevity, assumed to be here) ---

    # NOTE: The methods for handling Excel, QR generation, client selection, etc. 
    # (e.g., self.load_excel_data, self.generate_qr_code, self.select_client)
    # are assumed to be present here from your original file.

    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    """Main function to run the application"""
    app = QRScannerApp()
    app.run()

if __name__ == "__main__":
    main()