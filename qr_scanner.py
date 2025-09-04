"""
Python QR Code Scanner and Generator Application
Supports barcode scanner (douchette) input and Excel integration
"""

import pandas as pd
import qrcode
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import io
import base64
import requests
from PIL import Image, ImageTk
import threading
import time
import os
from dataclasses import dataclass, asdict
from typing import Optional, Dict, Any
import json

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
        self.root.geometry("800x600")
        
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
        ttk.Label(scanner_frame, text="Scanner Input:").grid(row=0, column=0, sticky=tk.W)
        self.scanner_entry = ttk.Entry(scanner_frame, width=50, font=('Courier', 10))
        self.scanner_entry.grid(row=0, column=1, padx=(10, 0), sticky=(tk.W, tk.E))
        self.scanner_entry.bind('<Return>', self.process_scanned_data)
        self.scanner_entry.bind('<KeyRelease>', self.on_scanner_input)
        
        scanner_frame.columnconfigure(1, weight=1)
        
        # Buttons
        button_frame = ttk.Frame(scanner_frame)
        button_frame.grid(row=1, column=0, columnspan=2, pady=(10, 0))
        
        ttk.Button(button_frame, text="Process Scan", 
                  command=self.process_scanned_data).grid(row=0, column=0, padx=(0, 5))
        ttk.Button(button_frame, text="Clear", 
                  command=self.clear_scanner_input).grid(row=0, column=1, padx=5)
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
        
        # Treeview for data display
        columns = ('ID_Produit', 'Marque', 'Designation_Reference', 'Serial_Number', 
                  'Couleur', 'Matricule', 'Magasin', 'Photo')
        
        self.tree = ttk.Treeview(data_frame, columns=columns, show='headings', height=10)
        
        # Configure column headings and widths
        for col in columns:
            self.tree.heading(col, text=col.replace('_', ' '))
            self.tree.column(col, width=120, minwidth=80)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        data_frame.columnconfigure(0, weight=1)
        data_frame.rowconfigure(0, weight=1)
        
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
    
    def on_scanner_input(self, event):
        """Handle scanner input in real-time"""
        current_text = self.scanner_entry.get()
        
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
        current_text = self.scanner_entry.get().strip()
        if current_text and not self.scanning:
            self.scanning = True
            self.process_scanned_data(None)
    
    def check_scan_complete(self):
        """Check if scan is complete and process - DEPRECATED"""
        # This method is kept for backward compatibility but not used
        current_text = self.scanner_entry.get().strip()
        if current_text and not self.scanning:
            self.scanning = True
            self.root.after(100, lambda: self.process_scanned_data(None))
    
    def process_scanned_data(self, event=None):
        """Process the scanned QR code data"""
        qr_data = self.scanner_entry.get().strip()
        
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
        
        # Check if it's the old format (enclosed in asterisks)
        if qr_data.startswith('*') and qr_data.endswith('*'):
            # Old format: *VMSDZ06CUKI191858*
            clean_data = qr_data[1:-1]  # Remove asterisks
            product.ID_Produit = clean_data
            
            # Extract information from old format pattern
            if 'CUKI' in clean_data:
                product.Marque = "CUKI"
                product.Designation_Reference = "MOTOCYCLE CUKI -I-"
            
            # Try to extract other info based on pattern
            product.Serial_Number = clean_data
            
        else:
            # New structured format - split by lines
            lines = qr_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
            lines = [line.strip() for line in lines if line.strip()]
            
            # Map lines to fields based on position
            field_mapping = [
                'ID_Produit',
                'Designation_Reference', 
                'Marque',
                'Couleur',
                'Matricule',
                'Magasin',
                'Serial_Number'
            ]
            
            for i, line in enumerate(lines):
                if i < len(field_mapping):
                    setattr(product, field_mapping[i], line)
        
        return product
    
    def generate_qr_data(self, product_data: ProductData) -> str:
        """Generate QR code data in structured format"""
        qr_lines = [
            product_data.ID_Produit,
            product_data.Designation_Reference,
            product_data.Marque,
            product_data.Couleur,
            product_data.Matricule,
            product_data.Magasin,
            product_data.Serial_Number
        ]
        
        return '\n'.join(qr_lines)
    
    def clear_scanner_input(self):
        """Clear the scanner input field"""
        # Cancel any pending scan timer
        if self.scan_timer:
            self.root.after_cancel(self.scan_timer)
            self.scan_timer = None
            
        self.scanner_entry.delete(0, tk.END)
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
        """Update the treeview with current data"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        # Add all products
        for i, product in enumerate(self.products_data):
            values = (
                product.ID_Produit,
                product.Marque,
                product.Designation_Reference,
                product.Serial_Number,
                product.Couleur,
                product.Matricule,
                product.Magasin,
                product.Photo
            )
            self.tree.insert('', 'end', values=values, tags=(str(i),))
    
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
        """Display QR code in a new window"""
        qr_data = self.generate_qr_data(product_data)
        
        # Create QR code
        qr = qrcode.QRCode(
            version=1,
            error_correction=qrcode.constants.ERROR_CORRECT_L,
            box_size=10,
            border=4,
        )
        qr.add_data(qr_data)
        qr.make(fit=True)
        
        # Create QR code image
        qr_image = qr.make_image(fill_color="black", back_color="white")
        
        # Display in new window
        qr_window = tk.Toplevel(self.root)
        qr_window.title(f"QR Code - {product_data.ID_Produit}")
        qr_window.geometry("400x500")
        
        # Convert PIL image to PhotoImage
        photo = ImageTk.PhotoImage(qr_image)
        
        label = ttk.Label(qr_window, image=photo)
        label.image = photo  # Keep a reference
        label.pack(pady=20)
        
        # Display QR data
        text_frame = ttk.LabelFrame(qr_window, text="QR Code Data", padding="10")
        text_frame.pack(fill='both', expand=True, padx=20, pady=(0, 20))
        
        text_area = tk.Text(text_frame, height=8, wrap='word')
        text_area.pack(fill='both', expand=True)
        text_area.insert('1.0', qr_data)
        text_area.config(state='disabled')
        
        # Save button
        def save_qr():
            filename = filedialog.asksaveasfilename(
                defaultextension=".png",
                filetypes=[("PNG files", "*.png"), ("All files", "*.*")]
            )
            if filename:
                qr_image.save(filename)
                messagebox.showinfo("Saved", f"QR code saved as {filename}")
        
        ttk.Button(qr_window, text="Save QR Code", command=save_qr).pack(pady=10)
    
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
    
    def run(self):
        """Start the application"""
        self.root.mainloop()

def main():
    """Main function to run the application"""
    app = QRScannerApp()
    app.run()

if __name__ == "__main__":
    main()