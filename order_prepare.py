# -*- coding: utf-8 -*-
"""
Order Preparation QR Scanner Application
Scans QR codes and manages order preparation status
Based on QR Scanner App ideology
"""

import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from PIL import Image, ImageTk
import os
import sys
from dataclasses import dataclass, asdict
import logging
import builtins
import re
import requests
from datetime import datetime
import threading
import time
# Add urllib3 for SSL warnings
import urllib3

#! REMOVE AFTER SSL FIX - Disable SSL warnings for app.diardzair.com.dz
urllib3.disable_warnings(urllib3.exceptions.InsecureRequestWarning)

# --- Logging Configuration ---
def setup_logging():
    """Configure logging to redirect all prints to a log file"""
    if getattr(sys, 'frozen', False):
        log_dir = os.path.dirname(sys.executable)
    else:
        log_dir = os.path.dirname(os.path.abspath(__file__))
    
    log_file = os.path.join(log_dir, f"Order_Prepare_{datetime.now().strftime('%Y%m%d')}.log")
    
    logging.basicConfig(
        level=logging.INFO,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.FileHandler(log_file, encoding='utf-8'),
            logging.StreamHandler(sys.stdout)
        ]
    )
    
    def log_print(*args, **kwargs):
        message = ' '.join(str(arg) for arg in args)
        logging.info(message)
    
    builtins.print = log_print
    
    logging.info("=" * 50)
    logging.info("Order Preparation Application Started")
    logging.info("=" * 50)

# Initialize logging
setup_logging()

def resource_path(relative_path):
    """Get absolute path to resource, needed for PyInstaller compilation"""
    try:
        base_path = sys._MEIPASS
    except AttributeError:
        if getattr(sys, 'frozen', False) and hasattr(sys, '_MEIPASS'):
            base_path = sys._MEIPASS
        else:
            base_path = os.path.dirname(os.path.abspath(__file__))
    
    return os.path.join(base_path, relative_path)

@dataclass
class OrderData:
    """Data structure for order preparation information"""
    DATE: str = ""
    ID: str = ""
    DESIGNATION: str = ""
    REFERENCE: str = ""
    QTE: int = 1
    PREPARED: bool = False

class OrderPrepareApp:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("Préparation Commandes - Scanner QR")
        self.root.geometry("1200x700")
        
        # Configure application icon
        try:
            icon_path = resource_path('qrcodescan.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except Exception as e:
            print(f"Could not load icon: {e}")
        
        # Data storage
        self.orders_data = []
        self.excel_file = None
        
        # Scanner state
        self.scanning = False
        self.scan_buffer = ""
        self.scan_timer = None
        
        self.setup_ui()
        self.setup_scanner_listener()
        
    def setup_ui(self):
        """Setup the main user interface"""
        # Main frame
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Préparation des Commandes", 
                               font=('Arial', 16, 'bold'))
        title_label.pack(pady=(0, 10))
        
        # Scanner input frame
        scanner_frame = ttk.LabelFrame(main_frame, text="Scanner QR / Saisie Manuelle", padding="10")
        scanner_frame.pack(fill='x', pady=(0, 10))
        
        # Scanner input
        ttk.Label(scanner_frame, text="Scanner QR Code:").pack(anchor='w')
        
        # Create a frame for the text widget and scrollbar
        text_frame = ttk.Frame(scanner_frame)
        text_frame.pack(fill='x', pady=(5, 10))
        text_frame.columnconfigure(0, weight=1)
        
        # Text widget for multi-line input
        self.scanner_entry = tk.Text(text_frame, width=60, height=4, font=('Courier', 10), wrap='word')
        self.scanner_entry.grid(row=0, column=0, sticky=(tk.W, tk.E))
        
        # Scrollbar for the text widget
        scanner_scrollbar = ttk.Scrollbar(text_frame, orient="vertical", command=self.scanner_entry.yview)
        scanner_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        self.scanner_entry.configure(yscrollcommand=scanner_scrollbar.set)
        
        # Bind events (note: Text widget uses different event handling)
        self.scanner_entry.bind('<KeyRelease>', self.on_scanner_input)
        
        text_frame.columnconfigure(0, weight=1)
        
        # Manual entry button
        manual_frame = ttk.Frame(scanner_frame)
        manual_frame.pack(fill='x')
        
        ttk.Button(manual_frame, text="Ajouter Manuellement", 
                  command=self.open_manual_entry_dialog).pack(side='left', padx=(0, 10))
        
        ttk.Button(manual_frame, text="Effacer", 
                  command=self.clear_scanner_input).pack(side='left')
        
        # Search frame
        search_frame = ttk.Frame(main_frame)
        search_frame.pack(fill='x', pady=(0, 10))
        
        ttk.Label(search_frame, text="Rechercher:").pack(side='left', padx=(0, 5))
        self.search_var = tk.StringVar()
        self.search_var.trace('w', self.on_search_change)
        search_entry = ttk.Entry(search_frame, textvariable=self.search_var, width=30)
        search_entry.pack(side='left', padx=(0, 10))
        
        # Filter frame
        filter_frame = ttk.Frame(search_frame)
        filter_frame.pack(side='left', padx=(10, 0))
        
        ttk.Label(filter_frame, text="Filtre:").pack(side='left', padx=(0, 5))
        self.filter_var = tk.StringVar(value="Tous")
        filter_combo = ttk.Combobox(filter_frame, textvariable=self.filter_var, 
                                   values=["Tous", "Préparés", "Non Préparés"], 
                                   state="readonly", width=12)
        filter_combo.pack(side='left')
        filter_combo.bind('<<ComboboxSelected>>', self.on_filter_change)
        
        # Data display frame
        data_frame = ttk.LabelFrame(main_frame, text="Commandes", padding="10")
        data_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Treeview setup
        columns = ("DATE", "ID", "DESIGNATION", "REFERENCE", "QTE", "PREPARED")
        self.tree = ttk.Treeview(data_frame, columns=columns, show='headings', height=15)
        
        # Configure columns
        self.tree.heading("DATE", text="Date")
        self.tree.heading("ID", text="ID Client")
        self.tree.heading("DESIGNATION", text="Désignation")
        self.tree.heading("REFERENCE", text="Référence")
        self.tree.heading("QTE", text="Qté")
        self.tree.heading("PREPARED", text="Préparé")
        
        # Column widths
        self.tree.column("DATE", width=100)
        self.tree.column("ID", width=100)
        self.tree.column("DESIGNATION", width=200)
        self.tree.column("REFERENCE", width=150)
        self.tree.column("QTE", width=60)
        self.tree.column("PREPARED", width=80)
        
        # Scrollbars
        v_scrollbar = ttk.Scrollbar(data_frame, orient="vertical", command=self.tree.yview)
        h_scrollbar = ttk.Scrollbar(data_frame, orient="horizontal", command=self.tree.xview)
        self.tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Pack treeview and scrollbars
        self.tree.grid(row=0, column=0, sticky='nsew')
        v_scrollbar.grid(row=0, column=1, sticky='ns')
        h_scrollbar.grid(row=1, column=0, sticky='ew')
        
        data_frame.grid_rowconfigure(0, weight=1)
        data_frame.grid_columnconfigure(0, weight=1)
        
        # Bind events
        self.tree.bind('<Double-1>', self.on_item_double_click)
        self.tree.bind('<Delete>', self.on_delete_key)
        self.tree.bind('<KeyPress-Delete>', self.on_delete_key)
        
        # Focus the tree to enable keyboard events
        self.tree.focus_set()
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill='x')
        
        # File operations
        ttk.Button(button_frame, text="Charger Excel", 
                  command=self.load_excel_file).pack(side='left', padx=(0, 5))
        
        ttk.Button(button_frame, text="Sauvegarder Excel", 
                  command=self.save_excel_file).pack(side='left', padx=(0, 5))
        
        ttk.Button(button_frame, text="Nouveau Fichier", 
                  command=self.clear_all_data).pack(side='left', padx=(0, 10))
        
        # Data operations
        ttk.Button(button_frame, text="Modifier", 
                  command=self.edit_selected_record).pack(side='left', padx=(0, 5))
        
        ttk.Button(button_frame, text="Supprimer", 
                  command=self.delete_selected_record).pack(side='left', padx=(0, 5))
        
        ttk.Button(button_frame, text="Marquer Préparé", 
                  command=self.toggle_preparation_status).pack(side='left', padx=(0, 10))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Prêt", foreground='green')
        self.status_label.pack(pady=(10, 0))
        
    def setup_scanner_listener(self):
        """Setup the scanner input listener"""
        self.root.bind('<FocusIn>', lambda e: self.scanner_entry.focus_set())
        
    def on_scanner_input(self, event):
        """Handle scanner input with auto-processing"""
        if self.scan_timer:
            self.root.after_cancel(self.scan_timer)
        
        # Auto-process after 100ms of no input (typical for barcode scanners)
        self.scan_timer = self.root.after(100, self.auto_process_scan)
        
    def auto_process_scan(self):
        """Auto-process scanned data"""
        data = self.scanner_entry.get("1.0", tk.END).strip()
        if data and len(data) > 10:  # Minimum length for valid QR data
            self.process_scanned_data()
            
    def process_scanned_data(self, event=None):
        """Process the scanned QR code data"""
        qr_data = self.scanner_entry.get("1.0", tk.END).strip()
        if not qr_data:
            return
            
        try:
            self.status_label.config(text="Traitement en cours...", foreground='blue')
            self.root.update()
            
            # Parse QR data
            order_data = self.parse_qr_data(qr_data)
            
            if order_data:
                # Add current date
                order_data.DATE = datetime.now().strftime("%d/%m/%Y")
                
                # Check if client ID was found, if not, don't add the order
                if not order_data.ID:
                    self.status_label.config(text="Erreur: Aucun client trouvé pour ce QR code", foreground='red')
                    messagebox.showwarning("Client non trouvé", 
                                         f"Aucun client n'a été trouvé pour la référence {order_data.REFERENCE}.\n"
                                         f"Veuillez utiliser l'ajout manuel pour sélectionner un client.")
                    self.clear_scanner_input()
                    return
                
                # Check for duplicates
                existing_index = self.find_existing_order(order_data)
                if existing_index is not None:
                    # Update quantity instead of adding duplicate
                    self.orders_data[existing_index].QTE += 1
                    messagebox.showinfo("Info", f"Quantité mise à jour pour {order_data.REFERENCE}")
                else:
                    # Add new order
                    self.orders_data.append(order_data)
                    messagebox.showinfo("Succès", f"Commande ajoutée: {order_data.REFERENCE} (Client ID: {order_data.ID})")
                
                self.update_tree_display()
                self.clear_scanner_input()
                self.status_label.config(text="Prêt", foreground='green')
            else:
                self.status_label.config(text="Erreur: Données QR invalides", foreground='red')
                messagebox.showerror("QR Invalide", "Les données du QR code ne peuvent pas être analysées.")
                
        except Exception as e:
            messagebox.showerror("Erreur", f"Erreur lors du traitement: {str(e)}")
            self.status_label.config(text="Erreur", foreground='red')
            
    def parse_qr_data(self, qr_data: str):
        """Parse QR code data to extract order information"""
        try:
            order_data = OrderData()
            chassis_number = ""
            
            # Check if data contains line breaks (structured format)
            if '\n' in qr_data or '\r' in qr_data:
                lines = qr_data.replace('\r\n', '\n').replace('\r', '\n').split('\n')
                lines = [line.strip() for line in lines]
                
                # Check if first line starts with asterisk (legacy format)
                if lines and lines[0].startswith('*') and lines[0].endswith('*'):
                    # Legacy format
                    if len(lines) >= 6:
                        order_data.REFERENCE = lines[0][1:-1] if lines[0].startswith('*') else lines[0]
                        order_data.DESIGNATION = lines[1] if len(lines) > 1 else ""
                        # Chassis number is in line 5 for legacy format
                        chassis_number = lines[5] if len(lines) > 5 else ""
                else:
                    # New structured format
                    if len(lines) >= 4:
                        order_data.REFERENCE = lines[0]
                        order_data.DESIGNATION = lines[2] if len(lines) > 2 else ""
                        chassis_number = lines[3] if len(lines) > 3 else ""
            else:
                # Single line input
                if '*' in qr_data:
                    # Extract reference between asterisks
                    import re
                    asterisk_match = re.search(r'\*([^*]+)\*', qr_data)
                    if asterisk_match:
                        order_data.REFERENCE = asterisk_match.group(1)
                        # Extract remaining data
                        remaining = qr_data[asterisk_match.end():].strip()
                        if remaining:
                            order_data.DESIGNATION = remaining
                        chassis_number = ""
                else:
                    # Simple format
                    order_data.REFERENCE = qr_data
                    chassis_number = ""
            
            # Fetch client ID from chassis number if available
            order_data.ID = ""  # Initialize as empty
            if chassis_number:
                try:
                    print(f"DEBUG - Attempting to fetch client info for chassis: {chassis_number}")
                    client_id = self.fetch_client_info_from_chassis(chassis_number)
                    if client_id:
                        order_data.ID = str(client_id)
                        print(f"DEBUG - Client ID found: {client_id}")
                    else:
                        print(f"DEBUG - No client ID found for chassis: {chassis_number}")
                except Exception as e:
                    print(f"DEBUG - Error fetching client info: {e}")
                    order_data.ID = ""
            else:
                print(f"DEBUG - No chassis number found in QR data")
            
            # Set default quantity
            order_data.QTE = 1
            order_data.PREPARED = False
            
            # Return order data only if we have a reference (minimum requirement)
            if order_data.REFERENCE:
                return order_data
            else:
                print(f"DEBUG - No reference found in QR data")
                return None
            
        except Exception as e:
            print(f"Error parsing QR data: {e}")
            return None
            
    def fetch_client_info_from_chassis(self, chassis_number):
        """Fetch client info using chassis number from the API"""
        try:
            url = f"https://app.diardzair.com.dz/api/vehicles/{chassis_number}"
            headers = {
                'accept': 'application/json',
                'Content-Type': 'application/json',
                'Authorization': 'Bearer eyJ0eXAiOiJKV1QiLCJhbGciOiJSUzI1NiJ9.eyJhdWQiOiIxIiwianRpIjoiNWM1ODJjMjY4OTQwMDAyNTFlNzUzOGVjZjI0NzBmZjE4YTJhMzJkNGI1ZGFkNjU0OWI0OGE0YzQyM2Y1NTA5NWEyZTFhNWI1NGExOGE1M2QiLCJpYXQiOjE3MTEwNzYzNDIuNDQyMDA0LCJuYmYiOjE3MTEwNzYzNDIuNDQyMDA2LCJleHAiOjE3NDI2MTIzNDIuNDM0Nzk1LCJzdWIiOiI1Iiwic2NvcGVzIjpbXX0.zZJznq9EhIDbWa2Gkh5Vv3-Pd7i0DXCiGGvJZYWt7l2rWQmvBpNmSqnrEFTz0cEz9E7xQh-4z1T0Wg9Bd2gT0gKNNW6NiEWbBzjqkXxK7LZq8_PfKqJ1VRvjzgNi1l8Z2G-mZXb1FqSI8fAXxAu-I3_6hBW4G2zY8P3mU1bQzBdZLRsKV5w3Q7pQp3zYKr8O9vH5O1Tg7CnFz4D9lIhGqKz5yF2xZ0nYqCVN3l8kQTFhXy5Y7J3QhF4pVJ0bNgV8M0QgK1LcZaT6M3zE5R8dP2wI7hL4rQ9sC0mB6vN9xGtH2fL8oK5qS4uY1pJ7eA3zW0xT6cFj9vB8nE4iM1uL7tP5gR2yO'
            }
            
            print(f"DEBUG - Sending GET request to: {url}")
            
            response = requests.get(url, headers=headers, timeout=10, verify=False)
            
            print(f"DEBUG - Response status code: {response.status_code}")
            print(f"DEBUG - Response headers: {response.headers}")
            
            if response.status_code == 200:
                data = response.json()
                print(f"DEBUG - Response data: {data}")
                
                if data and 'id' in data:
                    return data['id']
                else:
                    print(f"DEBUG - No 'id' found in response data")
                    return None
            else:
                print(f"DEBUG - API request failed with status code: {response.status_code}")
                print(f"DEBUG - Response text: {response.text}")
                return None
                
        except requests.RequestException as e:
            print(f"DEBUG - API request error: {e}")
            return None
        except Exception as e:
            print(f"DEBUG - Unexpected error: {e}")
            return None

    def find_existing_order(self, order_data):
        """Find existing order with same reference and client ID"""
        for i, existing_order in enumerate(self.orders_data):
            if (existing_order.REFERENCE == order_data.REFERENCE and 
                existing_order.ID == order_data.ID):
                return i
        return None
        
    def clear_scanner_input(self):
        """Clear the scanner input field"""
        self.scanner_entry.delete("1.0", tk.END)
        self.scanner_entry.focus_set()
        
    def update_tree_display(self):
        """Update the treeview display"""
        # Clear existing items
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        # Apply search filter
        search_term = self.search_var.get().lower()
        filter_value = self.filter_var.get()
        
        # Filter data
        filtered_data = []
        for order in self.orders_data:
            # Apply search filter
            if search_term:
                searchable_text = f"{order.DATE} {order.ID} {order.DESIGNATION} {order.REFERENCE}".lower()
                if search_term not in searchable_text:
                    continue
                    
            # Apply preparation filter
            if filter_value == "Préparés" and not order.PREPARED:
                continue
            elif filter_value == "Non Préparés" and order.PREPARED:
                continue
                
            filtered_data.append(order)
        
        # Add filtered data to tree
        for order in filtered_data:
            prepared_text = "Oui" if order.PREPARED else "Non"
            values = (order.DATE, order.ID, order.DESIGNATION, order.REFERENCE, order.QTE, prepared_text)
            
            # Color coding
            if order.PREPARED:
                item = self.tree.insert('', 'end', values=values, tags=('prepared',))
            else:
                item = self.tree.insert('', 'end', values=values, tags=('unprepared',))
        
        # Configure tags for color coding
        self.tree.tag_configure('prepared', background='#d4edda', foreground='#155724')
        self.tree.tag_configure('unprepared', background='#fff3cd', foreground='#856404')
        
    def on_search_change(self, *args):
        """Handle search input changes"""
        self.update_tree_display()
        
    def on_filter_change(self, event=None):
        """Handle filter changes"""
        self.update_tree_display()
        
    def on_item_double_click(self, event):
        """Handle double-click on tree item"""
        self.toggle_preparation_status()
        
    def toggle_preparation_status(self):
        """Toggle preparation status of selected item"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Attention", "Veuillez sélectionner une commande")
            return
            
        item = selected[0]
        values = self.tree.item(item, 'values')
        
        if values:
            # Find the order in data
            for order in self.orders_data:
                if (order.DATE == values[0] and order.ID == values[1] and 
                    order.REFERENCE == values[3]):
                    order.PREPARED = not order.PREPARED
                    status = "préparée" if order.PREPARED else "non préparée"
                    messagebox.showinfo("Succès", f"Commande marquée comme {status}")
                    self.update_tree_display()
                    break
                    
    def open_manual_entry_dialog(self):
        """Open dialog for manual entry"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Ajouter Commande Manuellement")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        
        # Center the dialog
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill='both', expand=True)
        
        # Variables to store client data
        selected_client_id = tk.StringVar()
        selected_client_name = tk.StringVar(value="Aucun client sélectionné")
        
        # Date field (editable)
        ttk.Label(form_frame, text="Date:").grid(row=0, column=0, sticky='w', pady=5)
        date_var = tk.StringVar(value=datetime.now().strftime("%d/%m/%Y"))
        ttk.Entry(form_frame, textvariable=date_var, width=30).grid(row=0, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        # Client selection row
        ttk.Label(form_frame, text="Client:").grid(row=1, column=0, sticky='w', pady=5)
        client_frame = ttk.Frame(form_frame)
        client_frame.grid(row=1, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        def choose_client():
            client = self.open_api_client_selection_dialog()
            if client:
                selected_client_id.set(str(client['ID_CLIENT']))
                selected_client_name.set(f"{client['NOM_PRENOM']} (ID: {client['ID_CLIENT']})")
        
        ttk.Button(client_frame, text="Choisir Client", command=choose_client).pack(side='left')
        client_label = ttk.Label(client_frame, textvariable=selected_client_name, fg="blue", wraplength=250)
        client_label.pack(side='left', padx=(10, 0))
        
        # Entry fields (all editable)
        ttk.Label(form_frame, text="Désignation:").grid(row=2, column=0, sticky='w', pady=5)
        designation_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=designation_var, width=30).grid(row=2, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Référence:").grid(row=3, column=0, sticky='w', pady=5)
        reference_var = tk.StringVar()
        ttk.Entry(form_frame, textvariable=reference_var, width=30).grid(row=3, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Quantité:").grid(row=4, column=0, sticky='w', pady=5)
        qte_var = tk.StringVar(value="1")
        ttk.Entry(form_frame, textvariable=qte_var, width=30).grid(row=4, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Préparé:").grid(row=5, column=0, sticky='w', pady=5)
        prepared_var = tk.BooleanVar()
        ttk.Checkbutton(form_frame, variable=prepared_var).grid(row=5, column=1, sticky='w', pady=5, padx=(10, 0))
        
        # Configure column weight
        form_frame.columnconfigure(1, weight=1)
        
        def add_manual_order():
            try:
                # Validate client selection
                client_id = selected_client_id.get().strip()
                if not client_id:
                    messagebox.showerror("Erreur", "Veuillez sélectionner un client")
                    return
                
                # Validate required fields
                if not reference_var.get().strip():
                    messagebox.showerror("Erreur", "La référence est obligatoire")
                    return
                
                # Validate quantity
                try:
                    qte = int(qte_var.get().strip()) if qte_var.get().strip() else 1
                    if qte <= 0:
                        raise ValueError("La quantité doit être positive")
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez entrer une quantité valide (nombre entier positif)")
                    return
                
                order = OrderData()
                order.DATE = date_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
                order.ID = client_id
                order.DESIGNATION = designation_var.get().strip()
                order.REFERENCE = reference_var.get().strip()
                order.QTE = qte
                order.PREPARED = prepared_var.get()
                
                # Check for duplicates
                existing_index = self.find_existing_order(order)
                if existing_index is not None:
                    self.orders_data[existing_index].QTE += order.QTE
                    messagebox.showinfo("Info", "Quantité mise à jour pour cette commande")
                else:
                    self.orders_data.append(order)
                    messagebox.showinfo("Succès", "Commande ajoutée avec succès")
                
                self.update_tree_display()
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de l'ajout: {str(e)}")
        
        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Ajouter", command=add_manual_order).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="Annuler", command=dialog.destroy).pack(side='left')
        
    def edit_selected_record(self):
        """Edit the selected record"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Attention", "Veuillez sélectionner une commande")
            return
            
        item = selected[0]
        values = self.tree.item(item, 'values')
        
        if values:
            # Find the order in data
            for i, order in enumerate(self.orders_data):
                if (order.DATE == values[0] and order.ID == values[1] and 
                    order.REFERENCE == values[3]):
                    self.open_edit_dialog(i)
                    break
                    
    def open_edit_dialog(self, index):
        """Open dialog to edit an order"""
        order = self.orders_data[index]
        
        dialog = tk.Toplevel(self.root)
        dialog.title("Modifier Commande")
        dialog.geometry("450x400")
        dialog.resizable(False, False)
        
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Form frame
        form_frame = ttk.Frame(dialog, padding="20")
        form_frame.pack(fill='both', expand=True)
        
        # Variables to store client data
        selected_client_id = tk.StringVar(value=order.ID)
        selected_client_name = tk.StringVar(value=f"Client ID: {order.ID}")
        
        # Entry fields with current values
        ttk.Label(form_frame, text="Date:").grid(row=0, column=0, sticky='w', pady=5)
        date_var = tk.StringVar(value=order.DATE)
        ttk.Entry(form_frame, textvariable=date_var, width=30).grid(row=0, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        # Client selection row with change button
        ttk.Label(form_frame, text="Client:").grid(row=1, column=0, sticky='w', pady=5)
        client_frame = ttk.Frame(form_frame)
        client_frame.grid(row=1, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        def change_client():
            client = self.open_api_client_selection_dialog()
            if client:
                selected_client_id.set(str(client['ID_CLIENT']))
                selected_client_name.set(f"{client['NOM_PRENOM']} (ID: {client['ID_CLIENT']})")
        
        ttk.Button(client_frame, text="Changer Client", command=change_client).pack(side='left')
        client_label = ttk.Label(client_frame, textvariable=selected_client_name, fg="blue", wraplength=250)
        client_label.pack(side='left', padx=(10, 0))
        
        ttk.Label(form_frame, text="Désignation:").grid(row=2, column=0, sticky='w', pady=5)
        designation_var = tk.StringVar(value=order.DESIGNATION)
        ttk.Entry(form_frame, textvariable=designation_var, width=30).grid(row=2, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Référence:").grid(row=3, column=0, sticky='w', pady=5)
        reference_var = tk.StringVar(value=order.REFERENCE)
        ttk.Entry(form_frame, textvariable=reference_var, width=30).grid(row=3, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Quantité:").grid(row=4, column=0, sticky='w', pady=5)
        qte_var = tk.StringVar(value=str(order.QTE))
        ttk.Entry(form_frame, textvariable=qte_var, width=30).grid(row=4, column=1, pady=5, padx=(10, 0), sticky='ew')
        
        ttk.Label(form_frame, text="Préparé:").grid(row=5, column=0, sticky='w', pady=5)
        prepared_var = tk.BooleanVar(value=order.PREPARED)
        ttk.Checkbutton(form_frame, variable=prepared_var).grid(row=5, column=1, sticky='w', pady=5, padx=(10, 0))
        
        # Configure column weight
        form_frame.columnconfigure(1, weight=1)
        
        def save_changes():
            try:
                # Validate required fields
                if not reference_var.get().strip():
                    messagebox.showerror("Erreur", "La référence est obligatoire")
                    return
                
                # Validate quantity
                try:
                    qte = int(qte_var.get().strip()) if qte_var.get().strip() else 1
                    if qte <= 0:
                        raise ValueError("La quantité doit être positive")
                except ValueError:
                    messagebox.showerror("Erreur", "Veuillez entrer une quantité valide (nombre entier positif)")
                    return
                
                self.orders_data[index].DATE = date_var.get().strip() or datetime.now().strftime("%d/%m/%Y")
                self.orders_data[index].ID = selected_client_id.get().strip()
                self.orders_data[index].DESIGNATION = designation_var.get().strip()
                self.orders_data[index].REFERENCE = reference_var.get().strip()
                self.orders_data[index].QTE = qte
                self.orders_data[index].PREPARED = prepared_var.get()
                
                self.update_tree_display()
                messagebox.showinfo("Succès", "Commande modifiée avec succès")
                dialog.destroy()
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la modification: {str(e)}")
        
        # Buttons
        button_frame = ttk.Frame(form_frame)
        button_frame.grid(row=6, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Sauvegarder", command=save_changes).pack(side='left', padx=(0, 10))
        ttk.Button(button_frame, text="Annuler", command=dialog.destroy).pack(side='left')
        
    def delete_selected_record(self):
        """Delete the selected record(s) - updated to handle multiple selection"""
        selected = self.tree.selection()
        if not selected:
            messagebox.showwarning("Attention", "Veuillez sélectionner une commande")
            return
            
        # Use the same logic as on_delete_key
        self.on_delete_key(None)
    
    def on_delete_key(self, event):
        """Handle delete key press to delete selected items"""
        selected = self.tree.selection()
        if not selected:
            return
            
        # Confirm deletion
        count = len(selected)
        if count == 1:
            message = "Êtes-vous sûr de vouloir supprimer cette commande?"
        else:
            message = f"Êtes-vous sûr de vouloir supprimer ces {count} commandes?"
            
        if messagebox.askyesno("Confirmation", message):
            # Collect orders to delete based on selection
            orders_to_delete = []
            
            for item in selected:
                values = self.tree.item(item, 'values')
                if values:
                    # Find the order in data
                    for i, order in enumerate(self.orders_data):
                        if (order.DATE == values[0] and order.ID == values[1] and 
                            order.REFERENCE == values[3]):
                            orders_to_delete.append(i)
                            break
            
            # Delete orders in reverse order to maintain indices
            for index in sorted(orders_to_delete, reverse=True):
                del self.orders_data[index]
            
            self.update_tree_display()
            
            if count == 1:
                messagebox.showinfo("Succès", "Commande supprimée")
            else:
                messagebox.showinfo("Succès", f"{count} commandes supprimées")
        
    def load_excel_file(self):
        """Load data from Excel file"""
        file_path = filedialog.askopenfilename(
            title="Charger fichier Excel",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        
        if file_path:
            try:
                df = pd.read_excel(file_path)
                
                # Verify columns
                required_columns = ["DATE", "ID", "DESIGNATION", "REFERENCE", "QTE", "PREPARED"]
                missing_columns = [col for col in required_columns if col not in df.columns]
                
                if missing_columns:
                    messagebox.showerror("Erreur", f"Colonnes manquantes: {', '.join(missing_columns)}")
                    return
                
                # Load data
                self.orders_data = []
                for _, row in df.iterrows():
                    order = OrderData()
                    order.DATE = str(row["DATE"]) if pd.notna(row["DATE"]) else ""
                    order.ID = str(row["ID"]) if pd.notna(row["ID"]) else ""
                    order.DESIGNATION = str(row["DESIGNATION"]) if pd.notna(row["DESIGNATION"]) else ""
                    order.REFERENCE = str(row["REFERENCE"]) if pd.notna(row["REFERENCE"]) else ""
                    order.QTE = int(row["QTE"]) if pd.notna(row["QTE"]) else 1
                    order.PREPARED = bool(row["PREPARED"]) if pd.notna(row["PREPARED"]) else False
                    
                    self.orders_data.append(order)
                
                self.excel_file = file_path
                self.update_tree_display()
                messagebox.showinfo("Succès", f"Fichier chargé: {len(self.orders_data)} commandes")
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors du chargement: {str(e)}")
                
    def save_excel_file(self):
        """Save data to Excel file"""
        if not self.orders_data:
            messagebox.showwarning("Attention", "Aucune donnée à sauvegarder")
            return
            
        file_path = self.excel_file
        if not file_path:
            file_path = filedialog.asksaveasfilename(
                title="Sauvegarder fichier Excel",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")]
            )
            
        if file_path:
            try:
                # Convert data to DataFrame
                data_dicts = []
                for order in self.orders_data:
                    data_dict = asdict(order)
                    data_dicts.append(data_dict)
                
                df = pd.DataFrame(data_dicts)
                
                # Save to Excel
                df.to_excel(file_path, index=False)
                
                self.excel_file = file_path
                messagebox.showinfo("Succès", f"Fichier sauvegardé: {file_path}")
                
            except Exception as e:
                messagebox.showerror("Erreur", f"Erreur lors de la sauvegarde: {str(e)}")
                
    def clear_all_data(self):
        """Clear all data and start new file"""
        if self.orders_data:
            if messagebox.askyesno("Confirmation", "Êtes-vous sûr de vouloir effacer toutes les données?"):
                self.orders_data = []
                self.excel_file = None
                self.update_tree_display()
                messagebox.showinfo("Succès", "Données effacées")
        else:
            messagebox.showinfo("Info", "Aucune donnée à effacer")
            
    def run(self):
        """Run the application"""
        self.root.mainloop()

    def open_api_client_selection_dialog(self):
        """Open dialog to search and select a client using API"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Rechercher et Sélectionner un Client")
        dialog.geometry("900x700")
        dialog.resizable(True, True)
        dialog.transient(self.root)
        dialog.grab_set()
        
        # Center the dialog
        dialog.update_idletasks()
        x = (dialog.winfo_screenwidth() // 2) - (dialog.winfo_width() // 2)
        y = (dialog.winfo_screenheight() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        selected_client = None
        current_page = 1
        total_pages = 1
        clients_per_page = 200
        
        # Main frame
        main_frame = ttk.Frame(dialog, padding="10")
        main_frame.pack(fill='both', expand=True)
        
        # Title
        ttk.Label(main_frame, text="Rechercher un Client dans la Base de Données", 
                font=('Arial', 14, 'bold')).pack(pady=(0, 20))
        
        # Search frame
        search_frame = ttk.LabelFrame(main_frame, text="Critères de Recherche", padding="10")
        search_frame.pack(fill='x', pady=(0, 10))
        
        # Search fields
        ttk.Label(search_frame, text="ID Client:").grid(row=0, column=0, sticky=tk.W, padx=(0, 5))
        search_id_var = tk.StringVar()
        search_id_entry = ttk.Entry(search_frame, textvariable=search_id_var, width=15)
        search_id_entry.grid(row=0, column=1, padx=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Label(search_frame, text="Nom:").grid(row=0, column=2, sticky=tk.W, padx=(0, 5))
        search_nom_var = tk.StringVar()
        search_nom_entry = ttk.Entry(search_frame, textvariable=search_nom_var, width=20)
        search_nom_entry.grid(row=0, column=3, padx=(0, 20), sticky=(tk.W, tk.E))
        
        ttk.Label(search_frame, text="Prénom:").grid(row=0, column=4, sticky=tk.W, padx=(0, 5))
        search_prenom_var = tk.StringVar()
        search_prenom_entry = ttk.Entry(search_frame, textvariable=search_prenom_var, width=20)
        search_prenom_entry.grid(row=0, column=5, padx=(0, 10), sticky=(tk.W, tk.E))
        
        # Search and Clear buttons
        button_row = ttk.Frame(search_frame)
        button_row.grid(row=1, column=0, columnspan=6, pady=(10, 0), sticky=(tk.W, tk.E))
        
        search_button = ttk.Button(button_row, text="Rechercher", command=lambda: search_clients())
        search_button.pack(side=tk.LEFT, padx=(0, 10))
        
        clear_button = ttk.Button(button_row, text="Effacer", command=lambda: clear_search())
        clear_button.pack(side=tk.LEFT)
        
        search_frame.columnconfigure(1, weight=1)
        search_frame.columnconfigure(3, weight=1)
        search_frame.columnconfigure(5, weight=1)
        
        # Results frame
        results_frame = ttk.LabelFrame(main_frame, text="Résultats de Recherche", padding="10")
        results_frame.pack(fill='both', expand=True, pady=(0, 10))
        
        # Treeview for results
        result_columns = ('ID', 'Nom', 'Prénom', 'Mobile', 'Email')
        result_tree = ttk.Treeview(results_frame, columns=result_columns, show='headings', height=15)
        
        # Configure column headings
        for col in result_columns:
            result_tree.heading(col, text=col)
            if col == 'ID':
                result_tree.column(col, width=80, minwidth=60)
            elif col == 'Mobile':
                result_tree.column(col, width=120, minwidth=100)
            elif col == 'Email':
                result_tree.column(col, width=200, minwidth=150)
            else:
                result_tree.column(col, width=150, minwidth=100)
        
        # Scrollbars for result tree
        v_scrollbar = ttk.Scrollbar(results_frame, orient="vertical", command=result_tree.yview)
        h_scrollbar = ttk.Scrollbar(results_frame, orient="horizontal", command=result_tree.xview)
        result_tree.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        # Grid layout for result tree
        result_tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        v_scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        h_scrollbar.grid(row=1, column=0, sticky=(tk.W, tk.E))
        
        results_frame.columnconfigure(0, weight=1)
        results_frame.rowconfigure(0, weight=1)
        
        # Pagination frame
        pagination_frame = ttk.Frame(results_frame)
        pagination_frame.grid(row=2, column=0, columnspan=2, sticky=(tk.W, tk.E), pady=(10, 0))
        
        # Pagination controls
        prev_button = ttk.Button(pagination_frame, text="◄ Précédent", command=lambda: change_page(-1))
        prev_button.pack(side=tk.LEFT)
        
        page_label = ttk.Label(pagination_frame, text="Page 1 sur 1")
        page_label.pack(side=tk.LEFT, padx=20)
        
        next_button = ttk.Button(pagination_frame, text="Suivant ►", command=lambda: change_page(1))
        next_button.pack(side=tk.LEFT)
        
        # Status label
        status_label = ttk.Label(results_frame, text="Chargement des clients...", 
                                foreground="blue")
        status_label.grid(row=3, column=0, columnspan=2, pady=10)
        
        # Variables to store current search criteria
        current_search_id = ""
        current_search_nom = ""
        current_search_prenom = ""
        
        # Function to change page
        def change_page(direction):
            nonlocal current_page
            new_page = current_page + direction
            if 1 <= new_page <= total_pages:
                current_page = new_page
                # Use stored search criteria for pagination
                load_clients(current_page, current_search_id, current_search_nom, current_search_prenom)
        
        # Function to update pagination controls
        def update_pagination():
            page_label.config(text=f"Page {current_page} sur {total_pages}")
            prev_button.config(state=tk.NORMAL if current_page > 1 else tk.DISABLED)
            next_button.config(state=tk.NORMAL if current_page < total_pages else tk.DISABLED)
        
        # Function to load clients (with or without filters)
        def load_clients(page=1, client_id="", nom="", prenom=""):
            nonlocal current_page, total_pages
            
            # Clear existing results
            for item in result_tree.get_children():
                result_tree.delete(item)
            
            status_label.config(text="Chargement en cours...", foreground="blue")
            dialog.update()
            
            try:
                # Call API to get clients
                result = self.fetch_clients_from_api_with_pagination(client_id, nom, prenom, page, clients_per_page)
                clients = result['clients']
                total_pages = result['total_pages']
                current_page = page
                
                if clients:
                    # Populate results
                    for client in clients:
                        result_tree.insert('', 'end', values=(
                            client["id"],
                            client["nom"],
                            client["prenom"],
                            client.get("mobile", ""),
                            client.get("email", "")
                        ), tags=(str(client["id"]),))
                    
                    status_label.config(text=f"{len(clients)} client(s) affiché(s) - Total: {result['total_clients']}", foreground="green")
                else:
                    status_label.config(text="Aucun client trouvé", foreground="orange")
                
                update_pagination()
                    
            except Exception as e:
                status_label.config(text=f"Erreur lors du chargement: {str(e)}", foreground="red")
                messagebox.showerror("Erreur API", f"Erreur lors du chargement des clients:\n{str(e)}")
        
        # Function to search clients via API
        def search_clients():
            nonlocal current_search_id, current_search_nom, current_search_prenom
            
            # Update stored search criteria
            current_search_id = search_id_var.get().strip()
            current_search_nom = search_nom_var.get().strip()
            current_search_prenom = search_prenom_var.get().strip()
            
            # Load with filters
            load_clients(1, current_search_id, current_search_nom, current_search_prenom)
        
        # Function to clear search and reload all clients
        def clear_search():
            nonlocal current_search_id, current_search_nom, current_search_prenom
            
            # Clear variables
            search_id_var.set("")
            search_nom_var.set("")
            search_prenom_var.set("")
            
            # Clear stored search criteria
            current_search_id = ""
            current_search_nom = ""
            current_search_prenom = ""
            
            load_clients(1)  # Load all clients from page 1
        
        # Button frame
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=20)
        
        def select_client():
            nonlocal selected_client
            selection = result_tree.selection()
            if not selection:
                messagebox.showwarning("Sélection", "Veuillez sélectionner un client dans la liste")
                return
            
            item = selection[0]
            values = result_tree.item(item, 'values')
            tags = result_tree.item(item, 'tags')
            
            if values and tags:
                client_id = tags[0]
                
                # Show loading message
                status_label.config(text="Récupération des informations détaillées...", foreground="blue")
                dialog.update()
                
                try:
                    # Fetch detailed client info including wilaya from the API
                    client_details = self.fetch_client_details_from_api(client_id)
                    
                    selected_client = {
                        "ID_CLIENT": str(values[0]),
                        "NOM_PRENOM": f"{values[1]} {values[2]}".strip(),
                        "WILAYA": client_details.get("wilaya", "Alger")  # Default to Alger if not found
                    }
                    
                    status_label.config(text="Client sélectionné avec succès", foreground="green")
                    dialog.destroy()
                    
                except Exception as e:
                    # If API fails, ask user to select wilaya manually as fallback
                    status_label.config(text="Erreur lors de la récupération de la wilaya", foreground="orange")
                    messagebox.showwarning(
                        "Erreur API", 
                        f"Impossible de récupérer la wilaya automatiquement: {str(e)}\n\n"
                        f"Veuillez sélectionner la wilaya manuellement."
                    )
                    # For order_prepare, we don't need wilaya, so we'll just use the client info
                    selected_client = {
                        "ID_CLIENT": str(values[0]),
                        "NOM_PRENOM": f"{values[1]} {values[2]}".strip()
                    }
                    dialog.destroy()
        
        def cancel_selection():
            nonlocal selected_client
            selected_client = None
            dialog.destroy()
        
        ttk.Button(button_frame, text="Sélectionner Client", 
                command=select_client).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Annuler", 
                command=cancel_selection).pack(side=tk.LEFT, padx=5)
        
        # Double-click to select
        def on_double_click(event):
            select_client()
        
        result_tree.bind('<Double-1>', on_double_click)
        
        # Bind Enter key to search
        def on_enter(event):
            search_clients()
        
        search_id_entry.bind('<Return>', on_enter)
        search_nom_entry.bind('<Return>', on_enter)
        search_prenom_entry.bind('<Return>', on_enter)
        
        # Focus on first search field
        search_id_entry.focus_set()
        
        # Load initial data (all clients from page 1)
        dialog.after(100, lambda: load_clients(1))
        
        # Wait for dialog to close
        dialog.wait_window()
        
        return selected_client

    def fetch_clients_from_api_with_pagination(self, client_id="", nom="", prenom="", page=1, per_page=200):
        """Fetch clients from the API with search parameters and pagination info"""
        try:
            api_url = "https://app.diardzair.com.dz/api/commandes"
            
            # Build parameters
            params = {
                "page": page,
                "perPage": per_page
            }
            headers={
                "Authorization": "Bearer f8peRksDOtpBRE6UAoJhC6kP3gPg5JOUJVsi9fhsJCn8sBgjE6C/2rUo3PEYCmYG"
            }
            # Add search parameters if provided
            if client_id:
                params["id"] = client_id
            if nom:
                params["nom"] = nom
            if prenom:
                params["prenom"] = prenom
            
            print(f"Fetching clients from API with params: {params}")
            
            # Make API call
            response = requests.get(api_url, params=params, headers=headers, timeout=10, verify=False)
            response.raise_for_status()
            
            data = response.json()
            print(f"API response: {data}")
            
            if data.get('error', True):
                raise Exception("API returned error")
            
            clients = data.get('data', [])
            total_clients = len(clients)  # For simplicity, assume all data is returned
            total_pages = max(1, (total_clients + per_page - 1) // per_page)
            
            return {
                'clients': clients,
                'total_clients': total_clients,
                'total_pages': total_pages,
                'current_page': page
            }
            
        except Exception as e:
            print(f"Error fetching clients from API: {e}")
            return {
                'clients': [],
                'total_clients': 0,
                'total_pages': 1,
                'current_page': 1
            }

    def fetch_client_details_from_api(self, client_id):
        """Fetch detailed client information from API"""
        try:
            api_url = f"https://albaraka.fun/api/orders/info/{client_id}"
            headers = {
                "Authorization": "Bearer U3wXgPLvreiyv5JRJsxVU4Tlbyakt7MLFzTjWq8DaPjLbGXgbMELK5xsMRqvHcOtc0H2obwVK4OGqJbfsgIo2hgakbxi5Sk4mWRKv1IOYr42qtOiDiyd3f8fexCLe9m"
            }
            
            response = requests.get(api_url, headers=headers, timeout=10, verify=False)
            response.raise_for_status()
            
            data = response.json()
            
            if data.get('error', True):
                raise Exception("Client details API returned error")
            
            return data.get('data', {})
            
        except Exception as e:
            print(f"Error fetching client details: {e}")
            return {}

def main():
    """Main function to run the application"""
    app = OrderPrepareApp()
    app.run()


if __name__ == "__main__":
    main()