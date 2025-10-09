#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Updater Stub - A lightweight updater for Mouvement Stock Application
This program handles the update process when the main application is closed.
"""

import sys
import os
import zipfile
import shutil
import subprocess
import tkinter as tk
from tkinter import ttk
import threading
import time


class UpdaterStubWindow:
    def __init__(self, zip_path, install_dir, main_exe_path):
        self.zip_path = zip_path
        self.install_dir = install_dir
        self.main_exe_path = main_exe_path
        
        # Create the progress window
        self.root = tk.Tk()
        self.root.title("Installation de la Mise à Jour")
        self.root.geometry("400x200")
        self.root.resizable(False, False)
        
        # Center the window
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() // 2) - (self.root.winfo_width() // 2)
        y = (self.root.winfo_screenheight() // 2) - (self.root.winfo_height() // 2)
        self.root.geometry(f"+{x}+{y}")
        
        # Prevent closing during update
        self.root.protocol("WM_DELETE_WINDOW", lambda: None)
        
        # Make it stay on top
        self.root.attributes('-topmost', True)
        
        self.setup_ui()
        
    def setup_ui(self):
        """Setup the updater UI"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill='both', expand=True)
        
        # Title
        title_label = ttk.Label(main_frame, text="Mise à jour en cours...", 
                               font=('Arial', 14, 'bold'), foreground='#2c3e50')
        title_label.pack(pady=(0, 20))
        
        # Status label
        self.status_label = ttk.Label(main_frame, text="Préparation de l'installation...", 
                                     font=('Arial', 10))
        self.status_label.pack(pady=(0, 10))
        
        # Progress bar
        self.progress_bar = ttk.Progressbar(main_frame, mode='indeterminate', length=300)
        self.progress_bar.pack(pady=(0, 20))
        self.progress_bar.start(10)
        
        # Warning label
        warning_label = ttk.Label(main_frame, 
                                 text="⚠️ Ne fermez pas cette fenêtre pendant l'installation",
                                 font=('Arial', 9), foreground='#e67e22')
        warning_label.pack()
        
    def update_status(self, message, color='#2c3e50'):
        """Update the status message"""
        self.status_label.config(text=message, foreground=color)
        self.root.update()
        
    def _safe_remove_file(self, file_path, max_attempts=5):
        """Safely remove a file with multiple attempts"""
        for attempt in range(max_attempts):
            try:
                if os.path.exists(file_path):
                    os.remove(file_path)
                return
            except PermissionError:
                if attempt < max_attempts - 1:
                    time.sleep(1)  # Wait before retry
                    continue
                else:
                    raise
    
    def _safe_remove_directory(self, dir_path, max_attempts=5):
        """Safely remove a directory with multiple attempts"""
        for attempt in range(max_attempts):
            try:
                if os.path.exists(dir_path):
                    shutil.rmtree(dir_path)
                return
            except PermissionError:
                if attempt < max_attempts - 1:
                    time.sleep(1)  # Wait before retry
                    continue
                else:
                    raise

    def perform_update(self):
        """Perform the actual update process"""
        try:
            # Step 0: Wait for main application to fully close
            self.update_status("Attente de la fermeture complète de l'application...")
            time.sleep(3)  # Give time for main app to close and release file locks
            
            # Step 1: Verify ZIP file exists
            self.update_status("Vérification du fichier de mise à jour...")
            if not os.path.exists(self.zip_path):
                raise Exception(f"Fichier de mise à jour introuvable: {self.zip_path}")
            
            time.sleep(0.5)
            
            # Step 2: Create backup directory
            self.update_status("Création de la sauvegarde...")
            backup_dir = os.path.join(self.install_dir, "_backup_" + str(int(time.time())))
            if not os.path.exists(backup_dir):
                os.makedirs(backup_dir)
            
            # Backup critical files (main exe and _internal folder)
            main_exe_name = os.path.basename(self.main_exe_path)
            if os.path.exists(self.main_exe_path):
                shutil.copy2(self.main_exe_path, os.path.join(backup_dir, main_exe_name))
            
            internal_dir = os.path.join(self.install_dir, "_internal")
            if os.path.exists(internal_dir):
                shutil.copytree(internal_dir, os.path.join(backup_dir, "_internal"))
            
            time.sleep(0.5)
            
            # Step 3: Extract new version
            self.update_status("Extraction de la nouvelle version...")
            temp_extract_dir = os.path.join(self.install_dir, "temp_update_extract")
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir)
            
            with zipfile.ZipFile(self.zip_path, 'r') as zip_ref:
                zip_ref.extractall(temp_extract_dir)
            
            time.sleep(0.5)
            
            # Step 4: Find source directory in extracted files
            self.update_status("Localisation des fichiers source...")
            source_dir = None
            
            # Check for "Mouvement Stock" folder first
            mouvement_stock_dir = os.path.join(temp_extract_dir, "Mouvement Stock")
            if os.path.exists(mouvement_stock_dir):
                source_dir = mouvement_stock_dir
            elif os.path.exists(os.path.join(temp_extract_dir, "_internal")):
                source_dir = temp_extract_dir
            else:
                source_dir = temp_extract_dir
            
            time.sleep(0.5)
            
            # Step 5: Replace files with better PyInstaller support
            self.update_status("Installation des nouveaux fichiers...")
            self.progress_bar.config(mode='determinate', maximum=100, value=0)
            
            # Special handling for PyInstaller structure
            items_to_copy = []
            
            # Get list of items from source, excluding the updater directory
            for item in os.listdir(source_dir):
                # Skip the updater directory to avoid self-modification
                if item.lower() == "updater":
                    print(f"Skipping updater directory: {item}")
                    continue
                    
                source_path = os.path.join(source_dir, item)
                items_to_copy.append((source_path, item))
            
            total_items = len(items_to_copy)
            
            for i, (source_path, item) in enumerate(items_to_copy):
                dest_path = os.path.join(self.install_dir, item)
                
                try:
                    if os.path.isdir(source_path):
                        # Special handling for _internal directory (PyInstaller)
                        if item == "_internal" and os.path.exists(dest_path):
                            # Remove existing _internal completely before copying
                            self.update_status(f"Remplacement du dossier {item}...")
                            self._safe_remove_directory(dest_path)
                        elif os.path.exists(dest_path):
                            self._safe_remove_directory(dest_path)
                        
                        # Copy directory
                        shutil.copytree(source_path, dest_path)
                        
                    else:
                        # Regular file copy
                        # For .exe files, ensure they're not in use
                        if item.endswith('.exe') and os.path.exists(dest_path):
                            self.update_status(f"Remplacement de {item}...")
                            self._safe_remove_file(dest_path)
                        
                        shutil.copy2(source_path, dest_path)
                    
                    # Update progress
                    progress = int((i + 1) / total_items * 100)
                    self.progress_bar.config(value=progress)
                    self.root.update()
                    
                except Exception as copy_error:
                    print(f"Warning: Could not copy {item}: {copy_error}")
                    # For critical files, this might be a problem
                    if item == "_internal" or item.endswith('.exe'):
                        raise Exception(f"Failed to copy critical file {item}: {copy_error}")
                    # Continue with other files
            
            time.sleep(1)
            
            # Step 6: Cleanup
            self.update_status("Nettoyage...")
            self.progress_bar.config(mode='indeterminate')
            self.progress_bar.start(10)
            
            # Remove temporary extraction directory
            if os.path.exists(temp_extract_dir):
                shutil.rmtree(temp_extract_dir)
            
            # Remove downloaded ZIP file
            if os.path.exists(self.zip_path):
                os.remove(self.zip_path)
            
            # Remove old backup if installation was successful
            if os.path.exists(backup_dir):
                shutil.rmtree(backup_dir)
            
            time.sleep(0.5)
            
            # Step 7: Success
            self.update_status("Installation terminée avec succès!", '#28a745')
            self.progress_bar.stop()
            self.progress_bar.config(value=100)
            
            time.sleep(1)
            
            # Step 8: Relaunch main application
            self.update_status("Redémarrage de l'application...", '#2980b9')
            time.sleep(1)
            
            # Launch the main application
            if os.path.exists(self.main_exe_path):
                subprocess.Popen([self.main_exe_path], cwd=self.install_dir)
            else:
                print(f"Warning: Main executable not found at {self.main_exe_path}")
            
            # Close the updater
            self.root.destroy()
            
        except Exception as e:
            # Handle errors
            self.update_status(f"Erreur: {str(e)}", '#dc3545')
            self.progress_bar.stop()
            
            # Try to restore backup
            try:
                if 'backup_dir' in locals() and os.path.exists(backup_dir):
                    self.update_status("Restauration de la sauvegarde...", '#f39c12')
                    
                    # Restore main exe
                    backup_exe = os.path.join(backup_dir, os.path.basename(self.main_exe_path))
                    if os.path.exists(backup_exe):
                        shutil.copy2(backup_exe, self.main_exe_path)
                    
                    # Restore _internal folder
                    backup_internal = os.path.join(backup_dir, "_internal")
                    target_internal = os.path.join(self.install_dir, "_internal")
                    if os.path.exists(backup_internal):
                        if os.path.exists(target_internal):
                            shutil.rmtree(target_internal)
                        shutil.copytree(backup_internal, target_internal)
                    
                    self.update_status("Sauvegarde restaurée", '#f39c12')
                    
            except Exception as restore_error:
                print(f"Could not restore backup: {restore_error}")
            
            # Enable closing the window
            self.root.protocol("WM_DELETE_WINDOW", self.root.destroy)
            
            # Show error message and allow manual restart
            error_frame = ttk.Frame(self.root)
            error_frame.pack(pady=10)
            
            ttk.Button(error_frame, text="Fermer", 
                      command=self.root.destroy).pack(side=tk.LEFT, padx=5)
            
            if os.path.exists(self.main_exe_path):
                ttk.Button(error_frame, text="Redémarrer l'App", 
                          command=lambda: self.manual_restart()).pack(side=tk.LEFT, padx=5)
    
    def manual_restart(self):
        """Manually restart the application"""
        try:
            subprocess.Popen([self.main_exe_path], cwd=self.install_dir)
            self.root.destroy()
        except Exception as e:
            print(f"Could not restart application: {e}")
    
    def run(self):
        """Run the updater"""
        # Start the update process in a separate thread
        threading.Thread(target=self.perform_update, daemon=True).start()
        
        # Run the Tkinter main loop
        self.root.mainloop()


def main():
    """Main function for the updater stub"""
    if len(sys.argv) != 4:
        print("Usage: updater_stub.exe <zip_path> <install_dir> <main_exe_path>")
        sys.exit(1)
    
    zip_path = sys.argv[1]
    install_dir = sys.argv[2]
    main_exe_path = sys.argv[3]
    
    print(f"Updater Stub starting...")
    print(f"ZIP Path: {zip_path}")
    print(f"Install Dir: {install_dir}")
    print(f"Main EXE: {main_exe_path}")
    
    # Create and run the updater window
    updater = UpdaterStubWindow(zip_path, install_dir, main_exe_path)
    updater.run()


if __name__ == "__main__":
    main()
