# -*- coding: utf-8 -*-
"""
Test script pour valider le parsing des données QR
"""

import sys
import os
sys.path.append(os.path.dirname(__file__))

from qr_scanner import QRScannerApp, ProductData, SortieData, RetourData

def test_legacy_parsing():
    """Test le parsing du format legacy depuis scan.txt"""
    
    # Données de test depuis scan.txt
    test_data_1 = """*VMSDZ06CUKI191698*
MOTOCYCLE CUKI -I-
CUKI
bleu nuit/ blanc
Unitª Oued-Ghir
CUKI I 06/2025"""

    test_data_2 = """*VMSDZ06CUKI191858*
MOTOCYCLE CUKI -II-
CUKI
ICON BLEU XMAX
Unitª Oued-Ghir
CUKI II 07/2025"""

    # Créer une instance de l'app pour tester
    app = QRScannerApp()
    
    print("=== TEST PARSING LEGACY FORMAT ===\n")
    
    # Test pour type Entrée
    app.data_type = "Entrée"
    
    print("1. Test données QR - Entrée CUKI191698:")
    product1 = app.parse_qr_data(test_data_1)
    print(f"   Reference: '{product1.Reference}'")
    print(f"   Designation: '{product1.Designation}'")
    print(f"   Fournisseur: '{product1.Fournisseur}'")
    print(f"   Couleur: '{product1.Couleur}'")
    print(f"   Magasin: '{product1.Magasin}'")
    print(f"   Num_Chasse: '{product1.Num_Chasse}'")
    print(f"   Lot: '{product1.Lot}'")
    print()
    
    print("2. Test données QR - Entrée CUKI191858:")
    product2 = app.parse_qr_data(test_data_2)
    print(f"   Reference: '{product2.Reference}'")
    print(f"   Designation: '{product2.Designation}'")
    print(f"   Fournisseur: '{product2.Fournisseur}'")
    print(f"   Couleur: '{product2.Couleur}'")
    print(f"   Magasin: '{product2.Magasin}'")
    print(f"   Num_Chasse: '{product2.Num_Chasse}'")
    print(f"   Lot: '{product2.Lot}'")
    print()
    
    # Test pour type Sortie
    app.data_type = "Sortie"
    
    print("3. Test données QR - Sortie CUKI191858:")
    sortie1 = app.parse_qr_data(test_data_2)
    print(f"   Date: '{sortie1.Date}'")
    print(f"   Heure: '{sortie1.Heure}'")
    print(f"   DESIGNATION: '{sortie1.DESIGNATION}'")
    print(f"   N_CHASSIS: '{sortie1.N_CHASSIS}'")
    print(f"   ID_CLIENT: '{sortie1.ID_CLIENT}'")
    print(f"   NOM_PRENOM: '{sortie1.NOM_PRENOM}'")
    print(f"   WILAYA: '{sortie1.WILAYA}'")
    print()
    
    # Test pour type Retour
    app.data_type = "Retour"
    
    print("4. Test données QR - Retour CUKI191858:")
    retour1 = app.parse_qr_data(test_data_2)
    print(f"   Date: '{retour1.Date}'")
    print(f"   Heure: '{retour1.Heure}'")
    print(f"   DESIGNATION: '{retour1.DESIGNATION}'")
    print(f"   N_CHASSIS: '{retour1.N_CHASSIS}'")
    print(f"   ID_CLIENT: '{retour1.ID_CLIENT}'")
    print(f"   NOM_PRENOM: '{retour1.NOM_PRENOM}'")
    print(f"   WILAYA: '{retour1.WILAYA}'")
    print()
    
    # Test génération QR
    print("=== TEST GENERATION QR ===\n")
    
    print("4. Génération QR depuis ProductData:")
    qr_data_product = app.generate_qr_data(product1)
    print("   Données générées:")
    for i, line in enumerate(qr_data_product.split('\n')):
        print(f"   Ligne {i+1}: '{line}'")
    print()
    
    print("5. Génération QR depuis SortieData:")
    qr_data_sortie = app.generate_qr_data(sortie1)
    print("   Données générées:")
    for i, line in enumerate(qr_data_sortie.split('\n')):
        print(f"   Ligne {i+1}: '{line}'")
    print()
    
    print("6. Génération QR depuis RetourData:")
    qr_data_retour = app.generate_qr_data(retour1)
    print("   Données générées:")
    for i, line in enumerate(qr_data_retour.split('\n')):
        print(f"   Ligne {i+1}: '{line}'")
    print()
    
    print("=== TESTS TERMINÉS ===")

if __name__ == "__main__":
    test_legacy_parsing()
