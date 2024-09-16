#les trois types de PDFs : https://pdf.abbyy.com/fr/learning-center/pdf-types/"""

import fitz

def check_pdf_type(pdf_path: str) -> str:
    """Détermine le type de fichier PDF : numérique, scanné, ou mixte.

    Args:
        pdf_path (str): Le chemin du fichier PDF à analyser.

    Returns:
        str: Le type du PDF identifié, soit "PDF numérique", "PDF interrogeable (mixte)",
             "PDF scanné" ou "Inconnu" si aucun type n'a pu être déterminé.

    Raises:
        FileNotFoundError: Si le fichier PDF n'est pas trouvé ou ne peut être ouvert.
    """
    
    # Ouverture du fichier PDF en utilisant la bibliothèque `fitz` (PyMuPDF)
    doc = fitz.open(pdf_path)
    
    # Initialisation des indicateurs pour la présence de texte et d'images
    has_text = False
    has_images = False

    # Parcourir toutes les pages du PDF pour analyser la présence de texte ou d'images
    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)  # Charge la page actuelle
        
        # Récupère le texte de la page
        text = page.get_text()
        if text.strip():  # Vérifie si du texte significatif est présent
            has_text = True
        
        # Récupère la liste des images présentes sur la page
        image_list = page.get_images(full=True)
        if image_list:  # Si des images sont présentes, active le flag `has_images`
            has_images = True
    
    # Déterminer le type de PDF en fonction de la présence de texte et/ou d'images
    if has_text and not has_images:
        return "PDF numérique"  # Contient uniquement du texte
    elif has_text and has_images:
        return "PDF interrogeable (mixte)"  # Contient du texte et des images
    elif not has_text and has_images:
        return "PDF scanné"  # Contient uniquement des images (probablement scanné)
    else:
        return "Inconnu"  # Aucun contenu identifiable
    
# Test du script avec un fichier PDF spécifique
pdf_path = r"H:\Desktop\Safran\Suite-Code\pdf_tools_ocr-sekou\typePDF\testFred.pdf"

# Appel de la fonction pour déterminer le type du PDF
pdf_type = check_pdf_type(pdf_path)

# Affichage du résultat
print(f"Le type du PDF est : {pdf_type}")

