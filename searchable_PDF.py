"""
Site où j'ai trouvé le code : https://medium.com/@kate19058/using-python-to-generate-searchable-pdfs-d869130006dc
lien pour installer Poppler sur windows : https://github.com/oschwartz10612/poppler-windows/releases/

Ouvrez le lien ci-dessus dans votre navigateur.
Sous la section "Assets", cliquez sur le fichier .zip correspondant à votre architecture (par exemple, poppler-xx.x.x-x86_64.zip pour 64 bits).
Téléchargez et extrayez le contenu du fichier ZIP dans un répertoire de votre choix, par exemple C:\poppler.
Une fois extrait, vous pourrez trouver le dossier bin à l'intérieur du dossier Poppler, et c'est ce chemin que vous devrez utiliser dans votre script Python comme poppler_path.

"""



# Importation des bibliothèques nécessaires

# Bibliothèques générales
import os  # Pour manipuler les chemins de fichiers et les répertoires
import re  # Pour les expressions régulières
import glob  # Pour récupérer les fichiers correspondant à un motif
import io  # Pour gérer les flux de données en mémoire
import pandas as pd  # Pour manipuler les données tabulaires (bien que non utilisé ici)

# Bibliothèques pour la création de PDFs interrogeables
import PyPDF2  # Pour lire et écrire des fichiers PDF
import pytesseract  # Pour l'OCR (reconnaissance optique de caractères)
from pdf2image import convert_from_path  # Pour convertir des pages PDF en images

# Bibliothèque pour convertir les images en noir et blanc (réduction d'espace)
from PIL import Image   

# Définir les chemins nécessaires

# Chemin vers Poppler (utilisé pour convertir les PDF en images)
poppler_path = r'C:\Poppler\Release-24.07.0-0\poppler-24.07.0\Library\bin' 

# Chemin vers l'exécutable de Tesseract OCR
pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe' 

# Définir le répertoire d'entrée (dossier contenant les PDF non interrogeables)
input_folder = r'H:\Desktop\Safran\Suite-Code\pdf_tools_ocr-sekou\table_dataframe\non_searchable_PDF'

# Définir le répertoire de sortie (dossier où les PDF interrogeables seront sauvegardés)
output_folder = r'H:\Desktop\Safran\Suite-Code\pdf_tools_ocr-sekou\table_dataframe\searchable_PDF'

# Obtenir la liste des fichiers déjà présents dans le répertoire de sortie
# Cela permet de ne pas traiter les fichiers déjà convertis
existing_output_files = [os.path.basename(file) for file in glob.glob(os.path.join(output_folder, '*.pdf'))]

# Boucle pour parcourir tous les fichiers PDF dans le dossier d'entrée
for pdf_file in glob.glob(os.path.join(input_folder, '*.pdf')):
    # Extraire le nom de base du fichier (sans le chemin complet)
    base_filename = os.path.basename(pdf_file)

    # Vérifier si le fichier de sortie existe déjà pour éviter de le traiter à nouveau
    if base_filename in existing_output_files:
        print(f"Skipping '{base_filename}' as it already exists in the output folder.")
        continue  # Passer au fichier suivant

    # Convertir le PDF en images (une image par page)
    images = convert_from_path(pdf_file, poppler_path=poppler_path)

    # Créer un objet `PdfWriter` pour écrire les pages PDF interrogeables
    pdf_writer = PyPDF2.PdfWriter()

    # Boucle pour ajouter chaque image comme une nouvelle page dans le PDF interrogeable
    for image in images:
        # Utiliser Tesseract pour extraire le texte et générer un fichier PDF pour l'image
        page = pytesseract.image_to_pdf_or_hocr(image, extension='pdf')
        
        # Lire le fichier PDF généré à partir de l'image en mémoire
        pdf = PyPDF2.PdfReader(io.BytesIO(page))
        
        # Ajouter la page extraite au document final
        pdf_writer.add_page(pdf.pages[0])

    # Définir le nom du fichier de sortie dans le dossier de sortie
    output_file = os.path.join(output_folder, base_filename)

    # Écrire le PDF interrogeable dans le fichier de sortie
    with open(output_file, 'wb') as f:
        pdf_writer.write(f)

    # Indiquer que le fichier a été traité et sauvegardé
    print(f"'{base_filename}' processed and saved to '{output_file}'.")
