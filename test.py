from docx import Document
import re
import os
import shutil  # Importer shutil pour le déplacement de fichiers

# Fonction pour identifier si une ligne contient une date
def contains_date(text):
    return bool(re.search(r'\d{1,2} \w+ \d{4}', text))

# Fonction pour extraire le nombre juste avant la parenthèse fermante et l'incrémenter
def increment_filename(file_name):
    # Recherche d'un nombre juste avant la parenthèse fermante
    match = re.search(r'(\d+)\s*\)', file_name)
    if match:
        # Récupérer le nombre, l'incrémenter de 1
        num = int(match.group(1)) + 1
        # Remplacer l'ancien nombre par le nouveau
        return re.sub(r'(\d+)\s*\)', f'{num})', file_name)
    return file_name

# Chemin vers votre fichier .docx
file_path = r"C:\Users\natha\Desktop\vs\aa_(Unwind 34).docx"

# Ouvrir le document Word
doc = Document(file_path)
i = 0

# Parcourir les tables du document
for table in doc.tables:
    for row in table.rows:
        first_column_text = row.cells[0].text
        # Si la première colonne contient "Issue Date", on s'arrête
        i += 1
        if "Issue Date" in first_column_text:
            cellul = row.cells[1].text 
            new_line = 'USD 16,000,000 as of 5 April 2023'

            # On découpe le texte d'origine en lignes
            lines = cellul.split('\n')

            # Trouver l'index de la dernière ligne contenant une date
            last_date_index = -1
            for i, line in enumerate(lines):
                if contains_date(line):
                    last_date_index = i

            # Insertion de la nouvelle ligne après la dernière ligne contenant une date
            if last_date_index != -1:
                lines.insert(last_date_index + 1, new_line)

            # Reconstruire le texte mis à jour
            updated_row_text = '\n'.join(lines)

            print(updated_row_text)
            row.cells[1].text = updated_row_text
            print("lol")
            break

# Extraire le nom du fichier sans son extension
file_name = os.path.basename(file_path)
file_name_without_ext = os.path.splitext(file_name)[0]

# Incrémenter le nombre dans le nom du fichier
new_file_name = increment_filename(file_name_without_ext)

# Nouveau chemin pour enregistrer le fichier modifié
new_file_path = os.path.join(os.path.dirname(file_path), f"{new_file_name}.docx")

# Sauvegarder le fichier avec les modifications
doc.save(new_file_path)
print(f"Fichier enregistré sous {new_file_path}")

# Déplacer le fichier vers le nouvel emplacement
destination_path = r"C:\Users\natha\Desktop\Perso\uber eats"
shutil.move(new_file_path, os.path.join(destination_path, f"{new_file_name}.docx"))
print(f"Fichier déplacé vers {os.path.join(destination_path, f'{new_file_name}.docx')}")
