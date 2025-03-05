import os
import win32com.client

def change_author_in_directory(directory, new_author):
    # Initialiser l'application COM pour interagir avec les fichiers Office
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(directory)

    # Parcourir tous les fichiers dans le répertoire
    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        # Vérifier si le fichier est un fichier Office (Word, Excel, PowerPoint)
        if filename.endswith(('.docx', '.xlsx', '.pptx')):
            try:
                # Ouvrir le fichier Office
                if filename.endswith('.docx'):
                    app = win32com.client.Dispatch("Word.Application")
                    doc = app.Documents.Open(file_path)
                    doc.BuiltInDocumentProperties("Author").Value = new_author
                    doc.Save()
                    doc.Close()
                    app.Quit()
                elif filename.endswith('.xlsx'):
                    app = win32com.client.Dispatch("Excel.Application")
                    workbook = app.Workbooks.Open(file_path)
                    workbook.BuiltinDocumentProperties("Author").Value = new_author
                    workbook.Save()
                    workbook.Close()
                    app.Quit()
                elif filename.endswith('.pptx'):
                    app = win32com.client.Dispatch("PowerPoint.Application")
                    presentation = app.Presentations.Open(file_path, WithWindow=False)
                    presentation.BuiltInDocumentProperties("Author").Value = new_author
                    presentation.Save()
                    presentation.Close()
                    app.Quit()

                print(f"Auteur modifié pour : {filename}")
            except Exception as e:
                print(f"Erreur lors du traitement de {filename} : {e}")
        else:
            print(f"{filename} n'est pas un fichier Office supporté.")

if __name__ == "__main__":
    # Spécifie le répertoire contenant les fichiers Office
    directory = r"C:\Users\Maxence\Documents\IT\Python\ENV_changement_auteur"
    
    # Spécifie le nouvel auteur (défini une seule fois)
    new_author = "Nouvel Auteur"

    # Appeler la fonction pour changer l'auteur
    change_author_in_directory(directory, new_author)