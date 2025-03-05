import os
import win32com.client

def change_author_in_directory(directory, new_author):
    shell = win32com.client.Dispatch("Shell.Application")
    namespace = shell.Namespace(directory)

    for filename in os.listdir(directory):
        file_path = os.path.join(directory, filename)

        if filename.endswith(('.docx', '.xlsx', '.pptx')):
            try:
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
    directory = r"C:\Users\Maxence\Documents\IT\Python\ENV_changement_auteur"
    
    new_author = "Nouvel Auteur"

    change_author_in_directory(directory, new_author)
