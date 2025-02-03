import os
import shutil
import win32com.client as win32
from pathlib import Path

class QuantumShield:
    def __init__(self, source_folder, destination_folder):
        self.source_folder = Path(source_folder)
        self.destination_folder = Path(destination_folder)
        self.ensure_directories_exist()
        
    def ensure_directories_exist(self):
        if not self.source_folder.exists():
            self.source_folder.mkdir(parents=True, exist_ok=True)
        if not self.destination_folder.exists():
            self.destination_folder.mkdir(parents=True, exist_ok=True)
        
    def transfer_files(self):
        for file in self.source_folder.iterdir():
            if file.is_file():
                shutil.move(str(file), self.destination_folder)
                print(f"Transferred {file.name} to {self.destination_folder}")

    def open_in_word(self, filename):
        word = win32.Dispatch('Word.Application')
        word.Visible = True
        doc_path = self.destination_folder / filename
        if doc_path.exists():
            word.Documents.Open(str(doc_path))
            print(f"Opened {filename} in Microsoft Word")
        else:
            print(f"{filename} does not exist in the destination folder")

    def list_files(self):
        files = [f.name for f in self.destination_folder.iterdir() if f.is_file()]
        print("Files in destination folder:")
        for file in files:
            print(file)

if __name__ == "__main__":
    # Example usage
    source = r"C:\Users\YourUsername\Documents\Source"
    destination = r"C:\Users\YourUsername\Documents\Destination"

    qs = QuantumShield(source, destination)
    qs.transfer_files()
    qs.list_files()
    qs.open_in_word('example.docx')  # Replace 'example.docx' with your actual file name