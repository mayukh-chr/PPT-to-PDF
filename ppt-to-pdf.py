import os
from win32com import client
from pathlib import Path
import pywintypes

def convert_to_pdf(input_folder):
    powerpoint = None
    word = None
    
    try:
        powerpoint = client.Dispatch("Powerpoint.Application")
        word = client.Dispatch("Word.Application")
        
        files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.ppt', '.pptx', '.doc', '.docx'))]
        
        for file in files:
            input_path = os.path.join(input_folder, file)
            output_path = os.path.join(input_folder, f"{Path(file).stem}.pdf")
            
            try:
                if file.lower().endswith(('.ppt', '.pptx')):
                    presentation = powerpoint.Presentations.Open(input_path)
                    try:
                        presentation.SaveAs(output_path, 32)  # 32 is the PDF format code
                        print(f"Converted {file} to PDF successfully.")
                    finally:
                        presentation.Close()
                elif file.lower().endswith(('.doc', '.docx')):
                    document = word.Documents.Open(input_path)
                    try:
                        document.SaveAs(output_path, FileFormat=17)  # 17 is the PDF format code
                        print(f"Converted {file} to PDF successfully.")
                    finally:
                        document.Close()
            except Exception as e:
                print(f"Error converting {file}: {str(e)}")
    
    finally:
        if powerpoint:
            try:
                powerpoint.Quit()
            except:
                print("Error quitting PowerPoint application")
        
        if word:
            try:
                word.Quit()
            except:
                print("Error quitting Word application")

folder_path = r"D:\School\Manipal\Resources\Sem 5\FML\Na Slides"
convert_to_pdf(folder_path)