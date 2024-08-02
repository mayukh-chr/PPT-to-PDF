import os
from win32com import client
from pathlib import Path
import pywintypes

def convert_ppt_to_pdf(input_folder):
    powerpoint = client.Dispatch("Powerpoint.Application")
   
    try:
        ppt_files = [f for f in os.listdir(input_folder) if f.lower().endswith(('.ppt', '.pptx'))]
       
        for ppt_file in ppt_files:
            input_path = os.path.join(input_folder, ppt_file)
            output_path = os.path.join(input_folder, f"{Path(ppt_file).stem}.pdf")
           
            try:
                presentation = powerpoint.Presentations.Open(input_path)
                try:
                    presentation.SaveAs(output_path, 32)  # 32 is the PDF format code
                    print(f"Converted {ppt_file} to PDF successfully.")
                except Exception as e:
                    print(f"Error saving {ppt_file} as PDF: {str(e)}")
            except Exception as e:
                print(f"Error opening {ppt_file}: {str(e)}")

    finally:
        try:
            # Close all open presentations
            for presentation in powerpoint.Presentations:
                try:
                    presentation.Close()
                except:
                    print(f"Failed to close a presentation")
            powerpoint.Quit()
        except:
            print("Error quitting PowerPoint application")

folder_path = r"sample-path"
convert_ppt_to_pdf(folder_path)
