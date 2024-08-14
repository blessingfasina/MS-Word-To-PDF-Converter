import os
import time
from concurrent.futures import ThreadPoolExecutor
from tkinter import Tk, filedialog
from docx2pdf import convert
from tqdm import tqdm

def select_folder(prompt_text):
    """
    Opens a dialog for the user to select a folder.
    """
    root = Tk()
    root.withdraw()  # Hide the root window
    folder_path = filedialog.askdirectory(title=prompt_text)
    root.destroy()
    return folder_path

def convert_single_file(input_output_paths):
    input_path, output_path = input_output_paths
    try:
        convert(input_path, output_path)
    except Exception as e:
        print(f"Failed to convert {input_path}. Error: {e}")

def convert_documents(input_dir, output_dir):
    """
    Converts all Word documents in the input directory to PDF in the output directory.
    """
    files_to_convert = [
        (os.path.join(input_dir, f), os.path.join(output_dir, os.path.splitext(f)[0] + '.pdf'))
        for f in os.listdir(input_dir)
        if f.lower().endswith(('.docx', '.doc')) and not f.startswith('~$')  # Exclude temp files
    ]

    with ThreadPoolExecutor() as executor:
        list(tqdm(executor.map(convert_single_file, files_to_convert),
                  total=len(files_to_convert), desc="Converting Files", unit="file"))

    print(f"\nConversion completed! PDFs are saved in '{output_dir}'.")

def show_welcome_message():
    print("***********************************************")
    print("*                                             *")
    print("*         MS Word to PDF Converter            *")
    print("*                                             *")
    print("*            Developed by: Blessing Fasina    *")
    print("*            Twitter: @genius_ng              *")
    print("*            GitHub: https://github.com/blessingfasina *")
    print("*                                             *")
    print("***********************************************")
    time.sleep(5)  # Pause for 5 seconds

def main():
    show_welcome_message()

    input_dir = select_folder("Select the folder containing Word documents")
    if not input_dir:
        print("No input folder selected. Exiting...")
        return

    output_dir = select_folder("Select the folder to save PDFs")
    if not output_dir:
        print("No output folder selected. Exiting...")
        return

    convert_documents(input_dir, output_dir)

if __name__ == "__main__":
    main()
