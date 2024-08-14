# MS Word to PDF Batch Converter

This tool converts one or multiple MS Word documents (`.docx` or `.doc`) into PDF format efficiently and quickly. It features a user-friendly interface, a progress bar, and a welcome message displayed in the terminal.

## Features

- **Batch Conversion**: Convert multiple Word documents to PDF in one go.
- **Speed**: Multi-threaded processing for faster conversions, especially with large files.
  
## Getting Started

Follow these steps to set up and run the tool on your local machine.

### Prerequisites

- Python 3.6 or higher
- `pip` (Python package installer)

### Installation

1. **Clone the Repository**

    ```bash
    git clone https://github.com/blessingfasina/ms-word-to-pdf-converter.git
    cd ms-word-to-pdf-converter
    ```

2. **Install the Required Packages**

    Navigate to the directory containing the `requirements.txt` file and install the dependencies:

    ```bash
    pip install -r requirements.txt
    ```

### Usage

1. **Run the Tool**

    After installing the required packages, run the tool by executing the following command:

    ```bash
    python converter.py
    ```

2. **Step-by-Step Workflow**

    - **Welcome Message**: Upon running the tool, a welcome message will be displayed in the terminal for 5 seconds, introducing the developer.
    - **Select Input Folder**: A dialog box will prompt you to select the folder containing the Word documents you wish to convert.
    - **Select Output Folder**: After selecting the input folder, you will be prompted to choose a destination folder where the converted PDFs will be saved.
    - **Conversion**: The tool will then start converting the Word documents to PDFs, showing a progress bar to keep you informed of the process.
    - **Completion**: Once done, the PDFs will be available in the output folder.

### Contributing

Contributions are welcome! If you have suggestions for improvements or find a bug, feel free to fork the repository and submit a pull request.

### Author

- **Blessing Fasina**
  - **Twitter**: [@genius_ng](https://twitter.com/genius_ng)
  - **GitHub**: [blessingfasina](https://github.com/blessingfasina)

### License

This project is licensed under the MIT License. See the `LICENSE` file for more details.

---

### Final Notes

This tool is designed to be simple yet powerful, providing a seamless experience for converting Word documents to PDF. Whether you're dealing with a single file or an entire directory of documents, this tool has you covered.
