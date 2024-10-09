# Word_Translator
A Python script for extracting, splitting, and translating text from Word documents.

## Project Description

This project reads text from Word documents, splits the content into individual sentences, and translates them using Google Translate. Additionally, users can choose to enable ChatGPT for translation, allowing them to compare the results between Google Translate and ChatGPT side by side. This tool is useful for users who want quick translations of Word documents with the option to see results from two different translation engines.

---

## Project Structure

- **/Word_Translator**
  - **src**
    - `word_translator.py` - The main script responsible for reading Word files, splitting sentences, and performing translations using Google Translate and ChatGPT.
    - `excel_formatter.py` - A utility script that helps format Excel files (optional, based on project needs).
    - `requirements.txt` - Lists the dependencies required to run the project.
  - `ForWindowsOpen.bat` - A Windows batch file for easy execution of the translator script on Windows systems.
  - `README.md` - This file, which explains the project’s functionality, usage, and structure.
  
---

## Installation

To install and run the project, follow these steps:

1. Clone or download the repository.
2. Install the necessary dependencies with the following command:
   ```bash
   pip install -r requirements.txt
3. Ensure that you have a .docx Word file ready for translation.
4. Run the word_translator.py script:
  ```bash
  python3 word_translator.py
  ```

---

## Usage
When you run the script, select the Word file you want to translate.
Choose whether to translate using only Google Translate or both Google Translate and ChatGPT.
The script will generate both translation outputs (if ChatGPT is enabled) and display the results.
Optionally, you can also format Excel files using the excel_formatter.py script.

---

## Requirements
The project uses the following Python packages:

openai
python-docx
deep-translator
pandas
inquirer
PyGetWindow
openpyxl

To install these dependencies, run:
  ```bash
  pip install -r requirements.txt
  ```

---


## Batch File (Windows)
For easier execution on Windows, a ForWindowsOpen.bat file is provided. Simply double-click the batch file to start the translator.

---

## Contact
For any questions or issues, feel free to reach out:

## Project Owner: [Halil Şafak Şimşek]
Email: [halil_tafak@hotmail.com]
