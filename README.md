# Word_Translator
A Python script for extracting, splitting, and translating text from Word documents.

## Project Description

This project reads text from Word documents, splits the content into individual sentences, and translates them using Google Translate. Additionally, users can choose to enable ChatGPT for translation, allowing them to compare the results between Google Translate and ChatGPT side by side. This tool is useful for users who want quick translations of Word documents with the option to see results from two different translation engines.

---

## Project Structure

- **/Word_Translator**
  - **`word_translator.py`** - The main script that handles Word document processing and translation.
  - **`excel_formatter.py`** - A script for formatting Excel files if needed.
  - **`ForWindowsOpen.bat`** - A batch file for easy execution on Windows systems.
  - **`requirements.txt`** - List of necessary dependencies to run the project.
  - **`README.md`** - Project documentation.

---

## Installation

To install and run the project, follow these steps:

1. Clone or download the repository.
2. Install the necessary dependencies with the following command:
   ```bash
   pip install -r requirements.txt
Ensure that you have a .docx Word file ready for translation.
Run the word_translator.py script:
  ```bash
  python3 word_translator.py

Usage
When you run the script, select the Word file you want to translate.
Choose whether to translate using only Google Translate or both Google Translate and ChatGPT.
The script will generate both translation outputs (if ChatGPT is enabled) and display the results.
Optionally, you can also format Excel files using the excel_formatter.py script.
