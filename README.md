# Automated Word Document Segregation for Press Releases

This project provides a Python-based solution to automate the segregation of Word documents (specifically press releases) into three distinct sections: **Intro (S1)**, **Main Body (S2)**, and **About the Company (S3)**. The code processes a folder of Word documents and extracts these sections based on predefined rules, saving each section into separate Word files.

## Features

- **S1 (Intro)**: Extracts the introductory paragraph. If a 'DATELINE:' label is found, it prioritizes the first paragraph after the 'DATELINE:' that meets the minimum length condition. Otherwise, it extracts the first paragraph that meets the length criteria.
- **S2 (Main Body)**: Extracts all content between the S1 and S3 sections (if S3 exists). If no 'About' section is found (S3), it extracts all paragraphs after S1 until the last paragraph that meets the minimum length.
- **S3 (About the Company)**: Extracts paragraphs starting with "About" and all subsequent paragraphs, collecting them until another paragraph meeting the length condition is found.
- **Empty File Handling**: If no valid content is found for S1, S2, or S3, the respective output files are created but left empty.

## Folder Structure

```plaintext
project-folder/
│
├── input-folder-path/           # Folder containing the Word (.docx) files to be processed.
│   ├── file1.docx
│   ├── file2.docx
│   └── ...
│
├── output-folder-path/          # Folder where the extracted S1, S2, and S3 files will be saved.
│   ├── file1_S1.docx
│   ├── file1_S2.docx
│   ├── file1_S3.docx
│   └── ...
│
├── main.py                      # Main script that processes the documents.
└── README.md                    # Project documentation.
```

## How It Works

1. The script reads each `.docx` file in the specified input folder.
2. It processes each document to extract the three sections:
    - **S1 (Intro)**: Extracts the first paragraph of sufficient length (150 characters by default). If 'DATELINE:' is found, S1 starts after that.
    - **S2 (Main Body)**: Extracts all paragraphs between the S1 and S3 sections. If no S3 exists, it extracts paragraphs until the end of the document.
    - **S3 (About the Company)**: Extracts paragraphs that start with "About" and collects all following paragraphs.
3. Each extracted section is saved as a separate `.docx` file in the output folder, named using the format `<original_filename>_S1.docx`, `<original_filename>_S2.docx`, and `<original_filename>_S3.docx`.

### S1 Extraction
- **Dateline Priority**: If a 'DATELINE:' label is found, S1 will be the first paragraph that meets the length condition after 'DATELINE:'.
- **Fallback**: If no 'DATELINE:' is present, it extracts the first paragraph that meets the minimum length.

### S3 Extraction
- S3 is extracted by searching for paragraphs starting with "About" and collecting all following paragraphs until another paragraph meeting the length condition is found.

### S2 Extraction
- If S3 exists, S2 is all content between S1 and S3.
- If S3 does not exist, S2 is all content after S1 until the last paragraph meeting the length condition.

## Installation and Usage

1. **Requirements**: Install the necessary Python libraries.
    ```bash
    pip install python-docx
    ```

2. **Run the script**:
    ```bash
    python main.py
    ```

    Modify the `input_folder_path` and `output_folder_path` in the script to your folder paths.

3. **Example Usage**:

    ```python
    input_folder_path = "input-folder-path"
    output_folder_path = "output-folder-path"
    min_length = 150
    process_folder(input_folder_path, output_folder_path, min_length)
    ```

    This will process all `.docx` files in `input-folder-path` and save the extracted S1, S2, and S3 sections into the `output-folder-path`.

## Configuration

- **Minimum Length (`min_length`)**: This parameter defines the minimum number of characters required to consider a paragraph valid for extraction. By default, this is set to 150 characters. You can change this value based on your document structure.
  
  Example:
  ```python
  min_length = 100  # Change minimum length
  ```

## File Naming Convention

- For each input `.docx` file, three output files will be generated (if content is available) with the following naming pattern:
    - `<original_filename>_S1.docx`: Contains the Intro section.
    - `<original_filename>_S2.docx`: Contains the Main Body.
    - `<original_filename>_S3.docx`: Contains the About the Company section.

## Limitations

- The script assumes the presence of certain patterns (e.g., 'DATELINE:', 'About') in the document. Documents without these patterns may not produce accurate sectioning.
- The script may not handle complex document structures (e.g., multi-column layouts) perfectly.
  
## License

This project is licensed under the MIT License. See the [LICENSE](LICENSE) file for details.
```
