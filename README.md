# Excel Image Downloader Inserter v2

**Excel Image Downloader Inserter v2** is an enhanced Python script that automates the task of downloading images from URLs stored in an Excel file and embedding those images into the Excel sheet. The script allows users to:
1. Select an Excel file.
2. Select columns containing image URLs and filenames.
3. Download images from URLs.
4. Embed those images into the Excel file in a new column.
5. Save the updated Excel file with embedded images.

---

### Features:
- **Excel File Handling:** Easily interact with Excel files using `openpyxl`.
- **Automatic Image Downloading:** Download images from URLs and save them locally.
- **Image Embedding:** Embed the downloaded images directly into Excel cells.
- **Retry on Error:** Automatically retries downloading images in case of internet connection issues.
- **User-Friendly Interface:** Prompts the user to select the required columns via a simple graphical interface.

---

### Requirements:
To run this script, make sure you have Python installed along with the following libraries:
- `requests`: To download the images from the web.
- `openpyxl`: To interact with Excel files and embed images.
- `tkinter`: For a graphical interface to select files and columns.
- `Pillow`: For Fatch image objects.
You can install these dependencies using `pip`:

```bash
pip install requests openpyxl Pillow
```

Note: `tkinter` is usually included with Python installations, but if it's missing, you may need to install it separately.

---

### How It Works:
1. **User Input**: The script will prompt you to select the Excel file that contains the image URLs and the corresponding column for filenames.
2. **Download Images**: The script will download each image from the URL in the selected column.
3. **Embed Images**: The images are embedded in a new column in the Excel file.
4. **Save**: The updated Excel file with embedded images is saved with the name `<original_filename>_with_images.xlsx`.

---

### Usage:

1. **Clone the repository:**

```bash
git clone https://github.com/jahirulislamdms/excel-image-downloader-inserter-v2.git
cd excel-image-downloader-inserter-v2
```

2. **Run the script:**

Make sure you have installed the necessary libraries, and run the script:

```bash
python download_images_and_embed_in_excel.py
```

3. **User Interaction:**
   - Select the Excel file using the graphical file dialog.
   - Choose the column number containing the image URLs.
   - Choose the column number containing the filenames to save images.
   - The images will be downloaded and embedded in a new column in the Excel sheet.
   - The updated Excel file will be saved in the same directory with the suffix `_with_images`.

---

### Example Workflow:

- Open the file selection dialog to select your Excel file.
- You'll see a list of columns and their names with corresponding numbers.
- Select the column number that contains the image URLs.
- Select the column number that contains the filenames for the images.
- The script will download the images, save them locally, and embed them in the Excel sheet.
- The updated file will be saved in the same directory as `<original_filename>_with_images.xlsx`.

---

### Example Output:

```plaintext
No file selected.
Invalid image URL column selected.
Invalid naming column selected.
Downloaded & Saved: example_image.jpg
Excel file saved: /path/to/your/excel/file_with_images.xlsx
```

---

### Error Handling:
- **ConnectionError:** The script will retry the download if there's a connection error (e.g., no internet).
- **RequestException:** If there’s an issue with the URL or a request problem, it skips that URL and continues.
- **HTTP Errors:** If the server returns a non-200 status code, it logs and skips the URL.

---

### License:
This repository is licensed under the MIT License. See the [LICENSE](LICENSE) file for more information.

---

### Contributions:
Feel free to fork this repository, make changes, and submit pull requests! All contributions are welcome.

---

## Folder Structure:

```
excel-image-downloader-inserter-v2/
├── downloaded_images/       # Folder where downloaded images will be stored
├── download_images_and_embed_in_excel.py  # Main script file
└── README.md                # Documentation file
```

---

### To Upload to GitHub:

1. **Initialize the Repository Locally:**

```bash
git init
```

2. **Add Files to Git:**

```bash
git add download_images_and_embed_in_excel.py
git add README.md
git add .gitignore
git commit -m "Initial commit: Upload version 2 of Excel Image Downloader Inserter"
```

3. **Push to GitHub:**

Create a new repository on GitHub (if not already done) and link it to your local repository:

```bash
git remote add origin https://github.com/jahirulislamdms/excel-image-downloader-inserter-v2.git
git branch -M main
git push -u origin main
```

---

### Notes:
- This script is designed to work with Excel files (.xlsx).
- The directory `downloaded_images` is created automatically in the same folder as the script to store all downloaded images.
- Ensure the URLs in the Excel file are direct image links (e.g., ending with `.jpg`, `.png`, etc.).

---

### Contact:
If you have any questions or issues, feel free to reach out via the [GitHub Issues](https://github.com/jahirulislamdms/excel-image-downloader-inserter-v2/issues) page.

--- 
