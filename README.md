
# AutoFill DL OCR

This project automates the extraction of text from driver’s license images using **Google Cloud Vision OCR** and logs the parsed details into an **Excel file**.  
It also includes a **watcher script** that automatically detects new images dropped into a folder, processes them, and appends the results into the spreadsheet.

Live video of the project here: <https://devpost.com/software/autofill-dl-ocr>

---

## ✨ Features
- Extracts **First Name, Last Name, DOB, Address, City, State, and Zip Code** from license images.
- Appends results to an **Excel sheet (`data.xlsx`)** with a `View Image` link.
- Supports **macOS** (tested), and also works on Linux/Windows.
- Auto-processes new files dropped into the target folder using a **file system watcher**.

---

## 📂 Project Structure

The following files are **NOT included** in the repo for security/privacy:
- `DLOCR.json` → your **Google Vision credentials**
- `data.xlsx` → the Excel sheet generated at runtime
- image samples (driver’s licenses)

---

## 🚀 Setup Instructions

### 1. Clone or Download
You can either:
- Clone with SSH:
  ```bash
  git clone git@github.com:ZishanVahora1/AutoFill-DL-OCR.git
(RECOMMENDED) Just download the repo as a .ZIP file

## Install Dependencies

Make sure you have Python 3.8+ installed.
Install dependencies with:

pip install google-cloud-vision openpyxl pandas watchdog pillow

## Add Your Google Vision Credentials

Download your service account JSON file from the Google Cloud Console
.

Place it inside the project folder (e.g., Excel_Test/DLOCR.json).

Set the environment variable (replace /Users/yourname/ with your own Mac username):

export GOOGLE_APPLICATION_CREDENTIALS="/Users/yourname/Desktop/Excel_Test/DLOCR.json"

You must do this step each time you open a new terminal, unless you add it permanently to your shell config (~/.zshrc or ~/.bashrc).


## Run the Watcher

Navigate into the folder:

cd ~/Desktop/Excel_Test
python3 watch_and_process.py

## ✅ Example Workflow

Start watcher → python3 watch_and_process.py

Drop in DL_Test.jpg

Open data.xlsx → see extracted details + clickable “View Image” link

## 📌 Requirements

Python 3.8+

Google Cloud Vision API enabled

Excel (or LibreOffice) to view data.xlsx

MacOS (Possible on Windows but instructions for MacOS only)
