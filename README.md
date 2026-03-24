# 🚀 SmartLED Automation Script (Parallel Processing)

## 📌 Project Overview

This project is a **Python-based automation tool** that automates the Luminary Installation process on the SmartLED portal.

It reads data from an Excel file and automatically fills web forms using Selenium, reducing manual effort and increasing speed significantly.

---

## ⚙️ Features

* 🌐 Automated login to SmartLED portal
* 📊 Reads Excel data dynamically
* 🔍 Validates Luminary Serial Number before submission
* ⚡ Parallel processing (multi-threading for speed)
* 🔁 Auto retry on failure (robust execution)
* 📝 Logs errors in `fault_log.txt`
* 📁 Adds remarks for each processed row
* ⏱️ Tracks script version & timestamp

---

## 🛠️ Technologies Used

* Python 🐍
* Selenium (Web Automation)
* pandas (Data Processing)
* numpy
* tkinter (File Picker GUI)
* threading & concurrent.futures

---

## 📂 Project Structure

```
📁 Project Folder
 ├── script.py
 ├── fault_log.txt
 ├── input.xlsx (user selected)
 ├── final.xlsx (output file)
```

---

## ▶️ How to Use

### Step 1: Install dependencies

```bash
pip install selenium pandas numpy tqdm
```

---

### Step 2: Run the script

```bash
python script.py
```

---

### Step 3: Select Excel file

* A file picker will open
* Select your Excel dataset

---

## 🔄 Workflow

1. Opens SmartLED portal
2. Logs in automatically
3. Reads each row from Excel
4. Matches Luminary Serial No
5. Fills:

   * District
   * Block
   * Panchayat
   * Network
6. Submits form
7. Records success/failure in Excel

---

## 📥 Output

* Final file saved as:

```
final.xlsx
```

* Includes:

  * Remarks column (status)
  * Script version
  * Timestamp

---

## 🚀 Performance

* Uses **multi-threading (3 workers)** ⚡
* Handles large datasets efficiently
* Automatic retry on failure

---

## ⚠️ Important Notes

* Ensure stable internet connection
* Website structure should not change

---

## 👨‍💻 Author

**Akshay Kumar Sharma**
B.S Computer Science & Data Analytics | IIT Patna

---

## 📌 Future Improvements

* GUI dashboard
* Dynamic worker scaling
* Better error reporting
* CAPTCHA handling

---

## ⭐ If you like this project

Give it a ⭐ on GitHub!
