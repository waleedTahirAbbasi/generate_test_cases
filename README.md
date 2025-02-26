# 📘 Guide: Running the generate_test_cases.py Script on Windows PowerShell

This guide provides step-by-step instructions for setting up and running the generate_test_cases.py script to generate test cases from an Excel file using **Windows PowerShell**.

---

## 📌 Prerequisites

### 1️⃣ Install Python (if not installed)
- Download and install Python from the official website:  
  - Link: https://www.python.org/downloads/windows/
- Ensure you check the box **"Add Python to PATH"** during installation.

### 2️⃣ Install Required Dependencies
Open **Windows PowerShell** and run:

```
pip install google-generativeai pandas openpyxl xlsxwriter
```

- google-generativeai → Used for AI-generated test cases  
- pandas → Handles reading and writing Excel files  
- openpyxl → Reads .xlsx input files  
- xlsxwriter → Writes .xlsx output files  

---

## 🔑 How to Get the Gemini API Key

To use the Gemini AI model, you need an API key from Google.

1. **Go to Google AI Studio**  
   - Link: https://aistudio.google.com

2. **Sign in with your Google account**  

3. **Generate an API Key**  
   - Click on **"Get API Key"** (or go to **Settings > API Keys**)  
   - Copy the API key  

4. **Store the API Key Securely**  
   - You must set this key in your environment to run the script  

---

## 🔐 Set Up Your Gemini API Key

### ✅ Option 1: Set API Key as an Environment Variable (Recommended)
1. Open **Windows PowerShell**  
2. Run this command (replace "your_api_key_here" with your actual API key)  

   ```
    $env:GEMINI_API_KEY="your_api_key_here"
   ```

3. Restart your PowerShell session  

### 🛠 Option 2: Use a .env File (Alternative)
1. Create a new file named .env in the script’s folder  
2. Add this line inside the .env file  

   ```
   GEMINI_API_KEY=your_api_key_here
   ```

3. Update the script to **load the API key from .env**  

---

## 📂 Preparing Input File

- The **input file must be an Excel file (.xlsx)**  
- The **first row** must contain these **headers**  

```
  User Story ID | User Story  
  --------------|------------  
  ABC-123      | As an employee, I want to submit...  
  XYZ-456      | As an admin, I want to approve...  
```

---

## 🚀 Running the Script

### Basic Usage (Auto-Named Output File)
```
python generate_test_cases.py user_stories.xlsx
```

**Output file will be automatically named**  
Test_Cases_DDMMYYYYHHmmss.xlsx  

### Custom Output File Name
```
python generate_test_cases.py user_stories.xlsx my_test_cases.xlsx
```

Make sure the file name **ends with .xlsx**, or the script will show an error.  

---

## 📌 Troubleshooting

### ❌ "Python Not Found" When Running the Script
- Try running:  

  ```
  Get-Command python
  ```

- If Python is not recognized, **reinstall Python** and ensure **"Add to PATH"** is selected.  

### ❌ "No module named X" Error
- Run the dependency installation again:  

  ```
  pip install google-generativeai pandas openpyxl xlsxwriter
  ```

### ❌ "API Key Not Found"
- Ensure you set the API key as an **environment variable** or use a .env file.  

---

## ✅ Summary

1. **Install Python & dependencies**  
2. **Get your Gemini API key from Google AI Studio**  
3. **Store the API key in an environment variable or .env file**  
4. **Prepare an Excel file with "User Story ID" & "User Story" columns**  
5. **Run the script via Windows PowerShell**  

---


