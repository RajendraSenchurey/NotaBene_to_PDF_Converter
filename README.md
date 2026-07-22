# Nota Bene to PDF Converter

This repository contains a PowerShell script that automates the batch conversion of Nota Bene (.nb) files into PDF format.

## Requirements
- Windows
- Nota Bene 15 installed

## Usage
1. Download the PowerShell script.
2. Open the script in a text editor.
3. Change the folder path on line 3 to the folder containing your `.nb` files.
4. Open Windows PowerShell.
5. If PowerShell blocks the script, run:

```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

6. Run the script.

The script will search the specified folder and its subfolders for `.nb` files and create PDF versions.

## Author
Rajendra Senchurey  
University of Arizona  
Research Assistant to Professor Emeritus J. Ronald Engel, PhD
