#Step 1: Since Windows blocks custom scripts by default, run this one first to allow execution:

Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser



#Step 2: Set your main parent folder here. It will search this folder and all subfolders inside it and converts all .nb files. Copy everything below this:

$folderPath = "C:\Users\Rajendra\Desktop\Test_Engel Manuscripts"

#Get all .nb files in the folder and all nested subfolders (-Recurse)
$nbFiles = Get-ChildItem -Path $folderPath -Filter "*.nb" -Recurse

if ($nbFiles.Count -eq 0) {
    Write-Host "No .nb files found in '$folderPath' or its subfolders." -ForegroundColor Red
    return
}

Write-Host "Found $($nbFiles.Count) .nb files across all subfolders. Starting conversion..." -ForegroundColor Cyan

foreach ($file in $nbFiles) {
    #Gets the exact directory of the specific file being processed
    $targetFolder = $file.DirectoryName
    $pdfPath      = [System.IO.Path]::ChangeExtension($file.FullName, ".pdf")
    
    Write-Host "Processing: $($file.Name) [Folder: $targetFolder]" -ForegroundColor Yellow

    try {
        #1. Read raw text with ANSI / CP1252 encoding
        $bytes = [System.IO.File]::ReadAllBytes($file.FullName)
        $text  = [System.Text.Encoding]::GetEncoding(1252).GetString($bytes)

        #2. Strip Nota Bene formatting codes, tags, and control characters
        $text = $text -replace '[\x00-\x09\x0B\x0C\x0E-\x1F\x7F-\x9F]', ''
        $text = $text -replace '®[A-Z0-9\+\-]+\¯', ''
        $text = $text -replace '[®¯]', ''
        $text = $text -replace '[ｮｯョッツヨ]', ''
        
        # Clean up specific Mojibake / encoding artifacts
        $text = $text -replace '鍍he', '"the'
        $text = $text -replace '的', '"'
        $text = $text -replace '痴', "'s"

        # 3. Create temporary UTF-8 text file right in the file's current directory
        $tempTxt = Join-Path $targetFolder "$($file.BaseName)_temp.txt"
        [System.IO.File]::WriteAllText($tempTxt, $text, [System.Text.Encoding]::UTF8)

        # 4. Open temp text in Word quietly and export directly to PDF
        $word = New-Object -ComObject Word.Application
        $word.DisplayAlerts = 0
        $word.Visible = $false

        $doc = $word.Documents.Open($tempTxt, $false, $true)
        $doc.ExportAsFixedFormat($pdfPath, 17) # 17 = wdExportFormatPDF
        $doc.Close(0)
        $word.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($word) | Out-Null

        #Delete the temporary text file
        Remove-Item $tempTxt -Force

        Write-Host "  [OK] Created: $($file.BaseName).pdf" -ForegroundColor Green
    }
    catch {
        Write-Host "  [ERROR] Failed to convert $($file.Name): $_" -ForegroundColor Red
    }
}

Write-Host "`nAll done! Every .nb file in all folders/subfolders has been converted to PDF!" -ForegroundColor Cyan
