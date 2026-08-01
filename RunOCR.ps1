$FineReaderExe = "C:\Program Files\ABBYY FineReader 16\finereaderocr.exe"
$InputFolder   = "Z:\01 Download"

$files = @(
    Get-ChildItem -LiteralPath $InputFolder -Filter "*.pdf" -File |
        Sort-Object Name
)

if ($files.Count -eq 0) {
    Write-Host "No PDF files found in: $InputFolder"
    exit 0
}

foreach ($file in $files) {
    Write-Host "Processing: $($file.FullName)"

    $process = Start-Process `
        -FilePath $FineReaderExe `
        -ArgumentList @(
            "`"$($file.FullName)`""
            "/lang"
            "English"
            "/send"
            "PDFViewer"
        ) `
        -Wait `
        -PassThru

    if ($process.ExitCode -ne 0) {
        Write-Warning "FineReader returned exit code $($process.ExitCode) for $($file.Name)"
    }
    else {
        Write-Host "Completed: $($file.Name)"
    }
}

Write-Host "All PDF files have been processed."
