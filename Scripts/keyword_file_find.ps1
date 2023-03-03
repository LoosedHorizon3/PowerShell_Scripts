# This script allows a user to search for specific documents containing a keyword by specifying a UserID and searching within two possible paths. The script then 
# searches through the specified paths for documents with certain extensions (pdf, docx, xlsx, etc.), and opens each file to search for the specified keyword. 
# If a document contains the keyword, the script writes the file path to the console. The script also displays a progress bar to indicate how many documents have 
# been searched so far, and whether or not the script has finished.
# Author: Nicholas Stevenson
# Created: Friday, 24 Feburary 2023

##################################################################################################################################################################


Write-Host "Z drive is [name of mapped share]" -ForegroundColor DarkCyan

# Prompt user to enter the UserID and keyword to search for
$userID = Read-Host "Enter the UserID"
$keyword = Read-Host "Enter keyword here"

# Define paths to search for documents
$userPath = 'Z:\Users\'
$userPath2 = 'Z:\Users2\'

$source = $userPath+$userID
$source2 = $userPath2+$userID

# Determine which path to search based on whether the specified UserID exists in $userPath
if(Test-Path $source){
    $docs = Get-ChildItem -Recurse -Path $source |
    Where-Object {$_.Extension -match '\.(pdf|docx|docm|doc|dctx|dotm|dot|odt|xlsx|xlsm|xlsb|xls|csv|xltx|xltm|xlt|ods|pptx|ppt|potx|potm|pot|ppsx|ppsm|pps)$'}
}
else{
    $docs = Get-ChildItem -Recurse -Path $source2 |
    Where-Object {$_.Extension -match '\.(pdf|docx|docm|doc|dctx|dotm|dot|odt|xlsx|xlsm|xlsb|xls|csv|xltx|xltm|xlt|ods|pptx|ppt|potx|potm|pot|ppsx|ppsm|pps)$'}
}

# Count the number of documents to search
$totalCount = $docs.Count
$progressCount = 0

# Loop through each document and search for the keyword
foreach($doc in $docs){
    # Try to open and search the document, and catch any errors
    try{
        if($doc.Extension -match '\.pdf$'){
            $content = Get-Content $doc.FullName -Raw
            if($content | Select-String $keyword){
                Write-Host "$($doc.FullName) contains the word '$keyword'" -ForegroundColor Green
            }
        }
        elseif($doc.Extension -match '\.(docx|docm|doc|dctx|dotm|dot|odt)$'){
            # Search Word documents using the Word COM object
            $word = New-Object -ComObject Word.Application
            $document = $word.Documents.Open($doc.FullName)
            if($document.Content.Find.Execute($keyword)){
                Write-Host "$($doc.FullName) contains the word '$keyword'" -ForegroundColor Green
            }
            $document.Close()
            if($word){
                try{
                    $word.Quit()
                }
                catch{
                    Write-Host "Error occurred while trying to quit Word: $_" -ForegroundColor Red
                }
            }
        }
        elseif($doc.Extension -match '\.(xlsx|xlsm|xlsb|xls|csv|xltx|xltm|xlt|ods)$'){
            # Search Word documents using the Excel COM object
            $excel = New-Object -ComObject Excel.Application
            $workbook = $excel.Workbooks.Open($doc.FullName)
            foreach($worksheet in $workbook.Worksheets) {
                if($worksheet.Cells.Find($keyword)){
                    Write-Host "$($doc.FullName) contains the word '$keyword'" -ForegroundColor Green
                    break
                }
            }
            $workbook.Close($false)
            if($excel){
                try{
                    $excel.Quit()
                }
                catch{
                    Write-Host "Error occurred while trying to quit Excel: $_" -ForegroundColor Red
                }
            }
        }
        elseif($doc.Extension -match '\.(pptx|ppt|potx|potm|pot|ppsx|ppsm|pps)$') {
            # Search Word documents using the Powerpoint COM object
            $powerpoint = New-Object -ComObject PowerPoint.Application
            $presentation = $powerpoint.Presentations.Open($doc.FullName)
            foreach ($slide in $presentation.Slides) {
                if ($slide.Shapes.TextFrame.TextRange.Find($keyword))
                {
                    Write-Host "$($doc.FullName) contains the word '$keyword'" -ForegroundColor Green
                    break
                }
            }
            $presentation.Close()
            if($powerpoint){
                try{
                    $powerpoint.Quit()
                }
                catch{
                    Write-Host "Error occurred while trying to quit PowerPoint: $_" -ForegroundColor Red
                }
            }
        }
    }
    # The catch block is used to catch any errors that occur while the script is trying to search through the documents. If an error is caught,
    # the script will write a message to the console that indicates which document caused the error.
    catch{
        Write-Host "Error opening or searching document: $($doc.FullName)" -ForegroundColor Red
    }
    # The finally block is used to ensure that the Microsoft Office applications used by the script (Word, Excel, and PowerPoint) are closed even if an error occurs.

    # After each document is processed, the $progressCount variable is incremented by one to keep track of how many documents have been searched so far.
    $progressCount++
    # The $percentComplete variable is then calculated as a percentage of the total number of documents
    $percentComplete = ($progressCount / $totalCount) * 100
    # Write-Progress cmdlet is used to display a progress bar indicating the current progress of the script
    Write-Progress -Activity "Searching documents" -PercentComplete $percentComplete -Status "$progressCount of $totalCount documents searched"
}
    # This will tell you that the script has now finished
    Write-Host "Script has now been completed"  -ForegroundColor Cyan