## This powershellscript will list all .doc and .docx files on a specified location local or remote 
## and then for each file open it and search for a specific string
##

$SearchText = 'White Rabbit'
$Word = New-Object -ComObject Word.Application
$docs = Get-ChildItem "\\sharename\subfolder\*" -Recurse -Include "*.doc*" | Select -Expand FullName  # .doc* gives makes the search go through both doc and docx
$counter = 1

ForEach($doc in $docs){
    Write-Progress -Activity "Checking: $doc" -Status "File $counter of $($docs.count)" -PercentComplete ($counter*100/$docs.count)
    try {
    $Document = $Word.Documents.Open($doc)
    If($Document.Content.Find.Execute("$SearchText")){
        $Document.Close($true)
        Write-Host "Term found in $doc!"

    }
    $Document.close($true)

    }
    catch {
    }
    $counter++
}
