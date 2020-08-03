## This powershellscript will list all .xls and .xlsx files on a specified location local or remote 
## and then for each file open the file, go through cells vertical and horizontal from column a-z and search for 
## a specific string


$SearchText = 'White Rabbit'
$Excel = New-Object -ComObject Excel.Application
$Files = Get-ChildItem "\\sharename\subfolder\*" -Recurse -Include "*.xls*" | Select -Expand FullName
# .xls* gives makes the search list both .xls and .xlsx
$counter = 1


ForEach($File in $Files){
    Write-Progress -Activity "Checking: $file" -Status "File $counter of $($files.count)" -PercentComplete ($counter*100/$files.count)
    try {
    $Workbook = $Excel.Workbooks.Open($File)
    If($Workbook.Sheets.Item(1).Range("A:Z").Find($SearchText)){
        $Workbook.Close($false)
        Write-Host "Term found in $file!"
    }
    $workbook.close($false)
    }
    catch {
    }
    $counter++
}
