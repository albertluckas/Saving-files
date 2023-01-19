

#interface w/ outlook
Add-Type -assembly "Microsoft.Office.Interop.Outlook"
add-type -assembly "System.Runtime.Interopservices"
try
{
$outlook = [Runtime.Interopservices.Marshal]::GetActiveObject('Outlook.Application')
}
catch
{
    try
    {
        $Outlook = New-Object -comobject Outlook.Application
    }
    catch
    {
        write-host "You must exit Outlook first."
        exit
    }
}

$namespace = $Outlook.GetNameSpace("MAPI")
#get emails from target inbox, target folder, with target phrase in the subject line
$subjectComparisonExpression="*Trigger phrase*"
$inbox =$namespace.Folders.Item('inbox name').Folders.Item('folder name')
$Emails=$inbox.Items| Where-Object {$_.Subject -like $subjectComparisonExpression}

#Names of the locations in the emails
$namesIn = @("Main office", "remote workers", "EU offices", "Canadian offices", "field workers")
#Desired names of the files
$namesOut = @( "US.xlsx", "RM.xlsx", "EU.xlsx", "CA.xlsx", "FD.xlsx")
#number of files found for each locale
$filesFound=@(0,0,0,0,0)

#Name of folder to save files to
$target="D:\"
$target+=Get-Date -format "yyyMMdd"

#make the target folder if it dosnt exist. will only make the lowest level folder.
if(!(Test-Path $target)){
md $target
}





$Emails |foreach{
    $s = $_.Subject




    for ($i = 0; $i -lt $namesIn.count; $i++){
        #iterate through all the names
        if($s -like "*" + $namesIn[$i] + "*"){
            #when we find the right one,count the file as found and save it.
            filesFound[i]+=1
            $_.Attachments|foreach{
                $_.saveasfile($target+$namesOut[i])
            }
        }
    }



}
#find all files that are missing or duplicated. if there are any, make a report to a txt file and open it
$issues="        "
for ($i = 0; $i -lt $namesIn.count; $i++){
    if($filesFound[$i] -ne 1){
        $issues+=$namesIn[$i]+" had "+$filesFound[$i]+" files
        "
    }
}
if($issues -ne "        "){
    $issues|Out-File $target'\issues.txt'
    start $target"\issues.txt"
}
