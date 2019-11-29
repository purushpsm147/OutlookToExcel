#Params
$Account = "hkannan@techsoup.org"
$Folder = "Inbox"


#specify the month and year to get the monthly timesheet
#later We can get the current mont and year automatically
#$RegexSubjMatch = "Offshore Work Status -\sNov\s\d+,\s2019"



#Create outlook COM object to search folders
$Outlook = New-Object -ComObject Outlook.Application
$OutlookNS = $Outlook.GetNamespace("MAPI")



#Word Com object
#$Word=NEW-Object –comobject Word.Application
#$Word.Visible = $True  # setting visibility to true


#Excel Com object
$Excel=new-object -ComObject Excel.Application
$Excel.Visible = $True  # setting visibility to true




#adding and selecting new doc page
#$doc = $word.documents.Add()
#$myDoc = $word.Selection


#adding and and selecting new excel page
$xl = $Excel.Workbooks.Add()
$myXl = $Excel.Selection
$myXlSheet = $Excel.Worksheets.Item("sheet1")
$myXlSheet.activate()


#Get all emails from specific account and folder
$AllEmails = $OutlookNS.Folders.Item($Account).Folders.Item($Folder).Items



#Filter to emails with date range
$ReportsEmails = $AllEmails | where { ($_.ReceivedTime -ge [datetime]"11-01-2019" -AND $_.ReceivedTime -lt ` [datetime]"11-26-2019") -and ($_.SenderName -match "Siva Murugaperumal" )}
# (($_.Subject -match $RegexSubjMatch) -or ($_.Subject -match "E-Scrum"))



$LatestReportEmail = $ReportsEmails | Sort ReceivedTime


$row2 = 2
$row1 = 2

foreach ($mail in $LatestReportEmail){
$mailbody = $mail | Select-Object -Property Body | Format-List
$bodytext = $mailbody | Out-String
$bodytext = $bodytext |  ForEach-Object {$_.Trim() -replace "\s+", "_" }
$bodytext = $bodytext.Replace("Body_:_Hi_All,_Good_Morning!_Off-shore_scrum_status:_","")
$bodytext = [regex]::Replace($bodytext,"Blocker.*Off-Shore_Team","")

$TaskArray = [System.Collections.ArrayList]@()
$hourslist = [System.Collections.ArrayList]@()


foreach ($task in $bodytext) {
  # Split the list element by '/' (if present) and output the 1st token.
 $TaskArray.AddRange([regex]::Split($task,"_\d_"))
}

$hourslist = $bodytext | Select-String "(?<!\d)\d(?!\d)" -AllMatches | Foreach {$_.Matches | Foreach {$_.Value}}

$myXlSheet.Cells.Item(1,2) = "Tasks"
$myXlSheet.Cells.Item(1,3) = "Hours"



foreach($item in $TaskArray){
 $myXlSheet.Cells.Item($row1,2) = $item.Replace("_"," ") 
 $row1++
}



foreach($item in $hourslist){
 $myXlSheet.Cells.Item($row2,3) = $item
 $row2++
}


#$myDoc.TypeText($bodytext)
#$myDoc.TypeParagraph()
#$myDoc.TypeParagraph()

}










#$Outlook.Quit()
#$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
#[gc]::Collect()
#[gc]::WaitForPendingFinalizers()
#Remove-Variable word

#$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Excel)
#[gc]::Collect()
#[gc]::WaitForPendingFinalizers()
#Remove-Variable Excel
