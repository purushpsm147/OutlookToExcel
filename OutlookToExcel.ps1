#Params
$Account = "YourMail@account.com"
$Folder = "Inbox"

#specify the month and year to get the monthly timesheet
#later We can get the current mont and year automatically
#$RegexSubjMatch = "Offshore Work Status -\sNov\s\d+,\s2019"



#Create outlook COM object to search folders
$Outlook = New-Object -ComObject Outlook.Application
$OutlookNS = $Outlook.GetNamespace("MAPI")



#Word Com object
$Word=NEW-Object –comobject Word.Application
$Word.Visible = $True  # setting visibility to true


#Excel Com object
$Excel=new-object -ComObject Excel.Application
$Excel.Visible = $True  # setting visibility to true




#adding and selecting new doc page
$doc = $word.documents.Add()
$myDoc = $word.Selection


#adding and and selecting new excel page
$xl = $Excel.Workbooks.Add()
#$myXl = $Excel.Selection
$myXlSheet = $Excel.Worksheets.Item("sheet1")
$myXlSheet.activate()


#Get all emails from specific account and folder
$AllEmails = $OutlookNS.Folders.Item($Account).Folders.Item($Folder).Items



#Filter to emails with date range
$ReportsEmails = $AllEmails | where { ($_.ReceivedTime -ge [datetime]"11-01-2019" -AND $_.ReceivedTime -lt ` [datetime]"11-30-2019") -and ($_.SenderName -match "Siva Murugaperumal" -or $_.SenderName -match "Hari Kannan" -or $_.SenderName -match "Puviyarasu Ganesan" -or $_.Subject -match "E-Scrum" )}
# (($_.Subject -match $RegexSubjMatch) -or ($_.Subject -match "E-Scrum"))


#sorting based on received time
$LatestReportEmail = $ReportsEmails | Sort ReceivedTime


#row variables
$row2 = 2
$row1 = 2


foreach ($mail in $LatestReportEmail){
#getting mail date
$mailDate = $mail | Select-Object -Property ReceivedTime 

#getting mail body and replacing unwanted characters
$mailbody = $mail | Select-Object -Property Body | Format-List
$bodytext = $mailbody | Out-String
$bodytext = $bodytext |  ForEach-Object {$_.Trim() -replace "\s+", "_" }
$bodytext = $bodytext.Replace("Body_:_Hi_All,_Good_Morning!_Off-shore_scrum_status:_","")
$bodytext = [regex]::Replace($bodytext,"Blocker.*Off-Shore_Team","")


#arraylist to store hours and body details
$TaskArray = [System.Collections.ArrayList]@()
$hourslist = [System.Collections.ArrayList]@()


foreach ($task in $bodytext) {
  # Split the list element by '_\d_' (a number) (if present) and output the 1st token.
 $TaskArray.AddRange([regex]::Split($task,"_\d_"))
}

$hourslist = $bodytext | Select-String "(?<!\d)_\d_(?!\d)" -AllMatches | Foreach {$_.Matches | Foreach {$_.Value}}

$myXlSheet.Cells.Item(1,1) = "Date"
$myXlSheet.Cells.Item(1,2) = "Tasks"
$myXlSheet.Cells.Item(1,3) = "Hours"


$myDoc.TypeText($mailDate.ReceivedTime)

$myDoc.TypeParagraph()

foreach($item in $hourslist){
    IF($item){
        $item = $item.Replace("_","")
        $myXlSheet.Cells.Item($row2,1) = $mailDate.ReceivedTime.ToString("yyyy/MM/dd") #mail date
        $myXlSheet.Cells.Item($row2,3) = $item #task hours
        $myDoc.TypeText($item + "`t")
        $row2++
    }
}

$myDoc.TypeParagraph()

foreach($item in $TaskArray){
    IF($item){
        $myXlSheet.Cells.Item($row1,2) = $item.Replace("_"," ") 
        $myDoc.TypeText($item.Replace("_"," "))
        $myDoc.TypeParagraph()
        $row1++
    }
}

# Set the width of the columns automatically
$myXlSheet.columns.item("A:J").EntireColumn.AutoFit() | out-null

$myDoc.TypeParagraph()
$myDoc.TypeParagraph()





#$myDoc.TypeText($bodytext)
#$myDoc.TypeParagraph()
#$myDoc.TypeParagraph()

}








#$workbook.Close($true)
#$excel.quit()

#$Outlook.Quit()
#$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$word)
#[gc]::Collect()
#[gc]::WaitForPendingFinalizers()
#Remove-Variable word

#$null = [System.Runtime.InteropServices.Marshal]::ReleaseComObject([System.__ComObject]$Excel)
#[gc]::Collect()
#[gc]::WaitForPendingFinalizers()
#Remove-Variable Excel
