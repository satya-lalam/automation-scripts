$DirectoryToSave='<direcotory to save excel>' 
$Filename='<report name>' 
$currentDateFileFormat= Get-Date -Format yyyyMMddTHHmmss
$currentDate= Get-Date -DisplayHint DateTime
$From ='<from email>' 
$myToEmails = <list all the emails in single quotes and seperated by comma>
$SMTP= '<smtp address>' 
$DSN='windows datasource name' 
 
# constants. 
$xlCenter=-4108 
$xlTop=-4160 
$xlOpenXMLWorkbook=[int]51 

# You can replace the SQL  
$SQL=@" 
<SQL query here>
"@ 
 
#Create a Excel file to save the data 
# if the directory doesn't exist, then create it 
if (!(Test-Path -path "$DirectoryToSave")) #create it if not existing 
  { 
  New-Item "$DirectoryToSave" -type directory | out-null 
  }  
$excel = New-Object -Com Excel.Application #open a new instance of Excel 
$excel.Visible = $False #make it visible (for debugging more than anything) 
$wb = $excel.Workbooks.Add() #create a workbook 
$currentWorksheet=1 #there are three open worksheets you can fill up 
      if ($currentWorksheet-lt 4)  
      { 
        $ws = $wb.Worksheets.Item($currentWorksheet) 
      } 
      else   
      { 
        $ws = $wb.Worksheets.Add() 
      } #add if it doesn't exist 
      $currentWorksheet += 1 #keep a tally     
  # You can refresh it  
      $qt = $ws.QueryTables.Add("ODBC;DSN=$DSN", $ws.Range("A1"), $SQL) 
      # and execute it 
      if ($qt.Refresh()) #if the routine works OK 
            { 
            $ws.Activate() 
            $ws.Select() 
            $excel.Rows.Item(1).HorizontalAlignment = $xlCenter 
            $excel.Rows.Item(1).VerticalAlignment = $xlTop 
            $excel.Rows.Item("1:1").Font.Name = "Calibri" 
            $excel.Rows.Item("1:1").Font.Size = 11 
            $excel.Rows.Item("1:1").Font.Bold = $true 
            $Excel.Columns.Item(1).Font.Bold = $true 
            }       
$filename = "$DirectoryToSave$filename$currentDateFileFormat.xlsx" #save it according to its title 
if (test-path $filename ) { rm $filename } #delete the file if it already exists 
$wb.SaveAs($filename,  $xlOpenXMLWorkbook) #save as an XML Workbook (xslx) 
$wb.Saved = $True #flag it as being saved 
$wb.Close() #close the document 
$excel.Quit() #and the instance of Excel 
$wb = $Null #set all variables that point to Excel objects to null 
$ws = $Null #makes sure Excel deflates 
#$excel=$Null #let the air out 
[System.Runtime.Interopservices.Marshal]::ReleaseComObject($excel)
Remove-Variable excel
#Function to send email with an attachment 
#[string]$emailTo1, [string]$emailTo2,
Function sendEmail([string]$emailFrom, [string[]]$myToEmails, [string]$subject,[string]$body,[string]$smtpServer,[string]$filePath) 
{ 
#initate message 
$email = New-Object System.Net.Mail.MailMessage  
$email.From = $emailFrom 
foreach ($emailTo in $myToEmails)
  {
    $email.To.Add($emailTo) 
  }
$email.Subject = $subject 
$email.Body = $body 
# initiate email attachment  
$emailAttach = New-Object System.Net.Mail.Attachment $filePath 
$email.Attachments.Add($emailAttach)  
#initiate sending email  
$smtp = new-object Net.Mail.SmtpClient($smtpServer) 
$smtp.Send($email) 
} 

#Call Function  
sendEmail -emailFrom $from -myToEmails $myToEmails -subject "<email subject>" -body "<email body `n>" -smtpServer $SMTP -filePath $filename 
 
