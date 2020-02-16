#See https://stackoverflow.com/questions/6239647/using-powershell-credentials-without-being-prompted-for-a-password
#INTRODUCTION
#Set the parameter as either success or failure
#[string]$status
Param([String]$status='Success', [String] $body = 'The VPFISMA script ran successfully.')

#Set the root publishing path to the parent of the current directory
$publishingproject_path = "$PSScriptRoot\.."
Set-Location $publishingproject_path

#Set the properties file location
$properties_location = "$publishingproject_path\scripts\variables.properties"

#Retrieve and prep the properties data
$properties_content = Get-Content "$properties_location" -Raw
$properties_content = $properties_content -replace "(\\)","/"
$properties_content = [regex]::Escape($properties_content)
$properties_content = $properties_content -replace "(\\r)?\\n",[Environment]::NewLine

$props = ConvertFrom-StringData ($properties_content)
$SMTPServer = $props.SMTPServer
$SMTPPort=$props.SMTPPort
$fromAddress=$props.fromAddress
[string[]]$toAddress=$props.toAddress.split(',')

#Create the password using this code (copy/paste in Powershell to generate) :
# read-host -assecurestring | convertfrom-securestring | out-file "$publishingproject_path/scripts/mysecurestring.txt"

$password = Get-Content "./scripts/mysecurestring.txt" | ConvertTo-SecureString
$cred = new-object -typename System.Management.Automation.PSCredential -argumentlist $fromAddress, $password
 $SMTPCred = $cred
 $SMTPServer = $SMTPServer
 $Attachments = "$publishingproject_path/log_summary.txt"
 
 $hName = [System.Net.Dns]::GetHostName().ToUpper()
 $MailParams = @{
        To = $toAddress
        From = $fromAddress
        Credential = $SMTPCred
        SmtpServer = $SMTPServer
        Port = $SMTPPort
        Subject = "$status in VPFISMA from server $hName"
        Body = $body
        Attachments = $Attachments
        }

Send-MailMessage @MailParams -UseSsl
