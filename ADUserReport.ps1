Import-module ActiveDirectory

# Modify These Variables accordingly.

$CompanyName = "CompanyNameHere"

$PCMOrganizationalUnit = "OU=AnotherSubContractorOU,OU=Contractors,OU=AnotherSubOU Users,OU=Users,OU=MyBusiness,DC=MyDomain,DC=local"

$ReportDate = get-date -Format g

$SenderName = "sender@example.com"
$Recipient = 'receiver@somedomain.com'

$EmailSubject = "User Report - $Reportdate"

$BodyMessage = "Attached is Users Audit Report for $CompanyName"

# Generate your own account password hash by running the command in powershell >>> read-host -assecurestring | convertfrom-securestring | out-string
# Then type in your actual password and hit ENTER.
# Copy the output series of strings and paste it in $hashpassword below.
$hashpassword = "0000000036654b828926d875ec921e1c624d1db0f5b2f79140000009e8179d66d258a7a1f4c31d139d92bc53d3ccda4"

$GeneratedFile = "C:\temp\adinfo.html"


### Table Build Up
$ListOfUsers = Get-ADUser -Filter * -SearchBase $PCMOrganizationalUnit | select-object -ExpandProperty SamAccountName

$TableOutput = @()

foreach ($SampleUser in $ListOfUsers) {

    $outputdata = New-Object psobject

    $SampleADInfo = Get-ADUser $SampleUser -Properties *

    $UserGroupMemberships = ((Get-ADPrincipalGroupMembership -Identity $SampleUser | select-object -ExpandProperty name) -join ',')

    if ($SampleADInfo.Enabled -eq $True) {
       $AccountActive = "<div class=green>YES</div>"
    } else {
        $AccountActive = "<div class=red>NO</div>"
    }


    if ($SampleADInfo.PasswordExpired -eq $false) {

        $IsPasswordExpired = "<div class=green>NO</div>"

    } else {$IsPasswordExpired = "<div class=red>YES</div>"}

    if ($SampleADInfo.PasswordNeverExpires -eq $false) {

        $PasswordNeverExpires = "<div class=green>NO</div>"

    } else {$PasswordNeverExpires = "<div class=red>YES</div>"}

    if ($SampleADInfo.LockedOut -eq $false){

        $LockedOut = "<div class=green>NO</div>"

    } else {$LockedOut = "<div class=red>YES</div>"}

    $outputdata | Add-Member NoteProperty -name "Username" -value $SampleADInfo.SamAccountName
    $outputdata | Add-Member NoteProperty -name "Display Name" -value $SampleADInfo.DisplayName
    $outputdata | Add-Member NoteProperty -name "Is Enabled?" -value $AccountActive
    $outputdata | Add-Member NoteProperty -name "Last Logon Date" -value $SampleADInfo.LastLogonDate
    $outputdata | Add-Member NoteProperty -name "Creation Date" -value $SampleADInfo.Created
    $outputdata | Add-Member NoteProperty -name "Last Password Reset" -value $SampleADInfo.PasswordLastSet
    $outputdata | Add-Member NoteProperty -name "Is Password Expired?" -value $IsPasswordExpired
    $outputdata | Add-Member NoteProperty -name "Password Never Expires?" -value $PasswordNeverExpires
    $outputdata | Add-Member NoteProperty -name "Is LockedOut?" -value $LockedOut
    $outputdata | Add-Member NoteProperty -name "Membership" -value $UserGroupMemberships

    $tableoutput += $outputdata

}

# For HTML CSS Formatting

$Header = @"
<style type='text/css'>
table {
  "table-layout: fixed";
  border-collapse: collapse;
  "width: 100%";
  white-space: nowrap;
  font-family: "Trebuchet MS", Arial, Helvetica, sans-serif;
}
td {
  border: 1px solid black;
  "white-space: nowrap";
  "width: 100%";
  "overflow: visible";
}
th {
  border: 1px solid black;
  "width: 100%"; 
  "white-space: nowrap";
  "overflow: visible";
  background-color: #0059FF;
}
div {
  
    text-align: center;
  }
  div.red {
    background-color: #FF0000;
  }
  div.green {
    background-color: #008000;
  
  }
</style>
"@

# HTML Generation part

$preReplacement = $tableoutput | Sort-Object "Display Name" | ConvertTo-Html -Fragment

$htmlreport = $preReplacement -replace "&lt;","<" -replace "&gt;",">"

ConvertTo-Html -Body $htmlreport -Head $Header -PostContent "Report Generated: $ReportDate" | Out-File $GeneratedFile

# Emailing Part

$securePwd = $hashpassword | convertto-securestring
$credObject = New-Object System.Management.Automation.PSCredential -ArgumentList $username, $securePwd


$mailParams = @{
    SmtpServer                 = 'smtp.office365.com'
    Port                       = '587'
    From                       = $SenderName
    To                         = $Recipient
    Subject                    = $EmailSubject
    Body                       = $BodyMessage
    Attachments                = $GeneratedFile
    BodyAsHTML                 = $true
    DeliveryNotificationOption = 'OnFailure', 'OnSuccess'
}

Send-MailMessage @mailParams -UseSSL -Credential $credObject
