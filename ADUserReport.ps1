Import-module ActiveDirectory

$ListOfUsers = Get-ADUser -Filter * -SearchBase "OU=AnotherSubContractorOU,OU=Contractors,OU=AnotherSubOU Users,OU=Users,OU=MyBusiness,DC=MyDomain,DC=local" | select-object -ExpandProperty SamAccountName

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

$preReplacement = $tableoutput | Sort-Object "Username" | ConvertTo-Html -Fragment

$htmlreport = $preReplacement -replace "&lt;","<" -replace "&gt;",">"

ConvertTo-Html -Body $htmlreport -Title "Staff AD Account Report" -Head $Header | Out-File adinfo.html
