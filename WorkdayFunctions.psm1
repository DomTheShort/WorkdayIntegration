#WorkdayFunctions.psm1
#This File contains the Functions used in WorkdayToADIntegration.ps1
#Author: Curtis Blackley, Dominic Douglas

#Helpdesk Functions
function Connect-Helpdesk
{
    param([string]$user, [string]$password)

    #Generate SessionID
    $sessionParams = @{username=$user;password=$password} #Change to not use my Account
    $session = Invoke-WebRequest -uri https://helpdesk.kengarff.com/staffapi/index.php?/Core/Default/Login -Method Post -Body $sessionParams

    $xml = [xml]$session.Content
    $sessionid = $xml.kayako_staffapi.sessionid.'#cdata-section'

    return $sessionid
}

function Check-HelpdeskTicketPosted
{
    param($User,$TicketType)

    $EmployeeNumber = $User.EmployeeNumber
    $Dealership = $User.Dealership

    $SQLQuery = "SELECT TOP 1 * FROM SubmittedTickets WHERE EmployeeNum = $EmployeeNumber AND DealershipDescription = '$Dealership' AND TicketType = '$TicketType'"

    $Results = Run-SQLCommand $SQLQuery

    if ($Results.Tables.TicketID -eq $null)
    {
        return $false
    }
    else
    {
        return $true
    }
}

function Post-TicketUpdate
{
    param([string]$SessionID, [string]$Response, [string]$TicketID)

    $details = [xml]('<kayako_staffapi><modify ticketid="' + $TicketID + '"><reply><contents>'+ $Response + '</contents></reply></modify></kayako_staffapi>')
   
    $replyParams = @{sessionid=$sessionid;payload=$details.InnerXml}
    $reply = Invoke-WebRequest -uri https://helpdesk.kengarff.com/staffapi/index.php?/Tickets/Push/Index -Method Post -Body $replyParams
}

function Post-NewTicket
{
    param([string]$Response,[array]$CC,[string]$Subject,$HDDepartment)

    $SessionID = Connect-Helpdesk -user autosupport -password '&koh#Ng082S'

    $XMLString = '<kayako_staffapi><create><subject>'+$Subject+'</subject><fullname>Automated Support</fullname><email>autosupport@kengarff.com</email>
    <departmentid>'+$HDDepartment+'</departmentid><ticketstatusid>1</ticketstatusid><ticketpriorityid>1</ticketpriorityid>
    <tickettypeid>1</tickettypeid><ownerstaffid>0</ownerstaffid><emailqueueid>0</emailqueueid><creator>user</creator><userid>0</userid>
    <type>default</type><sendautoresponder>0</sendautoresponder><flagtype>0</flagtype>'

    foreach($item in $CC) 
    {
        $XMLString += '<ccto>'+ $item +'</ccto>'
    }

    $XMLString += '<watch>0</watch><reply><contents>'+ $Response + '</contents></reply></create></kayako_staffapi>'

    $details = [xml]$XMLString
   
    $replyParams = @{sessionid=$sessionid;payload=$details.InnerXml}
    $reply = Invoke-WebRequest -uri https://helpdesk.kengarff.com/staffapi/index.php?/Tickets/Push/Index -Method Post -Body $replyParams

}

function Sanitize-UserInfo
{
    param($User)

    $User.Title = $User.Title -replace '&','&amp;'
    $User.Printers = $User.Printers -replace '&','&amp;'
    $User.DistributionLists = $User.DistributionLists -replace '&','&amp;'
}

function Get-TerminationTicketResponse
{
    param($User,$ADUser)

    Sanitize-UserInfo -User $User

    $Response = "Terminated User: " + $User.FirstName + " " + $User.LastName + "&#xA;"
    $Response += "Username: " + $AdUser.UserPrincipalName + "&#xA;"
    $Response += "Dealership: " + $AdUser.Company + "&#xA;"
    $Response += "Department: " + $User.Department + "&#xA;"
    $Response += "Title: " + $User.Title + "&#xA;"
    $Response += "Term Date:" + $User.TermDate + "&#xA;"
    $Response += "Employee Number:" + $User.employeeNumber + "&#xA;"
    
    return $Response 
}

function Create-NewHireHelpdeskResponse {

    param($User,$Username,$Password)

    Sanitize-UserInfo -User $User

    $Response = "Name: " + $User.PreferredName + "&#xA;"
    $Response += "Title: " + $User.Title + "&#xA;"
    $Response += "Department: " + $User.Department + "&#xA;"
    $Response += "Dealership: " + $User.Dealership + "&#xA;"
    $Response += "Manager: " + $User.ManagerName + "&#xA;"
    $Response += "Employee Number: " + $User.EmployeeNumber + "&#xA;"
    $Response += "Hire Date: " + $User.HireDate + "&#xA;&#xA;"
    $Response += "Login: " + $Username + "&#xA;"
    $Response += "Password: " + $Password + "&#xA;"

    $CopyUser = Get-UserForReynoldsCopy -user $User

    $Response += "User for Reynolds Copy: " + $CopyUser + "&#xA;"

    return $Response

}

function Create-UpdateHelpdeskResponse
{
    param($PreUser,$PostUser)

    Sanitize-UserInfo -User $User

    $Response = "Name: " + $PreUser.DisplayName + (Compare-UserFields -Prefield $PreUser.DisplayName -PostField $PostUser.DisplayName)  + "&#xA;"
    $Response += "Title: " + $PreUser.Title + (Compare-UserFields -Prefield $PreUser.Title -PostField $PostUser.Title)  + "&#xA;"
    $Response += "Department: " + $PreUser.Department + (Compare-UserFields -Prefield $PreUser.Department -PostField $PostUser.Department)  + "&#xA;"
    $Response += "Dealership: " + $PreUser.Company + (Compare-UserFields -Prefield $PreUser.Company -PostField $PostUser.Company)  + "&#xA;"

    $PreManager = (($PreUser.Manager.Split(","))[0].Split("="))[1]
    $PostManager = (($PostUser.Manager.Split(","))[0].Split("="))[1]

    $Response += "Manager: " + $PreManager + (Compare-UserFields -Prefield $PreManager -PostField $PostManager)  + "&#xA;"


    $Response += "Login: " + $PreUser.UserPrincipalName + (Compare-UserFields -Prefield $PreUser.UserPrincipalName -PostField $PostUser.UserPrincipalName)  + "&#xA;"

    $CopyUser = Get-UserForReynoldsCopy -user $User

    $Response += "User for Reynolds Copy: " + $CopyUser + "&#xA;"


    return $Response
}

function Post-QuestionnaireTickets
{
    param($User)

    Sanitize-UserInfo -User $User

    if($User.DistributionLists -eq "")
    {
    }
    else
    {
        $CC = $User.ManagerEmail

        if(($User.DistributionLists -eq "No" -or $User.DistributionLists -eq "None" -or $User.DistributionLists -eq "n/a" -or $User.DistributionLists -eq "na") -and  ($User.SharedFolders -eq "No" -or $User.SharedFolders -eq "None" -or $User.SharedFolders -eq "n/a" -or $User.SharedFolders -eq "na") -and  ($User.Printers -eq "No" -or $User.Printers -eq "None" -or $User.Printers -eq "n/a" -or $User.Printers -eq "na") -and  ($User.Reynolds -eq "No" -or $User.Reynolds -eq "None" -or $User.Reynolds -eq "n/a" -or $User.Reynolds -eq "na") -and  ($User.WorkstationName -eq "No" -or $User.WorkstationName -eq "None" -or $User.WorkstationName -eq "n/a" -or $User.WorkstationName -eq "na") -and  ($User.AdditionalDetails -eq "No" -or $User.AdditionalDetails -eq "None" -or $User.AdditionalDetails -eq "n/a" -or $User.AdditionalDetails -eq "na"))
        {
        }
        else
        {
                $Subject = "Software Update - " + $User.PreferredName + " - " + $User.Dealership
                Post-NewTicket -Response (Create-SoftwareRepsonse -User $User) -CC $CC -Subject $Subject -HDDepartment 21
                
                $EmployeeNumber = $User.EmployeeNumber
                $Dealership = $User.Dealership

                $SQLQuery = "INSERT INTO SubmittedTickets (EmployeeNum, DealershipDescription, TicketType, TicketDate) VALUES ($EmployeeNumber, '$Dealership', 'S', GETDATE())"
                Run-SQLCommand $SQLQuery
        }

        if(($User.PhoneExtension -eq "No" -or $User.PhoneExtension -eq "None" -or $User.PhoneExtension -eq "n/a"  -or $User.PhoneExtension -eq "na") -and  ($User.CurrentExtension -eq "No" -or $User.CurrentExtension -eq "None" -or $User.CurrentExtension -eq "n/a"  -or $User.CurrentExtension -eq "na") -and  ($User.Fax -eq "No" -or $User.Fax -eq "None" -or $User.Fax -eq "n/a" -or $User.Fax -eq "na") -and  ($User.Forwarding -eq "No" -or $User.Forwarding -eq "None" -or $User.Forwarding -eq "n/a"  -or $User.Forwarding -eq "na") -and  ($User.SharedDesk -eq "No" -or $User.SharedDesk -eq "None" -or $User.SharedDesk -eq "n/a"  -or $User.SharedDesk -eq "na") -and  ($User.Voicemail -eq "No" -or $User.Voicemail -eq "None" -or $User.Voicemail -eq "n/a" -or $User.Voicemail -eq "na"))
        {
        }
        else
        {
                $CC = @($User.ManagerEmail,"allcrm@kengarff.com")
                $Subject = "Phone Update - " + $User.PreferredName + " - " + $User.Dealership
                Post-NewTicket -Response (Create-PhoneResponse -User $User) -CC $CC -Subject $Subject -HDDepartment 21

                $EmployeeNumber = $User.EmployeeNumber
                $Dealership = $User.Dealership

                $SQLQuery = "INSERT INTO SubmittedTickets (EmployeeNum, DealershipDescription, TicketType, TicketDate) VALUES ($EmployeeNumber, '$Dealership', 'P', GETDATE())"
                Run-SQLCommand $SQLQuery
        }   
    }
}

function Create-SoftwareRepsonse
{
    param($User)

    $Response = "Name: " + $User.PreferredName + "&#xA;"
    $Response += "Employee Number: " + $User.EmployeeNumber + "&#xA;"
    $Response += "Dealership: " + $User.Dealership + "&#xA;"
    $Response += "Department: " + $User.Department + "&#xA;"
    $Response += "Title: " + $User.Title + "&#xA;"
    $Response += "Email Address: " + $User.Mail + "&#xA;&#xA;"


    $Response += "Distribution Lists: " + $User.DistributionLists + "&#xA;"
    $Response += "Shared Folders: " + $User.SharedFolders + "&#xA;"
    $Response += "Printers: " + $User.Printers + "&#xA;"
    $Response += "Reynolds User to Copy: " + $User.Reynolds + "&#xA;"
    $Response += "Workstation Name: " + $User.WorkstationName + "&#xA;"
    $Reposnse += "Additional Details: " + $User.AdditionalDetails + "&#xA;"

    return $Response

}

function Create-PhoneResponse
{
    param($User)

    $Response = "Name: " + $User.PreferredName + "&#xA;"
    $Response += "Employee Number: " + $User.EmployeeNumber + "&#xA;"
    $Response += "Dealership: " + $User.Dealership + "&#xA;"
    $Response += "Department: " + $User.Department + "&#xA;"
    $Response += "Manager: " + $User.Manager + "&#xA;"
    $Response += "Title: " + $User.Title + "&#xA;&#xA;"

    $Response += "New Extension: " + $User.PhoneExtension + "&#xA;"
    $Response += "Desired Extension: " + $User.CurrentExtension + "&#xA;"
    $Response += "Fax: " + $User.Fax + "&#xA;"
    $Response += "Allowed to Forward: " + $User.Forwarding + "&#xA;"
    $Response += "Cell Phone: " + $User.CellPhone + "&#xA;"
    $Response += "Do they Share a Desk: " + $User.SharedDesk + "&#xA;"
    $Response += "Voicemail: " + $User.Voicemail + "&#xA;"

    return $Response
}

#Email Functions

function Get-SMTPServer
{
    $smtpServer = "smtp.office365.com"
    $Login = "autosupport@kengarff.com"
    $Password = '^&AOo456nA$ohsedgf^*U&456d'
    $smtp = new-object Net.Mail.SmtpClient($smtpServer, 587)
    $smtp.EnableSsl = $true
    $smtp.Credentials = New-Object System.Net.NetworkCredential($Login,$Password)

    return $smtp
}

function Send-Email
{
    param($To,$From,$Subject,$Body,$ReplyTo)

    $smtp = Get-SMTPServer

    $MailMessage = new-object Net.Mail.MailMessage($From, $To, $Subject, $Body)
    $MailMessage.IsBodyHtml = $false
    $MailMessage.ReplyTo = $ReplyTo
    $smtp.Send($MailMessage)
}

function Notify-NewHire
{
    param($User,$ADUser)

    $NotificationSettings = Get-NotificationSettings -Title $User.Title
    $NotificationEmails = Get-NotificationEmails

    if($NotificationSettings.notifyServiceDirector -eq $true)
    {
        $Email = $NotificationEmails | Where {$_.description -eq "ServiceDirector"} | Select emailAddress
        Send-NewHireEmail -User $User -ADUser $ADUser -To $Email.emailAddress
    }
    
    if($NotificationSettings.notifyCRM -eq $true)
    {
        $Email = $NotificationEmails | Where {$_.description -eq "CRM"} | Select emailAddress
        Send-NewHireEmail -User $User -ADUser $ADUser -To $Email.emailAddress
    }
}

function Send-CRMEmail
{
    param($User)

    $emailFrom = "autosupport@kengarff.com"
    $emailTo = "allcrm@kengarff.com"
    $subject = "CRM Questionnaire Notification - " + $User.PreferredName

    $body = Get-CRMResponse -User $User
    if($User.DistributionLists -eq "")
    {
    }
    else
    {
        Send-Email -To $emailTo -From $emailFrom -Subject $Subject -Body $Body -ReplyTo $User.ManagerEmail
    }
}

function Get-CRMResponse
{
    param($User,$ADUser)

    $Response = "Name: " + $User.PreferredName + "`n"
    $Response += "Email: " + $ADUser.UserPrincipalName + "`n"
    $Response += "Title: " + $User.Title + "`n"
    $Response += "Dealership: " + $User.Dealership + "`n"
    $Response += "Department: " + $User.Department + "`n"
    $Response += "Employee Number: " + $User.EmployeeNumber + "`n`n"

    $Response += "MXIE Extension: " + $User.CurrentExtension + "`n"
    $Response += "Dealersocket Team: " + $User.DealersocketName + "`n"
    $Response += "Dealersocket Manager: " + $User.DealersocketManager + "`n"
    $Response += "Cell Phone: " + $User.CellPhone + "`n"
    $Response += "Cell Provider: " + $User.CellProvider + "`n"
    $Response += "Notifications: " + $User.Notifications + "`n"
    $Response += "Rehire: " + $User.Rehire + "`n"
    $Response += "Additional Details: " + $User.AdditionalDetails + "`n"

    return $Response
}

function Send-NewHireEmail
{
    param($User,$ADUser,$To)

    $emailFrom = "autosupport@kengarff.com"
    $subject = "New Hire Notification - " + $User.PreferredName + " @ " + $user.Dealership

    $body = Get-CRMNewHireResponse -User $User -ADUser $ADUser

    Send-Email -To $To -From $emailFrom -Subject $Subject -Body $Body -ReplyTo $User.ManagerEmail
}



function Send-CRMNewHireEmail
{
    param($User,$ADUser)

    $emailFrom = "autosupport@kengarff.com"
    $emailTo = "allcrm@kengarff.com"
    $subject = "CRM New Hire Notification - " + $User.PreferredName + " @ " + $user.Dealership

    $body = Get-CRMNewHireResponse -User $User -ADUser $ADUser

    Send-Email -To $emailTo -From $emailFrom -Subject $Subject -Body $Body -ReplyTo $User.ManagerEmail

}

function Get-CRMNewHireResponse
{
    param($User,$ADUser)

    $Response = "Name: " + $User.PreferredName + "`n"
    $Response += "Email: " + $ADUser.UserPrincipalName + "`n"
    $Response += "Title: " + $User.Title + "`n"
    $Response += "Department: " + $User.Department + "`n"
    $Response += "Dealership: " + $User.Dealership + "`n"
    $Response += "Manager: " + $User.ManagerName + "`n"
    $Response += "Employee Number: " + $User.EmployeeNumber + "`n"
    $Response += "Hire Date: " + $User.HireDate + "`n`n"
    


    return $Response
}

function Send-TermEmail
{
    param($User,$ADUser)

    $From = "autosupport@Kengarff.com"
    $ToArray = @("allcrm@kengarff.com","ceceliaa@kornerstoneadmin.com","joelu@kengarff.com","molson@kengarff.com","mbigler@kengarff.com")
    $Subject = $Subject = "Termination - " + $user.preferredName + " - " + $ADUser.Company

    $Body = Get-TerminationEmailResponse -User $User -ADUser $ADUser

    foreach($To in $ToArray)
    {
        Send-Email -To $To -From $From -Subject $Subject -Body $Body -ReplyTo $User.ManagerEmail
    }

}

function Get-TerminationEmailResponse
{
    param($User,$ADUser)

    $Response = "Terminated User: " + $User.FirstName + " " + $User.LastName + "`n"
    $Response += "Username: " + $AdUser.UserPrincipalName + "`n"
    $Response += "Dealership: " + $AdUser.Company + "`n"
    $Response += "Department: " + $User.Department + "`n"
    $Response += "Title: " + $User.Title + "`n"
    $Response += "Term Date:" + $User.TermDate + "`n"
    
    return $Response 
}

function Send-CRMUpdateEmail
{
    param($PreUser,$PostUser)

    $emailFrom = "autosupport@kengarff.com"
    $emailTo = "allcrm@kengarff.com"
    $subject = "CRM Update Notification - " + $User.PreferredName + " @ " + $user.Dealership

    $body = Get-UpdateEmailResponse -PreUser $PreUser -PostUser $PostUser

    Send-Email -To $emailTo -From $emailFrom -Subject $Subject -Body $Body -ReplyTo $User.ManagerEmail
}

function Get-UpdateEmailResponse
{
    param($PreUser,$PostUser)

    $Response = "Name: " + $PreUser.DisplayName + (Compare-UserFields -Prefield $PreUser.DisplayName -PostField $PostUser.DisplayName)  + "`n"
    $Response += "Title: " + $PreUser.Title + (Compare-UserFields -Prefield $PreUser.Title -PostField $PostUser.Title)  + "`n"
    $Response += "Department: " + $PreUser.Department + (Compare-UserFields -Prefield $PreUser.Department -PostField $PostUser.Department)  + "`n"
    $Response += "Dealership: " + $PreUser.Company + (Compare-UserFields -Prefield $PreUser.Company -PostField $PostUser.Company)  + "`n"

    $PreManager = (($PreUser.Manager.Split(","))[0].Split("="))[1]
    $PostManager = (($PostUser.Manager.Split(","))[0].Split("="))[1]

    $Response += "Manager: " + $PreManager + (Compare-UserFields -Prefield $PreManager -PostField $PostManager)  + "`n"


    $Response += "Login: " + $PreUser.UserPrincipalName + (Compare-UserFields -Prefield $PreUser.UserPrincipalName -PostField $PostUser.UserPrincipalName)  + "`n"

    return $Response
}


#File Functions
function Check-WorkdayFTP
{
    return Test-Path \\vm-ftp\FTPFolders\Workday\ActiveDirectory\Active_*.csv
}

function Import-UserList
{
    param([string]$file)
    $userlist = Import-Csv -Delimiter "," -Header SSN,EmployeeNumber,FirstName,LastName,PreferredName,Dealership,Department,Title,ManagerName,ManagerEmail,HireDate,TermDate,DistributionLists,SharedFolders,Printers,Reynolds,WorkstationName,PhoneExtension,CurrentExtension,Fax,Forwarding,SharedDesk,Voicemail,DealersocketName,DealersocketManager,CellPhone,CellProvider,Notifications,Rehire,AdditionalDetails \\vm-ftp\FTPFolders\Workday\ActiveDirectory\$file

    return $userlist
}


function Get-FileList
{
    $allFiles = Get-ChildItem \\vm-ftp\FTPFolders\Workday\ActiveDirectory\ -File

    return $allFiles
}

function Move-Files
{
    param($Files)

    $Destination = "\\vm-ftp\FTPFolders\Workday\ActiveDirectory\Processed"

    foreach($File in $files)
    {
        $Source = "\\vm-ftp\FTPFolders\Workday\ActiveDirectory\" + $File.Name
    
        Move-Item $Source $Destination
    }
}

function Generate-UserList
{
    param($FileList)

    $userlist = [System.Collections.ArrayList]@()

    foreach($file in $FileList)
    {
        
        $list = Import-UserList -file $file.Name

        foreach($item in $list)
        {
            $userlist += $item
        }

    }

    Move-Files -Files $FileList

    return $userlist
}

#Verification Functions
function Check-EmployeeExists
{
    param($EmployeeNumber)

    $check = Get-ADUser -LDAPFilter "(employeeNumber=$EmployeeNumber)"

    if($check -ne $null)
    {
        return $true
    }
    else
    {
        return $false
    }
}

function Check-ADUser 
{
    param([string]$User)

    $adUserCheck = Get-ADUser -LDAPFilter "(SAMAccountName=$User)"

    if($adUserCheck -ne $null)
    {
        return $true
    }
    else
    {
        return $false
    }
}

function Compare-User
{
    param($ADUser,$WorkdayUser)

    $Manager = Get-ADUserByEmailAddress -EmailAddress $WorkdayUser.ManagerEmail
    $Department = ($WorkdayUser.Department.Split(' ('))[0]

    if($AdUser.GivenName -ne $WorkdayUser.FirstName)
    {
        return $false
    }
    elseif($AdUser.Surname -ne $WorkdayUser.LastName)
    {
        return $false
    }
    elseif($AdUser.Name -ne $WorkdayUser.PreferredName)
    {
        return $false
    }
    elseif($AdUser.DisplayName -ne $WorkdayUser.PreferredName)
    {
        return $false
    }
    elseif($AdUser.Company -ne $WorkdayUser.Dealership)
    {
        return $false
    }
    elseif($AdUser.Department -ne $Department)
    {
        return $false
    }
    elseif($AdUser.Title -ne $WorkdayUser.Title)
    {
        return $false
    }
    elseif($AdUser.Manager -ne $Manager.DistinguishedName)
    {
        return $false
    }
    else
    {
        return $true
    }

}

function Compare-UserFields
{
    param($PreField,$PostField)

    if($PreField -ne $PostField)
    {
        return " --> " + $PostField
    }
    else
    {
        return ""
    }
}

function Check-AccentedCharacters
{
    param($String)

    $AccentedCharacters = "^[-a-zA-Z0-9\s]+$"

    return !($String -match $AccentedCharacters)
}

#User Groups Functions
function Update-ADUserSecurityGroup
{
    param([string]$User,[string]$SecurityGroup)

    Add-AdGroupMember -Identity $SecurityGroup -Members $User
}

function Remove-AllUserGroups
{
    param([string]$User)

    Get-ADPrincipalGroupMembership -Identity $User | Where-Object {$_.name -ne "Domain Users" } | foreach { Remove-ADPrincipalGroupMembership -Identity $User -MemberOf $_ -Confirm:$false}

}

#Retrival Functions
function Get-ADUserByEmployeeNumber
{
    param($EmployeeNumber)

    $user = Get-ADUser -LDAPFilter "(EmployeeNumber=$EmployeeNumber)" -Properties *

    return $user
}

function Get-ADUserByEmailAddress
{
    param($EmailAddress)

    if($EmailAddress -ne $null)
    {
        $user = Get-ADUser -LDAPFilter "(mail=$EmailAddress)"
    }
    else
    {
       $user = ""
    }
    
    return $user
}

function Get-UserForReynoldsCopy
{
    param($user)

    $Users = Get-Aduser -Filter {(Title -eq $user.Title) -and (Company -eq $user.Dealership) }

    return $Users[0].Name
}

#Creation Functions
function Create-ADUser
{
    param($inputUser,$inputUsername)

    if($inputUser.SSN -eq $null)
    {
        $inputUser.SSN = (Get-Random -Maximum 9999 -Minimum 1000)
    }

    $Password = "K3NG4RFF" + $inputUser.SSN + "!!"

    $Domain = Get-EmailDomain -Dealership $inputUser.Dealership

    New-ADUser -SamAccountName $inputUsername -UserPrincipalName ($inputUsername + $Domain) -AccountPassword (ConvertTo-SecureString -AsPlainText $Password -Force) -Name $inputUser.PreferredName -DisplayName $inputUser.PreferredName -GivenName $inputUser.FirstName -Surname $inputUser.LastName -EmployeeNumber $inputUser.EmployeeNumber -Company $inputUser.Dealership -Title $inputUser.Title -Path (Get-OU -Title $inputUser.Title -Dealership $inputUser.Dealership) -PassThru | Enable-ADAccount

    #-Path (Get-OU -Title $inputUser.Title -Dealership $inputUser.Dealership)

    #Update-ADUser -ADUser (Get-AdUserByEmployeeNumber -EmployeeNumber $inputUser.EmployeeNumber) -UpdateInfo $inputUser

    $newUser = Get-adUser $inputUsername

    $newUser.Department = ($inputUser.Department.Split(' ('))[0]
    $newUser.mail = $newUser.UserprincipalName
    $newUser.ProxyAddresses = "SMTP:" + $newUser.UserprincipalName
    
    $manager = Get-ADUserByEmailAddress -EmailAddress $inputUser.ManagerEmail

    $newUser.Manager = $manager.DistinguishedName

    Set-Aduser -Instance $newUser

    $Subject = "New Hire - " + $user.PreferredName + " @ " + $user.Dealership
    Post-NewTicket -CC $inputUser.ManagerEmail -Subject $Subject -Response (Create-NewHireHelpdeskResponse -User $inputUser -Username ($inputUsername+$Domain) -Password $Password) -HDDepartment 20   

   

    #Send-CRMNewHireEmail -User $inputUser -ADUser $newUser
    Notify-NewHire -User $inputUser -ADUser $newUser

    if($inputUser.Distributionlists -ne "")
    {
        Send-CRMEmail -User $inputUser
        Post-QuestionnaireTickets -User $inputUser
    }
}

function Create-SharedMailbox
{
    param([string]$MailboxName)

    $ExistingMailboxes = Get-Mailbox -RecipientTypeDetails SharedMailbox | Select Name

    if(!$ExistingMailboxes.Name.Contains($MailboxName))
    {
        New-Mailbox -Shared -Name $MailboxName -DisplayName $MailboxName -Alias $MailboxName

        Set-Mailbox $MailboxName -MessageCopyForSendOnBehalfEnabled $true
        Set-Mailbox $MailboxName -MessageCopyForSentAsEnabled $true
    }
    else
    {
        Write-Out "$MailboxName already exists."
    }
    
}

function Create-EmailAddress
{
    param($Name,$Dealership)

    $domain = Get-EmailDomain -Dealership $Dealership

    $preference = Get-EmailPreference -Dealership $Dealership

    return Generate-EmailAddress -Name $Name -Preference $preference -Domain $domain
}

function Generate-EmailAddress
{
    param($Name,$Preference,$Domain)

    $nameArray = $Name.ToLower().Split(" ")

    $email = ""

    if($Preference -eq "FirstL")
    {
        $email = $nameArray[0]+$nameArray[1][0]+$Domain
    }
    elseif($Preference -eq "FLast")
    {
        $email = ($nameArray[0][0]+$nameArray[1..($nameArray.Count-1)]).Replace(' ','')+$Domain
    }

    return $email
}

#Update Functions
function Update-ADUser 
{
    param($ADUser,$UpdateInfo)

    $preUser = $ADUser

    $UpdateInfo.Department = ($UpdateInfo.Department.Split(' ('))[0]

    #If User has changed Dealership, Department, or Title, Move user to correct OU then reload info.
    if($ADUser.Company -ne $UpdateInfo.Dealership -or $ADUser.Department -ne $UpdateInfo.Department -or $ADUser.Title -ne $UpdateInfo.Title)
    {
        $Target = Get-OU -Title $UpdateInfo.Title -Dealership $UpdateInfo.Dealership

        Move-ADObject -Identity $ADUser.DistinguishedName -TargetPath $Target

        $ADUser = Get-ADUser -Identity $ADUser.SamAccountName
    }

    #Add Rename Object to Fields for Name
    #Update Userinfo
    $AdUser.GivenName = $UpdateInfo.FirstName
    $AdUser.Surname = $UpdateInfo.LastName
    $AdUser.DisplayName = $UpdateInfo.PreferredName
    $AdUser.Company = $UpdateInfo.Dealership
    $AdUser.Department = $UpdateInfo.Department
    $AdUser.Title = $UpdateInfo.Title

    $manager = Get-ADUserByEmailAddress -EmailAddress $UpdateInfo.ManagerEmail

    $AdUser.Manager = $manager.DistinguishedName

    $domain = Get-EmailDomain $UpdateInfo.Dealership

    #Update Useremail Address if necessary
    if($ADUser.mail -notlike "*$domain")
    {
        $username = Create-EmailAddress -Name $AdUser.Name -Dealership $AdUser.Company

        if($username -ne "")
        {
            $AdUser.UserPrincipalName = $username
            $AdUser.mail = $username
            $AdUser.proxyAddresses = "SMTP:$username"
        }
    }

    Set-AdUser -Instance $ADUser

    Rename-ADObject $adUser.DistinguishedName -NewName $UpdateInfo.PreferredName

    #Post ticket for Update

    $Subject = "Update User - " + $user.PreferredName + " - " + $User.Dealership
    Send-CRMUpdateEmail -PreUser $preUser -PostUser $ADUser 
    Post-NewTicket -CC $UpdateInfo.ManagerEmail -Subject $Subject -Response (Create-UpdateHelpdeskResponse -PreUser $preUser -PostUser $ADUser) -HDDepartment 21 
}

#Disable Functions
function Disable-ADUser
{
    param($AdUser)

    Disable-ADAccount $AdUser
    Move-ADObject $AdUser -TargetPath "OU=Users,OU=Disabled,DC=ad,DC=kengarff,DC=com"

}

#SQL Functions
function ConnectTo-SQL
{

    param([string]$SQLServer,[string]$Database)
    #Server Connection Settings
    
    $SqlConnection = New-Object System.Data.SqlClient.SqlConnection
    $SqlConnection.ConnectionString = "Server = $SQLServer; Database = $Database; Integrated Security = True"
    
    return $SqlConnection 

}

function Run-SQLCommand
{
    param([string]$SqlQuery)

    $SqlConnection = ConnectTo-SQL KGSQL ADWorkday

    $SqlCmd = New-Object System.Data.SqlClient.SqlCommand
    $SqlCmd.CommandText = $SqlQuery
    $SqlCmd.Connection = $SqlConnection
     
    $SqlAdapter = New-Object System.Data.SqlClient.SqlDataAdapter
    $SqlAdapter.SelectCommand = $SqlCmd
     
    $DataSet = New-Object System.Data.DataSet
    [void]$SqlAdapter.Fill($DataSet)

    return $Dataset
}

function Get-EmailDomain
{
    param($Dealership)

    $SQLQuery = "SELECT d.domainDescription FROM Domain d, DealershipOU de WHERE de.domainID = d.domainID AND de.workdayDescription = '$Dealership'"

    $Results = Run-SQLCommand $SQLQuery

    return $Results.Tables.domainDescription
}

function Get-EmailPreference
{
    param($Dealership)

    $SQLQuery = "SELECT p1.preferenceDescription AS FirstPreference,p2.preferenceDescription AS SecondPreference FROM DealershipOU d INNER JOIN UsernamePreference p1 ON d.firstPreferenceID = p1.preferenceID INNER JOIN UsernamePreference p2 ON d.secondPreferenceID = p2.preferenceID WHERE d.workdayDescription = '$Dealership'"

    $Results = Run-SQLCommand $SQLQuery

    return $Results.Tables
}

function Get-OU
{
    param($Title, $Dealership)

    if($Title -eq "Service BDC Support Rep" -and $Dealership -eq "Garff Enterprises, Inc.")
    {
        $OU = "OU=Users,OU=Inbound,OU=Call Center,OU=Departments,OU=Service Call Center,OU=Draper,OU=Utah,OU=Automotive,OU=Businesses,OU=Ken Garff,DC=ad,DC=kengarff,DC=com"
    }
    elseif($Dealership -eq "Arivo Acceptance, LLC")
    {
        $OU = "OU=Users,OU=Southtowne Used,OU=Sandy,OU=Utah,OU=Arivo,OU=Businesses,OU=Ken Garff,DC=ad,DC=kengarff,DC=com"
    }
    else
    {
        $SQLQuery = "SELECT d.DepartmentOUDescription, s.dealershipOUDescription, s.cityOUDescription, s.regionOUDescription, sc.SubCompanyOUDescription, c.CompanyOUDescription  FROM DepartmentOU d, JobProfile p, SubCompanyOU sc, CompanyOU c, (SELECT de.dealershipOUDescription, r.regionOUDescription, c.cityOUDescription FROM DealershipOU de, RegionOU r, CityOU c WHERE de.cityOUID = c.cityOUID AND c.regionOUID = r.regionOUID AND de.workdayDescription = '$Dealership') s WHERE p.departmentOUID = d.departmentOUID AND p.jobProfileDescription = '$title' AND p.subCompanyOUID = sc.SubCompanyOUID AND sc.CompanyOUID = c.CompanyOUID"

        $Results = Run-SQLCommand $SqlQuery

        $Table = $Results.Tables

        #Add End of OU structure
        $OU = "OU=Users," + $Table.DepartmentOUDescription + "OU=" + $Table.DealershipOUDescription + ",OU=" + $Table.cityOUDescription + ",OU=" + $Table.regionOUDescription + ",OU=" + $Table.SubCompanyOUDescription + ",OU=" + $Table.CompanyOUDescription + ",OU=Ken Garff,DC=ad,DC=kengarff,DC=com"
    }

    

    return $OU
}

function Get-DealershipInfo 
{
    param([string]$Dealership)

    $SqlQuery = "Select d.dealershipOUDescription, o.domainDescription, c.cityOUDescription, r.regionOUDescription, u.preferenceDescription AS 'FirstPreference', u2.preferenceDescription AS 'SecondPreference' FROM DealershipOU d INNER JOIN UsernamePreference u ON d.firstPreferenceID = u.preferenceID INNER JOIN UsernamePreference u2 ON d.secondPreferenceID = u2.preferenceID INNER JOIN CityOU c ON d.cityOUID = c.cityOUID INNER JOIN RegionOU r ON c.regionOUID = r.regionOUID INNER JOIN Domain o ON d.domainID = o.domainID WHERE d.dealershipOUDescription = '$Dealership'" #Fill In Query Details
    
    $Results = Run-SQLCommand $SqlQuery

    return $Results.Tables
}

function Get-NotificationSettings
{
    param($Title)

    $SqlQuery = "SELECT notifyCRM,notifyApps,notifyKornerstone,notifyEmployeeSite,notifyServiceDirector FROM JobProfile Where jobProfileDescription = '$Title'"

    $Results = Run-SQLCommand $SqlQuery

    return $results.Tables
}

function Get-NotificationEmails
{
    $SqlQuery = "SELECT description,emailAddress FROM NotificationEmails"

    $Results = Run-SQLCommand $SqlQuery

    return $results.Tables
}

#Exchange Functions

function Connect-Exchange
{
    $user = "autosupport@kengarff.com"
    $Password = ConvertTo-SecureString -String "^&AOo456nA`$ohsedgf^*U&456d" -AsPlainText -Force
    
    $UserCredential = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $User, $Password

    $Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $UserCredential -Authentication Basic -AllowRedirection

    Import-PSSession $Session

    Connect-MsolService -Credential $UserCredential

    Connect-AzureAD -Credential $UserCredential
}

function Set-ExchangeLicense
{
    param($user,$licenseType)

    $logon = "svcAzureAD@kengarffit.com"
    $password = ConvertTo-SecureString -String "%tNh3@ki5A" -AsPlainText -Force
    $credential = New-Object -TypeName "System.Management.Automation.PSCredential" -ArgumentList $logon, $password

    Connect-MsolService -Credential $credential

    Set-MsolUserLicense -UserPrincipalName $user -AddLicenses $licenseType
}

function Remove-ExchangeLicense
{
    param($ADUser)

    Connect-Exchange

    $ExchangeUser = Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName

    Foreach($license in $ExchangeUser.Licenses.AccountSkuID)
    {
        Set-MsolUserLicense -UserPrincipalName $ADUser.UserPrincipalName -RemoveLicenses $license
    }

    

}

function ConvertTo-SharedMailbox 
{
    param($User)

    Connect-Exchange

    $ADUser = Get-ADUserByEmployeeNumber -EmployeeNumber $User.EmployeeNumber
    
    #$MailboxName = $ADUser.UserPrincipalName 

    Set-Mailbox $ADUser.UserPrincipalName -Type Shared

    Add-MailboxPermission $ADUser.UserPrincipalName -User $User.ManagerEmail -AccessRights FullAccess

    $ExchangeUser = Get-MsolUser -UserPrincipalName $ADUser.UserPrincipalName

    Set-MsolUserLicense -UserPrincipalName $ADUser.UserPrincipalName -RemoveLicenses $ExchangeUser.Licenses.AccountSkuID
}
