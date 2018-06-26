#WorkdayToADIntegration.ps1
#This File contains the Main Loop that takes information from Workday and Parses it to Active Directory and the Helpdesk
#Functions can be found in WorkdayFunctions.psm1
#Author: Curtis Blackley, Dominic Douglas

#Import Funcitons
Import-Module WorkdayFunctions

#Check to see if any Files exist on the FTP site

if(Check-WorkdayFTP)
{
    #Get files, Generate a list of users, Move Files to Processed Folder
    $files = Get-FileList

    $userList = Generate-UserList -FileList $files

    #For Each User: Determine if New Hire, Update, Or Termination and Perform correct Actions
    foreach($user in $userList)
    {
        #If Employee Number Does not exist, user is New Hire
        if(!(Check-EmployeeExists -EmployeeNumber $user.EmployeeNumber) -and ($user.TermDate -eq "" -or $user.TermDate -eq $null) )
        {
            #If Title, Dealership, and Department are empty, Skip and Continue loop
            if(($user.Title -eq "" -or $user.Title -eq $null) -and ($user.Dealership -eq "" -or $user.Dealership -eq $null) -and($user.Department -eq "" -or $user.Department -eq $null))
            {
                $Body = "Employee Number: " + $user.EmployeeNumber + "`n" + "Employee Name: " + $user.PreferredName + "`n" + "Start Date: " + $user.HireDate
                
                Send-Email -To "itnotifications@kengarffit.com" -From "Autosupport@kengarff.com" -Subject ("Empty User from Workday - " + $user.PreferredName) -Body $Body -ReplyTo "autosupport@kengarff.com"
                
                Continue
            }

            #Nullify $Username
            $username = $null

            $FirstName = $user.PreferredName.Split(" ")[0]

            #Get possible Username Combos
            $firstL = $FirstName + $user.LastName[0]
            $fLast = $FirstName[0] + $user.LastName
            $firstL = $firstL.ToLower()
            $fLast = $fLast.ToLower()

            $firstLAvailable = Check-ADUser -User $firstL
            $fLastAvailable = Check-ADUser -User $fLast

            $preferences = Get-EmailPreference -Dealership $user.Dealership

            #Set Username to first available Preference
            if(!$firstLAvailable -and $preferences.FirstPreference -eq "FirstL")
            {
                $username = $firstL
            }
            elseif(!$fLastAvailable -and $preferences.SecondPreference -eq "FLast")
            {
                $username = $fLast
            }
            elseif(!$fLastAvailable -and $preferences.FirstPreference -eq "FLast")
            {
                $username = $fLast
            }
            elseif(!$firstLAvailable -and $preferences.SecondPreference -eq "FirstL")
            {
                $username = $firstL
            }
            else
            {
                #No Username available, Post Ticket to Create User Manually
                $Subject = "New Hire - " + $user.PreferredName + " - " + $user.Dealership + " - Manual Creation Required"

                if(Check-HelpdeskTicketPosted -Subject $Subject -EmployeeNumber $user.EmployeeNumber)
                {
                }
                else
                {
                    Post-NewTicket -CC @($user.ManagerEmail,"allcrm@kengarff.com") -Subject $Subject -Response (Create-NewHireHelpdeskResponse -User $User -Username $Username -Password $Password) -HDDepartment 20
                }
            }

            #Automatically Create AD User and Email if $Username does not equal $null
            if($username -ne $null) 
            {
                $username = $username -replace '\s',''
                Create-ADUser -inputUser $user -inputUsername $username  
            }
        }
        elseif(($user.TermDate -ne "" -and $user.TermDate -ne $null)) #Else, If Termination Date is set, user is Termination
        {
            $ADUser = Get-ADUserByEmployeeNumber -EmployeeNumber $user.EmployeeNumber

            if($ADUser.enabled -or $ADUser.enabled -eq $null)
            {
                Disable-ADUser -AdUser $ADUser

                #Make Exchange Changes (Convert to Shared Mailbox, assign to Manager)
                #ConvertTo-SharedMailbox -User $user
                Remove-ExchangeLicense -ADUser $ADUser

                #Post Helpdesk Ticket
                $Subject = "Termination - " + $user.preferredName + " - " + $ADUser.Company
                $CC = $user.ManagerEmail
                $Response = Get-TerminationTicketResponse -User $user -ADUser $ADUser

                Send-TermEmail -User $User -ADUser $Aduser

                Post-NewTicket -Subject $Subject -CC $CC -Response $Response -HDDepartment 22                
            }
        }
        else #Otherwise, user is Update.
        {


            $ADUser = Get-ADUserByEmployeeNumber -EmployeeNumber $user.EmployeeNumber

            if(Compare-User -ADUser $ADUser -WorkdayUser $user)
            {
                if ((Check-HelpdeskTicketPosted -User $User -TicketType 'S') -or (Check-HelpdeskTicketPosted -User $User -TicketType 'P'))
                {
                }
                else
                {
                    Send-CRMEmail -User $User -ADUser $AdUser
                
                    #AD User hasn't changed. Post Questionaire Data to Helpdesk
                    Post-QuestionnaireTickets -User $User
                }
            }
            else
            {
                #AD User has changed, Update AD User
                Update-ADUser -ADUser (Get-AdUserByEmployeeNumber -EmployeeNumber $user.EmployeeNumber) -UpdateInfo $user

                #Check if Phone/Software Ticket has Posted
                #If Not, Run Questionnaire Stuff
                if ((Check-HelpdeskTicketPosted -User $User -TicketType 'S') -or (Check-HelpdeskTicketPosted -User $User -TicketType 'P'))
                {
                }
                else
                {
                    Post-QuestionnaireTickets -User $User
                }
            }    
        }
    }
}