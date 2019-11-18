# ****************************************************************
# ***   Password Expiration Email - Daily v1.2                 ***
# ****************************************************************
#  
# v1.0 - 30/08/2019 - Jacob Curulli - First version
# v1.1 - 02/09/2019 - Jacob Curulli - Added holiday detection and special message for passwords expiring during the holiday break
# v1.2 - 12/11/2019 - Jacob Curulli - expiresOn data was in USA format, changed to dd/MM/yyyy. Updated holiday message tosay no access to SEQTA or email if user doesn't reset password.
#
# Based on the work of Robert Pearman and Fernando Pérez 
#  
# Script to Automated Email using Office 365 account to remind users Passwords Expiry. 
# Office 365 require SSL 
# Requires: Windows PowerShell Module for Active Directory 
# 
# 
######################################################################################################
$smtpServer="127.0.0.1"
$emailFrom = "helpdesk@contoso.com" 
$logging = "Enabled" # Set to Disabled to disable logging 
$logFile = "\\file\Pathhere\PasswordChangeNotification.csv" 
$testing = "Disabled" # Set to Enabled to email the test recipient instead of the actual users
$adminRecipient = "jacob@contoso.com"  # This is the user that will be CC'd on all emails and if testing is enabled they will receive the email instead of the user
$date = Get-Date -format ddMMyyyy
$logDate = Get-Date -UFormat "%d/%m/%Y"
$dateDayMonth = Get-Date -UFormat "%d/%m"
$maxPasswordAge = "180.00:00:00"
######################################################################################################
# 
# Get Users From AD who are Enabled, Passwords Expire and are Not Currently Expired 
Import-Module ActiveDirectory 
$users = get-aduser -SearchBase 'OU=Staff,DC=contoso,DC=local' -filter * -properties Name, PasswordNeverExpires, PasswordExpired, PasswordLastSet, EmailAddress |where {$_.Enabled -eq "True"} | where { $_.PasswordNeverExpires -eq $false } | where { $_.passwordexpired -eq $false } 
 
  Function LogUser {
 param($firstName, $lastName, $emailaddress, $daysToExpire, $expiresOn, $message)
 
 # Check Logging Settings 
	if (($logging) -eq "Enabled") 
	{ 
		# Test Log File Path 
		$logfilePath = (Test-Path $logFile) 
		if (($logFilePath) -ne "True") 
		{ 
			# Create CSV File and Headers 
			New-Item $logfile -ItemType File 
			Add-Content $logfile "Date,FirstName,LastName,EmailAddress,DaystoExpire,ExpiresOn,Message" 
		} 
	} # End Logging Check 
 
 # Write to the file
 Add-Content $logfile "$logDate,$firstName,$lastName,$emailaddress,$daysToExpire,$expiresOn,$message" 
 }
 
 Function EmailUser {
 param($firstName, $lastName, $emailaddress, $daysToExpire, $expiresOn, $message, $specialBody)
 
 # If Testing Is Enabled - Email Administrator
    if (($testing) -eq "Enabled")
    {
        $emailaddress = $adminRecipient
    } # End Testing

    # If a user has no email address listed
    if (($emailaddress) -eq $null)
    {
        $emailaddress = "$adminRecipient"    
    }# End No Valid Email

 # Set content of Email
 # Email Subject Set Here
    $subject="$firstName your password will expire $message"
      
  # Email Body Set Here, Note You can use HTML, including Images.
  # Check here if a special body was set in the expiration check
  if (($specialBody) -eq $null)
    {
	
	# No special body was sent, using the default below
	$body ="    
<p>Hi $firstName,<br>
</P>
<p>Your password will expire $message on $expiresOn.<br>
Please change your password <b>before</b> it expires to avoid problems accessing network services.</P>
<p> For instructions on how to change your password on a Mac OS X device <a href=`"https://www.google.com`" target=`"new`">please click here</a>. If you only use a Windows device 
  you can change your password from your desktop by pressing (control + alt + delete), then selecting Change a Password.<br>
  Please note you can only change your password whilst you are on campus at the College.<br>
  <br>
  <em>We will never ask for your password or account details in an email.</em></P>
<p>If you require assistance with resetting your password then you can reply to this email and it will create a Help Desk ticket.</P>
<p>Thanks,<br>
   IT Teams<br>
    "    
    }
	else{
	# Special body was sent so we'll use that
	$body = $specialBody}
     		
    # Send Email Message
	Send-MailMessage -From $emailFrom -To $emailaddress -Cc $adminRecipient -Subject $subject -body $body -BodyAsHtml -SmtpServer $smtpServer

	# Write-Host "First Name:		$firstName"
	# Write-Host "Last Name: 		$lastName"
	# Write-Host "Days to Expire: 	$daysToExpire"
	# Write-Host "Expiry: 		$expiresOn"
	# Write-Host "Message: 		$message"
	# Write-Host "********"
	# Write-Host ""
	
    } # End Send Message
  
# Process Each User for Password Expiry 
foreach ($user in $users) 
{ 
    $fullName = $user.Name
	$firstName = $user.givenName
	$lastName = $user.Surname
    $emailaddress = $user.emailaddress 
    $passwordSetDate = $user.PasswordLastSet
	
    $expiresOnRaw = $passwordsetdate + $maxPasswordAge
	$expiresOn = $expiresOnRaw.ToString("dd/MM/yyyy")
    $today = (get-date) 
    $daysToExpire = (New-TimeSpan -Start $today -End $expiresOn).Days 
         
		
    # Set message based on number of days till password expires 
 
	if (($daysToExpire -eq "21")) # 3 weeks till expiry
	{
	$message = "in 3 weeks"
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	}

	elseif (($daysToExpire -eq "14")) # 2 weeks till expiry
	{
	$message = "in 2 weeks"
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	}

	elseif (($daysToExpire -eq "7")) # 1 week till expiry
	{
	$message = "in 1 week"
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	}
	
	elseif (($daysToExpire -eq "3")) # 3 days till expiry
	{
	$message = "in 3 days"
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	}
	
	elseif (($daysToExpire -eq "0")) # Expires today
	{
	$message = "TODAY"
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	}

	if (($dateDayMonth -eq "12/11" -And $daysToExpire -lt "50")) # Expires 50 days after the 11th of December - so during the Summer Break
	{
	$message = "during the summer holidays"
	$specialBody = "    
<p>Hi $firstName,<br>
</P>
<p>Your password will expire during the summer school holidays on $expiresOn.<br>
<b>Please change your password before you leave for summer holidays. If you don't you won't be able to access network Services during the summer break.<br></b></P>
<p> For instructions on how to change your password on a Mac OS X device <a href=`"https://www.google.com`" target=`"new`">please click here</a>. If you only use a Windows device 
  you can change your password from your desktop by pressing (control + alt + delete), then selecting Change a Password.<br>
  Please note you can only change your password whilst you are on campus at the College.<br>
  <br>
  <em>We will never ask for your password or account details in an email.</em></P>
<p>If you require assistance with resetting your password then you can reply to this email and it will create a Help Desk ticket.</P>
<p>Thanks,<br>
  IT Team<br>
    "
	LogUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message
	EmailUser $firstName $lastName $emailaddress $daysToExpire $expiresOn $message $specialBody
	}
		 
} # End User Processing 
# End
