#Created to find mal-emails and show option to delete them. 
#Creation date : 1-14-2016
#Creator: Alix N Hoover


#Add exchange snap-in to allow the right click run
add-pssnapin Microsoft.Exchange.Management.PowerShell.E2010

#Start of main menu
$title = "Find (and delete) mal-email(s)"
$message = "How to do you want to search?"
$Email = New-Object System.Management.Automation.Host.ChoiceDescription "&Email", `
    "Find Email(s) by the Senders Address"
$Subject = New-Object System.Management.Automation.Host.ChoiceDescription "&Subject", `
    "Find Email(s) By Subject"	

$options = [System.Management.Automation.Host.ChoiceDescription[]]($Email, $Subject)
$result = $host.ui.PromptForChoice($title, $message, $options, 0) 


#Start of Switch
switch ($result)
    {#open switch
	
		#Start of the search by Sender
        0 {#open case 0 
			Write-host "Sender of Email you wish to find:" -foregroundcolor magenta -nonewline
			$Semail = read-host
			Write-host "Your Email to send report to:" -foregroundcolor magenta -nonewline
			$Oemail = read-host			
			Get-mailbox -resultsize Unlimited | Search-Mailbox -SearchQuery ("From:" +$Semail) -TargetMailbox $Oemail -TargetFolder "Powershell Report"
			Write-host "Completed: look in your inbox for the folder -Powershell Report-" -foregroundcolor magenta 
			Write-host "Once you have checked, do you want to Delete from the users mailbox?" -foregroundcolor magenta 

			#Popup yes or no to delete
			#Button Types  
			# 
			#Value  Description   
			#0 Show OK button. 
			#1 Show OK and Cancel buttons. 
			#2 Show Abort, Retry, and Ignore buttons. 
			#3 Show Yes, No, and Cancel buttons. 
			#4 Show Yes and No buttons. 
			#5 Show Retry and Cancel buttons. 

			$a = new-object -comobject wscript.shell 
			$intAnswer = $a.popup("Do you want to delete these email(s)?",0,"Delete Email(s)",4) 
			If ($intAnswer -eq 6) 
			{ # open IF to yes or no
			$a.popup("Email(s) are being Delete.")
			Get-mailbox -resultsize Unlimited | Search-Mailbox -SearchQuery ("From:" +$Semail) -TargetMailbox $Oemail -TargetFolder "Powershell Delete Report" -DeleteContent -Force -LogLevel Full 
			} # Close IF to yes or no
				else { # open Else to yes or no
						$a.popup("Email(s) will be kept.") 
						write-host "Email(s) will be kept, closing script"
					} # Close Else to yes or no
  

		}#close case 0  
		
		
		
		#Start of the Search by Subject
		 1 {#open case 1
		 Write-host "Subject of Email(s) you wish to find:" -foregroundcolor magenta -nonewline
			$Semail = read-host
			Write-host "Your Email to send report to:" -foregroundcolor magenta -nonewline
			$Oemail = read-host			
			Get-mailbox -resultsize Unlimited | Search-Mailbox -SearchQuery ("Subject:" +$Semail) -TargetMailbox $Oemail -TargetFolder "Powershell Report"
			Write-host "Completed: look in your inbox for the folder -Powershell Report-" -foregroundcolor magenta
			Write-host "Once you have checked, do you want to Delete from the users mailbox?" -foregroundcolor magenta

			#Popup yes or no to delete
			#Button Types  
			# 
			#Value  Description   
			#0 Show OK button. 
			#1 Show OK and Cancel buttons. 
			#2 Show Abort, Retry, and Ignore buttons. 
			#3 Show Yes, No, and Cancel buttons. 
			#4 Show Yes and No buttons. 
			#5 Show Retry and Cancel buttons. 

			$a = new-object -comobject wscript.shell 
			$intAnswer = $a.popup("Do you want to delete these email(s)?",0,"Delete Email(s)",4) 
			If ($intAnswer -eq 6) 
			{ # open IF to yes or no
			$a.popup("Email(s) are being Delete.")
			Get-mailbox -resultsize Unlimited | Search-Mailbox -SearchQuery ("Subject:" +$Semail) -TargetMailbox $Oemail -TargetFolder "Powershell Delete Report" -DeleteContent -Force -LogLevel Full 
			} # Close IF to yes or no
				else { # open Else to yes or no
						$a.popup("Email(s) will be kept.") 
						write-host "Email(s) will be kept, closing script"
					} # Close Else to yes or no
  
		 
		 }#close case 1
		 


	} # Close Switch
