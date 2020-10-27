########################################################################## 
#                                                                        #
#          Menu Script for Shared Mailbox Automatic Replies              #
#                                                                        #
#   Created Date: 29/06/2020	               Created By: Lee Williams  #
#  Modified Date: 23/10/2020	              Modified By: Lee Williams  #
#                                                                        #
########################################################################## 

Clear-Host
# Create Connection to Office 365 Platform
 Import-module MSOnline
   $Creds = Get-Credential 
 Connect-MsolService -Credential $Creds
  $O365 = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://outlook.office365.com/powershell-liveid/ -Credential $Creds -Authentication Basic -AllowRedirection
 Import-PSSession $O365

$Mailboxes = Get-Content "OOOMailbox.txt" 
$ARMessage = Get-Content "OOOMessage.txt"

Function LoadMainMenu()
{
    [bool]$LoopMainMenu = $TRUE
    While ($LoopMainMenu)
    {
    Clear-Host
    Write-Host "`n▂▃▅▇█▓▒░ Shared Mailbox Automatic Replies ░▒▓█▇▅▃▂`n”
    Write-Host “`t`t [1] ★ Check Mailbox Status  ★”
    Write-Host “`t`t [2] ★ Get AutoReply Message ★”
    Write-Host “`t`t [3] ★ Set AutoReply Message ★”
    Write-Host “`t`t [4] ★ Auto Replies Enabled  ★”
    Write-Host “`t`t [5] ★ Auto Replies Disabled ★”
    Write-Host “`n`t`t [Q] ★ Quit And Exit         ★`n”
    $MainMenu = Read-Host “`t`t Enter Option”
    Switch ($MainMenu)
        {
        1{MainMenuOption1}
        2{Get-ARMessage}
        3{MainMenuOption3}
        4{Set-ARStatus -ARStatus "Enabled"}
        5{Set-ARStatus -ARStatus "Disabled"}
        "q" {   $LoopMainMenu = $FALSE
                Clear-Host
            }
        Default {
            Write-Host -BackgroundColor Red -ForegroundColor White "You did not enter a valid selection. Please try again."
            sleep -Seconds 2
                }
        }
    }
Return
}

Function MainMenuOption1()
{
Clear-Host
 Foreach ($Mailbox in $Mailboxes) {
    $Status = Get-MailboxAutoReplyConfiguration -Identity $Mailbox 
    Write-Host $Status.Identity -foregroundcolor Magenta -nonewline; Write-Host " Automatic replies are currently set to " -foregroundcolor White -nonewline; Write-Host $status.AutoReplyState -BackgroundColor Red -ForegroundColor White }
Pause
}

Function MainMenuOption3()
{
Clear-Host
 Foreach ($Mailbox in $Mailboxes) {
    Set-MailboxAutoReplyConfiguration -Identity $MailBox –ExternalMessage $ARMessage –InternalMessage $ARMessage
    $Status = Get-MailboxAutoReplyConfiguration -Identity $Mailbox 
    Write-Host "Setting message for mailbox " -nonewline; Write-Host $Status.Identity -foregroundcolor Magenta}
Get-ARMessage
}

Function Set-ARStatus()
{
Param ($ARStatus)
Clear-Host
 Foreach ($Mailbox in $Mailboxes) {
    Set-MailboxAutoReplyConfiguration -Identity $MailBox –AutoReplyState $ARStatus 
    $Status = Get-MailboxAutoReplyConfiguration -Identity $Mailbox 
    Write-Host $Status.Identity -foregroundcolor Magenta -nonewline; Write-Host " Automatic replies are now " -nonewline -foregroundcolor White; Write-Host $status.AutoReplyState -BackgroundColor Red -ForegroundColor White }
Pause
}

Function Get-ARMessage()
{
Clear-Host
 Foreach ($Mailbox in $Mailboxes) {
    $Status = Get-MailboxAutoReplyConfiguration -Identity $Mailbox 
    $IntMsg = $Status.InternalMessage
    $IntMsg = $IntMsg -replace ‘<[^>]+>’,""
    Write-Host $Status.Identity "Message:" -nonewline;
    Write-Host $IntMsg -foregroundcolor Magenta -nonewline;}
Pause
}
LoadMainMenu
IF ($clearHost) {Clear-Host}
Remove-PSSession $O365