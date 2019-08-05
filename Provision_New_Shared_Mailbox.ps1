<#
.SYNOPSIS 
Use this script to provision a new shared mailbox in Exchange online.
 
.DESCRIPTION  
This script perform below steps:
1. create new shared mailbox
2. add routing address: UPN@o365.huntsman.com
3. set properties in the mailbox to control how messages sent as or on-half are handled
4. get mailbox Exchange GUID
5. create CSV file that contain new Groups details
6. open & create first CSV file and make sure the file created successfully
7. create the new mail-enabled security groups
8. define $CustomEDN & $CustomADN access permission
9. open & create second CSV file that contain group specific permission 
10. grant the mail-enabled security groups (created in step 7) access rights to all folders inside the shared mailbox
11. Set Owner of $GroupName_Group_Owner to $GroupName_Group_Owner
12. open & create third CSV file that control SendOnbehalf permission
13. grant the mail-enabled security groups (created in step 7) send on behalf permission
14. confirm the Send on Behalf permission is in place

.OUTPUTS 
Results are printed to the console.
 
.NOTES 
Written by: Aws Ayad

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS
CODE REMAINS WITH THE USER.
#>

write-host 
write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
write-host   'Provisioning a new shared mailbox in Exchange online'
write-host +++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
write-host

try 
    {
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"
    
    $MailboxName = Read-Host "Please enter the shared mailbox name"
    $MailboxDisplayName = Read-Host "Please enter the shared mailbox display name"
    $MailboxAlias = Read-Host "Please enter the shared mailbox alias"
    $MailboxPrimarySmtpAddress = Read-Host "Please enter the shared mailbox Primary Smtp Address"
    $MailboxRoutingAddress = Read-Host "Please enter the shared mailbox routing address"
    $MailboxArea = Read-Host "Please enter the mailbox's group area (AMER, EMEA or APAC)"

    Write-Host "Step1: Creating the New Shared Mailbox"
    New-Mailbox -Shared -Name $MailboxName -DisplayName $MailboxDisplayName -Alias $MailboxAlias -PrimarySmtpAddress $MailboxPrimarySmtpAddress

    Write-Host "Step2: adding routing address UPN@o365.huntsman.com"
    Set-Mailbox -Identity $MailboxPrimarySmtpAddress  -EmailAddresses @{Add="smtp:$MailboxRoutingAddress"}

    Write-Host "Step3: set properties in the mailbox to control how messages sent as or on-half are handled"
    Set-Mailbox -Identity $MailboxPrimarySmtpAddress -MessageCopyForSendOnBehalfEnabled:$True -MessageCopyForSentAsEnabled:$True

    Write-Host "Step4: get mailbox Exchange GUID & last five character of GUID"
    $ExchangeGUID = (Get-Mailbox -Identity $MailboxPrimarySmtpAddress).Exchangeguid
    $Last5GUID = ("$($ExchangeGUID)" -replace '.*?(?=.{1,5}$)').ToUpper()

    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    Write-Host "Step5: create CSV file that contain new Groups details"
    
    # Open & create first CSV file in PS
    $ExcelPath = 'C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Create_Groups.csv'
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # Array of Group Types
    $grouptype =@("GroupOwner", "R", "ADN", "ADY", "EDN", "EDY")
    
    # update Name, Alias & DisplayName columns
    $count = 2
    for ($j=0; $j -le 5; $j++){
        for ($i=1; $i -le 3;$i++){
            $workbook.ActiveSheet.Cells.Item($count,$i)  = '$'+ $MailboxArea + $Last5GUID+ '_' + $grouptype[$j]
        }
        $count+=1
    }
    
    # update PrimarySMTPAddress column
    $count2 = 0
    for ($i=2; $i -le 7;$i++){
        $workbook.ActiveSheet.Cells.Item($i,4)  = '$'+ $MailboxArea + $Last5GUID + '_' + $grouptype[$count2] +'@huntsman.com'
        $count2++
    }
    
    # update ManagedBy column
    $workbook.ActiveSheet.Cells.Item(2,6)  = 'aws_ayad@huntsman.com'
    for ($i=3; $i -le 7; $i++){
        $workbook.ActiveSheet.Cells.Item($i,6)  = '$'+ $MailboxArea + $Last5GUID + '_GroupOwner'
    }
    
    # update Notes column
    $AccessNote = @("","","","Read Only access", "Author access w/o Delete", "Author access w/Delete", "Editor access w/o Delete","Editor access w/Delete")
    $workbook.ActiveSheet.Cells.Item(2,8)  = 'This group is used to manage the membership of the other groups that control access to the shared mailbox named ' + $MailboxDisplayName + ' with an Exchange GUID of ' + $ExchangeGUID + ' .'
    for ($i=3; $i -le 7; $i++){
        $workbook.ActiveSheet.Cells.Item($i,8)  = 'Members of this group will have ' + $AccessNote[$i] + ' to the shared mailbox named ' + $MailboxDisplayName +' with an Exchange GUID of ' + $ExchangeGUID + ' .'
    }
    
    #Obtain Group_Owner cell value
    $workbook.sheets.item(1).activate()
    $WorkbookTotal=$workbook.Worksheets.item(1)
    $value = $WorkbookTotal.Cells.Item(2, 1)
    $GroupOwner = $value.Text

    # save & close CSV file
    $workbook.SaveAs($ExcelPath)
    $workbook.Close($false)
    $excel.Quit()
    
    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    Write-Host "Step6: import first CSV file and make sure the file created successfully"
    $CreatGroups = Import-Csv C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Create_Groups.csv
    $CreatGroups | Format-Table -AutoSize

    Write-Host "Step7: create the new mail-enabled security groups"
    foreach ($group in $CreatGroups) {New-DistributionGroup -Name $group.Name -Alias $group.Alias -DisplayName $group.DisplayName -PrimarySmtpAddress $group.PrimarySMTPAddress -Type $group.Type -ManagedBy $group.ManagedBy -MemberJoinRestriction $group.MJR -CopyOwnerToMember:$False -Notes $group.Notes}

    Write-Host "Step8: define CustomEDN & CustomADN access permission"
    $CustomEDN=@("CreateItems","EditAllItems","EditOwnedItems","FolderVisible","ReadItems")
    $CustomADN=@("CreateItems","EditOwnedItems","FolderVisible","ReadItems")

    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    Write-Host "Step9: create second CSV file that contain group name & associated specific permission"

    # Open & create second CSV file in PS
    $ExcelPath = 'C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Groups_Permission.csv'
    $excel = New-Object -ComObject Excel.Application
    $workbook = $excel.Workbooks.Open($ExcelPath)
    
    # Array of Group Types
    $grouptype =@("ADN", "ADY", "EDN", "EDY", "R")

    # update BOXEmail column
    for ($i=2; $i -le 6; $i++){
        $workbook.ActiveSheet.Cells.Item($i,1)  = $MailboxPrimarySmtpAddress
    }
    
    # update ACL columns
    $count = 0
    for ($i=2; $i -le 6;$i++){
        $workbook.ActiveSheet.Cells.Item($i,2)  = '$'+ $MailboxArea + $Last5GUID+ '_' + $grouptype[$count]
    $count+=1
    }
        
    # update Permisison column
    $Permission = @("","","CustomADN", "Author", "CustomEDN", "Owner" ,"Reviewer")
    for ($i=2; $i -le 6; $i++){
        $workbook.ActiveSheet.Cells.Item($i,3)  = $Permission[$i]
    }

    # save & close CSV file
    $workbook.SaveAs($ExcelPath)
    $workbook.Close($false)
    $excel.Quit()

    write-host "*****************************************************************************************************"
    write-host "*****************************************************************************************************"

    Write-Host "Step10: grant the mail-enabled security groups (created in step 7) access rights to all folders inside the shared mailbox"   
    import-csv C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Groups_Permission.csv | ForEach-Object { if ($_.permissions -eq "CustomEDN") {C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Set_Folder_Permissions_recursive_BULK.ps1 -Mailbox $_.boxemail -User $_.ACL -AccessRights $CustomEDN} elseif ($_.permissions -eq "CustomADN") {C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Set_Folder_Permissions_recursive_BULK.ps1 -Mailbox $_.boxemail -User $_.ACL -AccessRights $CustomADN} else {C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Set_Folder_Permissions_recursive_BULK.ps1 -Mailbox $_.boxemail -User $_.ACL -AccessRights $_.permissions} }

    Write-Host "Step11: set Ownership of Group_Owner group to itself"
    Set-DistributionGroup -Identity $GroupOwner -ManagedBy $GroupOwner -BypassSecurityGroupManagerCheck

    Write-Host "Step12: create third CSV file for SendOnBehalf permission"
    Import-Csv C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Groups_Permission.csv | Where-Object {$_.Permissions -notmatch 'Reviewer'} | Export-Csv C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Groups_Permission_SendOnBehalf.csv -NoTypeInformation
    
    Write-Host "Step13: assigning the Send On-Behalf (default action)"
    $groupperms1b = Import-Csv C:\Scripts\ToWorkon\Provision_New_Shared_Mailbox\Groups_Permission_SendOnBehalf.csv
    foreach ($smb21 in $groupperms1b) {set-mailbox -Identity $smb21.BOXEmail -GrantSendOnBehalfTo @{Add=$smb21.ACL} -Confirm:$False}

    Write-Host "Step14: confirm the Send on Behalf permission is in place"
    Get-Mailbox $MailboxPrimarySmtpAddress | Where-Object {$_.GrantSendOnBehalfTo -ne $null} | Select-Object PrimarySmtpAddress,GrantSendOnBehalfTo

    }

finally
    {
    write-host "*****************************************************************************************************"
    write-host "Done, script completed successfully" -ForegroundColor white -BackgroundColor Red
    write-host "*****************************************************************************************************"
    write-host "`n"       
    }