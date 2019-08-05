## SYNOPSIS 
Use this script to provision a new shared mailbox in Exchange online.
 
## DESCRIPTION  
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

## OUTPUTS 
Results are printed to the console.
 
## NOTES 
Written by: Aws Ayad

THIS CODE IS MADE AVAILABLE AS IS, WITHOUT WARRANTY OF ANY KIND. THE ENTIRE RISK OF THE USE OR THE RESULTS FROM THE USE OF THIS
CODE REMAINS WITH THE USER.