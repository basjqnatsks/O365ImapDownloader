# O365ImapDownloader
Generic program to download all emails from any specified inbox. This is o365 compatible and works with enterprise outlook, currently downloads all attachments from inbox and moves the emails to the deleted inbox.



o365_imap_client.py Commands: Layout: python o365_imap_client.py AUTH(STR) CLIENT_ID(STR) SECRET(STR) EMAIL LOCATION(STR)

        Required:  
            AUTH(STR):
            CLIENT_ID(STR):
            SECRET(STR):
            EMAIL (STR):

        Optional:
            LOCATION(STR):
                If empty program will default to local folder.

        example python o365_imap_client.py "AUTH" "client_id" "secret" "test@clouds.com" "C:\Users\User\Desktop\Files"







SOP FOR AZURE TENANT PORTAL 
OPEN POWERSHELL
 
THIS CONNECT TO EXCHANGE SERVER
Install-Module -Name ExchangeOnlineManagement -allowprerelease
Import-module ExchangeOnlineManagement 
Connect-ExchangeOnline -Organization <TENANT_ID>
 
GET SERVICEPRICIPAL
Get-ServicePrincipal | fl
 
ADD SERVICE PRINCIPAL
New-ServicePrincipal -AppId client_id(IN PY APP) -ObjectId enterprise_object_id (GOTO AZURE FIND THE SPECIFIC PRINCIPAL AND THE OBJECT ID IN THE OVERVIEW PAGE OF THE SERVICE)
 
REMOVE SERVICEPRINCIPAL
Remove-ServicePrincipal  -identity IDENTITY_FOR_GET_SERVICE_PRICIPAL
 
SEE MAILBOX PERMS
Get-MailboxPermission -Identity EMAIL | Format-List
 
REMOVE MAILBOX PERMS
Remove-MailboxPermission -Identity IDENTITY_FROM_GET_MAILBOX -user USER_FROM_GET_MAILBOX_OR_SERVICE_PRICIPAL  -AccessRights FullAccess
Remove-MailboxPermission -Identity IDENTITY_FROM_GET_MAILBOX -user USER_FROM_GET_MAILBOX_OR_SERVICE_PRICIPAL  -AccessRights ReadPermission
 
ADD MAILBOX 
Add-MailboxPermission -Identity THE_EMAIL  -User USER_FROM_GET_MAILBOX_OR_IDENTITY_FROM_SERVICE_PRICIPAL  -AccessRights FullAccess
Add-MailboxPermission -Identity THE_EMAIL  -User USER_FROM_GET_MAILBOX_OR_IDENTITY_FROM_SERVICE_PRICIPAL  -AccessRights ReadPermission
