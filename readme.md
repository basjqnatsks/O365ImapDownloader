o365_imap_client.py
    Commands:
            Layout:
                    python o365_imap_client.py AUTH(STR) CLIENT_ID(STR) SECRET(STR) EMAIL LOCATION(STR)

            Required:  
                AUTH(STR):
                CLIENT_ID(STR):
                SECRET(STR):
                EMAIL (STR):

            Optional:
                LOCATION(STR):
                    If empty program will default to local folder.

            example python o365_imap_client.py "AUTH" "client_id" "secret" "test@clouds.com" "C:\Users\User\Desktop\Files"
