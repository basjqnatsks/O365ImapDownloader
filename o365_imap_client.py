import imaplib
import msal
from itertools import chain
import email
import os
from email import utils
import sys
class O365:
    #Classes Procedure defined in Init returns nothing as client is independent
    def __init__(self,AUTH=None,CLIENT=None,SECRET=None,EMAIL=None,LOCATION=os.getcwd()) -> None:
        #Disables delete function from inbox if true
        #Will add file logging, print statements if on
        self.debug = False
        #static
        self.IP = 'outlook.office365.com'
        self.exclusions = []
        self.LOCATION = LOCATION
        self.__DoesLOCATIONExist__()
        self.EMAIL = EMAIL
        self.Authority= 'https://login.microsoftonline.com/'+str(AUTH)
        self.Client = CLIENT
        self.Scope = ['https://outlook.office365.com/.default']
        self.SECRET = SECRET
        print(self.LOCATION)
        if AUTH and CLIENT and SECRET and self.LOCATION and EMAIL:
            self.imap = self.CONNECT(self.IP,self.EMAIL)
            self.main()
    @staticmethod
    def generate_auth_string(user, token):
        return f"user={user}\x01auth=Bearer {token}\x01\x01"
    def __DoesLOCATIONExist__(self):
        if os.path.isdir(self.LOCATION):
            #LocationGood!
            pass
        else:
            self.LOCATION = None # <--- Will end Program
    def SelectFolder(self, date):
        TODAYFOLDER = self.LOCATION + '\\' + str(date.year)
        if os.path.isdir(TODAYFOLDER):
            #LocationGood!
            pass
        else:
            os.mkdir(TODAYFOLDER)
        TODAYFOLDER += '\\' + str(date.month)
        if os.path.isdir(TODAYFOLDER):
            #LocationGood!
            pass
        else:
            os.mkdir(TODAYFOLDER)
        TODAYFOLDER +=  '\\' + str(date.day)
        if os.path.isdir(TODAYFOLDER):
            #LocationGood!
            pass
        else:
            os.mkdir(TODAYFOLDER)
        return TODAYFOLDER
    def CONNECT(self, HOST,EMAIL):
        try:
            app = msal.ConfidentialClientApplication(self.Client, authority=self.Authority,client_credential=self.SECRET)
            result = app.acquire_token_silent(self.Scope, account=None)
            if not result:
                print("No suitable token in cache.  Getting a new one.")
                result = app.acquire_token_for_client(scopes=self.Scope)
            if "access_token" in result:
                print("Received Token")
                #print(result['token_type'])
                #pprint.pprint(result)
            else:
                print(result.get("error"))
                print(result.get("error_description"))
                print(result.get("correlation_id"))
            imap = imaplib.IMAP4(HOST)
            imap.starttls()
            imap.authenticate("XOAUTH2", lambda x: self.generate_auth_string(self.EMAIL, result['access_token']).encode("utf-8"))
            return imap
        except Exception as a :
            print(str(a))
    #saves all non multipart attachments
    def save_attachment(self,msg) -> tuple:
        rt = bytearray(b'')
        for part in msg.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            #print(type(rt))
            #print(part.get_payload(decode=True))
            rt += bytearray(part.get_payload(decode=True))
        return part.get_filename(),rt
    #search function to search for criteria in UID emulating outlooks search box
    #not in use but possibly used to sort data  
    @staticmethod
    def search_string(uid_max, criteria):
        c = list(map(lambda t: (t[0], '"'+str(t[1])+'"'), criteria.items())) + [('UID', '%d:*' % (uid_max+1))]
        return '(%s)' % ' '.join(chain(*c))
    def main(self):
        self.imap.select('inbox')
        criteria = {}
        uid_max = 0
        result, data = self.imap.uid('SEARCH', None, self.search_string(uid_max, criteria))
        uids = [int(s) for s in data[0].split()]
        if self.debug:
            print('--------UIDS------------')
            print(uids)
        for uid in uids:
            # Have to check again because sometimes does not obey UID criterion.
            if uid > uid_max:
                #print(self.imap.uid('fetch', '('+str(uid)+' BODY[TEXT])', '(RFC822)'))
                result, data = self.imap.uid('fetch', str(uid), '(RFC822)')
                for response_part in data:
                    if isinstance(response_part, tuple):
                        #message_from_string can also be use here
                        if self.debug:
                            print(email.message_from_bytes(response_part[1]))
                        MAIL = email.message_from_bytes(response_part[1])
                        Filename,SAVEDMAIl = self.save_attachment(MAIL)
                        
                        if self.debug:
                            print(MAIL['from'])
                        FROM = MAIL['from']
                        if self.debug:
                            print(MAIL.get_content_charset())
                        #if self.debug:
                            #print(MAIL['BODY'])
                       #Filename must exist
                       #filename must not contain any string containing filetype
                        if Filename is not None and all(x not in Filename for x in self.exclusions):
                            datetime_obj = utils.parsedate_to_datetime(str(MAIL['date']))
                            with open(self.SelectFolder(datetime_obj)+'\\'+ str(MAIL['date']).replace(' ', '').replace(':', '').replace('+','')+'-'+Filename,'wb') as f:
                                f.write(SAVEDMAIl)
                uid_max = uid
                #Delete message
                self.imap.uid('STORE',str(uid),'+FLAGS', '(\\Deleted)')
                self.imap.expunge()
if __name__ == '__main__':
    try:
        argAuth = sys.argv[1]
        argClient_Id = sys.argv[2]
        argSecret = sys.argv[3]
        argEMAIl = sys.argv[4] 
    except:
        pass
    else:
        try:
            argLOCATION = sys.argv[5] 
        except:
            O365(argAuth, argClient_Id, argSecret,argEMAIl )
        else:
            O365(argAuth, argClient_Id, argSecret,argEMAIl,argLOCATION)