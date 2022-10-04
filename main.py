import win32com.client
import os 
from datetime import datetime,timedelta
import dns.resolver
import re 
mapi = win32com.client.Dispatch('outlook.application').GetNamespace('MAPI')
DEFAULT_ACCOUNT = mapi.Accounts[0]
INBOX_FOLDER_ID = 6
 

inbox = mapi.GetDefaultFolder(INBOX_FOLDER_ID)

DOMAIN_SPF_MAP = {}

def filter_messages_past_n_days(messages, n): 
    received_dt = datetime.now() - timedelta(days=n)
    received_dt = received_dt.strftime('%m/%d/%Y %H:%M %p')
    return messages.Restrict("[ReceivedTime] >= '" + received_dt + "'")

def filter_messages_by_sender_email_address(messages, sender_email_address):    
    return messages.Restrict(f"[SenderEmailAddress] = '{sender_email_address}'")

def filter_messages_by_subject(messages, subject): 
    return messages.Restrict(f"[Subject] = '{subject}'") 

def save_attachments_from_messages(messages, save_to_folder_path='./attachments'):  
    try:
        for message in list(messages):
            print(message)
            try:
                s = message.sender
                for attachment in message.Attachments:
                    print(attachment)
                    attachment.SaveASFile(os.path.join(save_to_folder_path, attachment.FileName))
                    print(f"attachment {attachment.FileName} from {s} saved")
            except Exception as e:
                print("error when saving the attachment:" + str(e))
    except Exception as e:
            print("error when processing emails messages:" + str(e))

def get_message_headers(message):
    PR_TRANSPORT_MESSAGE_HEADERS = 'http://schemas.microsoft.com/mapi/proptag/0x007D001F'
    return message.PropertyAccessor.GetProperty(PR_TRANSPORT_MESSAGE_HEADERS)

def get_txt_dns_records_for_domain(domain):  
    return [r for r in dns.resolver.resolve(domain, 'TXT').response.answer[0]]  
def get_return_path_from_headers(headers): 
    for h in headers.split('\n'):
        if 'Return-Path' in h: 
            rp = h.split('Return-Path:')[-1].strip()
            break 
    return rp 


def get_domain_from_email_address(email_address):
    return email_address.split('@')[-1].strip()

def get_return_path_domain_from_headers(headers):
    """ The return path domain is the domain whose SPF records need to be checked 
    to ensure that the sender address is included as an authorized sender. """
    return get_domain_from_email_address(get_return_path_from_headers(headers))

#### SPF - Sender Policy Framework parsing ####

def get_spf_record_from_txt_records(txt_records_list):  
    return [str(r) for r in txt_records_list if 'v=spf1' in str(r)][0]

def get_spf_dns_record_for_domain(domain):
    """ There should only be one SPF DNS TXT record for a domain. Obtain and return it. """
    print(f'Obtaining SPF DNS TXT record for domain {domain}')
    txt_records_list = get_txt_dns_records_for_domain(domain) 
    return get_spf_record_from_txt_records(txt_records_list)

def get_all_authorized_senders_from_spf_record(domain, spf_record):
    print(f'Getting a list of all authorized senders for domain {domain}')
    _split = spf_record.split() 
    authorized_senders = []
    for x in _split:
        if 'ip4:' in x or 'a:' in x:
            auth_sender = x.split(':')[-1].strip() 
            authorized_senders.append(auth_sender)
        elif 'include:' in x: 
            included_spf_policy = x.split(':')[-1].strip()
            print(f'SPF policy for {domain} includes SPF policy from {included_spf_policy}; recursing...')
            spf_record_from_included_policy = get_spf_dns_record_for_domain(included_spf_policy) 
            authorized_senders.extend(get_all_authorized_senders_from_spf_record(domain, spf_record_from_included_policy))
    return authorized_senders

def save_domain_authorized_senders(domain):
    """ Obtain list of authorized senders for a domain and save to a local map """
    spf_for_domain = get_spf_dns_record_for_domain(domain)
    DOMAIN_SPF_MAP[domain] = get_all_authorized_senders_from_spf_record(domain, spf_for_domain)

def display_domain_spf_map(): 
    for domain, authorized_senders in DOMAIN_SPF_MAP.items():
        print(f'The following addresses are authorized senders for {domain}:')
        for addr in authorized_senders:
            print(f'\t{addr}')

def get_mailserver_headers_chain(headers): 
    """ Each element in the returned list represents the contribution of one mail server; broken
    into segments using the "Received:" header as the separator """
    chain = []
    server_headers = None 
    for line in headers.split('\n'):  
        if line.strip().startswith('Received: '):
            if server_headers:
                chain.append(server_headers)
            server_headers = []
        if line.strip().startswith('Subject: '):
            break 
        server_headers.append(line)      
    chain.append(server_headers)
    return ['\n'.join(server_headers) for server_headers in chain]
        
def get_addresses_of_mail_servers_from_mailserver_headers_chain(chain):
    regex = r"(Received: from ([a-zA-Z0-9\-\.]*)[ \r\n]*(\([a-zA-Z0-9\-\.:\r\n \(\)\[\]]*\))[ \r\n]*by[ \r\n]*([a-zA-Z0-9\-\.]*)([ \r\n]*\([0-9\.\:a-zA-Z]*\))?)"
    addresses = []  
    for server_headers in chain:  
        match = re.search(regex, server_headers)
    
        if match:
            match = match.group()
            match = match.replace('\n',' ').replace('\r', ' ') 
            received_from = match.split(' ')[2].strip()
            received_by = match.split(' ')[7].strip()
            
            addresses.append({
                'received_from': received_from,
                'received_by':  received_by
            }) 
    return addresses


### SPF ###
messages = inbox.Items
messages = filter_messages_by_sender_email_address(messages, 'newsletter@smashingmagazine.com') 

def validate_message(m): 
    headers = get_message_headers(m)   
    mailserver_chain = get_mailserver_headers_chain(headers) 
    addresses = get_addresses_of_mail_servers_from_mailserver_headers_chain(mailserver_chain)
    print(addresses)
    print(f'Total mailservers involved: {len(mailserver_chain)}')
    return_path_domain = get_return_path_domain_from_headers(headers) 
    print(f'Obtaining Authorized Senders for Return-Path domain {return_path_domain}...')
    save_domain_authorized_senders(return_path_domain) 
    display_domain_spf_map()
    with open(f'email.txt', 'w') as f:
        f.write(headers) 


if __name__ == "__main__":
    validate_message(messages.GetFirst())
 