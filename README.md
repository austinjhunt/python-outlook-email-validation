# Outlook Email Validation with Python
This repository was created to explore email validation with Python. It uses [pywin32](https://pypi.org/project/pywin32/) to read emails and their headers from the Outlook desktop client on Windows. It then analyzes the values of the headers to verify the emails against SPF and DKIM protocols. 

Still a work in progress. 

Sample output as of 10/4/2022 when run against an email I received from a [Smashing Magazine](https://www.smashingmagazine.com/) Newsletter subscription. The DKIM failure is actually what I would expect since it looks like a mailing list is used for this, and this problem is actually one of the driving reasons behind the creation of the [Authenticated Received Chain (ARC) protocol](https://en.wikipedia.org/wiki/Authenticated_Received_Chain) published in 2019. 
```
[{'received_from': 'BYAPR04MB6150.namprd04.prod.outlook.com', 'received_by': 'BYAPR04MB5750.namprd04.prod.outlook.com'}, {'received_from': 'DM6PR06CA0045.namprd06.prod.outlook.com', 'received_by': 'BYAPR04MB6150.namprd04.prod.outlook.com'}, {'received_from': 'DM6NAM11FT047.eop-nam11.prod.protection.outlook.com', 'received_by': 'DM6PR06CA0045.outlook.office365.com'}, {'received_from': 'mail212.sea81.mcsv.net', 'received_by': 'DM6NAM11FT047.mail.protection.outlook.com'}]
Total mailservers involved: 5
Obtaining Authorized Senders for Return-Path domain mail212.sea81.mcsv.net...
Obtaining SPF DNS TXT record for domain mail212.sea81.mcsv.net
Getting a list of all authorized senders for domain mail212.sea81.mcsv.net
SPF policy for mail212.sea81.mcsv.net includes SPF policy from spf.mandrillapp.com; recursing...
Obtaining SPF DNS TXT record for domain spf.mandrillapp.com
Getting a list of all authorized senders for domain mail212.sea81.mcsv.net
The following addresses are authorized senders for mail212.sea81.mcsv.net:
        148.105.14.212
        198.2.128.0/24
        198.2.132.0/22
        198.2.136.0/23
        198.2.145.0/24
        198.2.186.0/23
        205.201.131.128/25
        205.201.134.128/25
        205.201.136.0/23
        205.201.139.0/24
        198.2.177.0/24
        198.2.178.0/23
        198.2.180.0/24
Checking SPF...
SPF result recorded with Received-SPF header: Pass
SPF=Pass
Checking DKIM...
DKIM Signature: DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed;d=smashingmagazine.com; s=k1; t=1646838291;i=newsletter@smashingmagazine.com;bh=t3u1lX26nPRNYACI6VjIsaUTMgcHPedvNnTZIYZ7Nug=;h=Subject:From:Reply-To:To:Date:Message-ID:List-ID:List-Unsubscribe: List-Unsubscribe-Post:Content-Type:MIME-Version:CC:Date:Subject;b=eip3LtZ8NYi/F1SX2T5OkOTVIANusWI7jyUo/eH/Aq0XudKaQdmyuBl03IrYPvpVj rlftQxfgOAWN62qYIoGUXLoPwKgeYvK5ptS9SyslyeQeG/i10QomMmoC09zYK7WZcJ U5M9x0zqjGl8S0nKUo7OUeryQD99ON9pyA22PkUk=
Selector = k1
Domain = smashingmagazine.com
Verification Body Hash VW5Zr1z9DOxrUiC8NTwOQ0MbdnETkMQUAAuZvGUREHM= NOT EQUAL to signed body hash t3u1lX26nPRNYACI6VjIsaUTMgcHPedvNnTZIYZ7Nug=; possibly due to mailing list processing
DKIM verification failed
```