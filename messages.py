from linkedin_api import Linkedin
from openpyxl import Workbook
from linkedin_api import client

def get_conversations(timestamp,linkedin):
    params = {"keyVersion": "LEGACY_INBOX",
                  "createdBefore":str(timestamp)}

    res = linkedin._fetch(f"/messaging/conversations", params=params)

    return res.json()

if __name__ == '__main__':
    login="YourUserName"
    password="YourPassword"
    linkedin = Linkedin(login, password,refresh_cookies=True)
    conversations = linkedin.get_conversations()
    last_id=conversations['elements'][0]['entityUrn']
    dialogs=[]
    dialogs_count=0
    LIMIT = 16 # will be scraped LIMIT*20 messages
    if LIMIT !=0:
        for i in range(LIMIT):
            conversations = get_conversations(conversations['elements'][-1]['events'][-1]['createdAt'], linkedin)
            if last_id == conversations['elements'][0]['entityUrn']:
                break
            else:
                last_id = conversations['elements'][0]['entityUrn']
            dialogs += conversations['elements']
            dialogs_count += 20
            print(str(dialogs_count) + " dialogs id scrapped")
    else:
        while 1:
            conversations = get_conversations(conversations['elements'][-1]['events'][-1]['createdAt'], linkedin)
            if last_id == conversations['elements'][0]['entityUrn']:
                break
            else:
                last_id = conversations['elements'][0]['entityUrn']
            dialogs += conversations['elements']
            dialogs_count += 20
            print(str(dialogs_count) + " dialogs id scrapped")


    wb = Workbook()
    ws = wb.active
    scrapped_dialogs=0
    try:
        for index, conversation in enumerate(dialogs):
            try:
                name = conversation['participants'][0]['com.linkedin.voyager.messaging.MessagingMember']['miniProfile'][
                           'firstName'] + \
                       " " + \
                       conversation['participants'][0]['com.linkedin.voyager.messaging.MessagingMember']['miniProfile'][
                           'lastName']
                url = "https://www.linkedin.com/in/" + \
                      conversation['participants'][0]['com.linkedin.voyager.messaging.MessagingMember']['miniProfile'][
                           'publicIdentifier']
                messages = linkedin.get_conversation(conversation['entityUrn'].split(':')[-1])
                messages_and_names = []
                for message in messages['elements']:
                    messages_and_names.append(
                        message['from']['com.linkedin.voyager.messaging.MessagingMember']['miniProfile'][
                            'firstName'] + ":" +
                        message['eventContent']['com.linkedin.voyager.messaging.event.MessageEvent']['attributedBody'][
                            'text'])
                ws.append(
                    [f"https://www.linkedin.com/messaging/thread/{conversation['entityUrn'].split(':')[-1]}/"] + [
                        name] + [url]+[index + 1] + messages_and_names)
                scrapped_dialogs += 1
                print(f"scrapped {scrapped_dialogs}/{dialogs_count} messages")
            except Exception:
                try:
                    linkedin = Linkedin(login, password,refresh_cookies=True)
                    print("Cookies refreshed")
                    continue
                except client.ChallengeException:
                    print("login to linkedin in browser, then press any key")
                    input()
                    continue


    finally:
        wb.save("messages.xlsx")
