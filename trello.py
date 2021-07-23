import json 
import pandas as pd 

def parse_cards_data(): 
    with open('trello.json', 'r') as json_file:
        data = json.load(json_file)
        cards = data['cards']
        
        lists_data = []
        lists = data['lists']
        for _list in lists:
            if(_list['closed'] == True):
                continue  

            element = {
                "name": _list['name'],
                "id"  : _list['id'],
                "created": False 
            }
            lists_data.append(element)

        cards_data = []
        for card in cards:
            bucket = None 
            priority = None
            new_bucket = 'False' 
            emails = [] 

            try:  
                if card['labels'][0]['color'] == 'green':
                    priority = 5 
                elif card['labels'][0]['color'] == 'orange': 
                    priority = 3
                elif card['labels'][0]['color'] == 'red': 
                    priority = 1 
                else: 
                    priority = 9 
            except IndexError:
                priority = 9 

            for _list in lists_data: 
                if _list['id'] == card['idList']:
                    bucket = _list['name']
                    if not _list['created']: 
                        _list['created'] = True 
                        new_bucket = 'True'
            if bucket == None:
                continue 
  
            for id in card['idMembers']: 
                for member in data['members']: 
                    if member['id'] == id: 
                        user = member['fullName']
                        user = user.split()
                        if len(user) == 2: 
                            email = user[0][0] + user[1] + '@xenon-pharma.com'
                        else:
                            email = user[0] + '@xenon-pharma.com'
                        email = email.lower()
                        emails.append(email)
            while len(emails) < 5: 
                emails.append(None)

            _card = { 
                "Title": card['name'],
                "Priority": priority, 
                "Due Date Time": card['due'], 
                "Bucket Name": bucket,
                "New Bucket": new_bucket, 
                "Email 1" : emails[0],
                "Email 2" : emails[1], 
                "Email 3" : emails[2], 
                "Email 4" : emails[3],
                "Email 5" : emails[4] 
            }

            cards_data.append(_card)
        
    return cards_data 

def write_excel(data): 

    Titles = [ i['Title'] for i in data]
    Priorities = [i['Priority'] for i in data]
    Due = [i['Due Date Time'] for i in data]
    Bucket = [i['Bucket Name'] for i in data]
    First = [i['New Bucket'] for i in data] 
    Email1 = [i['Email 1'] for i in data]
    Email2 = [i['Email 2'] for i in data]
    Email3 = [i['Email 3'] for i in data]
    Email4 = [i['Email 4'] for i in data]
    Email5 = [i['Email 5'] for i in data]

    df = pd.DataFrame({
        'Title': Titles,
        'Priority': Priorities,
        'Due Date Time': Due, 
        'Bucket Name': Bucket, 
        'Create': First, 
        'Email1' : Email1,
        'Email2' : Email2, 
        'Email3' : Email3,
        'Email4' : Email4, 
        'Email5' : Email5 
    })

    writer = pd.ExcelWriter('trello.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name='Informatics', startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets['Informatics']

    (max_row, max_col) = df.shape 

    column_settings = [{'header':column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    worksheet.set_column(0, max_col - 1, 12)

    writer.save()

if __name__ == '__main__':
    data = parse_cards_data()
    write_excel(data)
