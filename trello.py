import json 
import pandas as pd 
import io 
import sys

def parse_cards_data(file_name): 
    with io.open(file_name, 'r', encoding='utf-8') as json_file:
        data = json.load(json_file)
        cards = data['cards']        
        checks = data['checklists']

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
            
            if card['closed'] == True: 
                continue 

            bucket = None 
            priority = None
            new_bucket = 'False' 
            emails = []
            colors = ['No'] * 10 

            #calculate priority of the task
            try:  
                if '3' in card['labels'][0]['name']:
                    priority = 5 
                elif 'Important' in card['labels'][0]['name'] or '2' in card['labels'][0]['name']: 
                    priority = 3
                elif 'Urgent' in card['labels'][0]['name'] or '1' in card['labels'][0]['name']: 
                    priority = 1 
                else: 
                    priority = 9 
            except IndexError:
                priority = 9 
            #check if new bucket needs to be created
            for _list in lists_data: 
                if _list['id'] == card['idList']:
                    bucket = _list['name']
                    if not _list['created']: 
                        _list['created'] = True 
                        new_bucket = 'True'
            if bucket == None:
                continue     
            #add users to task
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

            #if card has cover, use color of cover; if has both cover & label color, use cover color; if only label, use label color. 
            try: 
                card_color = None 
                if card['cover']['color'] == None: 
                    card_color = card['labels'][0]['color']
                else: 
                    card_color = card['cover']['color'] 

                if card_color == 'green': 
                    colors[0] = 'Yes' 
                elif card_color == 'yellow': 
                    colors[1] = 'Yes'
                elif card_color == 'orange': 
                    colors[2] = 'Yes'
                elif card_color == 'red': 
                    colors[3] = 'Yes'
                elif card_color == 'purple': 
                    colors[4] = 'Yes' 
                elif card_color == 'blue': 
                    colors[5] = 'Yes' 
                elif card_color == 'sky': 
                    colors[6] = 'Yes'
                elif card_color == 'lime': 
                    colors[7] = 'Yes'
                elif card_color == 'pink': 
                    colors[8] = 'Yes'
                elif card_color == 'black': 
                    colors[9] = 'Yes'
            except IndexError: 
                pass 

            _card = { 
                "Title": card['name'],
                "Priority": priority, 
                "Due Date Time": card['due'], 
                "Bucket Name": bucket,
                "New Bucket": new_bucket,
                "Description": card['desc'].rstrip(),  
                "Email 1" : emails[0],
                "Email 2" : emails[1], 
                "Email 3" : emails[2], 
                "Green": colors[0],
                "Yellow": colors[1], 
                "Orange": colors[2],
                "Red": colors[3],
                "Purple": colors[4], 
                "Dark Blue": colors[5], 
                "Light Blue": colors[6], 
                "Turqoise": colors[7],
                "Pink": colors[8],
                "Black": colors[9]
            }

            #get up to 15 checklist items
            names = [None] * 15
            states = [None] * 15
            i = 0 
            for check in checks: 
                for check_id in card['idChecklists']: 
                    if check['id'] == check_id: 
                        for j in check['checkItems']:
                            names[i] = j['name'] 
                            if j['state'] == 'complete': 
                                states[i] = 'Yes'
                            else:
                                states[i] = 'No'
                            i += 1
 
            for i in range(15): 
                _card[f'checkItem{i}Name'] = names[i]
                _card[f'checkItem{i}State'] = states[i]

            cards_data.append(_card)
        
    return cards_data 

def write_excel(data, sheet_name): 

    data_frame = { 
        "Title": [i['Title'] for i in data],
        "Priority": [i['Priority'] for i in data], 
        "Due Date Time": [i['Due Date Time'] for i in data], 
        "Bucket Name": [i['Bucket Name'] for i in data],
        "New Bucket": [i['New Bucket'] for i in data],
        "Description": [i['Description'] for i in data],  
        "Email 1" : [i['Email 1'] for i in data],
        "Email 2" : [i['Email 2'] for i in data], 
        "Email 3" : [i['Email 3'] for i in data], 
        "Green": [i['Green'] for i in data],
        "Yellow": [i['Yellow'] for i in data], 
        "Orange": [i['Orange'] for i in data],
        "Red": [i['Red'] for i in data],
        "Purple": [i['Purple'] for i in data], 
        "Dark Blue": [i['Dark Blue'] for i in data], 
        "Light Blue": [i['Light Blue'] for i in data], 
        "Turqoise": [i['Turqoise'] for i in data],
        "Pink": [i['Pink'] for i in data],
        "Black": [i['Black'] for i in data]
    }

    for i in range(15): 
        data_frame[f'checkItem{i}Name'] = [j[f'checkItem{i}Name'] for j in data]
        data_frame[f'checkItem{i}State'] = [j[f'checkItem{i}State'] for j in data] 

    df = pd.DataFrame(data_frame)

    writer = pd.ExcelWriter('trello.xlsx', engine='xlsxwriter')
    df.to_excel(writer, sheet_name=sheet_name, startrow=1, header=False, index=False)

    workbook = writer.book
    worksheet = writer.sheets[sheet_name]

    (max_row, max_col) = df.shape 

    column_settings = [{'header':column} for column in df.columns]
    worksheet.add_table(0, 0, max_row, max_col - 1, {'columns': column_settings})

    worksheet.set_column(0, max_col - 1, 12)

    writer.save()

if __name__ == '__main__':
    file_name = sys.argv[1]
    sheet_name = sys.argv[2]
    data = parse_cards_data(file_name)
    write_excel(data, sheet_name)
