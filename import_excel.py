import configparser
import requests
from openpyxl import load_workbook

# load the configuration
config = configparser.ConfigParser()
config.read('config.cfg')

api_config = config['API']
board_config = config['Board']
card_config = config['Card']
list_config = config['List']
input_config = config['Input']

API_KEY = api_config.get('Key')
API_TOKEN = api_config.get('Token')

BOARD_ID = board_config.get('Id')
FILE_PATH = input_config.get('FilePath')

# the only fields that we will get from the cards api
CARD_FIELDS = card_config.get('Fields')

# build the cards url
CARD_LIST_URL = 'https://api.trello.com/1/boards/{}/cards?' \
    'fields={}&key={}&token={}'.format(BOARD_ID, CARD_FIELDS, API_KEY,
        API_TOKEN)

CARD_URL = 'https://api.trello.com/1/cards'
COMMENT_URL = 'https://api.trello.com/1/cards/{}/actions/comments'

# build the labels url
LABEL_LIST_URL = 'https://api.trello.com/1/boards/{}/labels?key={}&token={}' \
    .format(BOARD_ID, API_KEY, API_TOKEN)

# backlog ID
BACKLOG_LIST_ID = list_config.get('BacklogId')
# closed ID
CLOSED_LIST_ID = list_config.get('ClosedId')

# get the cards from the url
cards = requests.get(CARD_LIST_URL).json()

# get the labels from the url
labels = requests.get(LABEL_LIST_URL).json()

# read the file to be imported to trello
# load workbook
wb = load_workbook(filename=FILE_PATH, read_only=True)
# get the issue log details worksheet
ws = wb[input_config.get('WorkbookName')]

# get the maximum rows
MAX_ROWS = ws.max_row

# read each row on the file
def read_rows():
    # start at row 5.
    row_number = 4
    closed_count = 0
    reopen_count = 0
    existing_count = 0
    new_count = 0

    for row in ws.iter_rows(row_offset=1, max_row=MAX_ROWS):
        # increment row number for the next issue
        row_number += 1

        # get the IR number
        ir = row[9].value
        status = row[11].value

        # check if the status is not closed. if the status is closed, skip this
        if status.lower() == 'closed':
            print(ir + ' is closed')

            card = get_card_by_ir(ir);

            if card is not None:
                if card['idList'] != CLOSED_LIST_ID:
                    move_card_list(card['id'], CLOSED_LIST_ID)

                    print(ir + ' is moved to closed')
            else:
                create_card(row_number, ir, row, CLOSED_LIST_ID)

                print(ir + ' is created in the closed list')

            closed_count += 1

            continue
        elif status.lower() == 're-open':
            print(ir + ' is re-opened')

            card = get_card_by_ir(ir);

            if card is not None:
                if card['idList'] != BACKLOG_LIST_ID:
                    move_card_list(card['id'], BACKLOG_LIST_ID)

                    print(ir + ' is moved to backlog')
            else:
                create_card(row_number, ir, row, BACKLOG_LIST_ID)

                print(ir + ' is created in the backlog list')

            reopen_count += 1

        # check if the IR number is already existing
        if has_ir_already(ir):
            print(ir + ' already exists')

            existing_count += 1

            continue

        # create the card in the API
        create_card(row_number, ir, row, BACKLOG_LIST_ID)

        new_count += 1

    total = closed_count + reopen_count + new_count + existing_count

    print('Closed : {}'.format(closed_count))
    print('Re-Opened : {}'.format(reopen_count))
    print('New : {}'.format(new_count))
    print('Total : {}'.format(total))

# create a card based on the ir and row
def create_card(row_num, ir, row, list_id):
    # get the column values
    module = row[3].value
    problem_statement = row[4].value
    severity = row[6].value

    # build the title. the format is
    # "{ir number}: {first line of the problem statement}"
    title = '{}: {}'.format(ir, problem_statement.split('\n', 1)[0])

    # build the description of the card
    description = '{}\n\n' \
        '==========================================================\n\n' \
        '**Module:** {}\n\n' \
        '**Line:** #{}'.format(problem_statement, module, str(row_num))

    # get the label using the severity
    label_id = ''

    for label in labels:
        name = label['name']

        if name.lower() == severity.lower():
            label_id = label['id']

            break

    # create the card object
    param = {
        'name': title,
        'desc': description,
        'pos': 'bottom',
        'idList': list_id,
        'idLabels': label_id,
        'key': API_KEY,
        'token': API_TOKEN
    }

    # post the card to the API
    card = requests.post(CARD_URL, params=param).json()

    # add comments if there are any
    supp_docu = row[5].value
    clg_notes = row[20].value
    sp_notes = row[21].value

    if supp_docu:
        create_comment(card['id'],
            'Supporting Documents:\n\n' + str(supp_docu))

    if clg_notes:
        create_comment(card['id'],
            'Investigation Notes - CLG Systems:\n\n' + str(clg_notes))

    if sp_notes:
        create_comment(card['id'],
            'Investigation Notes - Service Provider:\n\n' + str(sp_notes))



def move_card_list(card_id, list_id):
    param = {
        'idList': list_id,
        'pos': 'bottom',
        'key': API_KEY,
        'token': API_TOKEN
    }

    requests.put(CARD_URL + '/{}'.format(card_id), params=param)

# create a comment for the card
def create_comment(card_id, text):
    comment = {
        'text': text,
        'key': API_KEY,
        'token': API_TOKEN
    }

    requests.post(COMMENT_URL.format(card_id), comment)


# get the card by it's IR number
def get_card_by_ir(ir):
    for card in cards:
        name = card['name']

        if name.lower().startswith(ir.lower()):
            return card

    return None


# check if the list of cards has the IR number already
def has_ir_already(ir):
    return get_card_by_ir(ir) is not None


if __name__ == '__main__':
    read_rows()