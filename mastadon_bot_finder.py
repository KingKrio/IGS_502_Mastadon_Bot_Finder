import openpyxl
from spellchecker import SpellChecker
from names_dataset import NameDataset
from random_string_detector import RandomStringDetector

nd = NameDataset()
checker = SpellChecker(distance=2)
valid_name = RandomStringDetector(allow_numbers=True)

def read_xlsx(path):
    accounts = {}
    workbook = openpyxl.load_workbook(path)

    for sheet in workbook.worksheets:#sheet = workbook.active
        row_len = sheet.max_row
        col_len = sheet.max_column

        for row_i in range(2, row_len + 1):
            post = {}
            user = sheet.cell(row = row_i, column = 4).value

            if user not in accounts:
                accounts[user] = []

            for col_j in range(1, col_len + 1):
                key = sheet.cell(row = 1, column = col_j).value
                val = sheet.cell(row = row_i, column = col_j).value
                post[key] = val
            accounts[user].append(post)
    return accounts

def find_bots(accounts):
    scrambled_names, freq_pstr, hstg_spmr = set(), set(), set()

    for user in accounts:
        author = (accounts[user])[0]
        if valid_name(author['author']):
            scrambled_names.add(author['author'])

        if spams_hashtags(accounts[user]):
            hstg_spmr.add(author['author'])

        if posts_frequently(accounts[user]):
           freq_pstr.add(author['author'])

    return {'scrambled names': scrambled_names, 'frequent posters': freq_pstr, 'hashtag spammers': hstg_spmr}

def posts_frequently(user_posts):
    dates = {}

    for post in user_posts:
        date_time = post['date'].split('T')

        date = date_time[0]
        time = (date_time[1]).split(':')[0]
        if date not in dates:
            dates[date] = []
        (dates[date]).append(time)

    for date in dates:
        if len(dates[date]) >= 4: # 4 or more posts is frequent
            for time in dates[date]:
                if dates[date].count(time) >= 2:
                    return True
    return False

def spams_hashtags(user):
    hashtags = []

    for post in user:
        hashtag_list = list(part[1:] for part in post['text'].split() if part.startswith('#'))
        [hashtags.append(hashtag) for hashtag in hashtag_list]

    for hashtag in hashtags:#
        if hashtags.count(hashtag) > 6:
            return True
    return False

def main():
    path = "/data/defundthepolice-full.xlsx"

    account_list = read_xlsx(path)
    bots = find_bots(account_list)

    print('----Potential bots based on looking for scrambled names: ')
    for scrambled_name in bots['scrambled names']:
        print(scrambled_name)

    print('\n----Potential bots based on searching for frequent posters: ')
    for frequent_poster in bots['frequent posters']:
        print(frequent_poster)

    print('\n----Potential bots based on searching for hashtag spammers: ')
    for frequent_poster in bots['hashtag spammers']:
        print(frequent_poster)

    total_bots = len(bots['scrambled names'] | bots['frequent posters'] | bots['hashtag spammers'])
    print("\n----Total number of potential bots:", total_bots)

main()
