import re
import openpyxl
from spellchecker import SpellChecker
from names_dataset import NameDataset

nd = NameDataset()

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
    scrambled_names, freq_pstr, spams_hstgs = set(), set(), set()

    for user in accounts:
        author = (accounts[user])[0]

        if user not in scrambled_names and is_alphanum_scrmbl(author['author'], author['language']):
            scrambled_names.add(author['author'])

        if user not in freq_pstr and len(accounts[user]) > 3:
            dates = {}

            for post in accounts[user]:
                date_time = post['date'].split('T')

                date = date_time[0]
                time = (date_time[1]).split(':')[0]
                if date not in dates:
                    dates[date] = []
                (dates[date]).append(time)

            if posts_frequently(dates):
                freq_pstr.add(author['author'])

    #criteria =
    bots = scrambled_names.intersection(freq_pstr) #[user for user in accounts if ]
    return bots

def is_alphanum_scrmbl(author='', lang=''):
    if lang != 'en': return False

    author = author.replace('-', "_")
    name_parts = author.split("_") # char_part is short for character parts
    if len(name_parts) == 0:
        name_parts = re.findall('[a-zA-Z][^A-Z]*', author)
    else:
        for word in name_parts:
            multi_cased = re.findall('[a-zA-Z][^A-Z]*', word)
            if len(multi_cased) > 1:
                i = name_parts.index(word)
                name_parts = name_parts[:i] + name_parts[i + 1:] + multi_cased
    name_parts = list(filter(lambda wrd: len(wrd) > 1 or str(wrd).isdigit(), name_parts))

    digits = "0123456789"
    checker = SpellChecker(distance=2)
    if len(name_parts) > 1: # check if there are discernible parts in the author name
        unknowns = checker.unknown(name_parts)
        for word in name_parts:
            wrd_is_nme = [(nd.search(word))[key] for key in ["first_name", "last_name"]].count(None) != 2
            if any(char.isdigit() for char in word) and not word.isdigit(): # if we aren't dealing with a num but nums are present
                word = word.strip(digits)
                wrd_is_nme = [(nd.search(word))[key] for key in ["first_name", "last_name"]].count(None) != 2

                return (wrd_is_nme or word not in checker.unknown([word])) and not any(char.isdigit() for char in word)
            elif (word.isdigit() and len(word) > 1 and len(unknowns) > 1) or (not word.isalnum()) or (not wrd_is_nme and word in unknowns):
                return True
    elif len(name_parts) == 1:
        word = name_parts[0].strip(digits).lower()
        if any(char.isdigit() for char in word): return True

        wrd_is_nme = [(nd.search(word))[key] for key in ["first_name", "last_name"]].count(None) != 2
        if wrd_is_nme or word not in checker.unknown([word]):
            return False

        i = 2
        while i < len(word):
            str1, str2 = word[:i], word[i:]
            unknowns = checker.unknown([str1, str2])

            str1_candidates = [nd.search(str1)[key] for key in ["first_name", "last_name"]]
            str2_candidates = [nd.search(str2)[key] for key in ["first_name", "last_name"]]

            str1_is_name = str1_candidates.count(None) != 2
            str2_is_name = str2_candidates.count(None) != 2

            str1_rank = list(filter(None, [list(filter(None, dic['rank'].values())) for dic in str1_candidates if dic is not None and str1_is_name]))
            str1_rank = min([j for i in str1_rank for j in i]) if len(str1_rank) > 0 else 1001

            str2_rank = list(filter(None, [list(filter(None, dic['rank'].values())) for dic in str2_candidates if dic is not None and str2_is_name]))
            str2_rank = min([j for i in str2_rank for j in i]) if len(str2_rank) > 0 else 1001

            str1_valid = (str1_is_name and str1_rank < 1000 and len(str1) > 2) or (str1 not in unknowns and len(str1) > 1)
            str2_valid = (str2_is_name and str2_rank < 1000 and len(str2) > 2) or (str2 not in unknowns and len(str2) > 1)

            if str1_valid and str2_valid:
                return False
            elif str1_valid and not str2_valid:
                word = str2
                i = 2
            else:
                i += 1
        return True
    return False

def posts_frequently(dates):
    for date in dates:
        if len(dates[date]) >= 5: # 5 or more posts is frequent
            return True
        else: # check if the user has posted frequently
            for time in dates[date]:
                if dates[date].count(time) >= 2:
                    return True
    return False

def frequent_hashtags(posts=None):
    if posts is None: return None

    return []

def main():
    path = "/home/ybrenin/PycharmProjects/IGS_502_Mastadon_Bot_Finder/data/defundthepolice-full.xlsx"
    account_list = read_xlsx(path)
    print(len(account_list))
    test = account_list['Natalia_Army_of_1@mstdn.party']

    #print(test)

    print(find_bots(account_list))

main()
