import re
import openpyxl
from spellchecker import SpellChecker
from names_dataset import NameDataset

nd = NameDataset()
checker = SpellChecker(distance=2)

def read_xlsx(path):
    accounts = {}

    workbook = openpyxl.load_workbook(path)
    sheet = workbook.active

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
    scrambled_names = []
    for user in accounts:
        if detect_alnum_scrmbl(user["author"]):
            scrambled_names.append(user["author"])

    bots = []
    return bots

def detect_alnum_scrmbl(author=""):
    if author == "": return False


    char_parts = author.split("_") + re.findall('[A-Z][^A-Z]*', author) # char_part is short for character parts

    # check if there are discernible parts in the author name
    if len(char_parts) > 1:
        unknowns = checker.unknown(char_parts)
        for word in char_parts:
            size = len(word)
            is_word_name = [(nd.search(word))[key] for key in ["first_name", "last_name"]].count(None) != 2
            is_not_valid = not is_word_name and word in unknowns

            if not word.isdigit() and any(char.isdigit() for char in word): # if we aren't dealing with a num but nums are present\
                rm_fst = (word[1:] if word[0].isdigit() else '')
                rm_lst = (word[:size - 1] if word[size - 1].isdigit() else '')
                rm_bth = ''.join(sorted(set(rm_fst) & set(rm_lst), key=rm_fst.index))

                if rm_bth != '':
                    is_word_name = [(nd.search(rm_bth))[key] for key in ["first_name", "last_name"]].count(None) != 2
                    is_not_valid = not is_word_name and len(checker.unknown([rm_bth])) > 0

                    return is_non_alnum_replacable(rm_bth) if rm_bth.isalnum() else is_not_valid

                if rm_fst == '':
                    return is_non_alnum_replacable(rm_fst)

                for char in word: # first check if the numbers are only on the end of th
                    if char.isdigit():
                        spl_wrd += word.split(char)
            elif word.isdigit() and len(word) > 1 and len(unknowns) > 1:
                return True
            elif not word.isalnum():
                return is_non_alnum_replacable(word)
            elif is_not_valid:
                    return True
            elif not


    return False

def is_non_alnum_replacable(txt):
    is_word_name = [(nd.search(txt))[key] for key in ["first_name", "last_name"]].count(None) != 2
    is_not_valid = not is_word_name and txt in checker.unknown([txt])

    if txt.isalnum() and :
        return True
    return True

def frequent_hashtags(posts=None):
    if posts is None: return None

    return []

def main():
    path = "//data/defundthepolice-full.xlsx"
    account_list = read_xlsx(path)
    test = account_list['Natalia_Army_of_1@mstdn.party']

    print(len(test))

main()
