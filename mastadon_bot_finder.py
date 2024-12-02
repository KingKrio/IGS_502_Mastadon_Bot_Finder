import openpyxl
from spellchecker import SpellChecker
import language_tool_python
from random_string_detector import RandomStringDetector

grammar_checker = language_tool_python.LanguageTool('en-US')
spell_checker = SpellChecker(distance=2)
invalid_name = RandomStringDetector(allow_numbers=True)

def read_xlsx(path, sheet_num):
    accounts = {}
    workbook = openpyxl.load_workbook(path)

    # if the number for the sheet in the workbook is invalid change the number to the last sheet in the page
    sheet_num = sheet_num if sheet_num <= len(workbook.sheetnames) else len(workbook.sheetnames)
    sheet = workbook[workbook.sheetnames[sheet_num - 1]]

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
    scrambled_names, freq_pstr, mny_mstk = set(), set(), set()
    for acc in accounts:
        user = accounts[acc]
        username = (user[0])['author']

        if invalid_name(username):
            scrambled_names.add(username)

        if many_mistakes_in_text(user):
            mny_mstk.add(username)

        if posts_frequently(user):
           freq_pstr.add(username)
    return {'scrambled names': scrambled_names, 'frequent posters': freq_pstr, 'many mistakes': mny_mstk}

def posts_frequently(user_posts):
    hours, days = [], []
    for post in user_posts:
        days.append((post['date'].split('T'))[0]) # days
        hours.append((post['date'].split(':'))[0]) # hours of the day
    return (True if any(days.count(day) >= 12 for day in days) else False) or (True if any(hours.count(hour) >= 4 for hour in hours) else False)

def many_mistakes_in_text(posts):
    for post in posts:
        word_list = []
        text = post['text']
        lang = post['language']

        for word in text.split():
            word_list.append(word if not any(not char.isalnum() for char in word) else None)
        word_list = list(filter(None, word_list))

        if lang is not None and lang[:2] == 'en':
            if len(spell_checker.unknown(word_list)) >= 6 or len(grammar_checker.check(text)) >= 6:
                return True
    return False

def main():
    path = input("Please provide the path to the file: ")
    sheet = int(input("Please provide the number of the sheet: "))

    account_list = read_xlsx(path, sheet)
    bots = find_bots(account_list)
    grammar_checker.close()

    for category in bots:
        print('\n----Potential bots based on looking for:', category)
        for bot in bots[category]: print(bot)

    total_bots = len(bots['scrambled names'] | bots['frequent posters'] | bots['many mistakes'])
    print("\n----Total number of potential bots:", total_bots)

main()
